import os
import re
import sys
import shutil
import threading
import queue
import subprocess
import unicodedata
from pathlib import Path
from datetime import datetime, timedelta

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import pdfplumber
import requests
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


# =========================
# Excel kolonları
# =========================
COL_B = "B"  # prefix
COL_D = "D"  # fatura tarihi
COL_K = "K"  # fatura no (eşleşme)
COL_M = "M"  # KDV'siz TL tutar
COL_N = "N"  # etiket
COL_T = "T"  # USD tutarı (M/U)
COL_U = "U"  # USD kur (TCMB alış)
COL_V = "V"  # kur (kullanılmıyor, geriye dönük uyumluluk için)

VALID_TAGS = ["anten", "endüstriyel", "genel", "üretim", "savunma"]
FOLDER_UNMATCHED = "eşleşmedi"

# Etiket normalizasyon fonksiyonu
def normalize_tag(s: str) -> str:
    """
    Etiket normalizasyonu: Türkçe karakterleri ASCII'ye çevir, combining marks temizle.
    ÜRETİM → uretim, ENDÜSTİRİYEL → endustriyel
    Combining dot (İ/i̇) sorununu çözer.
    """
    if not s:
        return ""

    # 1) Unicode NFKD normalize (harf + combining işaretlere ayır)
    s = unicodedata.normalize("NFKD", str(s))

    # 2) Combining mark'ları (özellikle noktalı i'nin noktası) sil
    s = "".join(c for c in s if not unicodedata.combining(c))

    # 3) Türkçe özel harfleri ASCII'ye indir
    s = s.replace("ı", "i")
    s = s.replace("İ", "i")
    s = s.replace("ş", "s")
    s = s.replace("Ş", "s")
    s = s.replace("ğ", "g")
    s = s.replace("Ğ", "g")
    s = s.replace("ü", "u")
    s = s.replace("Ü", "u")
    s = s.replace("ö", "o")
    s = s.replace("Ö", "o")
    s = s.replace("ç", "c")
    s = s.replace("Ç", "c")

    # 4) lowercase
    s = s.lower()

    # 5) harf ve rakam dışındaki her şeyi sil
    s = re.sub(r"[^a-z0-9]+", "", s)

    return s


def parse_excel_date(date_value):
    """
    Excel'den gelen tarih değerini parse eder.
    Gün.ay.yıl formatını (örn: 01.11.2025) datetime'a çevirir.
    """
    if date_value is None:
        return None
    
    try:
        date_str = str(date_value).strip()
        if not date_str or date_str.lower() in ["", "nan", "none", "null"]:
            return None
        
        # gün.ay.yıl formatını parse et (örn: 01.11.2025)
        if "." in date_str and len(date_str.split(".")) == 3:
            parts = date_str.split(".")
            if len(parts) == 3:
                day, month, year = parts[0].strip(), parts[1].strip(), parts[2].strip()
                try:
                    return datetime(int(year), int(month), int(day))
                except ValueError:
                    pass
        
        # pandas dayfirst=True ile parse et
        d = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
        if pd.notna(d):
            return d.to_pydatetime()
    except Exception:
        pass
    
    return None


def read_tl_cell(cell):
    """
    Excel hücresinden gerçek TL tutarını güvenli okur.
    - Yüzde formatlı hücreleri TL'ye çevirmez (None döner)
    - String ise virgül/nokta temizler
    - Decimal / float ise aynen alır
    - Bilimsel notation (E+) güvenli parse eder
    """
    if cell.value is None:
        return None

    # Hücre yüzde formatlıysa -> gerçek TL olamaz, bu durumda değeri kullanma
    if cell.number_format and "%" in str(cell.number_format):
        return None

    val = cell.value

    if isinstance(val, str):
        v = val.strip().replace(" ", "")
        v = v.replace(".", "").replace(",", ".")
        try:
            return float(v)
        except:
            return None

    try:
        return float(val)
    except:
        return None

# Etiket mapping: normalize edilmiş key → klasör adı
TAG_MAP = {
    "anten": "anten",
    "endustriyel": "endüstriyel",
    "genel": "genel",
    "uretim": "üretim",
    "savunma": "savunma"
}

# Fatura numarası regex desenleri
INVOICE_PATTERNS = [
    r"FATURA\s*(NO|NUMARASI|NO\.|NUMARA)\s*[:\-]?\s*([A-Z0-9\-\/\. ]{5,})",
    r"FATURA\s*#\s*[:\-]?\s*([A-Z0-9\-\/\. ]{5,})",
    r"(INVOICE|INV)\s*(NO|NUMBER|#)\s*[:\-]?\s*([A-Z0-9\-\/\. ]{5,})",
]

# FX_LINE_PAT kaldırıldı - PDF'den döviz arama artık yok


# =========================
# Helpers
# =========================
def col_letter_to_idx(letter: str) -> int:
    letter = letter.upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def safe_str(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip()


def normalize_invoice(s):
    """
    Fatura numarasını normalize eder.
    Excel'den gelen sayısal değerleri (scientific notation, .0 eklemeleri) düzeltir.
    """
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u00A0", "").replace("\u200b", "")
    s = s.strip().upper()

    # Eğer Excel sayısal okuyup .0 eklediyse sil
    if s.endswith(".0"):
        s = s[:-2]

    # Scientific notation engelle
    if "E+" in s or "E-" in s:
        try:
            s = format(int(float(s)), "d")
        except:
            pass

    # Sadece harf ve rakam bırak
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def norm_invoice_no(s: str) -> str:
    """Deprecated: normalize_invoice kullanın."""
    return normalize_invoice(s)


def ensure_folders(out_dir: Path):
    out_dir.mkdir(parents=True, exist_ok=True)
    for t in VALID_TAGS + [FOLDER_UNMATCHED]:
        (out_dir / t).mkdir(parents=True, exist_ok=True)


def pdf_text_extract(pdf_path: Path, max_pages: int = 5) -> str:
    parts = []
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages[:max_pages]:
                txt = page.extract_text() or ""
                if txt:
                    parts.append(txt)
    except Exception:
        return ""
    return "\n".join(parts).strip()


def ocr_available() -> bool:
    try:
        import pytesseract  # noqa: F401
        from pdf2image import convert_from_path  # noqa: F401
        return True
    except Exception:
        return False


def pdf_ocr_extract(pdf_path: Path, max_pages: int = 3) -> str:
    """
    OCR fallback. Windows'ta Tesseract + Poppler yoksa boş döner.
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract

        images = convert_from_path(str(pdf_path), first_page=1, last_page=max_pages)
        out = []
        for img in images:
            out.append(pytesseract.image_to_string(img, lang="tur+eng"))
        return "\n".join(out).strip()
    except Exception:
        return ""


def find_invoice_no(text: str) -> str:
    """
    PDF'den fatura numarasını bulur (normalize edilmiş değil, ham değer).
    normalize_invoice ile normalize edilmeli.
    """
    t = (text or "").upper()

    for pat in INVOICE_PATTERNS:
        m = re.search(pat, t, flags=re.IGNORECASE)
        if m:
            cand = m.group(m.lastindex).strip()
            cand = cand.split("\n")[0].strip()
            cand = re.sub(r"[^A-Z0-9\-\/\. ]", "", cand).strip()
            return cand

    # Son çare: uzun alfanumerik aday (A612025000112251 gibi)
    candidates = re.findall(r"\b[A-Z]{0,3}\d{8,}\b", t)
    if candidates:
        candidates.sort(key=len, reverse=True)
        return candidates[0]

    return ""


# parse_fx_amount fonksiyonu kaldırıldı - PDF'den döviz arama artık yok


def tcmb_usd_rate_prev_business_day(invoice_date: datetime, log_callback=None):
    """
    Fatura tarihinden 1 gün önce başlar, yoksa geriye gider (max 10 iş günü).
    Hafta sonları (Cumartesi, Pazar) atlanır.
    TCMB XML ForexBuying (USD alış kuru).
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    for back in range(1, 11):
        d = invoice_date - timedelta(days=back)
        
        # Hafta sonu kontrolü: Cumartesi (5) veya Pazar (6) ise atla
        weekday = d.weekday()
        if weekday == 5 or weekday == 6:  # Cumartesi veya Pazar
            date_str_normalized = d.strftime("%Y-%m-%d")
            log(f"[TCMB] denendi: {date_str_normalized} → hafta sonu, atlandı")
            continue
        
        # Tarih formatı: YYYY-MM-DD (normalize, log için)
        date_str_normalized = d.strftime("%Y-%m-%d")
        
        # TCMB URL için format: YYYYMM/DDMMYYYY
        yyyymm = d.strftime("%Y%m")
        ddmmyyyy = d.strftime("%d%m%Y")
        url = f"https://www.tcmb.gov.tr/kurlar/{yyyymm}/{ddmmyyyy}.xml"

        try:
            r = requests.get(url, timeout=10)
            if r.status_code != 200:
                log(f"[TCMB] denendi: {date_str_normalized} → HTTP {r.status_code}, atlandı")
                continue
            
            root = etree.fromstring(r.content)
            nodes = root.xpath(".//Currency[@CurrencyCode='USD']/ForexBuying")
            
            if not nodes or not nodes[0].text:
                log(f"[TCMB] denendi: {date_str_normalized} → kur bulunamadı")
                continue
            
            val = nodes[0].text.strip()
            rate = float(val.replace(",", "."))
            log(f"[TCMB] denendi: {date_str_normalized} → bulundu: {rate:.4f}")
            return rate
            
        except Exception as e:
            log(f"[TCMB] denendi: {date_str_normalized} → hata: {e}")
            continue

    log(f"[TCMB] son 10 iş günü içinde kur bulunamadı (fatura tarihi: {invoice_date.strftime('%Y-%m-%d')})")
    return None


def open_folder(path: Path):
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


def calculate_rates_on_output_excel(excel_path: Path, log_q: queue.Queue) -> int:
    """
    Faz-2'nin oluşturduğu output Excel üzerinde kur ve USD tutarı hesaplar.
    
    Akış:
    1) new_excel aç
    2) wb.active (tek sheet) seç
    3) D sütunundan tarih oku (datetime olarak, parse yok)
    4) TCMB'den bir önceki iş günü USD alış kurunu bul
    5) U sütununa kur yaz
    6) T sütununa = M / U yaz
    7) kaydet
    """
    def log(msg: str):
        log_q.put(msg)
    
    log(f"[KUR-HESAP] Excel açılıyor: {excel_path}")
    
    # Faz-2'nin oluşturduğu Excel'i aç
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active  # Tek sheet (Sheet1)
    
    log(f"[KUR-HESAP] Sheet: {ws.title}")
    
    fx_rows = 0
    
    # Tüm satırları dolaş (header satır 1, data satırları 2'den başlar)
    for excel_row in range(2, ws.max_row + 1):
        # D sütunu: Fatura tarihi - string, datetime veya serial number olabilir
        date_cell = ws[f"{COL_D}{excel_row}"]
        raw = date_cell.value
        
        d = None
        
        if isinstance(raw, datetime):
            d = raw.date()
        
        elif isinstance(raw, (int, float)):
            try:
                d = from_excel(raw).date()
            except:
                d = None
        
        elif isinstance(raw, str):
            s = raw.strip().replace("\u00a0", "")
            if s and s != "-":
                try:
                    d = datetime.strptime(s, "%d.%m.%Y").date()
                except:
                    d = None
        
        if not d:
            log(f"[KUR-HESAP] Satır {excel_row}: Tarih okunamadı (D sütunu='{raw}')")
            continue
        
        # TCMB fonksiyonu datetime beklediği için date'i datetime'a çevir
        invoice_date = datetime.combine(d, datetime.min.time())
        
        log(f"[KUR-HESAP] Satır {excel_row}: Tarih={d.strftime('%d.%m.%Y')}")
        
        # M sütunu: KDV'siz TL tutar
        m_cell = ws[f"{COL_M}{excel_row}"]
        m_value = read_tl_cell(m_cell)
        
        if m_value is None or m_value <= 0:
            log(f"[KUR-HESAP] Satır {excel_row}: M tutarı geçersiz (M='{m_cell.value}')")
            continue
        
        log(f"[KUR-HESAP] Satır {excel_row}: M={m_value:.2f} TL")
        
        # TCMB'den bir önceki iş günü USD alış kurunu bul
        usd_rate = tcmb_usd_rate_prev_business_day(
            invoice_date, 
            log_callback=lambda msg: log(f"[KUR-HESAP] Satır {excel_row}: {msg}")
        )
        
        if usd_rate is None:
            log(f"[KUR-HESAP] Satır {excel_row}: TCMB'den kur bulunamadı (tarih: {invoice_date.strftime('%d.%m.%Y')})")
            continue
        
        log(f"[KUR-HESAP] Satır {excel_row}: Kur={usd_rate:.4f}")
        
        # U sütununa kur yaz
        ws[f"{COL_U}{excel_row}"].value = float(usd_rate)
        
        # T sütununa = M / U yaz (USD tutarı)
        usd_amount = m_value / usd_rate
        t_cell = ws[f"{COL_T}{excel_row}"]
        t_cell.value = float(usd_amount)
        # T sütununa 2 ondalık formatı ekle (0.00)
        t_cell.number_format = "0.00"
        
        log(f"[KUR-HESAP] Satır {excel_row}: ✅ U={usd_rate:.4f} kur yazıldı, T={usd_amount:.2f} USD yazıldı (M={m_value:.2f} TL / {usd_rate:.4f})")
        fx_rows += 1
    
    # En son veri satırını tespit et (T sütununda değer olan son satır)
    last_data_row = 1  # Header satır 1
    for row in range(2, ws.max_row + 1):
        t_cell = ws[f"{COL_T}{row}"]
        if t_cell.value is not None:
            last_data_row = row
    
    # Dip toplam ekle: Son veri satırından sonra 3 boş satır, sonra TOPLAM
    if fx_rows > 0 and last_data_row > 1:
        total_row = last_data_row + 4  # Son veri satırı + 3 boş + 1 (TOPLAM satırı)
        
        # Sol hücreye "TOPLAM" yaz (A sütununa veya ilk görünür sütuna)
        # Genelde T sütununun solunda bir hücre olur, burada B sütununu kullanabiliriz
        ws[f"B{total_row}"].value = "TOPLAM"
        
        # T sütununa toplam formülü ekle: =SUM(T2:T<last_data_row>)
        t_total_cell = ws[f"{COL_T}{total_row}"]
        t_total_cell.value = f"=SUM(T2:T{last_data_row})"
        # TOPLAM hücresine de 2 ondalık formatı ekle
        t_total_cell.number_format = "#,##0.00"
        
        log(f"[KUR-HESAP] Dip toplam eklendi: Satır {total_row}, T{total_row} =SUM(T2:T{last_data_row})")
    
    # Kaydet
    log(f"[KUR-HESAP] {fx_rows} satır için kur hesaplandı, Excel kaydediliyor...")
    wb.save(excel_path)
    wb.close()
    log(f"[KUR-HESAP] Excel kaydedildi: {excel_path}")
    
    return fx_rows


# =========================
# Core processing
# =========================
def process_all(excel_path: Path, pdf_paths: list[Path], out_dir: Path, log_q: queue.Queue):
    def log(msg: str):
        log_q.put(msg)

    ensure_folders(out_dir)

    log("[INIT] Excel okunuyor...")
    
    # Excel'i openpyxl ile açarak K sütununu string olarak oku
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    
    idx_B = col_letter_to_idx(COL_B)
    idx_D = col_letter_to_idx(COL_D)
    idx_K = col_letter_to_idx(COL_K)
    idx_M = col_letter_to_idx(COL_M)
    idx_N = col_letter_to_idx(COL_N)
    
    # Excel K => Excel satır numarası map (openpyxl ile oku, string olarak)
    invoice_to_excel_row: dict[str, int] = {}
    # Excel satır numarası => B sütunu değeri (openpyxl ile oku, string olarak)
    excel_row_to_b_value: dict[int, str] = {}
    
    for excel_row in range(2, ws.max_row + 1):  # 2'den başla (header satır 1)
        # K sütunu (fatura no) okuma
        k_cell = ws[f"{COL_K}{excel_row}"]
        if k_cell.value is not None:
            inv_raw = str(k_cell.value).strip() if k_cell.value else ""
            inv = normalize_invoice(inv_raw)
            if inv:
                invoice_to_excel_row[inv] = excel_row  # Excel satır numarası (1-based)
        
        # B sütunu (prefix) okuma - openpyxl ile string olarak
        b_cell = ws[f"{COL_B}{excel_row}"]
        b_raw = ""
        if b_cell.value is not None:
            b_raw = str(b_cell.value).strip()
        # Boşlukları temizle
        b_clean = " ".join(b_raw.split()) if b_raw else ""
        excel_row_to_b_value[excel_row] = b_clean
    
    # Diğer sütunlar için pandas kullan (dtype=str ile)
    df = pd.read_excel(excel_path, dtype=str)
    
    # Excel satır numarası => DataFrame index map oluştur
    excel_row_to_df_idx: dict[int, int] = {}
    # pandas header=0 varsayılan (ilk satır header), data satırları 1-based Excel'de 2'den başlar
    for df_idx in range(len(df)):
        excel_row = df_idx + 2  # DataFrame index 0 -> Excel satır 2
        excel_row_to_df_idx[excel_row] = df_idx

    can_ocr = ocr_available()
    log(f"[INIT] OCR {'AKTİF' if can_ocr else 'YOK (metinli PDF varsa çalışır)'}")

    total = len(pdf_paths)
    matched = 0

    for n, pdf_path in enumerate(pdf_paths, start=1):
        log(f"[PDF] ({n}/{total}) {pdf_path.name}")

        text = pdf_text_extract(pdf_path, max_pages=5)
        if len(text) < 30:
            if can_ocr:
                ocr_text = pdf_ocr_extract(pdf_path, max_pages=3)
                if len(ocr_text) > len(text):
                    text = ocr_text
            else:
                # OCR yoksa ve metin de yoksa eşleşmedi
                if len(text) < 10:
                    dest = out_dir / FOLDER_UNMATCHED / pdf_path.name
                    shutil.copy2(pdf_path, dest)
                    log("[MATCH] metin yok + OCR yok -> eşleşmedi")
                    continue

        inv_no_raw = find_invoice_no(text)
        inv_no = normalize_invoice(inv_no_raw) if inv_no_raw else ""
        
        # Eğer bulunan fatura no Excel'de yoksa, fallback regex ile tekrar dene
        if inv_no and inv_no not in invoice_to_excel_row:
            fallback = re.findall(r"[A-Z]{2,5}[0-9]{8,}", text.upper())
            if fallback:
                fallback.sort(key=len, reverse=True)
                inv_no_fallback = normalize_invoice(fallback[0])
                if inv_no_fallback in invoice_to_excel_row:
                    log(f"[FIX] Regex fallback kullanıldı: {inv_no} -> {inv_no_fallback}")
                    inv_no = inv_no_fallback
        
        if not inv_no:
            dest = out_dir / FOLDER_UNMATCHED / pdf_path.name
            shutil.copy2(pdf_path, dest)
            log("[MATCH] fatura no bulunamadı -> eşleşmedi")
            continue

        excel_row = invoice_to_excel_row.get(inv_no)
        
        # Son çare: PDF dosya adından fatura no ile dene
        if excel_row is None:
            filename = os.path.splitext(os.path.basename(str(pdf_path)))[0]  # AVA2025000441018
            filename_norm = normalize_invoice(filename)
            
            if filename_norm in invoice_to_excel_row:
                excel_row = invoice_to_excel_row[filename_norm]
                log(f"[FIX-FILENAME] PDF adı ile eşleşti: {inv_no} -> {filename_norm} (satır {excel_row})")
        
        if excel_row is None:
            dest = out_dir / FOLDER_UNMATCHED / pdf_path.name
            shutil.copy2(pdf_path, dest)
            log(f"[MATCH] Excel'de yok inv={inv_no} -> eşleşmedi")
            continue

        # DataFrame index'ini bul
        df_idx = excel_row_to_df_idx.get(excel_row)
        if df_idx is None:
            dest = out_dir / FOLDER_UNMATCHED / pdf_path.name
            shutil.copy2(pdf_path, dest)
            log(f"[MATCH] Excel satır {excel_row} DataFrame'de bulunamadı -> eşleşmedi")
            continue

        # N sütunu (etiket) - PDF kopyalanmadan hemen önce doğrudan oku ve klasör hesapla
        n_cell = ws[f"{COL_N}{excel_row}"]
        tag_raw = ""
        if n_cell.value is not None:
            tag_raw = str(n_cell.value)
        
        # Normalize etiket: Türkçe karakter sadeleştirme + boşluk temizleme
        tag_key = normalize_tag(tag_raw)
        
        # Mapping sözlüğü ile klasör adını bul
        folder = TAG_MAP.get(tag_key, FOLDER_UNMATCHED)
        
        log(f"[ETIKET] Excel satır={excel_row} raw='{tag_raw}' normalized='{tag_key}' → folder='{folder}'")

        # B sütunu (prefix) - openpyxl mapping'den (string, trim edilmiş)
        b_value = excel_row_to_b_value.get(excel_row, "")
        
        # PDF isimlendirme: B_sütunu + "-" + normalize_invoice(fatura_no) + ".pdf"
        # inv_no zaten normalize_invoice ile normalize edilmiş
        new_name = f"{b_value}-{inv_no}.pdf"
        
        dest = out_dir / folder / new_name
        shutil.copy2(pdf_path, dest)

        matched += 1
        log(f"[NAME] B={b_value} inv={inv_no} -> {new_name}")
        log(f"[MATCH] OK inv={inv_no} Excel satır={excel_row} tag={folder} -> {new_name}")

    # Faz-2: Yeni Excel dosyası oluştur (input Excel'in kopyası)
    log("[FAZ-2] Yeni Excel dosyası oluşturuluyor...")
    
    # Çıktı Excel'de sadece tek sekme olacak
    # Aktif sekmenin dışındaki tüm sekmeleri sil
    sheet_names = wb.sheetnames[:]
    active_sheet_name = ws.title
    
    for sheet_name in sheet_names:
        if sheet_name != active_sheet_name:
            wb.remove(wb[sheet_name])
    
    # Aktif sekmenin adını "Sheet1" yap (default tek sheet)
    ws.title = "Sheet1"
    
    # A sütunundan ay bilgisini al (A2 hücresi)
    month_cell = ws["A2"]
    month = ""
    if month_cell.value is not None:
        month = str(month_cell.value).strip().upper()
    else:
        log("[FAZ-2] UYARI: A2 hücresi boş, varsayılan isim kullanılıyor")
        month = "AY"
    
    # Output Excel dosya adı: gider_kalemleri_<AY>_2025_faz2.xlsx
    output_filename = f"gider_kalemleri_{month}_2025_faz2.xlsx"
    out_excel = out_dir / output_filename
    wb.save(out_excel)
    log(f"[FAZ-2] Yeni Excel kaydedildi: {out_excel}")

    # Input Excel'i kapat
    wb.close()

    log("--------------------------------------------------")
    log(f"[DONE] PDF toplam: {total}")
    log(f"[DONE] Eşleşen: {matched}")
    log(f"[DONE] Yeni Excel: {out_excel}")

    # Faz-3: Yeni Excel üzerinde kur ve USD hesaplama
    log("[FAZ-3] Kur ve USD hesaplama başlıyor...")
    fx_rows = calculate_rates_on_output_excel(out_excel, log_q)
    
    log("--------------------------------------------------")
    log(f"[DONE] Döviz yazılan satır: {fx_rows}")
    log("[DONE] İşlem bitti.")


# =========================
# Tkinter UI
# =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gider Hazırlama Faz-2")
        self.geometry("900x620")

        self.excel_path: Path | None = None
        self.pdf_paths: list[Path] = []
        self.out_dir: Path | None = None

        self.log_q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None

        self._build_ui()
        self._poll_log_queue()
        self._update_checks()

    def _build_ui(self):
        pad = 10
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=pad, pady=pad)

        # Excel
        f1 = ttk.LabelFrame(root, text="1) Excel Seç")
        f1.pack(fill="x", pady=(0, pad))
        self.excel_label = ttk.Label(f1, text="Seçilmedi")
        self.excel_label.pack(side="left", fill="x", expand=True, padx=pad, pady=pad)
        self.excel_ok = ttk.Label(f1, text="✗", width=3)
        self.excel_ok.pack(side="right", padx=(0, pad), pady=pad)
        ttk.Button(f1, text="Excel Seç", command=self.pick_excel).pack(side="right", padx=(0, pad), pady=pad)

        # PDFs
        f2 = ttk.LabelFrame(root, text="2) PDF'leri Seç (çoklu)")
        f2.pack(fill="x", pady=(0, pad))
        self.pdf_label = ttk.Label(f2, text="Seçilmedi")
        self.pdf_label.pack(side="left", fill="x", expand=True, padx=pad, pady=pad)
        self.pdf_ok = ttk.Label(f2, text="✗", width=3)
        self.pdf_ok.pack(side="right", padx=(0, pad), pady=pad)
        ttk.Button(f2, text="PDF'leri Seç", command=self.pick_pdfs).pack(side="right", padx=(0, pad), pady=pad)

        # Output
        f3 = ttk.LabelFrame(root, text="3) Çıktı Klasörü Seç")
        f3.pack(fill="x", pady=(0, pad))
        self.out_label = ttk.Label(f3, text="Seçilmedi")
        self.out_label.pack(side="left", fill="x", expand=True, padx=pad, pady=pad)
        self.out_ok = ttk.Label(f3, text="✗", width=3)
        self.out_ok.pack(side="right", padx=(0, pad), pady=pad)
        ttk.Button(f3, text="Çıktı Seç", command=self.pick_output).pack(side="right", padx=(0, pad), pady=pad)

        # Actions
        act = ttk.Frame(root)
        act.pack(fill="x", pady=(0, pad))
        self.convert_btn = ttk.Button(act, text="Dönüştür", command=self.start_convert, state="disabled")
        self.convert_btn.pack(side="left")
        self.progress = ttk.Progressbar(act, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=pad)

        # Log area
        lf = ttk.LabelFrame(root, text="Log")
        lf.pack(fill="both", expand=True)
        self.log_text = tk.Text(lf, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)
        scr = ttk.Scrollbar(lf, command=self.log_text.yview)
        scr.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scr.set)

    def _set_check(self, label_widget, ok_widget, is_ok: bool, text: str):
        label_widget.config(text=text)
        if is_ok:
            ok_widget.config(text="✓", foreground="green")
        else:
            ok_widget.config(text="✗", foreground="red")

    def _update_checks(self):
        self._set_check(
            self.excel_label,
            self.excel_ok,
            bool(self.excel_path),
            str(self.excel_path) if self.excel_path else "Seçilmedi"
        )
        self._set_check(
            self.pdf_label,
            self.pdf_ok,
            bool(self.pdf_paths),
            f"{len(self.pdf_paths)} PDF seçildi" if self.pdf_paths else "Seçilmedi"
        )
        self._set_check(
            self.out_label,
            self.out_ok,
            bool(self.out_dir),
            str(self.out_dir) if self.out_dir else "Seçilmedi"
        )

        ready = bool(self.excel_path) and bool(self.pdf_paths) and bool(self.out_dir)
        busy = self.worker is not None and self.worker.is_alive()
        self.convert_btn.config(state=("normal" if ready and not busy else "disabled"))

    def pick_excel(self):
        fp = filedialog.askopenfilename(
            title="Excel seç",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]
        )
        if fp:
            self.excel_path = Path(fp)
        self._update_checks()

    def pick_pdfs(self):
        fps = filedialog.askopenfilenames(
            title="PDF seç (çoklu)",
            filetypes=[("PDF", "*.pdf *.PDF")]
        )
        if fps:
            self.pdf_paths = [Path(p) for p in fps]
        self._update_checks()

    def pick_output(self):
        dp = filedialog.askdirectory(title="Çıktı klasörü seç")
        if dp:
            self.out_dir = Path(dp)
        self._update_checks()

    def log(self, msg: str):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def _poll_log_queue(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.log(msg)
        except queue.Empty:
            pass
        self.after(150, self._poll_log_queue)

    def start_convert(self):
        if not (self.excel_path and self.pdf_paths and self.out_dir):
            return
        if self.worker and self.worker.is_alive():
            return

        self.log("=== BAŞLADI ===")
        self.progress.start(10)
        self._update_checks()

        def run():
            try:
                process_all(self.excel_path, self.pdf_paths, self.out_dir, self.log_q)
            except Exception as e:
                self.log_q.put(f"[ERROR] {e}")
            finally:
                self.after(0, self._on_done)

        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()

    def _on_done(self):
        self.progress.stop()
        self._update_checks()
        if self.out_dir:
            if messagebox.askyesno("Bitti", "İşlem bitti. Çıktı klasörünü açayım mı?"):
                open_folder(self.out_dir)


if __name__ == "__main__":
    App().mainloop()
