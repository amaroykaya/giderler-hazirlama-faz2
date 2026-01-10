# Gider Hazırlama Faz-2 (EXE)

## Kurulum (Geliştirici için)
```powershell
cd gider-hazırlama-faz2
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Kullanım

1. Excel Seç
2. PDF'leri Seç (çoklu)
3. Çıktı Klasörü Seç
4. Dönüştür

## Çıktı klasörleri:
- `anten` / `endüstriyel` / `genel` / `üretim` / `savunma` / `eşleşmedi`

## Excel çıktısı:
- `data_guncel.xlsx` (T/U/V yazılır)

## OCR Notu (Opsiyonel)
Metinli olmayan PDF'ler için OCR gerekir.
- `pytesseract` çalışması için Windows'ta Tesseract kurulumu gerekir.
- `pdf2image` için Poppler gerekebilir.
- OCR yoksa: metinsiz PDF'ler "eşleşmedi"ye gider.

## EXE Build

PowerShell ile:
```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Çıktı:
```
dist\GiderHazirlamaFaz2.exe
```
