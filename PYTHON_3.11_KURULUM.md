# ⚠️ ÖNEMLİ: Python 3.11 Kurulumu Gerekiyor

## Sorun
Mevcut sistem Python sürümü: **Python 3.14.1**

Bu proje için Python 3.14 uygun değildir çünkü:
- Pandas, lxml, numpy, pillow gibi paketler Python 3.14'te düzgün derlenemiyor
- EXE build sırasında hatalar oluşuyor

## Çözüm: Python 3.11 Kurulumu

### 1. Python 3.11 İndir ve Kur
**İndirme Linki:** https://www.python.org/downloads/release/python-3118/

**Windows Installer'ı indir:**
- Windows 64-bit için: `python-3.11.8-amd64.exe`

### 2. Kurulum Sırasında ÖNEMLİ Ayarlar
⚠️ **Mutlaka işaretle:**
- ☑️ **"Add Python to PATH"** 
- ☑️ **"Install for all users"** (isteğe bağlı ama önerilir)

### 3. PATH Ayarları (Manuel)
Python 3.11 kurulduktan sonra:

1. **System Properties** aç (Win + Pause/Break)
2. **"Environment Variables"** tıkla
3. **System Variables** altında **"Path"** seç ve **Edit**
4. **Python 3.11 yolunu en üste taşı** (Python 3.14'ün üstüne)
   - Örnek: `C:\Python311\` ve `C:\Python311\Scripts\`
5. **Python 3.14 yollarını listeden kaldır veya en alta taşı**

### 4. Doğrulama
Yeni bir PowerShell/Terminal penceresi aç ve kontrol et:
```powershell
python --version
```
**Çıktı şu olmalı:** `Python 3.11.8`

### 5. Projeyi Hazırla
Python 3.11 kurulduktan ve PATH düzeltildikten sonra:

```powershell
cd "C:\Users\Asus.DESKTOP-9F6EQVL\Desktop\gider-hazırlama-faz2"
python -m venv .venv
.\.venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
```

### 6. EXE Build
```powershell
.\build_exe.bat
```

---

**NOT:** Python 3.14'ü sistemden kaldırmanıza gerek yok, sadece PATH'te Python 3.11'i öncelikli yapmanız yeterli.
