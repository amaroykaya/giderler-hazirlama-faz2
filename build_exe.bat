@echo off
echo === Gider Hazirlama Faz-2 EXE Build ===
if not exist .venv (
    python -m venv .venv
)
call .venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --onefile --noconsole --name "GiderHazirlamaFaz2" ^
 --collect-all numpy ^
 --collect-all pandas ^
 --collect-all lxml ^
 --collect-all openpyxl ^
 --collect-all pillow ^
 --collect-all pdfplumber ^
 app.py
echo.
echo EXE hazir: dist\GiderHazirlamaFaz2.exe
pause
