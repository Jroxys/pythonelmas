@echo off
chcp 65001 >nul
python -m pip install -q openpyxl pyinstaller pillow
cd /d "%~dp0"
if exist logo.png (
  python -m PyInstaller --onefile --windowed --name MaasKarsilastir --icon logo.png --add-data "logo.png;." maas_karsilastir_gui.py
) else (
  python -m PyInstaller --onefile --windowed --name MaasKarsilastir maas_karsilastir_gui.py
)
echo.
echo Hazir: dist\MaasKarsilastir.exe
pause
