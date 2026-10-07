@echo off
chcp 65001 >nul
python -m pip install -q openpyxl pyinstaller
python -m PyInstaller --onefile --windowed --name MaasKarsilastir "%~dp0maas_karsilastir_gui.py"
echo.
echo Hazir: dist\MaasKarsilastir.exe
pause
