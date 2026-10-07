@echo off
chcp 65001 >nul
python -m pip install -q openpyxl
python "%~dp0maas_karsilastir_gui.py"
pause
