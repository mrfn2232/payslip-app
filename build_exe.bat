@echo off
REM Double-click on a Windows PC that has Python 3.10+ installed (python.org). Produces dist\PayslipMaker.exe
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl
python -m PyInstaller --onefile --windowed --name PayslipMaker app.py
echo.
echo Done: dist\PayslipMaker.exe
pause
