@echo off
setlocal
pushd "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    py -3.12 -m venv .venv
    if errorlevel 1 goto :failed
)
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed
".venv\Scripts\python.exe" -m unittest discover -s tests -v
if errorlevel 1 goto :failed
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name ExcelProjectGrader --collect-data tkinterdnd2 --add-data "country_codes.csv:." desktop.py
if errorlevel 1 goto :failed
popd
exit /b 0
:failed
popd
exit /b 1
