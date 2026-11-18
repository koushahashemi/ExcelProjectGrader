@echo off
setlocal
pushd "%~dp0"
if exist "ExcelProjectGrader.exe" (
    start "" "%CD%\ExcelProjectGrader.exe" %*
    popd
    exit /b 0
)
if exist "dist\ExcelProjectGrader.exe" (
    start "" "%CD%\dist\ExcelProjectGrader.exe" %*
    popd
    exit /b 0
)
if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv
    if errorlevel 1 goto :failed
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :failed
".venv\Scripts\pythonw.exe" desktop.py %*
if errorlevel 1 goto :failed
popd
exit /b 0
:failed
echo Setup failed. Install Python 3.12 with the Windows Python launcher and Tcl/Tk, or download the standalone release.
pause
popd
exit /b 1
