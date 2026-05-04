@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %ERRORLEVEL%==0 (set PYTHON_CMD=py -3) else (set PYTHON_CMD=python)
if not exist ".venv\Scripts\python.exe" %PYTHON_CMD% -m venv .venv
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
pause
