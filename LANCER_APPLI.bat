@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo =====================================================
echo   Lancement de l'application de routage Excel - V3
echo =====================================================
echo.

REM On evite Python 3.14 si possible car certains modules ne sont pas encore stables dessus.
set PYTHON_CMD=
py -3.12 --version >nul 2>nul
if %ERRORLEVEL%==0 set PYTHON_CMD=py -3.12
if "%PYTHON_CMD%"=="" (
    py -3.13 --version >nul 2>nul
    if %ERRORLEVEL%==0 set PYTHON_CMD=py -3.13
)
if "%PYTHON_CMD%"=="" (
    py -3.11 --version >nul 2>nul
    if %ERRORLEVEL%==0 set PYTHON_CMD=py -3.11
)
if "%PYTHON_CMD%"=="" (
    where python >nul 2>nul
    if %ERRORLEVEL%==0 set PYTHON_CMD=python
)
if "%PYTHON_CMD%"=="" (
    echo ERREUR : Python n'est pas installe ou n'est pas dans le PATH.
    echo Installe Python 3.12 depuis python.org et coche "Add Python to PATH".
    pause
    exit /b 1
)

echo Python utilise : %PYTHON_CMD%
%PYTHON_CMD% --version

echo.
echo Creation / verification de l'environnement local...
if not exist ".venv\Scripts\python.exe" (
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo ERREUR : impossible de creer l'environnement Python.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"

echo.
echo Mise a jour de pip...
python -m pip install --upgrade pip

echo.
echo Installation / verification des modules necessaires...
echo Important : on force l'installation de versions Windows precompilees pour eviter Visual Studio.
python -m pip install --no-cache-dir --only-binary=:all: -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERREUR : installation des modules impossible.
    echo Cause probable : version de Python trop recente, souvent Python 3.14.
    echo Solution la plus fiable : installer Python 3.12, puis relancer ce fichier.
    echo Lien : https://www.python.org/downloads/release/python-31210/
    echo Pendant l'installation, coche bien "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

echo.
echo Ouverture de l'application...
echo Si le navigateur ne s'ouvre pas, va sur : http://localhost:8501
echo.
python -m streamlit run app.py

pause
