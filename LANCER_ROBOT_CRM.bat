@echo off
chcp 65001 >nul
echo ===============================================
echo Robot CRM FROID24 - export RDV experimental
echo ===============================================
if not exist .venv_robot (
  py -3.12 -m venv .venv_robot
  if errorlevel 1 python -m venv .venv_robot
)
call .venv_robot\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements_robot.txt
python -m playwright install chromium
python robot_crm_froid24.py
pause
