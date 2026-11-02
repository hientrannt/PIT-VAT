@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  echo Lan dau chay: dang cai dat thu vien, vui long doi...
  python -m venv .venv || (echo Can cai Python 3.10+ tu https://www.python.org/downloads/ va tick "Add python.exe to PATH" & pause & exit /b 1)
  .venv\Scripts\python -m pip install --upgrade pip
  .venv\Scripts\python -m pip install -r requirements.txt
)
.venv\Scripts\python -m streamlit run app.py
pause
