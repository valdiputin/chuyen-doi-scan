@echo off
cd /d "%~dp0"
title Web App Chuyen Doi Dinh Dang

if not exist .venv (
  echo [1/2] Dang khoi tao moi truong ao Python...
  python -m venv .venv
)
echo [1/2] Dang kiem tra thu vien...
.venv\Scripts\pip install -q -r requirements.txt

echo [2/2] Dang khoi dong dev server tai http://localhost:8000 ...
start "" "http://localhost:8000"
.venv\Scripts\python dev_server.py --port 8000
