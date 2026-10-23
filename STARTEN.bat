@echo off
cd /d "%~dp0"
py --version >nul 2>&1
if errorlevel 1 (
  echo Python-Launcher py wurde nicht gefunden.
  pause
  exit /b 1
)
py -c "import faster_whisper, docx" >nul 2>&1
if errorlevel 1 (
  echo Installiere fehlende Python-Pakete...
  py -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Paketinstallation fehlgeschlagen.
    pause
    exit /b 1
  )
)
py agent.py
if errorlevel 1 pause
