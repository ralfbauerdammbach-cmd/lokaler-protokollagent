@echo off
cd /d "%~dp0"
python -c "import faster_whisper, docx" >nul 2>&1
if errorlevel 1 (
  echo Installiere fehlende Python-Pakete...
  python -m pip install -r requirements.txt
  if errorlevel 1 (pause & exit /b 1)
)
python agent.py
if errorlevel 1 pause
