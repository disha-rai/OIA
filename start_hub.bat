@echo off
REM ===========================================================
REM  OIA Hub - start script for the office PC
REM  Runs the Hub so every computer on the network can open it,
REM  and restarts it automatically if it ever stops.
REM ===========================================================
cd /d "%~dp0"

REM Listen on the network (not just this PC)
set OIA_HUB_HOST=0.0.0.0
set OIA_HUB_PORT=5000

:loop
echo [%date% %time%] Starting OIA Hub >> hub.log
python app.py >> hub.log 2>&1
echo [%date% %time%] Hub stopped - restarting in 10 seconds >> hub.log
timeout /t 10 /nobreak > nul
goto loop
