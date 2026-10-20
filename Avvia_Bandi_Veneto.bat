@echo off
cd /d "%~dp0"
where py >nul 2>nul && (py -3 avvia.py %*) || (python avvia.py %*)
echo.
pause
