@echo off
REM Called by Windows Task Scheduler every day at 10:30.
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" auto_report.py %*
) else (
    python auto_report.py %*
)
exit /b %ERRORLEVEL%
