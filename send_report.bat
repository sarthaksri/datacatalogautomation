@echo off
REM Wrapper that Task Scheduler invokes every 15 min.
REM Activates the venv (named "sarthak" on your machine) and runs the
REM email script. cd /d makes the script work no matter who launches it.

cd /d "%~dp0"

if exist "sarthak\Scripts\python.exe" (
    "sarthak\Scripts\python.exe" "%~dp0email_report.py" >> "%~dp0logs\email_report.log" 2>&1
) else if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" "%~dp0email_report.py" >> "%~dp0logs\email_report.log" 2>&1
) else (
    echo No venv found in either "sarthak" or "venv" >> "%~dp0logs\email_report.log"
    exit /b 1
)
