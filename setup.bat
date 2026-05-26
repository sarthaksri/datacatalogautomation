@echo off
echo Setting up virtual environment...

python -m venv venv
if errorlevel 1 (
    echo ERROR: Could not create venv. Is Python installed?
    exit /b 1
)

call venv\Scripts\activate.bat

echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

echo Installing Playwright browsers...
playwright install chromium

echo.
echo Setup complete! Run the automation with:
echo   venv\Scripts\activate
echo   python main.py
