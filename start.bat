@echo off
echo ================================================
echo   VT to EcoTEA Conversion Tool
echo ================================================
echo.

:: Install dependencies (only needed first time)
pip install -r requirements.txt --quiet

echo Starting tool...
python app.py
pause
