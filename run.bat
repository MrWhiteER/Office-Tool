@echo off
REM ---- Office Tool (Windows) ----
where python >nul 2>nul || (echo Python is not installed. Get it from https://python.org & pause & exit /b)
python -m pip install -r requirements.txt
python app.py
pause
