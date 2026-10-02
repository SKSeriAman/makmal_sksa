@echo off
echo Starting Python Web Application for SKSA Computer Lab...
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py
) else (
    python app.py
)
pause
