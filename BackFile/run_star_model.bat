@echo off
cd /d "%~dp0"
if exist "%~dp0\.venv\Scripts\python.exe" (
    "%~dp0\.venv\Scripts\python.exe" "%~dp0\star_model_generator.py"
) else (
    python "%~dp0\star_model_generator.py"
)
pause
