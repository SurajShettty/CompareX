@echo off
REM Launch the Excel Diff Checker UI using the local virtual environment.

cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment not found.
    echo Run the following command once to create it:
    echo   python -m venv .venv
    echo   .venv\Scripts\pip install tkinterdnd2 pandas openpyxl
    pause
    exit /b 1
)

".venv\Scripts\python.exe" "diff_checker_ui.py"
