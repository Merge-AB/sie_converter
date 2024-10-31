@echo off
REM Change to the directory where your virtual environment is located
cd /d "%USERPROFILE%\OneDrive - Merge Group\Documents - Merge admin\SIE"

REM Activate the virtual environment
call "venv\Scripts\activate.bat"

REM Run the Python script using the virtual environment's python.exe
python "%USERPROFILE%\OneDrive - Merge Group\Documents - Merge admin\SIE\sie_converter.py"

REM Deactivate the virtual environment
pause