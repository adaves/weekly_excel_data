@echo off
echo ===== Circana Excel Processor =====
echo Finding and processing Excel files...

:: Activate virtual environment
call venv\Scripts\activate.bat

:: Run the Python script
python circana_data_script.py

:: Deactivate virtual environment
call deactivate

echo.
echo Process completed! Press any key to exit...
pause 