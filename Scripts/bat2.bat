@echo off

REM Activate conda environment
call conda activate projects

REM Change directory to the script folder
cd "C:\Users\Ricar\Desktop\script"

REM Run the Python script
python sitemap-to-excel.py


REM Pause to keep the command window open (optional)
pause
