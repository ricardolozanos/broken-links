@echo off

REM Activate conda environment
call conda activate projects

REM Change directory to the script folder
cd "C:\Users\Ricar\Desktop\script"

REM Run the Python script
python main.py

REM Pause the batch script to keep the window open
pause
