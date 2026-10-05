@echo off
setlocal

echo ========================================================
echo Starting ReachOut App...
echo ========================================================

if exist "C:\Users\DELL\anaconda3\python.exe" (
    set "PYTHON_EXE=C:\Users\DELL\anaconda3\python.exe"
) else (
    set "PYTHON_EXE=python"
)

echo Using Python: %PYTHON_EXE%
%PYTHON_EXE% app.py

pause
