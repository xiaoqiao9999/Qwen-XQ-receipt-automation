@echo off
rem One-click entry for the receipt automation pipeline.
rem run.py picks the Python interpreter that has the dependencies installed.
cd /d "%~dp0"
python run.py
pause
