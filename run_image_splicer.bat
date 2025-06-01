@echo off
REM Change directory to the script's location
cd /d "%~dp0"

REM Run the Python application without a console window
echo Starting Image Splicer (no console)...
start "" pythonw image_splicer_app.py

REM The 'start ""' part is used to run pythonw in a way that the .bat script itself can exit immediately
REM without waiting for pythonw to finish, and it helps prevent issues with paths containing spaces.