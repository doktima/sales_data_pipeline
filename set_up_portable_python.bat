@echo off
echo Setting up portable Python environment...
echo This only needs to be run once by an administrator.
echo.

powershell -ExecutionPolicy Bypass -File "setup_portable_python.ps1"

echo.
echo Setup completed. Press any key to exit...
pause > nul