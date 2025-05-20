@echo off
title SPMS PET Form Processor
color 1F
cls

echo ================================================
echo         SPMS PET Form Processor
echo ================================================
echo.

REM Use the portable Python - no installation needed
set PYTHON_PATH=J:\SPMS_Registration_Structured\python-portable\python.exe

REM Check if portable Python exists
if not exist "%PYTHON_PATH%" (
    echo ERROR: Portable Python not found at %PYTHON_PATH%
    echo Please contact the administrator to set up the portable Python environment.
    pause
    exit /b 1
)

REM Ask for team member 
echo Please select your name:
echo.
echo    [1] Asa
echo    [2] Natalja
echo    [3] Shyam
echo    [4] Tima
echo.
choice /c 1234 /n /m "Enter number (1-4): "

if %ERRORLEVEL%==1 set MEMBER_DIR=Asa
if %ERRORLEVEL%==2 set MEMBER_DIR=Natalja
if %ERRORLEVEL%==3 set MEMBER_DIR=Shyam
if %ERRORLEVEL%==4 set MEMBER_DIR=Tima

echo.
echo Processing PET forms for %MEMBER_DIR%...
echo.

REM Run the Python script with the portable Python
"%PYTHON_PATH%" J:\SPMS_Registration_Structured\Bugatti\main.py %MEMBER_DIR%

echo.
echo ================================================
echo Process complete!
echo.
echo Check these locations for your output files:
echo - J:\SPMS_Registration_Structured\Team Members\%MEMBER_DIR%\CombinedExtractedColumns.xlsx
echo - J:\SPMS_Registration_Structured\Team Members\%MEMBER_DIR%\Uploads\MassUpload.xlsx
echo ================================================
echo.
pause