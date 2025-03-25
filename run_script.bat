# Update the run_script.bat file in the Environment folder
$batchContent = @'
@echo off
REM Universal script runner for the pet registration system

REM Get team member name
set /p MEMBER_NAME=Enter your name (Asa, Natalja, Shyam, Tima): 

REM List available scripts
echo.
echo Available scripts:
echo -----------------
dir /b J:\SPMS_Pet_Registration\Bugatti\*.py

REM Get script to run
set /p SCRIPT_NAME=Enter script name (e.g. Headers.py): 

REM Run the script with the team member name
echo.
echo Running %SCRIPT_NAME% for %MEMBER_NAME%...
J:\SPMS_Pet_Registration\Environment\team_env\Scripts\python.exe J:\SPMS_Pet_Registration\Bugatti\%SCRIPT_NAME% %MEMBER_NAME%

pause

$batchContent | Out-File -FilePath "J:\SPMS_Pet_Registration\Environment\run_script.bat" -Encoding ASCII