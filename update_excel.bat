@echo off
setlocal
chcp 65001 >nul
echo Close both Excel workbooks before continuing.
echo This keeps user data and runs layout updates, refresh, and validation.
choice /C YN /N /M "Continue? [Y/N]: "
if errorlevel 2 exit /b 1
call "%~dp0task_management_runner.bat" batch-folder upgrade-layout "%~dp0."
set "RESULT=%ERRORLEVEL%"
echo.
if "%RESULT%"=="0" (
    echo Update completed.
) else (
    echo Update failed. The pre-update backup has been restored.
)
pause
exit /b %RESULT%
