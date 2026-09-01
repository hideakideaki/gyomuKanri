@echo off
setlocal
chcp 65001 >nul
echo Checking the task-management environment...
call "%~dp0task_management_runner.bat" batch-folder validate "%~dp0."
set "RESULT=%ERRORLEVEL%"
echo.
if "%RESULT%"=="0" (
    echo PASS: Python, pywin32, workbooks, and workbook structures are available.
) else (
    echo FAIL: See the error above.
    echo Set the Python path in config.cmd.
    echo Keep the whole folder available offline in OneDrive.
)
echo.
pause
exit /b %RESULT%
