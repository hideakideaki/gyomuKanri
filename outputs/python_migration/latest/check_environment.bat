@echo off
setlocal
echo Checking the configured task-management environment...
call "%~dp0task_management_runner.bat" batch-folder validate "%~dp0."
set "RESULT=%ERRORLEVEL%"
echo.
if "%RESULT%"=="0" (
    echo PASS: Configured Python, pywin32, project, and both workbooks are available.
) else (
    echo FAIL: See the error above.
    echo Edit PYTHON_EXE and PROJECT_ROOT in task_management_env.cmd.
)
echo.
pause
exit /b %RESULT%
