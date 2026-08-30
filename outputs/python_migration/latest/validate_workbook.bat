@echo off
if "%~1"=="" (
    call "%~dp0task_management_runner.bat" batch-folder validate "%~dp0."
    pause
    exit /b %ERRORLEVEL%
)
call "%~dp0task_management_runner.bat" validate %*
exit /b %ERRORLEVEL%
