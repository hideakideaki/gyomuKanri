@echo off
if "%~1"=="" (
    call "%~dp0task_management_runner.bat" batch-folder backup-auto "%~dp0."
    pause
    exit /b %ERRORLEVEL%
)
call "%~dp0task_management_runner.bat" backup-auto %*
exit /b %ERRORLEVEL%
