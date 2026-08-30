@echo off
if "%~1"=="" (
    call "%~dp0task_management_runner.bat" batch-folder refresh-all "%~dp0."
    pause
    exit /b %ERRORLEVEL%
)
call "%~dp0task_management_runner.bat" refresh-all %*
exit /b %ERRORLEVEL%
