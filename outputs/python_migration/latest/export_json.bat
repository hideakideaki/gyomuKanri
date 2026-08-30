@echo off
if "%~1"=="" (
    call "%~dp0task_management_runner.bat" batch-folder export-auto "%~dp0."
    pause
    exit /b %ERRORLEVEL%
)
call "%~dp0task_management_runner.bat" export-auto %*
exit /b %ERRORLEVEL%
