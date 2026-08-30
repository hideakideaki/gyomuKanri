@echo off
call "%~dp0task_management_runner.bat" import-json %*
exit /b %ERRORLEVEL%
