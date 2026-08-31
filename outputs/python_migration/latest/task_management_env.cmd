@echo off
rem PC-specific settings used by task_management_runner.bat.
rem Edit these two values when moving the release folder to another PC.
set "PYTHON_EXE=%USERPROFILE%\miniconda3\python.exe"
for %%I in ("%~dp0..\..\..") do set "PROJECT_ROOT=%%~fI"
