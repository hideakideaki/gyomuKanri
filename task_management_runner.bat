@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "ROOT_DIR=%~dp0"
set "PYTHON_EXE="
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not exist "%ROOT_DIR%config.cmd" (
    echo ERROR: config.cmd was not found.
    echo Copy config.example.cmd to config.cmd and set PYTHON_EXE.
    exit /b 2
)
call "%ROOT_DIR%config.cmd"

if not defined PYTHON_EXE (
    echo ERROR: PYTHON_EXE is not set in config.cmd.
    exit /b 2
)
if not exist "%PYTHON_EXE%" (
    echo ERROR: The configured Python was not found.
    echo Python: "%PYTHON_EXE%"
    exit /b 3
)
if not exist "%ROOT_DIR%task_management\__init__.py" (
    echo ERROR: task_management was not found next to this BAT.
    echo Root: "%ROOT_DIR%"
    exit /b 2
)

"%PYTHON_EXE%" -X utf8 -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo ERROR: pywin32 is not available in the configured Python.
    echo Python: "%PYTHON_EXE%"
    exit /b 3
)

pushd "%ROOT_DIR%"
"%PYTHON_EXE%" -X utf8 -m task_management.cli %*
set "RESULT=%ERRORLEVEL%"
popd
exit /b %RESULT%
