@echo off
setlocal EnableExtensions
set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE="
set "PROJECT_ROOT="

if not exist "%SCRIPT_DIR%task_management_env.cmd" (
    echo ERROR: task_management_env.cmd was not found.
    echo Copy task_management_env.example.cmd as task_management_env.cmd and edit it.
    exit /b 2
)
call "%SCRIPT_DIR%task_management_env.cmd"

if not defined PYTHON_EXE (
    echo ERROR: PYTHON_EXE is not set in task_management_env.cmd.
    exit /b 2
)
if not exist "%PYTHON_EXE%" (
    echo ERROR: PYTHON_EXE does not exist.
    echo Configured path: "%PYTHON_EXE%"
    exit /b 3
)
if not defined PROJECT_ROOT (
    echo ERROR: PROJECT_ROOT is not set in task_management_env.cmd.
    exit /b 2
)
if not exist "%PROJECT_ROOT%\task_management\__init__.py" (
    echo ERROR: task_management package was not found under PROJECT_ROOT.
    echo Configured root: "%PROJECT_ROOT%"
    exit /b 2
)

"%PYTHON_EXE%" -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo ERROR: pywin32 is not available in the configured Python environment.
    echo Python: "%PYTHON_EXE%"
    exit /b 3
)

pushd "%PROJECT_ROOT%"
"%PYTHON_EXE%" -m task_management.cli %*
set "RESULT=%ERRORLEVEL%"
popd
exit /b %RESULT%
