@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPT_DIR%"

if exist "%PROJECT_ROOT%task_management\__init__.py" goto project_found
for %%I in ("%SCRIPT_DIR%..\..\..") do set "PROJECT_ROOT=%%~fI\"

:project_found
if not exist "%PROJECT_ROOT%task_management\__init__.py" (
    echo ERROR: task_management package was not found.
    echo Expected project root: "%PROJECT_ROOT%"
    exit /b 2
)

set "CONDA_PYTHON=%USERPROFILE%\miniconda3\python.exe"
if not exist "%CONDA_PYTHON%" (
    echo ERROR: Miniconda Python was not found.
    echo Expected Python: "%CONDA_PYTHON%"
    exit /b 3
)

pushd "%PROJECT_ROOT%"
"%CONDA_PYTHON%" -m task_management.cli %*
set "RESULT=%ERRORLEVEL%"
popd
exit /b %RESULT%
