@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_CMD="
where py >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
  where python >nul 2>&1
  if not errorlevel 1 set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
  echo Python was not found.
  echo Install Python 3.10 or newer from https://www.python.org/downloads/windows/
  echo During setup, enable "Add python.exe to PATH". Then reopen this file.
  pause
  exit /b 1
)
%PYTHON_CMD% -c "import pypdf, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Required libraries are missing. Installing them now...
  %PYTHON_CMD% -m pip install --user pypdf openpyxl
  if errorlevel 1 (
    echo Installation failed. Check that Python and pip are installed and internet access is available.
    pause
    exit /b 1
  )
)
%PYTHON_CMD% pdf_splitter_by_cell.py
if errorlevel 1 pause
endlocal

