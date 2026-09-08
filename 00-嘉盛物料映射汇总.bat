@echo off
cd /d "%~dp0"

echo ============================================================
echo  Excel Material Summary Tool
echo ============================================================
echo.

set "PY_CMD=py"

echo Python: %PY_CMD%
echo.
echo Checking dependencies...
%PY_CMD% -m pip install openpyxl xlrd
echo.
echo ============================================================
echo  Running...
echo ============================================================
echo.

%PY_CMD% 01-ºŒ ¢ŒÔ¡œ”≥…‰ª„◊‹.py

echo.
pause
