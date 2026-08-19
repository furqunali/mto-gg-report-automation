@echo off
REM ===================================================================
REM  MTO & G&G - FMO monthly report builder  (double-click to run)
REM  1. Save the new "Product Group Ranking" from Focal Point into the
REM     report folder (the folder ABOVE this one).
REM  2. Make sure TEMPLATE_FILE in update_report.py points at last
REM     month's finished workbook.
REM  3. Double-click this file.
REM ===================================================================
cd /d "%~dp0"
echo Building this month's MTO ^& G^&G report...
echo.
python update_report.py
echo.
echo ===================================================================
echo Done. Review the new workbook + the "Dashboard ....html" file in the
echo report folder. This window stays open so you can read the log above.
echo ===================================================================
pause
