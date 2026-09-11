@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===============================================
echo   Draft Data Extractor V1
echo   PDF -> accurate voter Excel
echo ===============================================
echo.
python draft_extractor.py
if errorlevel 1 (
  echo.
  echo Extraction failed. Read the error above.
  pause
  exit /b 1
)
echo.
echo ===============================================
echo   Extraction complete.
echo   Open output\Master_Voter_List.xlsx
echo ===============================================
pause
