@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===============================================
echo   Draft Data Extractor V1.2 - Setup
 echo ===============================================
echo.
python -m pip install numpy pandas pytesseract pdf2image pillow openpyxl
if errorlevel 1 (
  echo.
  echo Python package installation failed.
  pause
  exit /b 1
)
echo.
echo Python packages are ready.
echo Tesseract and Poppler are checked by the extractor when it runs.
echo ===============================================
pause
