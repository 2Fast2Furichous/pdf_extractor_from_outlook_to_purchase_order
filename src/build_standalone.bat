@echo off
echo =====================================================
echo Building Standalone Executables with PyInstaller
echo =====================================================
echo.

echo Building PDF_Extractor.exe...
python -m PyInstaller PDF_Extractor.spec --clean
if %errorlevel% neq 0 (
    echo ERROR: Build failed for PDF_Extractor.exe
    pause
    exit /b 1
)

echo.
echo =====================================================
echo BUILD SUCCESSFUL!
echo =====================================================
pause
