@echo off
echo =====================================================
echo Building Standalone Executables with PyInstaller
echo =====================================================
echo.

REM Clean previous builds
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

echo Installing requirements...
pip install -r ../requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install requirements
    pause
    exit /b 1
)

echo Building with fixed configuration...
python -m PyInstaller PDF_Extractor.spec --clean
if %errorlevel% equ 0 (
    echo.
    echo BUILD SUCCESSFUL!
    echo Executable: dist\PDF_Extractor.exe
) else (
    echo.
    echo Build failed. Check error messages above.
)
pause


if %errorlevel% neq 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo.
echo =====================================================
echo BUILD COMPLETE!
echo.
echo Executable: dist\PDF_Extractor.exe
echo.
echo To run: dist\PDF_Extractor.exe
echo =====================================================
pause