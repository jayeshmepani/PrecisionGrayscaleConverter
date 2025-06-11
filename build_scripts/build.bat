@echo off
echo Building Precision Grayscale Converter...

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Install required packages
echo Installing required packages...
pip install -r requirements.txt
pip install pyinstaller

REM Create executable with PyInstaller
echo Creating executable...
pyinstaller --onefile --windowed --name "PrecisionGrayscaleConverter" --icon=icon.ico main.py

REM Copy executable to main directory
if exist "dist\PrecisionGrayscaleConverter.exe" (
    copy "dist\PrecisionGrayscaleConverter.exe" "PrecisionGrayscaleConverter.exe"
    echo.
    echo Build completed successfully!
    echo Executable: PrecisionGrayscaleConverter.exe
) else (
    echo Build failed!
)

pause