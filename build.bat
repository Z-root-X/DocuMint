@echo off
echo =================================================
echo        DocuMint Build Script
echo =================================================

REM Check if PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

echo.
echo Building executable...
pyinstaller DocuMint.spec --clean --noconfirm

echo.
echo =================================================
if exist "dist\DocuMint\DocuMint.exe" (
    echo Build SUCCESS!
    echo Executable location: dist\DocuMint\DocuMint.exe
) else (
    echo Build FAILED. Check errors above.
)
echo =================================================
pause