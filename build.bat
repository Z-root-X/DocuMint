@echo off

echo Building DocuMint...

REM Ensure you have PyInstaller installed: pip install pyinstaller

pyinstaller --onefile --windowed --icon=app.ico --name "DocuMint" gui.py

echo.
echo =================================================
echo Build complete!

echo The executable can be found in the 'dist' folder.

echo.
echo Preparing final package...

REM Create a distribution folder
if exist .\dist\DocuMint_Package rmdir /s /q .\dist\DocuMint_Package
mkdir .\dist\DocuMint_Package

REM Copy the necessary files to the package folder
copy .\dist\DocuMint.exe .\dist\DocuMint_Package\
copy .\example_data.csv .\dist\DocuMint_Package\
copy .\example_template.txt .\dist\DocuMint_Package\
copy .\README.txt .\dist\DocuMint_Package\

echo.
echo =================================================
echo Final package created in the 'dist\DocuMint_Package' folder.

echo You can now zip the 'DocuMint_Package' folder and distribute it.

pause
