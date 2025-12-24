@echo off
chcp 65001 >nul
echo ========================================
echo    Build TraceLD.exe (Python)
echo ========================================

if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist TraceLD_Py.spec del /f TraceLD_Py.spec

echo Compiling with PyInstaller...
pyinstaller --noconfirm --onefile --windowed --icon="favicon.ico" --name "TraceLD_Py" --add-data "trace;trace" TraceLD.py

if %errorlevel% equ 0 (
    echo.
    echo Build Success!
    echo Output: dist\TraceLD_Py.exe
    
    echo Copying to Release folder...
    if not exist Release_TraceLD_Py mkdir Release_TraceLD_Py
    copy /y dist\TraceLD_Py.exe Release_TraceLD_Py\TraceLD.exe
    
    echo Release updated: Release_TraceLD_Py\TraceLD.exe
) else (
    echo.
    echo Build Failed!
)
pause
