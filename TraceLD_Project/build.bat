@echo off
chcp 65001 >nul
echo ========================================
echo    Build TraceLD.exe x64
echo ========================================

set CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe

if not exist "%CSC_PATH%" (
    set CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
)

if not exist "%CSC_PATH%" (
    echo Error: C# Compiler not found.
    exit /b 1
)

if exist TraceLD.exe del /f TraceLD.exe

echo Compiling...
"%CSC_PATH%" /target:winexe /out:TraceLD.exe /platform:x64 /unsafe /optimize+ /win32icon:favicon.ico /r:OpenCvSharp.dll /r:OpenCvSharp.Extensions.dll TraceLD.cs

if %errorlevel% equ 0 (
    echo Build Success!
    
    if exist "packages\OpenCvSharp4.runtime.win.4.5.5.20211231\runtimes\win-x64\native\OpenCvSharpExtern.dll" (
        copy /y "packages\OpenCvSharp4.runtime.win.4.5.5.20211231\runtimes\win-x64\native\OpenCvSharpExtern.dll" . >nul
        echo Updated OpenCvSharpExtern.dll x64
    )
    
    echo Output: TraceLD.exe
) else (
    echo Build Failed!
)
pause
