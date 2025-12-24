@echo off
chcp 65001 >nul
echo ========================================
echo    編譯 DiaryLD.exe
echo ========================================

REM 使用 .NET Framework 4.0 編譯器
set CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe

REM 檢查編譯器是否存在
if not exist "%CSC_PATH%" (
    echo 錯誤：找不到 C# 編譯器
    echo 路徑：%CSC_PATH%
    echo.
    echo 嘗試尋找其他版本...
    set CSC_PATH=C:\Windows\Microsoft.NET\Framework\v2.0.50727\csc.exe
    if not exist "!CSC_PATH!" (
        echo 仍然找不到編譯器，請確認已安裝 .NET Framework
        pause
        exit /b 1
    )
)

echo 使用編譯器：%CSC_PATH%
echo.

REM 刪除舊的執行檔
if exist DiaryLD.exe (
    echo 刪除舊的 DiaryLD.exe...
    del /f DiaryLD.exe
)

echo 開始編譯...
"%CSC_PATH%" /target:winexe /out:DiaryLD.exe /unsafe /optimize+ /win32icon:favicon.ico /r:OpenCvSharp.dll /r:OpenCvSharp.Extensions.dll DiaryLD.cs

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    編譯成功！
    echo    輸出檔案：DiaryLD.exe
    echo ========================================
    
    if exist "packages\OpenCvSharp4.runtime.win.4.5.5.20211231\runtimes\win-x64\native\OpenCvSharpExtern.dll" (
        copy /y "packages\OpenCvSharp4.runtime.win.4.5.5.20211231\runtimes\win-x64\native\OpenCvSharpExtern.dll" . >nul
        echo 已更新 OpenCvSharpExtern.dll
    )
) else (
    echo.
    echo ========================================
    echo    編譯失敗！請檢查錯誤訊息
    echo ========================================
)

echo.
pause
