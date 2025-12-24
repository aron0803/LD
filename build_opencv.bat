@echo off
chcp 65001 >nul
echo ========================================
echo    編譯 DiaryLD.exe (with OpenCvSharp)
echo ========================================

REM 使用 .NET Framework 4.8 編譯器 (OpenCvSharp 需要至少 .NET 4.8)
set CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe

REM 檢查編譯器是否存在
if not exist "%CSC_PATH%" (
    echo 錯誤：找不到 C# 編譯器
    echo 路徑：%CSC_PATH%
    echo.
    echo 嘗試尋找其他版本...
    set CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
    if not exist "!CSC_PATH!" (
        echo 仍然找不到編譯器，請確認已安裝 .NET Framework
        pause
        exit /b 1
    )
)

echo 使用編譯器：%CSC_PATH%
echo.

REM OpenCvSharp DLL 路徑 (使用 net48 版本)
set OPENCV_DLL=packages\OpenCvSharp4.4.5.5.20211231\lib\net48\OpenCvSharp.dll
set OPENCV_EXT_DLL=packages\OpenCvSharp4.Extensions.4.5.5.20211231\lib\net48\OpenCvSharp.Extensions.dll

REM 檢查 OpenCvSharp DLL 是否存在
if not exist "%OPENCV_DLL%" (
    echo 錯誤：找不到 OpenCvSharp.dll
    echo 請先執行 NuGet 安裝套件
    pause
    exit /b 1
)

REM 刪除舊的執行檔
if exist DiaryLD.exe (
    echo 刪除舊的 DiaryLD.exe...
    del /f DiaryLD.exe
)

echo 開始編譯...
"%CSC_PATH%" /target:winexe /out:DiaryLD.exe /unsafe /optimize+ ^
    /langversion:5 ^
    /reference:"%OPENCV_DLL%" ^
    /reference:"%OPENCV_EXT_DLL%" ^
    DiaryLD.cs

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    編譯成功！
    echo    輸出檔案：DiaryLD.exe
    echo ========================================
    echo.
    echo 複製 OpenCvSharp DLLs 到當前目錄...
    
    REM 複製 OpenCvSharp DLLs
    copy /y "%OPENCV_DLL%" . >nul
    copy /y "%OPENCV_EXT_DLL%" . >nul
    
    REM 複製 Native DLLs (OpenCvSharpExtern.dll)
    copy /y "packages\OpenCvSharp4.runtime.win.4.5.5.20211231\runtimes\win-x64\native\OpenCvSharpExtern.dll" . >nul
    
    echo OpenCvSharp DLLs 已複製完成
) else (
    echo.
    echo ========================================
    echo    編譯失敗！請檢查錯誤訊息
    echo ========================================
)

echo.
pause
