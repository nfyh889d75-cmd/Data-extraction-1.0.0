@echo off
setlocal enabledelayedexpansion
chcp 936 >nul 2>&1

:: ============================================================
:: 【用户配置区】
:: ============================================================
set "SOURCE_FOLDER=C:\Users\YUN\Desktop\ceshi"
set "OUTPUT_TARGET=DESKTOP"
set "WEBSITE_URL=https://nfyh889d75-cmd.github.io/Data-extraction-1.0.2/table-extractor.html"
:: ============================================================

if /i "%OUTPUT_TARGET%"=="DESKTOP" (
    set "OUT_DIR=%USERPROFILE%\Desktop"
) else (
    set "OUT_DIR=%SOURCE_FOLDER%"
)

echo ========================================
echo   WPS xls to xlsx Batch Converter
echo ========================================
echo [Source] %SOURCE_FOLDER%
echo [Output] %OUT_DIR%
echo.

if not exist "%SOURCE_FOLDER%" (
    echo [ERROR] Folder not found: %SOURCE_FOLDER%
    pause
    exit /b 1
)

:: 生成 VBS 脚本（使用逐行 echo 追加模式，兼容性最强）
set "VBS_FILE=%TEMP%\wps_convert_%RANDOM%.vbs"
if exist "%VBS_FILE%" del "%VBS_FILE%"

>>"%VBS_FILE%" echo Dim wpsApp, fso, srcFolder, outDir
>>"%VBS_FILE%" echo Set wpsApp = CreateObject("Ket.Application")
>>"%VBS_FILE%" echo Set fso = CreateObject("Scripting.FileSystemObject")
>>"%VBS_FILE%" echo srcFolder = "%SOURCE_FOLDER%"
>>"%VBS_FILE%" echo outDir = "%OUT_DIR%"
>>"%VBS_FILE%" echo If Not wpsApp Is Nothing Then
>>"%VBS_FILE%" echo     wpsApp.Visible = False
>>"%VBS_FILE%" echo     wpsApp.DisplayAlerts = 0
>>"%VBS_FILE%" echo     Dim folder: Set folder = fso.GetFolder(srcFolder)
>>"%VBS_FILE%" echo     Dim file, count: count = 0
>>"%VBS_FILE%" echo     For Each file In folder.Files
>>"%VBS_FILE%" echo         If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
>>"%VBS_FILE%" echo             Dim fullPath, baseName, newPath
>>"%VBS_FILE%" echo             fullPath = file.Path
>>"%VBS_FILE%" echo             baseName = fso.GetBaseName(file.Name)
>>"%VBS_FILE%" echo             newPath = outDir ^& "\" ^& baseName ^& ".xlsx"
>>"%VBS_FILE%" echo             On Error Resume Next
>>"%VBS_FILE%" echo             Dim wb: Set wb = wpsApp.Workbooks.Open(fullPath, False, True)
>>"%VBS_FILE%" echo             If Err.Number = 0 And Not wb Is Nothing Then
>>"%VBS_FILE%" echo                 wb.SaveAs newPath, 51
>>"%VBS_FILE%" echo                 wb.Close False
>>"%VBS_FILE%" echo                 count = count + 1
>>"%VBS_FILE%" echo                 WScript.Echo "[OK] " ^& baseName ^& ".xls -^> .xlsx"
>>"%VBS_FILE%" echo             Else
>>"%VBS_FILE%" echo                 WScript.Echo "[FAIL] " ^& file.Name ^& " - " ^& Err.Description
>>"%VBS_FILE%" echo                 Err.Clear
>>"%VBS_FILE%" echo             End If
>>"%VBS_FILE%" echo             On Error GoTo 0
>>"%VBS_FILE%" echo         End If
>>"%VBS_FILE%" echo     Next
>>"%VBS_FILE%" echo     wpsApp.Quit
>>"%VBS_FILE%" echo     WScript.Echo ""
>>"%VBS_FILE%" echo     WScript.Echo "Done! Converted " ^& count ^& " file(s)"
>>"%VBS_FILE%" echo Else
>>"%VBS_FILE%" echo     WScript.Echo "[ERROR] Cannot create Ket.Application"
>>"%VBS_FILE%" echo     WScript.Echo "Please ensure WPS Office is installed"
>>"%VBS_FILE%" echo End If
>>"%VBS_FILE%" echo Set wpsApp = Nothing
>>"%VBS_FILE%" echo Set fso = Nothing

echo [Converting...]
cscript //nologo "%VBS_FILE%"
del "%VBS_FILE%" >nul 2>&1

echo.
echo [Opening Website] %WEBSITE_URL%
start "" "%WEBSITE_URL%"

echo.
echo ========================================
echo   All tasks completed
echo ========================================
pause
exit /b 0