@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: ============================================================
:: SHA256_checksum_tool.bat
:: - Generate or verify SHA-256 checksums for all files in a folder
:: - Drag-and-drop friendly or run from command line
:: - Adds folder-level SHA256 hash at the end of the manifest
:: ============================================================

:: --------- Parse arguments ----------
set "MODE=%~1"
set "TARGET=%~2"

:: If no target, assume drag-and-drop or current folder
if "%TARGET%"=="" set "TARGET=%CD%"

:: If no mode and a folder is dropped, prompt
if "%MODE%"=="" (
    set /p MODE=Select mode [G=Generate, V=Verify]: 
    if /I "%MODE%"=="G" set MODE=generate
    if /I "%MODE%"=="V" set MODE=verify
)

:: Check folder exists
if not exist "%TARGET%\NUL" (
    echo [ERROR] Path not found: "%TARGET%"
    pause
    exit /b 1
)

:: Normalize base path
set "BASE=%~f2"
if "%BASE%"=="" set "BASE=%CD%"
if not exist "%BASE%\NUL" set "BASE=%~dp2"
if "%BASE:~-1%"=="\" set "BASE=%BASE:~0,-1%"

set "MANIFEST=%BASE%\SHA256SUMS.txt"
set "TMP=%BASE%\SHA256SUMS.tmp"

certutil >NUL 2>&1
if errorlevel 1 (
    echo [ERROR] 'certutil' not found. This script requires Windows certutil.
    pause
    exit /b 1
)

:: ============================================================
:: GENERATE MODE
:: ============================================================
if /I "%MODE%"=="generate" (
    > "%TMP%" (
        echo # SHA256 checksums
        echo # Folder: %BASE%
        echo # Generated: %DATE% %TIME%
        echo.
    )

    echo.
    echo ===========================================
    echo Generating SHA-256 checksums under:
    echo   %BASE%
    echo Writing to:
    echo   %MANIFEST%
    echo ===========================================
    echo.

    set "COUNT=0"
    for /r "%BASE%" %%F in (*) do (
        if /I not "%%~fF"=="%MANIFEST%" if /I not "%%~fF"=="%TMP%" (
            set "REL=%%~fF"
            set "REL=!REL:%BASE%\=!"
            set "HASH="
            for /f "usebackq tokens=* delims=" %%H in (`
                certutil -hashfile "%%~fF" SHA256 ^| findstr /R /C:"^[0-9A-F][0-9A-F]"
            `) do (
                if not defined HASH set "HASH=%%H"
            )
            >> "%TMP%" echo !HASH! *!REL!
            set /a COUNT+=1 >NUL
            echo [OK] !REL!
        )
    )

    :: Compute folder-level hash of the manifest
    set "FOLDER_HASH="
    for /f "usebackq tokens=* delims=" %%H in (`
        certutil -hashfile "%TMP%" SHA256 ^| findstr /R /C:"^[0-9A-F][0-9A-F]"
    `) do (
        if not defined FOLDER_HASH set "FOLDER_HASH=%%H"
    )

    if defined FOLDER_HASH (
        >> "%TMP%" echo.
        >> "%TMP%" echo # Folder-level SHA256 (of this manifest):
        >> "%TMP%" echo # %FOLDER_HASH%
        echo.
        echo Folder-level SHA256: %FOLDER_HASH%
    )

    move /y "%TMP%" "%MANIFEST%" >NUL
    echo.
    echo Completed. Files hashed: %COUNT%
    echo Manifest saved: %MANIFEST%
    echo.
    pause
    exit /b 0
)

:: ============================================================
:: VERIFY MODE
:: ============================================================
if /I "%MODE%"=="verify" (
    if not exist "%MANIFEST%" (
        echo [ERROR] Manifest not found: %MANIFEST%
        pause
        exit /b 1
    )

    echo.
    echo ===========================================
    echo Verifying files using:
    echo   %MANIFEST%
    echo ===========================================
    echo.

    set "ERRORCOUNT=0"
    for /f "usebackq tokens=1,* delims= " %%H in ("%MANIFEST%") do (
        set "EXPECTED=%%H"
        set "FILE=%%I"
        if "!FILE!"=="" (
            rem skip empty lines or comments
        ) else (
            if not exist "%BASE%\!FILE!" (
                echo [MISSING] !FILE!
                set /a ERRORCOUNT+=1
            ) else (
                set "ACTUAL="
                for /f "usebackq tokens=* delims=" %%A in (`
                    certutil -hashfile "%BASE%\!FILE!" SHA256 ^| findstr /R /C:"^[0-9A-F][0-9A-F]"
                `) do (
                    if not defined ACTUAL set "ACTUAL=%%A"
                )
                if /I "!EXPECTED!"=="!ACTUAL!" (
                    echo [OK] !FILE!
                ) else (
                    echo [FAIL] !FILE!
                    set /a ERRORCOUNT+=1
                )
            )
        )
    )

    echo.
    if "%ERRORCOUNT%"=="0" (
        echo ===========================================
        echo All files verified successfully.
        echo ===========================================
    ) else (
        echo ===========================================
        echo Verification completed with %ERRORCOUNT% error(s).
        echo ===========================================
    )
    echo.
    pause
    exit /b 0
)

:: ============================================================
:: INVALID MODE
:: ============================================================
echo [ERROR] Invalid mode: %MODE%
echo Must be 'generate' or 'verify'
pause
exit /b 1
