@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: ============================================================
:: SHA256_checksum_generator.bat
:: - Generates SHA-256 checksums for all files in a folder (recursively)
:: - Output: SHA256SUMS.txt in the target folder
:: - Usage:
::     1) Drag & drop a folder onto this .bat
::     2) Or: SHA256_checksum_generator.bat "C:\path\to\folder"
:: Requires: Windows built-in 'certutil'
:: ============================================================

:: -------- Validate input --------
if "%~1"=="" (
  echo.
  echo Usage:
  echo   Drag and drop a folder onto this script
  echo   or run: %~nx0 "C:\path\to\folder"
  echo.
  pause
  exit /b 1
)

if not exist "%~1" (
  echo [ERROR] Path not found: "%~1"
  exit /b 1
)

:: If a file was passed, switch to its parent folder
set "TARGET=%~1"
if exist "%TARGET%\NUL" (
  rem It's a folder
) else (
  echo [INFO] You passed a file. Using its parent folder...
  set "TARGET=%~dp1"
)

:: Normalize base path (remove trailing backslash if present)
set "BASE=%~f1"
if not exist "%BASE%\NUL" set "BASE=%~dp1"
if "%BASE:~-1%"=="\" set "BASE=%BASE:~0,-1%"

:: Prepare output file
set "OUT=%BASE%\SHA256SUMS.txt"
set "TMP=%BASE%\SHA256SUMS.tmp"

:: Confirm certutil exists
certutil >NUL 2>&1
if errorlevel 1 (
  echo [ERROR] 'certutil' not found. This script needs Windows certutil.
  exit /b 1
)

:: Header
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
echo   %OUT%
echo ===========================================
echo.

set "COUNT=0"
for /r "%BASE%" %%F in (*) do (
  rem Skip the output files themselves if re-running
  if /I not "%%~fF"=="%OUT%" if /I not "%%~fF"=="%TMP%" (
    set "FULL=%%~fF"
    set "REL=%%~fF"
    set "REL=!REL:%BASE%\=!"

    rem Compute SHA256 with certutil, capture the hex line
    set "HASH="
    for /f "usebackq tokens=* delims=" %%H in (`
        certutil -hashfile "%%~fF" SHA256 ^| findstr /R /C:"^[0-9A-F][0-9A-F]"
    `) do (
      if not defined HASH set "HASH=%%H"
    )

    if not defined HASH (
      echo [WARN] Could not compute hash: "%%~fF"
    ) else (
      >> "%TMP%" echo !HASH! *!REL!
      set /a COUNT+=1 >NUL
      echo [OK] !REL!
    )
  )
)

rem Finalize
move /y "%TMP%" "%OUT%" >NUL
echo.
echo ===========================================
echo Completed. Files hashed: %COUNT%
echo Manifest: %OUT%
echo ===========================================
echo.

:: Optional: compute a single "folder hash" by hashing the manifest itself
set "FOLDER_HASH="
for /f "usebackq tokens=* delims=" %%H in (`
    certutil -hashfile "%OUT%" SHA256 ^| findstr /R /C:"^[0-9A-F][0-9A-F]"
`) do (
  if not defined FOLDER_HASH set "FOLDER_HASH=%%H"
)

if defined FOLDER_HASH (
  echo Folder-level SHA256 (of manifest):
  echo   %FOLDER_HASH%
  echo.
  >> "%OUT%" echo.
  >> "%OUT%" echo # Folder-level SHA256 (of this manifest):
  >> "%OUT%" echo # %FOLDER_HASH%
)

endlocal
exit /b 0
