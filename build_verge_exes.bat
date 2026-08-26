@echo off
setlocal enabledelayedexpansion
title Build All Verge EXEs

set "BASEDIR=%USERPROFILE%\Downloads\GitHub\_src"
set "OUTDIR=%USERPROFILE%\Downloads\GitHub"
if not exist "%BASEDIR%" mkdir "%BASEDIR%"
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

echo.
echo  ============================================================
echo   Build All Verge Desk Solutions EXEs
echo  ============================================================
echo.
echo    Source:  %BASEDIR%
echo    Output:  %OUTDIR%
echo.

REM ── Check prerequisites ──
python --version >nul 2>&1
if errorlevel 1 (
    echo    ERROR: Python not found in PATH. Install Python 3.11+ first.
    pause
    exit /b 1
)
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo    PyInstaller not found. Installing...
    python -m pip install --upgrade pyinstaller
)
git --version >nul 2>&1
if errorlevel 1 (
    echo    ERROR: Git not found in PATH.
    pause
    exit /b 1
)
echo    Prerequisites OK
echo.

REM ── Verge repos and their spec files ──
set REPO_1=verge-xls-to-xlsx
set SPEC_1=verge_xls_to_xlsx.spec
set EXE_1=verge_xls_to_xlsx.exe

set REPO_2=verge-rebate-tools
set SPEC_2=rebate_tools.spec
set EXE_2=rebate_tools.exe

set REPO_3=verge-ups-tracking-checker
set SPEC_3=UPS_tracking_checker.spec
set EXE_3=UPS_tracking_checker.exe

set SUCCESS=0
set FAIL=0

for /L %%I in (1,1,3) do (
    call :BUILD_REPO %%I
)

echo.
echo  ============================================================
echo   Build Summary
echo  ============================================================
echo    Successful: !SUCCESS!
echo    Failed:     !FAIL!
echo    Output:     %OUTDIR%
echo  ============================================================
echo.

if exist "%OUTDIR%\*.exe" (
    echo  Built EXE files:
    dir /b "%OUTDIR%\*.exe" 2>nul
) else (
    echo  WARNING: No EXE files found in output directory.
)

echo.
pause
endlocal
exit /b 0

:BUILD_REPO
call set "REPO=%%REPO_%1%%"
call set "SPEC=%%SPEC_%1%%"
call set "EXE=%%EXE_%1%%"

echo  ────────────────────────────────────────────────────────────
echo    Building: !REPO! (!SPEC!)
echo  ────────────────────────────────────────────────────────────
echo.

REM ── Clone or pull ──
if exist "%BASEDIR%\!REPO!" (
    echo    Pulling latest...
    cd "%BASEDIR%\!REPO!"
    git pull 2>&1
) else (
    echo    Cloning !REPO!...
    git clone "https://github.com/abaduchanna/!REPO!.git" "%BASEDIR%\!REPO!" 2>&1
)

cd "%BASEDIR%\!REPO!"

REM ── Clean previous build ──
echo    Cleaning previous build...
if exist "build" rmdir /s /q "build"
if exist "dist"  rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"
del /s /q *.pyc 2>nul

REM ── Redirect workpath to TEMP (avoid OneDrive sync issues) ──
set "WORKBASE=%TEMP%\pyi_build\!REPO!"
if exist "%WORKBASE%" rmdir /s /q "%WORKBASE%"
mkdir "%WORKBASE%" 2>nul

REM ── Install deps ──
if exist "requirements.txt" (
    echo    Installing requirements...
    python -m pip install -r requirements.txt --quiet 2>nul
)

REM ── Build ──
echo    Building !SPEC!...
python -m PyInstaller "!SPEC!" --noconfirm --clean --workpath "%WORKBASE%" 2>&1

if errorlevel 1 (
    echo.
    echo    FAILED: !REPO!
    set /a FAIL+=1
    goto :EOF
)

echo    SUCCESS: !REPO!

REM ── Copy .exe to output ──
if exist "dist\!EXE!" (
    copy /Y "dist\!EXE!" "%OUTDIR%\!EXE!" >nul
    echo    Collected: %OUTDIR%\!EXE!
    set /a SUCCESS+=1
) else (
    echo    WARNING: dist\!EXE! not found
    set /a FAIL+=1
)

echo.
goto :EOF
