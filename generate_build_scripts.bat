@echo off
setlocal enabledelayedexpansion
REM ==========================================================================
REM  Generate Build Scripts — Creates build_*.bat for all repos
REM  Developed by Abad Umair Channa
REM
REM  Run this ONCE. It creates a folder of build scripts that you can
REM  double-click to build any app's .exe locally.
REM
REM  Output: %USERPROFILE%\Downloads\GitHub\build_all\build_*.bat
REM ==========================================================================

set "BASEDIR=%USERPROFILE%\Downloads\GitHub"
set "OUTDIR=%BASEDIR%\build_all"
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

echo.
echo  ============================================================
echo   Generating Build Scripts for All Repos
echo  ============================================================
echo.
echo    Output: %OUTDIR%
echo.

set COUNT=0

REM ── GFH repos ──
call :WRITE gfh-accessories-order-history-scraper gfh_accessories_order_history_scraper gfh_accessories_order_history_scraper
call :WRITE gfh-inventory-aging-processor GFH_Inventory_Aging_Processor GFH_Inventory_Aging_Processor
call :WRITE gfh-rebate-tools gfh_rebate_tools gfh_rebate_tools
call :WRITE gfh-ups-tracking-checker gfh_ups_tracking_checker gfh_ups_tracking_checker
call :WRITE gfh-xls-to-xlsx gfh_xls_to_xlsx gfh_xls_to_xlsx

REM ── VidaPay repos ──
call :WRITE vidapay-extractor VidaPay_Incentive_Extractor_FULL VidaPay_Incentive_Extractor_FULL
call :WRITE vidapay-extractor VidaPay_Incentive_Extractor_TRIAL VidaPay_Incentive_Extractor_TRIAL
call :WRITE vidapay-extractor VidaPay_Incentive_Extractor_TRIAL_1Y VidaPay_Incentive_Extractor_TRIAL_1Y
call :WRITE vidapay-gfh GFH_Accessories_Ordering GFH_Accessories_Ordering
call :WRITE vidapay-gfh GFH_Inventory_Audit GFH_Inventory_Audit
call :WRITE vidapay-gfh GFH_Inventory_Audit_Timesheet GFH_Inventory_Audit_Timesheet
call :WRITE vidapay-ordering VidaPay_Device_Ordering_FULL VidaPay_Device_Ordering_FULL
call :WRITE vidapay-ordering VidaPay_Device_Ordering_FULL_MULTIBROWSER VidaPay_Device_Ordering_FULL_MULTIBROWSER
call :WRITE vidapay-ordering VidaPay_Device_Ordering_FULL_NOCLONE VidaPay_Device_Ordering_FULL_NOCLONE
call :WRITE vidapay-ordering VidaPay_Device_Ordering_FULL_NOCLONE_MULTIBROWSER VidaPay_Device_Ordering_FULL_NOCLONE_MULTIBROWSER
call :WRITE vidapay-ordering VidaPay_Device_Ordering_TRIAL VidaPay_Device_Ordering_TRIAL
call :WRITE vidapay-ordering VidaPay_Device_Ordering_TRIAL_1YEAR VidaPay_Device_Ordering_TRIAL_1YEAR
call :WRITE vidapay-ordering VidaPay_Device_Ordering_TRIAL_1Y_MULTIBROWSER VidaPay_Device_Ordering_TRIAL_1Y_MULTIBROWSER
call :WRITE vidapay-ordering VidaPay_Device_Ordering_TRIAL_1Y_NOCLONE VidaPay_Device_Ordering_TRIAL_1Y_NOCLONE
call :WRITE vidapay-ordering VidaPay_Device_Ordering_TRIAL_MULTIBROWSER VidaPay_Device_Ordering_TRIAL_MULTIBROWSER
call :WRITE vidapay-ordering VidaPay_Device_Ordering_TRIAL_NOCLONE VidaPay_Device_Ordering_TRIAL_NOCLONE
call :WRITE vidapay-transfer-bot VidaPay_Transfer_Bot VidaPay_Transfer_Bot

REM ── Verge repos ──
call :WRITE verge-xls-to-xlsx verge_xls_to_xlsx verge_xls_to_xlsx
call :WRITE verge-rebate-tools verge_rebate_tools verge_rebate_tools
call :WRITE verge-ups-tracking-checker verge_ups_tracking_checker verge_ups_tracking_checker

echo.
echo  ============================================================
echo   Generated %COUNT% build scripts
echo  ============================================================
echo   Location: %OUTDIR%
echo.
echo   Double-click any build_*.bat to build that app.
echo  ============================================================
echo.
pause
endlocal
exit /b 0

:WRITE
set "REPO=%~1"
set "SPEC=%~2"
set "EXE=%~3"
set /a COUNT+=1
set "BATPATH=%OUTDIR%\build_%EXE%.bat"

(
echo @echo off
echo setlocal enabledelayedexpansion
echo title Build %EXE%
echo.
echo set "SRCDIR=%BASEDIR%\%REPO%"
echo set "OUTDIR=%BASEDIR%"
echo set "WORKBASE=%%TEMP%%\pyi_build\%EXE%"
echo.
echo echo.
echo echo  ============================================================
echo echo   Building: %EXE%.exe
echo echo  ============================================================
echo echo.
echo.
echo REM Check prerequisites
echo python --version ^>nul 2^>^&1
echo if errorlevel 1 ^(
echo     echo    ERROR: Python not found in PATH.
echo     pause
echo     exit /b 1
echo ^)
echo python -m PyInstaller --version ^>nul 2^>^&1
echo if errorlevel 1 ^(
echo     echo    PyInstaller not found. Installing...
echo     python -m pip install --upgrade pyinstaller
echo ^)
echo git --version ^>nul 2^>^&1
echo if errorlevel 1 ^(
echo     echo    ERROR: Git not found in PATH.
echo     pause
echo     exit /b 1
echo ^)
echo echo    Prerequisites OK
echo echo.
echo.
echo REM Clone or pull
echo if exist "%%SRCDIR%%" ^(
echo     echo    Pulling latest...
echo     cd "%%SRCDIR%%"
echo     git pull 2^>^&1
echo ^) else ^(
echo     echo    Cloning %REPO%...
echo     git clone "https://github.com/abaduchanna/%REPO%.git" "%%SRCDIR%%" 2^>^&1
echo ^)
echo.
echo cd "%%SRCDIR%%"
echo.
echo REM Clean previous build
echo echo    Cleaning previous build...
echo if exist "build" rmdir /s /q "build"
echo if exist "dist" rmdir /s /q "dist"
echo if exist "__pycache__" rmdir /s /q "__pycache__"
echo del /s /q *.pyc 2^>nul
echo.
echo REM Redirect workpath to TEMP
echo if exist "%%WORKBASE%%" rmdir /s /q "%%WORKBASE%%"
echo mkdir "%%WORKBASE%%" 2^>nul
echo.
echo REM Install deps
echo if exist "requirements.txt" ^(
echo     echo    Installing requirements...
echo     python -m pip install -r requirements.txt --quiet 2^>nul
echo ^)
echo.
echo REM Build
echo echo    Building %SPEC%.spec...
echo python -m PyInstaller "%SPEC%.spec" --noconfirm --clean --workpath "%%WORKBASE%%" 2^>^&1
echo.
echo if errorlevel 1 ^(
echo     echo    FAILED: %EXE%
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo    SUCCESS: %EXE%
echo.
echo REM Copy .exe to output
echo if exist "dist\%EXE%.exe" ^(
echo     if not exist "%%OUTDIR%%" mkdir "%%OUTDIR%%"
echo     copy /Y "dist\%EXE%.exe" "%%OUTDIR%%\%EXE%.exe" ^>nul
echo     echo    Collected: %%OUTDIR%%\%EXE%.exe
echo ^) else ^(
echo     echo    WARNING: dist\%EXE%.exe not found
echo ^)
echo.
echo echo.
echo echo  ============================================================
echo echo   Done: %EXE%.exe
echo echo  ============================================================
echo echo.
echo pause
echo endlocal
echo exit /b 0
) > "%BATPATH%"

echo   %COUNT%. build_%EXE%.bat  (%REPO%)
goto :EOF
