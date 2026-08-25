@echo off
setlocal enabledelayedexpansion
title Download All Repos

set "BASEDIR=%USERPROFILE%\Downloads\GitHub\_src"
if not exist "%BASEDIR%" mkdir "%BASEDIR%"

echo.
echo  ============================================================
echo   Downloading All Repositories
echo  ============================================================
echo.
echo    Target: %BASEDIR%
echo.

REM ── List of repos to download (excluding pos, reportify, readme) ──
REM  GFH repos
set REPOS_GFH=gfh-accessories-order-history-scraper gfh-inventory-aging-processor gfh-rebate-tools gfh-ups-tracking-checker gfh-xls-to-xlsx

REM  VidaPay repos
set REPOS_VIDAPAY=vidapay-extractor vidapay-gfh vidapay-ordering vidapay-transfer-bot

REM  Verge repos
set REPOS_VERGE=verge-xls-to-xlsx verge-rebate-tools verge-ups-tracking-checker

set SUCCESS=0
set FAIL=0

echo  ── GFH Repos ──────────────────────────────────────────────
for %%R in (%REPOS_GFH%) do (
    echo.
    echo    Cloning: %%R
    if exist "%BASEDIR%\%%R" (
        echo      Already exists — pulling latest...
        cd "%BASEDIR%\%%R"
        git pull 2>&1
    ) else (
        git clone "https://github.com/abaduchanna/%%R.git" "%BASEDIR%\%%R" 2>&1
    )
    if errorlevel 1 (
        echo      FAILED: %%R
        set /a FAIL+=1
    ) else (
        echo      OK: %%R
        set /a SUCCESS+=1
    )
)

echo.
echo  ── VidaPay Repos ──────────────────────────────────────────
for %%R in (%REPOS_VIDAPAY%) do (
    echo.
    echo    Cloning: %%R
    if exist "%BASEDIR%\%%R" (
        echo      Already exists — pulling latest...
        cd "%BASEDIR%\%%R"
        git pull 2>&1
    ) else (
        git clone "https://github.com/abaduchanna/%%R.git" "%BASEDIR%\%%R" 2>&1
    )
    if errorlevel 1 (
        echo      FAILED: %%R
        set /a FAIL+=1
    ) else (
        echo      OK: %%R
        set /a SUCCESS+=1
    )
)

echo.
echo  ── Verge Repos ────────────────────────────────────────────
for %%R in (%REPOS_VERGE%) do (
    echo.
    echo    Cloning: %%R
    if exist "%BASEDIR%\%%R" (
        echo      Already exists — pulling latest...
        cd "%BASEDIR%\%%R"
        git pull 2>&1
    ) else (
        git clone "https://github.com/abaduchanna/%%R.git" "%BASEDIR%\%%R" 2>&1
    )
    if errorlevel 1 (
        echo      FAILED: %%R
        set /a FAIL+=1
    ) else (
        echo      OK: %%R
        set /a SUCCESS+=1
    )
)

echo.
echo  ============================================================
echo   Download Summary
echo  ============================================================
echo    Successful: !SUCCESS!
echo    Failed:     !FAIL!
echo    Location:   %BASEDIR%
echo  ============================================================
echo.

if %FAIL% gtr 0 (
    echo  Some repos failed to download. Check the output above.
) else (
    echo  All repos downloaded successfully!
)

echo.
echo  Repos in %BASEDIR%:
dir /b /ad "%BASEDIR%" 2>nul

echo.
pause
endlocal
exit /b 0
