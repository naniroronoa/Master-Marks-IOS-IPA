@echo off
setlocal enabledelayedexpansion

:: === Config ===
set GIT="C:\Program Files\Git\bin\git.exe"
set GH="C:\Program Files\GitHub CLI\gh.exe"
:: !IMPORTANT! Change this to your GitHub repository (e.g., username/repo-name)
set REPO=naniroronoa/Master-Marks-IOS-IPA
set BRANCH=master
set OUTPUT_DIR=%~dp0IOS
set APP_NAME=Master_Marks

echo ===========================================
echo    %APP_NAME% - iOS One-Click Builder
echo ===========================================
echo.

:: === Step 0: Increment Version ===
echo [0/6] Incrementing version number...
call npm version patch --no-git-tag-version
for /f "usebackq tokens=*" %%a in (`powershell -Command "(Get-Content package.json | ConvertFrom-Json).version"`) do set VERSION=%%a
set VERSION=!VERSION: ^=!
echo [OK] New version: v!VERSION!
echo.

:: === Step 1: Sync Web Assets ===
echo [1/6] Syncing web assets...
if not exist "www" mkdir "www"
xcopy /E /I /Y assets www\assets > nul 2>&1
copy /Y index.html www\index.html > nul 2>&1
copy /Y script.js www\script.js > nul 2>&1
copy /Y style.css www\style.css > nul 2>&1
copy /Y activation-system.js www\activation-system.js > nul 2>&1
echo [OK] Web assets ready.
echo.

:: === Step 2: Capacitor Sync ===
echo [2/6] Synchronizing Capacitor iOS project...
call npx cap sync ios
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Capacitor sync failed.
    pause & exit /b 1
)
echo [OK] Capacitor sync done.
echo.

:: === Step 3: Git Push to GitHub ===
echo [3/6] Pushing to GitHub to trigger the build...
%GIT% add .
%GIT% commit -m "Build %APP_NAME% v!VERSION!" --allow-empty
%GIT% push origin %BRANCH% --force
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Push to GitHub failed. 
    echo Make sure you have linked this folder to a GitHub repository.
    pause & exit /b 1
)
echo [OK] Code pushed to GitHub. Build triggered!
echo.

:: === Step 4: Wait for GitHub Actions Build ===
echo [4/6] Waiting for GitHub Actions to finish building...
echo       (This takes about 2-3 minutes...)
echo.

timeout /t 20 /nobreak > nul

:CHECK_BUILD
%GH% run list --repo %REPO% --limit 1 --json status,conclusion --jq ".[0].status" > %TEMP%\gh_status.txt 2>&1
set /p BUILD_STATUS=< %TEMP%\gh_status.txt

if "%BUILD_STATUS%"=="completed" goto BUILD_DONE
echo   Current status: !BUILD_STATUS!
timeout /t 20 /nobreak > nul
goto CHECK_BUILD

:BUILD_DONE
%GH% run list --repo %REPO% --limit 1 --json conclusion --jq ".[0].conclusion" > %TEMP%\gh_conclusion.txt 2>&1
set /p BUILD_RESULT=< %TEMP%\gh_conclusion.txt

if "%BUILD_RESULT%"=="success" (
    echo [OK] Build succeeded!
) else (
    echo [ERROR] Build failed. Check: https://github.com/%REPO%/actions
    pause & exit /b 1
)
echo.

:: === Step 5: Download IPA ===
echo [5/6] Finalizing: Downloading IPA to IOS folder...
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

%GH% run list --repo %REPO% --limit 1 --json databaseId --jq ".[0].databaseId" > %TEMP%\gh_run_id.txt 2>&1
set /p RUN_ID=< %TEMP%\gh_run_id.txt

set TEMP_DOWNLOAD=%TEMP%\ipa_download_%RUN_ID%
if exist "%TEMP_DOWNLOAD%" rmdir /s /q "%TEMP_DOWNLOAD%"
mkdir "%TEMP_DOWNLOAD%"

%GH% run download %RUN_ID% --repo %REPO% --name ios-app-ipa --dir "%TEMP_DOWNLOAD%"

:: Rename and Move
set FINAL_NAME=%APP_NAME%_v!VERSION!.ipa
for /r "%TEMP_DOWNLOAD%" %%f in (*.ipa) do (
    copy "%%f" "%OUTPUT_DIR%\!FINAL_NAME!" > nul
)
rmdir /s /q "%TEMP_DOWNLOAD%"

echo [OK] Saved as: %OUTPUT_DIR%\!FINAL_NAME!
echo.
echo ===========================================
echo   SUCCESS! v!VERSION! is ready in IOS folder.
echo ===========================================
echo.
start "" "%OUTPUT_DIR%"
pause
