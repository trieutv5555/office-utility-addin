@echo off
setlocal
cd /d "%~dp0"
where git >nul 2>&1 || goto fail
if not exist .git (
 git init -b main || goto fail
)
git remote get-url origin >nul 2>&1
if errorlevel 1 (
 git remote add origin https://github.com/trieutv5555/office-utility-addin.git || goto fail
)
rem Set identity only in this repository if missing.
git config --get user.name >nul 2>&1
if errorlevel 1 (
 git config --local user.name "trieutv5555" || goto fail
)
git config --get user.email >nul 2>&1
if errorlevel 1 (
 git config --local user.email "102889723+trieutv5555@users.noreply.github.com" || goto fail
)
rem docs\manifest.xml is the only source manifest.
if not exist "docs\manifest.xml" (
 echo Missing docs\manifest.xml.
 goto fail
)
git add -A || goto fail
git diff --cached --quiet
if errorlevel 2 goto fail
if errorlevel 1 (
 git commit -m "Update Office Utility Add-in" || goto fail
)
git push -u origin main || goto fail
echo Upload successful.
pause
exit /b 0
:fail
echo Upload failed. Check Git installation, credentials and remote branch. Do not force push.
pause
exit /b 1
