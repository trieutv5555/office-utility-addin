@echo off
setlocal
powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $p=Join-Path $env:LOCALAPPDATA 'OfficeUtilityAddin'; New-Item -ItemType Directory -Force -Path $p | Out-Null; $f=Join-Path $p 'manifest.xml'; Invoke-WebRequest -UseBasicParsing 'https://trieutv5555.github.io/office-utility-addin/manifest.xml' -OutFile $f; [xml]$m=Get-Content -Raw $f; if($m.OfficeApp.Id -ne 'e3db0f53-a635-4e76-89c0-750f7713b103'){throw 'Invalid manifest'}; Write-Host ('Manifest downloaded: '+$f)"
if errorlevel 1 goto fail
 echo Share the folder shown above with your Windows account, then add its UNC path in Office Trust Center - Trusted Add-in Catalogs. Enable Show in Menu, restart Office, and choose Shared Folder in Add-ins.
 pause
 exit /b 0
:fail
 echo Download failed. Check GitHub Pages deployment.
 pause
 exit /b 1
