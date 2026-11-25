@echo off
echo Remove the add-in in Excel and Word, then remove its catalog in Trust Center.
if exist "%LOCALAPPDATA%\OfficeUtilityAddin\manifest.xml" del "%LOCALAPPDATA%\OfficeUtilityAddin\manifest.xml"
pause
