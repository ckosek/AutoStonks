@echo off
echo ================================================
echo         AutoStonks - First Time Setup
echo ================================================
echo.
echo Setting up permissions and scheduling tasks...
echo.

:: Fix execution policy for current user (scoped to this user only)
powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force"

:: Unblock the scripts so Windows allows them to run
powershell -Command "Unblock-File -Path '%~dp0Set-SPXWallpaper.ps1'"
powershell -Command "Unblock-File -Path '%~dp0Setup.ps1'"

:: Run the setup script from the same folder as this bat file
powershell -ExecutionPolicy Unrestricted -File "%~dp0Setup.ps1"

echo.
echo ================================================
echo  All done! Stonks will update at market close and open!
echo ================================================
pause