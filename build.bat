@echo off
setlocal

:: Clean old build
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

:: Build
echo Building...
pyinstaller build.spec
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

:: Copy user data files next to exe
set DIST=dist\pixel-pilot

if exist steps.json   copy /y steps.json   "%DIST%\"
if exist groups.json  copy /y groups.json  "%DIST%\"
if exist config.json  copy /y config.json  "%DIST%\"
if exist template     xcopy /e /y /i /q template "%DIST%\template\"

echo.
echo Done! Output: %DIST%\pixel-pilot.exe
echo Copy the entire dist\pixel-pilot\ folder to distribute.
echo.
pause
