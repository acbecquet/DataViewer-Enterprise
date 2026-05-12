@echo off
echo ============================================
echo  DataViewer Enterprise - Build Installer
echo ============================================
echo.

REM Check that the exe exists
if not exist "release\DataViewer.exe" (
    echo ERROR: release\DataViewer.exe not found.
    echo        Build the project first with qmake + mingw32-make.
    pause
    exit /b 1
)

REM Create dist output directory
if not exist "dist" mkdir dist

REM Extract VERSION from DataViewerEnterprise.pro (single source of truth).
for /f "tokens=3" %%v in ('findstr /B /C:"VERSION = " DataViewerEnterprise.pro') do set APP_VERSION=%%v
if "%APP_VERSION%"=="" (
    echo ERROR: could not parse VERSION from DataViewerEnterprise.pro
    pause
    exit /b 1
)
echo Building installer for v%APP_VERSION% ...

echo Copying libpq runtime DLLs to release\...
copy /Y vendor\libpq-16\libpq.dll release\ >nul
copy /Y vendor\libpq-16\libcrypto-3-x64.dll release\ >nul
copy /Y vendor\libpq-16\libssl-3-x64.dll release\ >nul
copy /Y vendor\libpq-16\libintl-9.dll release\ >nul
copy /Y vendor\libpq-16\libiconv-2.dll release\ >nul
if errorlevel 1 (
    echo ERROR: failed to copy libpq DLLs
    exit /b 1
)

"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=%APP_VERSION% installer.iss

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Installer build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  SUCCESS! Installer created at:
echo  dist\DataViewer-setup.exe
echo ============================================
pause
