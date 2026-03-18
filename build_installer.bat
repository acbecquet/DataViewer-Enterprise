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

REM Run Inno Setup compiler
echo Building installer...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Installer build failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  SUCCESS! Installer created at:
echo  dist\DataViewerEnterprise-v0.1.0-alpha-setup.exe
echo ============================================
pause
