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

REM Sanity check: confirm the exe's baked-in version matches the .pro VERSION.
REM Catches stale incremental builds where main.o was compiled before VERSION
REM was bumped (qmake doesn't auto-detect VERSION changes and rebuild .o files).
REM See CLAUDE.md "Release Workflow" for the recommended clean-build recipe.
for /f "tokens=3" %%v in ('findstr /B /C:"VERSION = " DataViewerEnterprise.pro') do set PRO_VERSION=%%v
for /f "tokens=*" %%v in ('powershell -NoProfile -Command "(Get-Item 'release\DataViewer.exe').VersionInfo.FileVersion"') do set EXE_VERSION=%%v
echo Checking version: .pro=%PRO_VERSION% exe=%EXE_VERSION%
echo %EXE_VERSION% | findstr /B "%PRO_VERSION%" >nul
if errorlevel 1 (
    echo.
    echo ERROR: release\DataViewer.exe FileVersion ^(%EXE_VERSION%^) does not match
    echo        .pro VERSION ^(%PRO_VERSION%^). The exe is built from stale .o files.
    echo        Fix: cd build ^&^& mingw32-make clean ^&^& mingw32-make -j8
    echo        Then re-run build_installer.bat.
    pause
    exit /b 1
)

REM Staleness check: the installer packages release\DataViewer.exe (the ROOT
REM in-tree build). A rebuild in build\release\ does NOT update it, and the
REM version gate above cannot catch that when VERSION was not bumped. Refuse
REM if any source file under src\ is newer than the exe being packaged.
for /f "tokens=*" %%v in ('powershell -NoProfile -Command "$exe=(Get-Item 'release\DataViewer.exe').LastWriteTime; $newer=Get-ChildItem src -Recurse -Include *.cpp,*.h | Where-Object LastWriteTime -gt $exe | Select-Object -First 1; if ($newer) { $newer.FullName } else { 'OK' }"') do set STALE_SRC=%%v
if not "%STALE_SRC%"=="OK" (
    echo.
    echo ERROR: release\DataViewer.exe is OLDER than source file:
    echo        %STALE_SRC%
    echo        The installer packages the ROOT release\ tree - rebuild it first:
    echo        qmake -spec win32-g++ CONFIG+=release DataViewerEnterprise.pro
    echo        mingw32-make -f Makefile.Release -j8
    echo        Then re-run build_installer.bat.
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

echo Copying QPSQL driver plugin to release\sqldrivers\...
if not exist release\sqldrivers mkdir release\sqldrivers
set "QSQLPSQL_SRC="
for %%v in (6.10.2 6.10.1) do (
    if exist "C:\Qt\%%v\mingw_64\plugins\sqldrivers\qsqlpsql.dll" (
        if not defined QSQLPSQL_SRC set "QSQLPSQL_SRC=C:\Qt\%%v\mingw_64\plugins\sqldrivers\qsqlpsql.dll"
    )
)
if not defined QSQLPSQL_SRC (
    echo ERROR: qsqlpsql.dll not found in C:\Qt\6.10.x\mingw_64\plugins\sqldrivers\
    exit /b 1
)
copy /Y "%QSQLPSQL_SRC%" release\sqldrivers\ >nul
if errorlevel 1 (
    echo ERROR: failed to copy qsqlpsql.dll
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
