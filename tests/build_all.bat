@echo off
set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH%

set PASS=0
set FAIL=0
set FAILED=

for /D %%d in (tst_*) do (
    echo.
    echo === Building %%d ===
    pushd %%d
    if exist Makefile.Release (
        mingw32-make.exe -f Makefile.Release >build.log 2>&1
        if errorlevel 1 (
            echo   FAIL
            set /a FAIL+=1
            set FAILED=!FAILED! %%d
            type build.log
        ) else (
            echo   OK
            set /a PASS+=1
        )
    ) else (
        echo   SKIP - no Makefile.Release
    )
    popd
)

echo.
echo ==========================================
echo Build Results: %PASS% passed, %FAIL% failed
echo ==========================================
if not "%FAILED%"=="" echo Failed: %FAILED%
