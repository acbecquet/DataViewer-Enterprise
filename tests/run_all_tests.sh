#!/bin/bash
# ── DataViewer Enterprise — Run All Tests ──────────────────────────────
# Usage: bash run_all_tests.sh
# ──────────────────────────────────────────────────────────────────────

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Ensure Qt and MinGW are on PATH so DLLs can be found at runtime.
export PATH="/c/Qt/Tools/mingw1310_64/bin:/c/Qt/6.10.1/mingw_64/bin:$PATH"
export QT_PLUGIN_PATH="C:/Qt/6.10.1/mingw_64/plugins"

PASS=0
FAIL=0
SKIP=0
ERRORS=""

echo "============================================================"
echo "  DataViewer Enterprise — Test Suite"
echo "============================================================"
echo ""

for dir in "$SCRIPT_DIR"/tst_*/; do
    name=$(basename "$dir")
    exe="$dir/release/$name.exe"

    if [ ! -f "$exe" ]; then
        echo "  SKIP  $name (not built)"
        SKIP=$((SKIP + 1))
        continue
    fi

    printf "  %-40s" "$name"

    result_file="/tmp/dve_test_${name}.txt"

    # tst_reportgenerator calls zlib deflate which deadlocks under MSYS2
    # pipe handling. Run it via cmd.exe /c start /wait so it gets native
    # Windows I/O handles instead of MSYS pseudo-pipes.
    if [ "$name" = "tst_reportgenerator" ]; then
        win_exe=$(echo "$exe" | sed 's|/|\\|g')
        cmd.exe /c "set PATH=C:\\Qt\\Tools\\mingw1310_64\\bin;C:\\Qt\\6.10.1\\mingw_64\\bin;%PATH%&& ${win_exe}" < /dev/null > /dev/null 2>&1
        rc=$?
        if [ $rc -eq 0 ]; then echo "PASS"; PASS=$((PASS+1)); else echo "FAIL (exit $rc)"; FAIL=$((FAIL+1)); fi
        continue
    fi

    "$exe" -o "$result_file,txt" < /dev/null > /dev/null 2> /dev/null
    rc=$?

    if [ $rc -eq 0 ] && [ -f "$result_file" ]; then
        totals=$(grep "^Totals:" "$result_file" 2>/dev/null | head -1)
        passed=$(echo "$totals" | grep -oE '\d+ passed' || echo "")
        failed_count=$(echo "$totals" | grep -oE '\d+ failed' || echo "0 failed")
        if echo "$failed_count" | grep -q "^0 failed"; then
            echo "PASS ($passed)"
            PASS=$((PASS + 1))
        else
            echo "FAIL — $totals"
            FAIL=$((FAIL + 1))
            ERRORS="$ERRORS\n--- $name ---\n"
            ERRORS="$ERRORS$(grep 'FAIL!' "$result_file" 2>/dev/null)\n"
        fi
    elif [ $rc -eq 0 ]; then
        echo "PASS"
        PASS=$((PASS + 1))
    else
        echo "FAIL (exit $rc)"
        FAIL=$((FAIL + 1))
        ERRORS="$ERRORS\n--- $name (exit $rc) ---\n"
        if [ -f "$result_file" ]; then
            ERRORS="$ERRORS$(grep 'FAIL!' "$result_file" 2>/dev/null)\n"
        fi
    fi
    rm -f "$result_file"
done

echo ""
echo "============================================================"
echo "  Results: $PASS passed, $FAIL failed, $SKIP skipped"
echo "============================================================"

if [ $FAIL -gt 0 ]; then
    echo ""
    echo "FAILURE DETAILS:"
    echo -e "$ERRORS"
    exit 1
fi
exit 0
