#!/usr/bin/env bash
# Wrapper to run PowerShell scripts from MSYS2
# Converts path arguments to Windows style and invokes powershell.exe

if [ $# -lt 1 ]; then
    echo "Usage: $(basename "$0") <script.ps1> [args...]" >&2
    exit 1
fi

SCRIPT=$1
shift

if command -v cygpath >/dev/null 2>&1; then
    SCRIPT=$(cygpath -w "$SCRIPT")
    args=()
    for a in "$@"; do
        args+=("$(cygpath -w "$a")")
    done
else
    args=("$@")
fi

cmd=(powershell.exe -ExecutionPolicy Bypass -File "$SCRIPT" "${args[@]}")

if command -v winpty >/dev/null 2>&1; then
    cmd=(winpty "${cmd[@]}")
fi

exec "${cmd[@]}"
