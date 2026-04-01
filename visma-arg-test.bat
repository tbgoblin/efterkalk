@echo off
setlocal

set "OUT=%~dp0visma-arg-test-output.txt"

(
    echo Timestamp: %date% %time%
    echo Script: %~f0
    echo Current dir: %cd%
    echo Raw args: %*
    echo Arg1: %1
    echo Arg2: %2
    echo Arg3: %3
) > "%OUT%"

start "" notepad "%OUT%"

endlocal