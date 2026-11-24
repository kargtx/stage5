@echo off
cd /d "%~dp0"
title ASUSN Lab 6 - Distributed Control System

echo ===========================================================================
echo   ASUSN Lab 6: Distributed Control System Launcher
echo   Authors: Gaynanov D.I., Garifullin K.R. (Group EAS-514S, 2026)
echo ===========================================================================
echo.
echo Launching distributed system via run_all.py ...
echo.

where python >nul 2>nul
if %ERRORLEVEL% equ 0 (
    python run_all.py
) else (
    py run_all.py
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo [FALLBACK] Direct launch of 4 components in individual windows...
    start "[1/4] Plant Simulator" python plant_simulator.py
    timeout /t 2 /nobreak > nul
    start "[2/4] PLC 1 (Modbus 5020)" python plc1.py
    timeout /t 2 /nobreak > nul
    start "[3/4] PLC 2 (Modbus 5021)" python plc2.py
    timeout /t 2 /nobreak > nul
    start "[4/4] SCADA Dispatch Terminal" python scada.py
)

echo.
echo ===========================================================================
pause
