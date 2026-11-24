@echo off
cd /d "%~dp0"
title ASUSN Lab 6 - Stop System

echo ===========================================================================
echo   Stopping all Lab 6 processes (Plant, PLC1, PLC2, SCADA)...
echo ===========================================================================

taskkill /FI "WINDOWTITLE eq *Plant Simulator*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq *PLC 1*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq *PLC 2*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq *SCADA*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq [1/4]*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq [2/4]*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq [3/4]*" /T /F > nul 2>&1
taskkill /FI "WINDOWTITLE eq [4/4]*" /T /F > nul 2>&1

powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match 'plant_simulator|plc1\.py|plc2\.py|scada\.py|run_all\.py' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" > nul 2>&1

echo.
echo All Lab 6 components have been stopped.
echo ===========================================================================
pause
