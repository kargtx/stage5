@echo off
chcp 65001 >nul
cd /d "%~dp0"
python modbus_plc_simulator.py 502
pause
