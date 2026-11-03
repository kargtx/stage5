# -*- coding: utf-8 -*-
import subprocess

ps_commands = [
    # 1. Отключаем общие правила File Sharing (SMB)
    "Get-NetFirewallRule -DisplayGroup 'Общий доступ к файлам и принтерам' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    "Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    
    # 2. Отключаем общие правила Remote Desktop (RDP)
    "Get-NetFirewallRule -DisplayGroup 'Дистанционное управление рабочим столом' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    "Get-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    "Set-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-TCP' -Enabled False -ErrorAction SilentlyContinue",
    "Set-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-UDP' -Enabled False -ErrorAction SilentlyContinue",
    
    # 3. Отключаем общие правила Remote Registry / RPC
    "Get-NetFirewallRule -DisplayGroup 'Удаленное управление службами' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    "Get-NetFirewallRule -DisplayGroup 'Remote Service Management' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    
    # 4. Отключаем общие правила ICMP (эхо-запросы)
    "Get-NetFirewallRule -Name 'FPS-ICMP*' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    "Get-NetFirewallRule -Name 'CoreNet-Diag-ICMP*' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False",
    
    # 5. Включаем НАШИ правила STAU_LR5 (они уже привязаны к 10.8.1.11)
    "Enable-NetFirewallRule -Name 'STAU_LR5_ICMP_Allow_Danil'",
    "Enable-NetFirewallRule -Name 'STAU_LR5_SMB_TCP_Allow_Danil'",
    "Enable-NetFirewallRule -Name 'STAU_LR5_SMB_UDP_Allow_Danil'",
    "Enable-NetFirewallRule -Name 'STAU_LR5_RemoteReg_TCP_Allow_Danil'",
    "Enable-NetFirewallRule -Name 'STAU_LR5_RDP_TCP_Allow_Danil'",
    "Enable-NetFirewallRule -Name 'STAU_LR5_RDP_UDP_Allow_Danil'"
]

for cmd in ps_commands:
    r = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    if r.returncode != 0:
        print("ERR on:", cmd, "->", r.stderr.decode('cp866', errors='replace'))
    else:
        print("OK:", cmd[:60] + "...")

print("\n=== ВСЕ ОБЩИЕ ПРАВИЛА ОТКЛЮЧЕНЫ. ДОСТУП ТОЛЬКО ДЛЯ 10.8.1.11! ===")
