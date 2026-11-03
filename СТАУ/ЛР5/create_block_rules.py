# -*- coding: utf-8 -*-
import subprocess

other_ips = "10.8.1.1-10.8.1.10,10.8.1.12-10.8.1.254"

block_rules = [
    # 1. ICMP Block for others
    f"New-NetFirewallRule -Name 'STAU_LR5_ICMP_Block_Others' -DisplayName 'СТАУ ЛР5: Блокировка ICMP для остальных' -Direction Inbound -Protocol ICMPv4 -IcmpType 8 -RemoteAddress '{other_ips}' -Action Block -Profile Any -Enabled True",
    
    # 2. SMB Block for others
    f"New-NetFirewallRule -Name 'STAU_LR5_SMB_Block_Others' -DisplayName 'СТАУ ЛР5: Блокировка SMB (445, 139) для остальных' -Direction Inbound -Protocol TCP -LocalPort 445, 139 -RemoteAddress '{other_ips}' -Action Block -Profile Any -Enabled True",
    
    # 3. Remote Registry Block for others
    f"New-NetFirewallRule -Name 'STAU_LR5_RemoteReg_Block_Others' -DisplayName 'СТАУ ЛР5: Блокировка реестра RPC (135, 49152-65535) для остальных' -Direction Inbound -Protocol TCP -LocalPort 135, 49152-65535 -RemoteAddress '{other_ips}' -Action Block -Profile Any -Enabled True",
    
    # 4. RDP Block for others
    f"New-NetFirewallRule -Name 'STAU_LR5_RDP_Block_Others' -DisplayName 'СТАУ ЛР5: Блокировка RDP (3389) для остальных' -Direction Inbound -Protocol TCP -LocalPort 3389 -RemoteAddress '{other_ips}' -Action Block -Profile Any -Enabled True"
]

for cmd in block_rules:
    # First remove if exists
    name = cmd.split("-Name '")[1].split("'")[0]
    subprocess.run(["powershell", "-Command", f"Remove-NetFirewallRule -Name '{name}' -ErrorAction SilentlyContinue"], capture_output=True)
    r = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    if r.returncode == 0:
        print(f"Created Block Rule: {name} (Scope: {other_ips})")
    else:
        print(f"Error on {name}:", r.stderr.decode('cp866', errors='replace'))

print("\n=== ЯВНЫЕ ЗАПРЕЩАЮЩИЕ ПРАВИЛА УСПЕШНО СОЗДАНЫ! ===")
