# -*- coding: utf-8 -*-
import subprocess
import sys

script = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 1. Remove any old STAU rules
Get-NetFirewallRule | Where-Object { $_.Name -like 'STAU_LR5*' } | Remove-NetFirewallRule

# 2. Allow rules for Danil (10.8.1.11)
New-NetFirewallRule -Name 'STAU_LR5_ICMP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Разрешить ICMP (ping) для Данила' `
    -Direction Inbound -Action Allow -Protocol ICMPv4 -IcmpType 8 `
    -RemoteAddress '10.8.1.11' -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_SMB_TCP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Разрешить SMB TCP (445, 139) для Данила' `
    -Direction Inbound -Action Allow -Protocol TCP -LocalPort 445, 139 `
    -RemoteAddress '10.8.1.11' -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_SMB_UDP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Разрешить NetBIOS UDP (137, 138) для Данила' `
    -Direction Inbound -Action Allow -Protocol UDP -LocalPort 137, 138 `
    -RemoteAddress '10.8.1.11' -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_RemoteReg_TCP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Разрешить удаленный реестр RPC TCP для Данила' `
    -Direction Inbound -Action Allow -Protocol TCP -LocalPort 135, 445, '49152-65535' `
    -RemoteAddress '10.8.1.11' -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_RDP_TCP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Разрешить RDP TCP (3389) для Данила' `
    -Direction Inbound -Action Allow -Protocol TCP -LocalPort 3389 `
    -RemoteAddress '10.8.1.11' -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_RDP_UDP_Allow_Danil' `
    -DisplayName 'СТАУ ЛР5: Разрешить RDP UDP (3389) для Данила' `
    -Direction Inbound -Action Allow -Protocol UDP -LocalPort 3389 `
    -RemoteAddress '10.8.1.11' -Profile Any -Enabled True

# 3. Explicit Block rules for all other subnet neighbors
$otherSubnets = @('10.8.1.1-10.8.1.10', '10.8.1.12-10.8.1.254')

New-NetFirewallRule -Name 'STAU_LR5_ICMP_Block_Others' `
    -DisplayName 'СТАУ ЛР5: Блокировка ICMP для остальных' `
    -Direction Inbound -Action Block -Protocol ICMPv4 -IcmpType 8 `
    -RemoteAddress $otherSubnets -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_SMB_Block_Others' `
    -DisplayName 'СТАУ ЛР5: Блокировка SMB (445, 139) для остальных' `
    -Direction Inbound -Action Block -Protocol TCP -LocalPort 445, 139 `
    -RemoteAddress $otherSubnets -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_RemoteReg_Block_Others' `
    -DisplayName 'СТАУ ЛР5: Блокировка реестра RPC (135, 49152-65535) для остальных' `
    -Direction Inbound -Action Block -Protocol TCP -LocalPort 135, '49152-65535' `
    -RemoteAddress $otherSubnets -Profile Any -Enabled True

New-NetFirewallRule -Name 'STAU_LR5_RDP_Block_Others' `
    -DisplayName 'СТАУ ЛР5: Блокировка RDP (3389) для остальных' `
    -Direction Inbound -Action Block -Protocol TCP -LocalPort 3389 `
    -RemoteAddress $otherSubnets -Profile Any -Enabled True

# 4. Disable generic built-in rules that might allow traffic from other IPs
Disable-NetFirewallRule -DisplayGroup 'Общий доступ к файлам и принтерам' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'File and Printer Sharing' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'Дистанционное управление рабочим столом' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'Удаленное управление службами' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'Remote Service Management' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayName '*эхо-запрос*' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayName '*Echo Request*' -ErrorAction SilentlyContinue
"""

print("Applying rules...")
res = subprocess.run(["powershell", "-NoProfile", "-Command", script], capture_output=True, text=True, encoding="utf-8")
if res.returncode != 0:
    print("PowerShell Error:", res.stderr)
else:
    print("Success applying rules.")
