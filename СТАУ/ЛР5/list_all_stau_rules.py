# -*- coding: utf-8 -*-
import subprocess

ps = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Get-NetFirewallRule | Where-Object { $_.DisplayName -match 'СТАУ' -or $_.Name -match 'STAU' } | ForEach-Object {
    $rule = $_
    $addr = $rule | Get-NetFirewallAddressFilter
    $port = $rule | Get-NetFirewallPortFilter
    [PSCustomObject]@{
        Name = $rule.Name
        DisplayName = $rule.DisplayName
        Action = $rule.Action
        Enabled = $rule.Enabled
        RemoteAddress = ($addr.RemoteAddress -join ', ')
        Protocol = $port.Protocol
        LocalPort = ($port.LocalPort -join ', ')
    }
} | Format-Table -AutoSize
"""

r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True, encoding="utf-8")
print(r.stdout)
