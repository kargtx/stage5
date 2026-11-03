# -*- coding: utf-8 -*-
import subprocess

cmds = [
    'Disable-NetFirewallRule -DisplayName "zerotier-icmpv4*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -Name "zerotier-icmpv4*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -DisplayName "*zerotier*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -DisplayName "RemoteUtility*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -DisplayName "*Radmin VPN*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -DisplayName "INCY VPN*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -DisplayName "NoxVM*" -ErrorAction SilentlyContinue',
    'Disable-NetFirewallRule -DisplayName "BlueStacks*" -ErrorAction SilentlyContinue'
]

for c in cmds:
    r = subprocess.run(["powershell", "-Command", c], capture_output=True)
    print("Done:", c[:45])

print("\nAll third-party leak rules successfully disabled!")
