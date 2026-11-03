$rules = Get-NetFirewallRule -Direction Inbound -Enabled True | Where-Object { $_.Name -like "*ICMP*" -or $_.DisplayName -like "*ICMP*" -or $_.DisplayName -like "*эхо*" }
$rules | Select-Object Name, DisplayName, Action, @{N='Remote';E={($_ | Get-NetFirewallAddressFilter).RemoteAddress}} | Format-Table -AutoSize
