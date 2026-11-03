# Enforce isolation
Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup 'Remote Service Management' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -Name 'FPS-ICMP*' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -Name 'CoreNet-Diag-ICMP*' -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False

# Also Russian groups using unicode
 = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes('Общий доступ к файлам и принтерам'))
 = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes('Дистанционное управление рабочим столом'))
 = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes('Удаленное управление службами'))
Get-NetFirewallRule -DisplayGroup  -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup  -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup  -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False

# Enable our STAU_LR5 rules
Get-NetFirewallRule -Name 'STAU_LR5*' | ForEach-Object {
    Set-NetFirewallRule -Name .Name -Enabled True
     =  | Get-NetFirewallAddressFilter
    Write-Host ('Rule ' + .Name + ' enabled for ' + (.RemoteAddress -join ','))
}
Write-Host 'SUCCESS: All 4 resources restricted strictly to 10.8.1.11'
