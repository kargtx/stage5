# Отключаем общие встроенные правила Windows для 4 ресурсов
Write-Host "Отключение встроенных правил File Sharing..."
Get-NetFirewallRule -DisplayGroup "Общий доступ к файлам и принтерам" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False

Write-Host "Отключение встроенных правил Remote Desktop..."
Get-NetFirewallRule -DisplayGroup "Дистанционное управление рабочим столом" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False

Write-Host "Отключение встроенных правил Remote Service Management / RPC..."
Get-NetFirewallRule -DisplayGroup "Удаленное управление службами" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -DisplayGroup "Remote Service Management" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False

Write-Host "Отключение встроенных правил ICMP Echo..."
Get-NetFirewallRule -Name "FPS-ICMP*" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False
Get-NetFirewallRule -Name "CoreNet-Diag-ICMP*" -ErrorAction SilentlyContinue | Set-NetFirewallRule -Enabled False

Write-Host "Включение наших правил STAU_LR5 строго для 10.8.1.11..."
Get-NetFirewallRule -Name "STAU_LR5*" | ForEach-Object {
    Set-NetFirewallRule -Name $_.Name -Enabled True
    $af = $_ | Get-NetFirewallAddressFilter
    Write-Host "Правило $($_.Name) включено (RemoteAddress: $($af.RemoteAddress))"
}
Write-Host "ГОТОВО: Все 4 ресурса открыты ТОЛЬКО для 10.8.1.11, для остальных закрыты!"
