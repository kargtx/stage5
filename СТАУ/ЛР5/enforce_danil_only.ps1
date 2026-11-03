# Скрипт очистки и жесткой привязки правил брандмауэра к 10.8.1.11
# Отключаем любые встроенные общие правила для защищаемых ресурсов,
# чтобы доступ был открыт ИСКЛЮЧИТЕЛЬНО для ноутбука Данила (10.8.1.11).

Write-Host "=== 1. Отключение встроенных общих правил File Sharing (SMB) ==="
Disable-NetFirewallRule -DisplayGroup "Общий доступ к файлам и принтерам" -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction SilentlyContinue

Write-Host "=== 2. Отключение встроенных общих правил Удаленного рабочего стола (RDP) ==="
Disable-NetFirewallRule -DisplayGroup "Дистанционное управление рабочим столом" -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue

Write-Host "=== 3. Отключение встроенных общих правил Удаленного реестра / RPC ==="
Disable-NetFirewallRule -DisplayGroup "Удаленное управление службами" -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup "Remote Service Management" -ErrorAction SilentlyContinue
Disable-NetFirewallRule -Name "WMI-WINMGMT-In-TCP" -ErrorAction SilentlyContinue

Write-Host "=== 4. Отключение фоновых правил ICMP ==="
Disable-NetFirewallRule -DisplayName "*эхо-запрос*" -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayName "*Echo Request*" -ErrorAction SilentlyContinue

Write-Host "=== 5. Гарантированное включение НАШИХ правил для Данила (10.8.1.11) ==="
$ourRules = Get-NetFirewallRule -Name "STAU_LR5*"
foreach ($r in $ourRules) {
    Enable-NetFirewallRule -Name $r.Name
    Set-NetFirewallRule -Name $r.Name -RemoteAddress "10.8.1.11" -Action Allow
    Write-Host "  Включено и привязано к 10.8.1.11: $($r.DisplayName)"
}

Write-Host "=== Готово! Теперь доступ ко всем 4 ресурсам открыт ТОЛЬКО для 10.8.1.11! ==="
