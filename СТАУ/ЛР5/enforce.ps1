# 1. Отключаем встроенные общие правила
Disable-NetFirewallRule -DisplayGroup 'Общий доступ к файлам и принтерам' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'File and Printer Sharing' -ErrorAction SilentlyContinue

Disable-NetFirewallRule -DisplayGroup 'Дистанционное управление рабочим столом' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue

Disable-NetFirewallRule -DisplayGroup 'Удаленное управление службами' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayGroup 'Remote Service Management' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-TCP' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-UDP' -ErrorAction SilentlyContinue

Disable-NetFirewallRule -DisplayName '*эхо-запрос*' -ErrorAction SilentlyContinue
Disable-NetFirewallRule -DisplayName '*Echo Request*' -ErrorAction SilentlyContinue

# 2. Включаем и жестко привязываем правила STAU_LR5 к 10.8.1.11
$rules = Get-NetFirewallRule -Name 'STAU_LR5*'
foreach ($r in $rules) {
    Enable-NetFirewallRule -Name $r.Name
    Set-NetFirewallRule -Name $r.Name -RemoteAddress '10.8.1.11'
    Write-Output "Rule $($r.Name) enabled for 10.8.1.11"
}
