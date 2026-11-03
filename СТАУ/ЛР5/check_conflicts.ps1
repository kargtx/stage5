$rules = Get-NetFirewallRule -Enabled True -Direction Inbound -Action Allow | Where-Object { $_.Name -notlike 'STAU_LR5*' }
$found = @()
foreach ($r in $rules) {
    $pf = $r | Get-NetFirewallPortFilter
    $af = $r | Get-NetFirewallAddressFilter
    $ports = ($pf.LocalPort -join ',')
    $proto = $pf.Protocol
    $icmp = $pf.IcmpType
    if ($ports -match '3389' -or $ports -match '445' -or $ports -match '139' -or $ports -match '135' -or $icmp -eq 8) {
        $found += [PSCustomObject]@{
            Name = $r.Name
            DisplayName = $r.DisplayName
            Remote = ($af.RemoteAddress -join ', ')
            Protocol = $proto
            Port = $ports
        }
    }
}
Write-Host "Found rules count: $($found.Count)"
$found | Select-Object -First 15 | Format-Table -AutoSize
