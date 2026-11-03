$rules = Get-NetFirewallRule -Direction Inbound -Enabled True
$echoRules = @()
foreach ($r in $rules) {
    $pf = $r | Get-NetFirewallPortFilter
    if ($pf.Protocol -eq 'ICMPv4' -and $pf.IcmpType -eq 8) {
        $af = $r | Get-NetFirewallAddressFilter
        $echoRules += [PSCustomObject]@{
            Name = $r.Name
            DisplayName = $r.DisplayName
            Action = $r.Action
            Remote = ($af.RemoteAddress -join ', ')
        }
    }
}
$echoRules | Format-Table -AutoSize
