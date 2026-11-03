$rules = Get-NetFirewallRule -Enabled True -Direction Inbound
$icmpRules = @()
foreach ($r in $rules) {
    $pf = $r | Get-NetFirewallPortFilter
    if ($pf.Protocol -eq 'ICMPv4' -or $pf.IcmpType) {
        $af = $r | Get-NetFirewallAddressFilter
        $icmpRules += [PSCustomObject]@{
            Name = $r.Name
            DisplayName = $r.DisplayName
            Action = $r.Action
            RemoteIP = ($af.RemoteAddress -join ', ')
            IcmpType = $pf.IcmpType
            Profile = $r.Profile
        }
    }
}
$icmpRules | Format-Table -AutoSize
