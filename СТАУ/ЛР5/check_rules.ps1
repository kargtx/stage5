$rules = Get-NetFirewallRule -Name 'STAU_LR5*'
$results = @()
foreach ($r in $rules) {
    $af = $r | Get-NetFirewallAddressFilter
    $pf = $r | Get-NetFirewallPortFilter
    $ic = $r | Get-NetFirewallIcmpFilter -ErrorAction SilentlyContinue
    $p = if ($pf.Protocol) { $pf.Protocol } else { "ICMP" }
    $pt = if ($pf.LocalPort) { ($pf.LocalPort -join ', ') } elseif ($ic.IcmpType) { "Type " + $ic.IcmpType } else { "Any" }
    $obj = [PSCustomObject]@{
        RuleName = $r.Name
        DisplayName = $r.DisplayName
        Action = $r.Action
        RemoteIP = ($af.RemoteAddress -join ', ')
        Protocol = $p
        Port_Type = $pt
        Enabled = $r.Enabled
    }
    $results += $obj
}
$results | Format-Table -AutoSize
