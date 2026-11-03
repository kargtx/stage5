
 = Get-NetFirewallRule -Name 'STAU_LR5*'
foreach ( in ) {
     =  | Get-NetFirewallAddressFilter
    Write-Output (.Name + ' : ' + (.RemoteAddress -join ','))
}
