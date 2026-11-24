import subprocess

ps = """
Write-Output "--- CHECK MPIUSER ---"
Get-LocalUser -Name "mpiuser" -ErrorAction SilentlyContinue | Select-Object Name, Enabled, PasswordRequired

Write-Output "--- CHECK SMPD SERVICE ---"
Get-Service *smpd* -ErrorAction SilentlyContinue | Select-Object Name, Status, StartType

Write-Output "--- CHECK FIREWALL RULES ---"
netsh advfirewall firewall show rule name="MPI_SMPD_TCP"
"""

res = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
print(res.stdout)
if res.stderr:
    print(res.stderr)
