import subprocess

ps_script = """
Get-NetConnectionProfile | Select-Object Name, InterfaceAlias, NetworkCategory
Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -notlike '169.254.*'
} | Select-Object InterfaceAlias, IPAddress, PrefixLength
"""

res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True, text=True)
print("STDOUT:")
print(res.stdout)
if res.stderr:
    print("STDERR:")
    print(res.stderr)
