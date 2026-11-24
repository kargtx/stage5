$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$p = New-Object Security.Principal.WindowsPrincipal($id)
$isAdmin = $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-Output "IsAdmin=$isAdmin"

# Check source cluster folder
$srcDir = "c:\Users\danch\OneDrive\Рабочий стол\stage5\ВАС\lr3\кластер_сеть"
Write-Output "Checking source directory: $srcDir"
if (Test-Path $srcDir) {
    Get-ChildItem $srcDir | Select-Object Name, Length
} else {
    Write-Output "Source directory not found!"
}

# Check if C:\cluster exists
Write-Output "Checking C:\cluster..."
if (Test-Path "C:\cluster") {
    Get-ChildItem "C:\cluster" | Select-Object Name, Length
} else {
    Write-Output "C:\cluster does not exist yet."
}
