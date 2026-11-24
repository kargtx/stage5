# Инструкция для AI-агента: подготовка ПК Данила (Узел 2) к кластерному запуску MPI (ЛР №3)

> **Задача агента:** автоматически подготовить этот компьютер (Узел 2) для распределённых вычислений MPICH2 v1.4.1p1 по локальной сети или VPN в паре с главным компьютером (`kargtx`, пользователь `mpiuser`).
>
> **Правила безопасности и точности:**
> - Не выдумывай результаты команд — проверяй статус каждого шага по реальному выводу.
> - Если терминал запущен **без прав администратора**, не пытайся обходить UAC; сразу остановись и попроси Данила подтвердить/запустить команды от имени Администратора.
> - После завершения настройки не меняй лишние системные настройки и обязательно выведи Данилу его активный **IPv4-адрес**, чтобы он отправил его Кариму.

---

## Шаг 1. Проверка прав администратора и поиск архива `кластер_сеть.zip`

1. Проверь права администратора в PowerShell:
   ```powershell
   $id = [Security.Principal.WindowsIdentity]::GetCurrent()
   $p = New-Object Security.Principal.WindowsPrincipal($id)
   $isAdmin = $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
   Write-Output "IsAdmin=$isAdmin"
   ```
   - Если `IsAdmin=False` — попроси Данила перезапустить среду/терминал от имени Администратора (или выполнять административные команды через окно с правами администратора).

2. Найди файл `кластер_сеть.zip` (проверь папку рядом с этим `.md` файлом, `Desktop`, `Downloads`).
   - Если архив не найден, попроси Данила положить `кластер_сеть.zip` в `Downloads` или указать путь к нему.

3. Распакуй архив строго в корень `C:\cluster` **без вложенной папки `кластер_сеть`** (чтобы `C:\cluster\smpd.exe` и `C:\cluster\mpiexec.exe` лежали прямо в `C:\cluster\`):
   ```powershell
   $zipPath = "<ПУТЬ_К_кластер_сеть.zip>"
   $tmpDir = "$env:TEMP\cluster_unpack"
   Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
   Expand-Archive -Path $zipPath -DestinationPath $tmpDir -Force
   New-Item -ItemType Directory -Path "C:\cluster" -Force | Out-Null
   if (Test-Path "$tmpDir\кластер_сеть\smpd.exe") {
       Copy-Item "$tmpDir\кластер_сеть\*" "C:\cluster\" -Recurse -Force
   } else {
       Copy-Item "$tmpDir\*" "C:\cluster\" -Recurse -Force
   }
   Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
   Get-ChildItem "C:\cluster\*.exe" | Select-Object Name
   ```
   - Убедись, что в `C:\cluster` присутствуют `smpd.exe`, `mpiexec.exe`, `mpich2.dll`, `mpich2mpi.dll`, `mpich2nemesis.dll`, `mpi1-hello.exe`, `mpi7-integral.exe`.

---

## Шаг 2. Создание локального пользователя `mpiuser`

Служба MPICH2 `smpd` требует на обоих компьютерах локального пользователя с **одинаковым логином и непустым паролем** (учётная запись Microsoft / вход по PIN не подходят).

Выполни в PowerShell (от Администратора):
```powershell
$u = Get-LocalUser -Name "mpiuser" -ErrorAction SilentlyContinue
if (-not $u) {
    net user mpiuser 123456 /add
}
$members = Get-LocalGroupMember -SID "S-1-5-32-544" | Where-Object { $_.Name -like "*\mpiuser" }
if (-not $members) {
    Add-LocalGroupMember -SID "S-1-5-32-544" -Member "mpiuser"
}
Get-LocalUser -Name "mpiuser" | Select-Object Name, Enabled, PasswordRequired
```
- Проверь, что `Enabled = True` и `mpiuser` состоит в группе администраторов (`S-1-5-32-544`).

---

## Шаг 3. Настройка Брандмауэра Windows (Firewall)

Для работы кластера нужно разрешить входящие `ping`, порты `8676`/`8677` для `smpd`, порт `445` (SMB, если используется сетевая папка) и входящие соединения для всех `mpi*.exe` в `C:\cluster` (MPI-процессы общаются напрямую через динамические TCP-порты).

Выполни в PowerShell (идемпотентно — старые одноимённые правила предварительно удаляются):
```powershell
$rules = @(
    @{ Name="MPI_Ping";        Args=@("protocol=icmpv4:8,any", "dir=in", "action=allow", "profile=any") },
    @{ Name="MPI_SMPD_TCP";    Args=@("dir=in", "action=allow", "protocol=TCP", "localport=8676,8677", "profile=any") },
    @{ Name="MPI_SMB_445";     Args=@("dir=in", "action=allow", "protocol=TCP", "localport=445", "profile=any") },
    @{ Name="MPI_SMPD_App";    Args=@("dir=in", "action=allow", "program=C:\cluster\smpd.exe", "profile=any") },
    @{ Name="MPI_Exec_App";    Args=@("dir=in", "action=allow", "program=C:\cluster\mpiexec.exe", "profile=any") },
    @{ Name="MPI_Hello";       Args=@("dir=in", "action=allow", "program=C:\cluster\mpi1-hello.exe", "profile=any") },
    @{ Name="MPI_Sum";         Args=@("dir=in", "action=allow", "program=C:\cluster\mpi4-sum.exe", "profile=any") },
    @{ Name="MPI_Reduce";      Args=@("dir=in", "action=allow", "program=C:\cluster\mpi5-reduce.exe", "profile=any") },
    @{ Name="MPI_Clock";       Args=@("dir=in", "action=allow", "program=C:\cluster\mpi6-clock.exe", "profile=any") },
    @{ Name="MPI_Integral";    Args=@("dir=in", "action=allow", "program=C:\cluster\mpi7-integral.exe", "profile=any") },
    @{ Name="MPI_Integral_D";  Args=@("dir=in", "action=allow", "program=C:\cluster\mpi7-integral-double.exe", "profile=any") },
    @{ Name="MPI_Integral_F";  Args=@("dir=in", "action=allow", "program=C:\cluster\mpi7-integral-float.exe", "profile=any") },
    @{ Name="MPI_Affinity";    Args=@("dir=in", "action=allow", "program=C:\cluster\mpi8-affinity.exe", "profile=any") }
)

foreach ($r in $rules) {
    netsh advfirewall firewall delete rule name="$($r.Name)" > $null 2>&1
    & netsh advfirewall firewall add rule name="$($r.Name)" @($r.Args) | Out-Null
}
netsh advfirewall firewall show rule name="MPI_SMPD_TCP"
```

---

## Шаг 4. Установка и запуск службы `smpd` и проверка учётных данных

1. Установи и запусти службу `smpd`:
   ```cmd
   C:\cluster\smpd.exe -install
   ```
2. Проверь статус службы в Windows (не зависает, в отличие от некоторых флагов в несвоём каталоге):
   ```powershell
   Get-Service *smpd* | Select-Object Name, Status, StartType
   ```
   - Должно быть `Status = Running`. Если служба остановлена — запусти `Start-Service mpich2_smpd` (или проверь точное `Name` из `Get-Service *smpd*`).
3. Попроси Данила один раз выполнить регистрацию учётной записи для локальных проверок (если `C:\cluster\mpiexec.exe -validate` ещё не возвращает `SUCCESS`):
   ```cmd
   C:\cluster\mpiexec.exe -remove
   C:\cluster\mpiexec.exe -register
   ```
   При запросе ввести:
   - `account (domain\user):` `mpiuser`
   - `password:` `123456`
   - `confirm password:` `123456`

   Затем проверить:
   ```cmd
   C:\cluster\mpiexec.exe -validate
   ```
   Ожидаемый ответ: `SUCCESS`.

   > Примечание для агента: чтобы главный компьютер Карима мог запускать процессы на этом ПК, достаточно работающей службы `smpd` и существующего пользователя `mpiuser`/`123456` — но `mpiexec -register` на этом ПК также позволит проверить локальный запуск `C:\cluster\mpiexec.exe -localonly 2 C:\cluster\mpi1-hello.exe`.

---

## Шаг 5. Определение IPv4-адреса и сетевой связности

Выведи активные IPv4-адреса (исключая `127.0.0.1` и `169.254.*`) и текущий профиль сети:
```powershell
Get-NetConnectionProfile | Select-Object Name, InterfaceAlias, NetworkCategory
Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -notlike '169.254.*'
} | Select-Object InterfaceAlias, IPAddress, PrefixLength
```

Подскажи Данилу:
- Если вы с Каримом в одной локальной сети / Wi-Fi / точке доступа — передай Кариму адрес интерфейса **Wi-Fi** или **Ethernet** (обычно `192.168.x.x` или `172.20.x.x`).
- Если вы подключаетесь через VPN-локалку (Radmin VPN / ZeroTier / Tailscale) — передай адрес виртуального адаптера VPN (например, `26.x.x.x`).
- Если `NetworkCategory` стоит `Public` — порекомендуй переключить на `Private` (или оставь правила `profile=any`, которые уже добавлены).

---

## Шаг 6. Итоговый отчёт Данилу (что сказать пользователю)

В финальном ответе кратко сообщи Данилу:
1. Создан ли пользователь `mpiuser` (`123456`) и добавлен ли в Администраторы.
2. Распакованы ли файлы в `C:\cluster\`.
3. Работает ли служба `smpd` (`Running`) и прошёл ли `mpiexec -validate` (или нужно ли ему вручную ввести `mpiuser`/`123456` в `mpiexec -register`).
4. Добавлены ли правила Брандмауэра (`MPI_*`).
5. **Точный IPv4-адрес**, который нужно отправить Кариму, и команду, которую Карим сможет сразу запустить у себя:
   ```cmd
   C:\cluster\mpiexec.exe -hosts 2 <IP_КАРИМА> 2 <IP_ДАНИЛА> 2 C:\cluster\mpi7-integral.exe 100000000
   ```

---

## Приложение: очистка компьютера после завершения лабораторной работы

Когда лабораторная будет сдана, агент или Данил может полностью вернуть всё в исходное состояние одной командой в PowerShell (от Администратора):
```powershell
C:\cluster\smpd.exe -uninstall
net user mpiuser /delete
foreach ($n in @("MPI_Ping","MPI_SMPD_TCP","MPI_SMB_445","MPI_SMPD_App","MPI_Exec_App","MPI_Hello","MPI_Sum","MPI_Reduce","MPI_Clock","MPI_Integral","MPI_Integral_D","MPI_Integral_F","MPI_Affinity")) {
    netsh advfirewall firewall delete rule name="$n" > $null 2>&1
}
Remove-Item "C:\cluster" -Recurse -Force -ErrorAction SilentlyContinue
```
