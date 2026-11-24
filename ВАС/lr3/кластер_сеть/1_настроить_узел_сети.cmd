@echo off
chcp 65001 >nul
title Настройка узла кластера MPI (Запуск от Администратора)
echo ===============================================================================
echo     НАСТРОЙКА УЗЛА КЛАСТЕРА ДЛЯ РАСПРЕДЕЛЕННЫХ ВЫЧИСЛЕНИЙ MPI
echo ===============================================================================
echo.
echo ВНИМАНИЕ: Запускайте этот файл от имени Администратора на КАЖДОМ компьютере!
echo.
echo 1. Открытие порта 8677 в Брандмауэре Windows (порт службы SMPD)...
netsh advfirewall firewall add rule name="MPICH2 SMPD Service (Port 8677)" dir=in action=allow protocol=TCP localport=8677 profile=any >nul 2>&1
netsh advfirewall firewall add rule name="MPICH2 SMPD App" dir=in action=allow program="%~dp0smpd.exe" enable=yes profile=any >nul 2>&1
netsh advfirewall firewall add rule name="MPICH2 mpiexec App" dir=in action=allow program="%~dp0mpiexec.exe" enable=yes profile=any >nul 2>&1
echo [OK] Правила брандмауэра успешно добавлены.
echo.

echo 2. Установка и запуск службы менеджера процессов smpd...
"%~dp0smpd.exe" -install
echo.

echo 3. Регистрация учетной записи для запуска удаленных процессов:
echo    (ВАЖНО: Введите имя пользователя и пароль, которые ОДИНАКОВЫ на всех ПК!)
"%~dp0mpiexec.exe" -register
echo.

echo 4. Проверка регистрации:
"%~dp0mpiexec.exe" -validate
echo.

echo 5. Проверка статуса службы:
"%~dp0smpd.exe" -status
echo.
echo ===============================================================================
echo Настройка узла завершена!
echo ===============================================================================
pause
