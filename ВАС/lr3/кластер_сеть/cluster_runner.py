# -*- coding: utf-8 -*-
"""
Интерактивный мастер сетевого запуска MPI на двух компьютерах
по методичке проф. Арькова В.Ю. (Раздел 11.2)
"""

import os
import subprocess
import socket
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MPIEXEC = os.path.join(BASE_DIR, "mpiexec.exe")

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def ping_host(host):
    print(f"\n[1/3] Проверка сетевой связи с {host} (ping)...")
    res = subprocess.run(["ping", "-n", "2", "-w", "1000", host], capture_output=True, text=True)
    if res.returncode == 0:
        print(f"[OK] Узел {host} доступен по сети!")
        return True
    else:
        print(f"[!] ВНИМАНИЕ: Узел {host} не ответил на ping.")
        print("    Возможные причины: включен Брандмауэр на удаленном ПК или неверный IP.")
        ans = input("    Продолжить запуск все равно? (д/н): ").strip().lower()
        return ans in ["д", "y", "yes", "да"]

def main():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("=" * 78)
    print("  РАСПРЕДЕЛЕННЫЙ ЗАПУСК MPI НА КЛАСТЕРЕ ИЗ ДВУХ КОМПЬЮТЕРОВ (Раздел 11.2)")
    print("  Студент: Гарифуллин К.Р. | Группа: ЭАС-514С | Вариант 4")
    print("=" * 78)
    
    local_ip = get_local_ip()
    print(f"\nТекущий компьютер (Узел 1, Хост): {local_ip}")
    
    remote_ip = input("\nВведите IP-адрес второго компьютера (Узел 2): ").strip()
    if not remote_ip:
        print("IP-адрес не указан. Отмена.")
        input("\nНажмите Enter для выхода...")
        return
        
    if not ping_host(remote_ip):
        return

    print("\n[2/3] Выберите программу для запуска на кластере:")
    print("  [1] mpi1-hello.exe       - Тест 'Hello World' от процессов с обоих компьютеров")
    print("  [2] mpi7-integral.exe    - Численное интегрирование (Вариант 4, N=10^8)")
    print("  [3] mpi8-affinity.exe    - Привязка ядер и исследование Hyper-Threading")
    print("  [4] mpi4-sum.exe         - Параллельное суммирование 10^8")
    
    choice = input("\nНомер программы [1-4] (по умолчанию 2): ").strip() or "2"
    
    p1 = input(f"Количество процессов на Узле 1 ({local_ip}) [по умолчанию 2]: ").strip() or "2"
    p2 = input(f"Количество процессов на Узле 2 ({remote_ip}) [по умолчанию 2]: ").strip() or "2"
    
    prog_map = {
        "1": ("mpi1-hello.exe", []),
        "2": ("mpi7-integral.exe", ["100000000"]),
        "3": ("mpi8-affinity.exe", ["100000000", "2"]),
        "4": ("mpi4-sum.exe", ["100000000"])
    }
    
    exe_name, extra_args = prog_map.get(choice, ("mpi7-integral.exe", ["100000000"]))
    
    print("\n[3/3] Где расположен файл программы на втором компьютере?")
    print("  [1] В локальной сетевой папке (Windows Share, например \\\\" + local_ip + "\\share\\" + exe_name + ")")
    print("  [2] Скопирован на второй компьютер в такую же папку (например, C:\\кластер\\" + exe_name + ")")
    share_choice = input("Выберите вариант [1/2] (по умолчанию 1): ").strip() or "1"
    
    if share_choice == "1":
        share_name = input("Имя сетевой папки на этом ПК [по умолчанию 'кластер']: ").strip() or "кластер"
        exe_path = f"\\\\{local_ip}\\{share_name}\\{exe_name}"
    else:
        exe_path = input("Полный путь к файлу на удаленном ПК [например C:\\кластер\\" + exe_name + "]: ").strip() or f"C:\\кластер\\{exe_name}"
    
    # Сборка команды mpiexec
    cmd = [
        MPIEXEC,
        "-hosts", "2",
        local_ip, str(p1),
        remote_ip, str(p2),
        exe_path
    ] + extra_args
    
    print("\n" + "=" * 78)
    print("СФОРМИРОВАННАЯ КОМАНДА ДЛЯ ЗАПУСКА:")
    print(" ".join(cmd))
    print("=" * 78)
    print("\nЗапуск процессов на узлах сети... Пожалуйста, подождите...\n")
    
    try:
        res = subprocess.run(cmd, cwd=BASE_DIR)
        print("\n" + "-" * 78)
        print(f"Выполнение завершено с кодом: {res.returncode}")
        print("-" * 78)
    except Exception as e:
        print(f"\nОшибка при выполнении: {e}")
        
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()
