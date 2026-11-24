# -*- coding: utf-8 -*-
"""
run_all.py
Мастер-запуск распределенной системы управления (АСУСН, ЛР №6).
Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.

Запускает все 4 подсистемы в отдельных независимых консольных окнах:
  1. Имитатор технологического объекта (порт 9999)
  2. Виртуальный ПЛК №1 (Modbus TCP порт 5020 - ПИД-регулятор)
  3. Виртуальный ПЛК №2 (Modbus TCP порт 5021 - ПАЗ, насос, клапан)
  4. SCADA-система диспетчерского контроля и управления
"""

import os
import sys
import time
import subprocess
import signal

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

COMPONENTS = [
    {
        'name': '[1/4] Имитатор объекта (порт 9999)',
        'script': os.path.join(BASE_DIR, 'plant_simulator.py'),
        'delay': 1.5
    },
    {
        'name': '[2/4] ПЛК №1 Регулятор (Modbus 5020)',
        'script': os.path.join(BASE_DIR, 'plc1.py'),
        'delay': 1.5
    },
    {
        'name': '[3/4] ПЛК №2 ПАЗ и Клапан (Modbus 5021)',
        'script': os.path.join(BASE_DIR, 'plc2.py'),
        'delay': 1.5
    },
    {
        'name': '[4/4] SCADA Диспетчерский терминал',
        'script': os.path.join(BASE_DIR, 'scada.py'),
        'delay': 0.5
    }
]


def main():
    print("=" * 72)
    print("  МАСТЕР-ЗАПУСК РАСПРЕДЕЛЕННОЙ СИСТЕМЫ УПРАВЛЕНИЯ (АСУСН ЛР №6)")
    print("  Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С, 2026 г.")
    print("=" * 72)

    processes = []
    python_exe = sys.executable

    # Флаг создания нового окна консоли в Windows
    creationflags = 0
    if sys.platform == 'win32':
        creationflags = subprocess.CREATE_NEW_CONSOLE

    try:
        for comp in COMPONENTS:
            name = comp['name']
            script = comp['script']
            delay = comp['delay']

            if not os.path.isfile(script):
                print(f"[ОШИБКА] Файл скрипта не найден: {script}")
                continue

            print(f"-> Запуск: {name} ...")
            p = subprocess.Popen(
                [python_exe, script],
                cwd=BASE_DIR,
                creationflags=creationflags
            )
            processes.append((name, p))
            time.sleep(delay)

        print("\n" + "=" * 72)
        print("  Все 4 компонента успешно запущены в отдельных окнах!")
        print("  Нажмите Ctrl+C в этом окне для завершения всех процессов.")
        print("=" * 72 + "\n")

        # Ожидание завершения или прерывания
        while True:
            time.sleep(1)
            # Проверяем, не закрыл ли пользователь SCADA
            scada_proc = processes[-1][1] if processes else None
            if scada_proc and scada_proc.poll() is not None:
                print("Окно SCADA закрыто. Завершение вспомогательных процессов...")
                break

    except KeyboardInterrupt:
        print("\nПолучен сигнал остановки (Ctrl+C)...")
    finally:
        print("Завершение запущенных подсистем...")
        for name, p in processes:
            if p.poll() is None:
                try:
                    p.terminate()
                    p.wait(timeout=1.0)
                except Exception:
                    p.kill()
                print(f"  Остановлен: {name}")
        print("Все компоненты успешно остановлены.")


if __name__ == '__main__':
    main()
