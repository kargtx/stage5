# -*- coding: utf-8 -*-
"""
Лабораторная работа 1. Часть 2.
Моделирование дискретной логики ПЛК на Python.
"""

import os
import matplotlib.pyplot as plt

# Настройка шрифта для корректного отображения кириллицы
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


class PLC_Discrete:
    """Модель программируемого логического контроллера (дискретная логика)."""

    def __init__(self):
        # Таблица дискретных входов (Process Image Input)
        self.DI = {0: False, 1: False, 2: False, 3: False}
        # Таблица дискретных выходов (Process Image Output)
        self.DQ = {0: False, 1: False, 2: False, 3: False}
        # Внутренний флаг состояния самоблокировки (RS-триггер с приоритетом сброса)
        self.self_lock = False
        # Счётчик нажатий кнопки DI0
        self.counter = 0
        # Предыдущее состояние DI0 для фиксации переднего фронта (R_TRIG)
        self.prev_DI0 = False

    def read_inputs(self, di0, di1, di2, di3):
        """Фаза 1: Чтение входов в область памяти отображения входов."""
        self.DI[0] = bool(di0)
        self.DI[1] = bool(di1)
        self.DI[2] = bool(di2)
        self.DI[3] = bool(di3)

    def execute(self):
        """Фаза 2: Выполнение прикладной программы управления."""
        # Логика 1: схема пуск/стоп с самоблокировкой (DI0 - Пуск, DI1 - Стоп)
        # Приоритет команды «Стоп» (DI1)
        if self.DI[0] and not self.DI[1]:
            self.self_lock = True
        if self.DI[1]:
            self.self_lock = False
        self.DQ[0] = self.self_lock

        # Логика 2: логическая операция «И» (AND): DI0 и DI2
        self.DQ[1] = self.DI[0] and self.DI[2]

        # Логика 3: логическая операция «ИЛИ» (OR): DI0 или DI3
        self.DQ[2] = self.DI[0] or self.DI[3]

        # Логика 4: логическая операция «НЕ» (NOT): инверсия сигнала аварии/стопа DI1
        self.DQ[3] = not self.DI[1]

        # Логика 5: счётчик импульсов с детектором переднего фронта DI0
        if self.DI[0] and not self.prev_DI0:
            self.counter += 1
        self.prev_DI0 = self.DI[0]

    def write_outputs(self):
        """Фаза 3: Запись результатов из памяти отображения в физические выходы."""
        return self.DQ.copy()

    def print_state(self):
        """Информационный вывод текущего состояния входов, выходов и счётчика."""
        print(f"Входы:  DI0 (Пуск)={int(self.DI[0])}, DI1 (Стоп)={int(self.DI[1])}, "
              f"DI2={int(self.DI[2])}, DI3={int(self.DI[3])}")
        print(f"Выходы: DQ0 (Самоблок)={int(self.DQ[0])}, DQ1 (И)={int(self.DQ[1])}, "
              f"DQ2 (ИЛИ)={int(self.DQ[2])}, DQ3 (НЕ)={int(self.DQ[3])}")
        print(f"Счётчик нажатий DI0: {self.counter}")
        print("-" * 55)


def build_timing_diagram(output_path='timing_diagram.png'):
    """Построение временных диаграмм работы дискретной логики."""
    time = list(range(10))
    # Временные диаграммы входных и выходных сигналов
    # Шаги 0..3: нажат пуск DI0; Шаг 4: пуск отпущен, но лампа горит (самоблокировка);
    # Шаги 5..6: нажат стоп DI1, лампа гаснет; Шаги 7..9: покой
    di0_signal = [0, 1, 1, 1, 0, 0, 0, 0, 0, 0]
    di1_signal = [0, 0, 0, 0, 0, 1, 1, 0, 0, 0]
    dq0_signal = [0, 1, 1, 1, 1, 0, 0, 0, 0, 0]

    plt.figure(figsize=(10, 7), dpi=150)

    # 1. DI0 (Пуск)
    plt.subplot(3, 1, 1)
    plt.step(time, di0_signal, where='post', color='#1f77b4', linewidth=2)
    plt.title('Дискретный вход DI0 («Пуск»)', fontsize=12, fontweight='bold')
    plt.ylabel('Сигнал (0/1)', fontsize=10)
    plt.ylim(-0.15, 1.25)
    plt.yticks([0, 1])
    plt.grid(True, linestyle='--', alpha=0.6)

    # 2. DI1 (Стоп)
    plt.subplot(3, 1, 2)
    plt.step(time, di1_signal, where='post', color='#d62728', linewidth=2)
    plt.title('Дискретный вход DI1 («Стоп»)', fontsize=12, fontweight='bold')
    plt.ylabel('Сигнал (0/1)', fontsize=10)
    plt.ylim(-0.15, 1.25)
    plt.yticks([0, 1])
    plt.grid(True, linestyle='--', alpha=0.6)

    # 3. DQ0 (Лампа / Самоблокировка)
    plt.subplot(3, 1, 3)
    plt.step(time, dq0_signal, where='post', color='#2ca02c', linewidth=2)
    plt.title('Дискретный выход DQ0 (Исполнительный механизм / Самоблокировка)', fontsize=12, fontweight='bold')
    plt.xlabel('Дискретные такты цикла ПЛК (t)', fontsize=11)
    plt.ylabel('Сигнал (0/1)', fontsize=10)
    plt.ylim(-0.15, 1.25)
    plt.yticks([0, 1])
    plt.grid(True, linestyle='--', alpha=0.6)

    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()
    print(f"Временная диаграмма успешно сохранена в файл: {output_path}")


if __name__ == "__main__":
    print("=" * 55)
    print("МОДЕЛИРОВАНИЕ ДИСКРЕТНОЙ ЛОГИКИ ПЛК")
    print("=" * 55)

    plc = PLC_Discrete()

    # Сценарий 1: Нажатие кнопки «Пуск»
    print("Сценарий 1: Нажатие кнопки «Пуск» (DI0=True, DI1=False)")
    plc.read_inputs(True, False, False, False)
    plc.execute()
    plc.print_state()

    # Сценарий 2: Отпускание кнопки «Пуск» (проверка самоблокировки)
    print("Сценарий 2: Отпускание кнопки «Пуск» (DI0=False, DI1=False) — самоблокировка")
    plc.read_inputs(False, False, False, False)
    plc.execute()
    plc.print_state()

    # Сценарий 3: Нажатие кнопки «Стоп»
    print("Сценарий 3: Нажатие кнопки «Стоп» (DI0=False, DI1=True) — сброс выхода")
    plc.read_inputs(False, True, False, False)
    plc.execute()
    plc.print_state()

    # Сценарий 4: Проверка логики «И»
    print("Сценарий 4: Логика «И» (DI0=True, DI2=True) -> DQ1=1")
    plc.read_inputs(True, False, True, False)
    plc.execute()
    plc.print_state()

    # Сценарий 5: Проверка логики «ИЛИ»
    print("Сценарий 5: Логика «ИЛИ» (DI0=False, DI3=True) -> DQ2=1")
    plc.read_inputs(False, False, False, True)
    plc.execute()
    plc.print_state()

    # Построение временной диаграммы
    script_dir = os.path.dirname(os.path.abspath(__file__))
    diagram_file = os.path.join(script_dir, 'timing_diagram.png')
    build_timing_diagram(diagram_file)
