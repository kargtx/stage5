# Моделирование дискретной логики ПЛК
import matplotlib.pyplot as plt


class PLC_Discrete:
    def __init__(self):
        # Дискретные входы
        self.DI = {0: False, 1: False, 2: False, 3: False}
        # Дискретные выходы
        self.DQ = {0: False, 1: False, 2: False, 3: False}
        # Состояние самоблокировки
        self.self_lock = False
        # Счётчик нажатий
        self.counter = 0
        self.prev_DI0 = False

    def read_inputs(self, DI0, DI1, DI2, DI3):
        """Чтение входов"""
        self.DI[0] = DI0
        self.DI[1] = DI1
        self.DI[2] = DI2
        self.DI[3] = DI3

    def execute(self):
        """Выполнение программы"""
        # Логика 1: самоблокировка (пуск DI0, стоп DI1)
        if self.DI[0] and not self.DI[1]:
            self.self_lock = True
        if self.DI[1]:
            self.self_lock = False
        self.DQ[0] = self.self_lock

        # Логика 2: И (DI0 и DI2)
        self.DQ[1] = self.DI[0] and self.DI[2]

        # Логика 3: ИЛИ (DI0 или DI3)
        self.DQ[2] = self.DI[0] or self.DI[3]

        # Логика 4: НЕ (инверсия DI1)
        self.DQ[3] = not self.DI[1]

        # Логика 5: счётчик нажатий DI0
        if self.DI[0] and not self.prev_DI0:
            self.counter += 1
        self.prev_DI0 = self.DI[0]

    def write_outputs(self):
        """Запись выходов"""
        return self.DQ

    def print_state(self):
        """Вывод состояния"""
        print(f"Входы: DI0={self.DI[0]}, DI1={self.DI[1]}, DI2={self.DI[2]}, DI3={self.DI[3]}")
        print(f"Выходы: DQ0={self.DQ[0]}, DQ1={self.DQ[1]}, DQ2={self.DQ[2]}, DQ3={self.DQ[3]}")
        print(f"Счётчик: {self.counter}")
        print("-" * 40)


# Тестирование
if __name__ == "__main__":
    plc = PLC_Discrete()

    # Сценарий 1: нажатие пуска
    print("Сценарий 1: Нажатие пуска (DI0=True)")
    plc.read_inputs(True, False, False, False)
    plc.execute()
    plc.print_state()

    # Сценарий 2: отпускание пуска (самоблокировка)
    print("Сценарий 2: Отпускание пуска (DI0=False)")
    plc.read_inputs(False, False, False, False)
    plc.execute()
    plc.print_state()

    # Сценарий 3: нажатие стопа
    print("Сценарий 3: Нажатие стопа (DI1=True)")
    plc.read_inputs(False, True, False, False)
    plc.execute()
    plc.print_state()

    # Сценарий 4: логика И
    print("Сценарий 4: Логика И (DI0=True, DI2=True)")
    plc.read_inputs(True, False, True, False)
    plc.execute()
    plc.print_state()

    # Сценарий 5: логика ИЛИ
    print("Сценарий 5: Логика ИЛИ (DI0=True, DI3=True)")
    plc.read_inputs(True, False, False, True)
    plc.execute()
    plc.print_state()

    # Временные диаграммы
    time = list(range(10))
    DI0 = [0, 1, 1, 1, 0, 0, 0, 0, 0, 0]
    DI1 = [0, 0, 0, 0, 0, 1, 1, 0, 0, 0]
    DQ0 = [0, 1, 1, 1, 1, 0, 0, 0, 0, 0]

    plt.figure(figsize=(10, 6))

    plt.subplot(3, 1, 1)
    plt.step(time, DI0, where='post', color='black', linewidth=1.5)
    plt.title('DI0 (Пуск)')
    plt.ylim(-0.1, 1.1)
    plt.grid(True)

    plt.subplot(3, 1, 2)
    plt.step(time, DI1, where='post', color='black', linewidth=1.5)
    plt.title('DI1 (Стоп)')
    plt.ylim(-0.1, 1.1)
    plt.grid(True)

    plt.subplot(3, 1, 3)
    plt.step(time, DQ0, where='post', color='black', linewidth=1.5)
    plt.title('DQ0 (Лампа, самоблокировка)')
    plt.ylim(-0.1, 1.1)
    plt.grid(True)

    plt.tight_layout()
    plt.savefig('timing_diagram.png', dpi=300)
    plt.close()
    print("Временные диаграммы сохранены в timing_diagram.png")
