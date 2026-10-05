# Моделирование аналогового сигнала ПЛК
import matplotlib.pyplot as plt
import numpy as np


class PLC_Analog:
    def __init__(self):
        self.MinRaw = 0
        self.MaxRaw = 27648
        self.MinScale = 0.0
        self.MaxScale = 100.0

    def scale_input(self, raw_value):
        """Масштабирование аналогового входа"""
        if raw_value < self.MinRaw:
            raw_value = self.MinRaw
        if raw_value > self.MaxRaw:
            raw_value = self.MaxRaw
        scaled = self.MinScale + (raw_value - self.MinRaw) / (self.MaxRaw - self.MinRaw) * (self.MaxScale - self.MinScale)
        return scaled

    def scale_output(self, scaled_value):
        """Обратное преобразование для аналогового выхода"""
        raw = int(scaled_value / 100.0 * self.MaxRaw)
        return raw


# Тестирование
if __name__ == "__main__":
    plc = PLC_Analog()
    print("Масштабирование аналогового сигнала")
    print("-" * 50)
    print(f"{'Сырое':>10} | {'Физическое':>12} | {'Обратно':>10}")
    print("-" * 50)
    for raw in [0, 6912, 13824, 20736, 27648]:
        scaled = plc.scale_input(raw)
        back = plc.scale_output(scaled)
        print(f"{raw:>10} | {scaled:>12.2f} | {back:>10}")

    # Построение графика
    raw_values = np.linspace(0, 27648, 100)
    scaled_values = raw_values / 27648 * 100

    plt.figure(figsize=(10, 6))
    plt.plot(raw_values, scaled_values, 'b-', linewidth=2)
    plt.xlabel('Сырое значение АЦП')
    plt.ylabel('Физическая величина, %')
    plt.title('Характеристика масштабирования аналогового сигнала')
    plt.grid(True)
    plt.savefig('analog_scaling.png', dpi=300)
    plt.close()
    print("График масштабирования сохранен в analog_scaling.png")
