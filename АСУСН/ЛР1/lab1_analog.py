# -*- coding: utf-8 -*-
"""
Лабораторная работа 1. Часть 3.
Моделирование аналогового сигнала ПЛК на Python.
"""

import os
import matplotlib.pyplot as plt
import numpy as np

# Настройка шрифта для корректного отображения кириллицы
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


class PLC_Analog:
    """Модель аналогового канала ПЛК (масштабирование и нормализация)."""

    def __init__(self, min_raw=0, max_raw=27648, min_scale=0.0, max_scale=100.0):
        # Диапазон целочисленных значений АЦП/ЦАП (стандарт Siemens S7-1200/1500)
        self.MinRaw = min_raw
        self.MaxRaw = max_raw
        # Инженерный диапазон физической величины (например, уровень 0..100% или давление 0..10 бар)
        self.MinScale = min_scale
        self.MaxScale = max_scale

    def scale_input(self, raw_value):
        """
        Масштабирование аналогового входа (АЦП -> инженерная величина).
        Аналог стандартной функции SCALE_X / NORM_X.
        """
        # Защита от выхода за пределы допустимого диапазона (clamping)
        if raw_value < self.MinRaw:
            raw_value = self.MinRaw
        elif raw_value > self.MaxRaw:
            raw_value = self.MaxRaw

        # Линейная интерполяция
        scaled = self.MinScale + (raw_value - self.MinRaw) / (self.MaxRaw - self.MinRaw) * (self.MaxScale - self.MinScale)
        return scaled

    def scale_output(self, scaled_value):
        """
        Обратное преобразование для аналогового выхода (инженерная величина -> код ЦАП).
        Аналог стандартной функции UNSCALE.
        """
        # Ограничение физической величины границами шкалы
        if scaled_value < self.MinScale:
            scaled_value = self.MinScale
        elif scaled_value > self.MaxScale:
            scaled_value = self.MaxScale

        # Расчет кода ЦАП
        raw = int(round(self.MinRaw + (scaled_value - self.MinScale) / (self.MaxScale - self.MinScale) * (self.MaxRaw - self.MinRaw)))
        return raw


def build_scaling_plot(output_path='analog_scaling.png'):
    """Построение статической характеристики масштабирования аналогового сигнала."""
    plc = PLC_Analog()
    raw_values = np.linspace(0, 27648, 200)
    scaled_values = [plc.scale_input(v) for v in raw_values]

    # Контрольные реперные точки
    test_points_raw = [0, 6912, 13824, 20736, 27648]
    test_points_scaled = [plc.scale_input(r) for r in test_points_raw]

    plt.figure(figsize=(10, 6), dpi=150)
    plt.plot(raw_values, scaled_values, 'b-', linewidth=2.5, label='Статическая характеристика (линейная)')
    plt.plot(test_points_raw, test_points_scaled, 'ro', markersize=7, label='Контрольные точки проверки (0%, 25%, 50%, 75%, 100%)')

    # Аннотации к контрольным точкам
    for raw, sc in zip(test_points_raw, test_points_scaled):
        plt.annotate(f"({raw}, {sc:.1f}%)",
                     xy=(raw, sc),
                     xytext=(raw - 1200, sc + 4),
                     fontsize=9,
                     bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.3))

    plt.xlabel('«Сырое» значение кода АЦП (Raw Value, 0..27648)', fontsize=11)
    plt.ylabel('Физическая величина (Engineering Value), %', fontsize=11)
    plt.title('Характеристика масштабирования аналогового сигнала ПЛК', fontsize=13, fontweight='bold')
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend(loc='lower right', fontsize=10)
    plt.xlim(-500, 29000)
    plt.ylim(-5, 110)

    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()
    print(f"График масштабирования успешно сохранен в файл: {output_path}")


if __name__ == "__main__":
    print("=" * 60)
    print("МОДЕЛИРОВАНИЕ АНАЛОГОВОГО СИГНАЛА И МАСШТАБИРОВАНИЯ ПЛК")
    print("=" * 60)

    plc = PLC_Analog()

    print(f"{'Сырое значение АЦП':>18} | {'Физическое (%):':>15} | {'Обратно в ЦАП':>15}")
    print("-" * 60)

    test_raw_values = [0, 6912, 13824, 20736, 27648]
    for raw in test_raw_values:
        scaled = plc.scale_input(raw)
        back = plc.scale_output(scaled)
        print(f"{raw:>18} | {scaled:>15.2f} | {back:>15}")

    # Построение и сохранение графика
    script_dir = os.path.dirname(os.path.abspath(__file__))
    plot_file = os.path.join(script_dir, 'analog_scaling.png')
    build_scaling_plot(plot_file)
