# -*- coding: utf-8 -*-
"""
Модуль имитатора датчика технологических параметров для АС СН.
Реализует генерацию нормального зашумленного сигнала и внесение типовых отказов:
- обрыв (open circuit)
- короткое замыкание (short circuit)
- дрейф нуля (zero drift)
- "залипание" (stuck value)
- случайные импульсные выбросы (spikes)

Поддерживает работу как в реальном времени (через time.time()),
так и в режиме дискретно-событийного синхронного моделирования (передача sim_time).
"""

import time
import numpy as np


class SensorSimulator:
    """
    Имитатор первичного измерительного преобразователя (датчика).
    Генерирует нормальный технологический сигнал с гауссовым шумом.
    Позволяет вносить регламентированные отказы по расписанию или по требованию.
    """

    def __init__(self, nominal=50.0, noise=0.5, min_val=-40.0, max_val=150.0,
                 drift_rate=0.5, spike_amplitude=30.0, unit="°C", name="Датчик температуры"):
        """
        Инициализация параметров датчика.
        
        :param nominal: номинальное (истинное) технологическое значение
        :param noise: среднеквадратическое отклонение (СКО) аддитивного белого гауссова шума
        :param min_val: нижняя граница физического диапазона измерения
        :param max_val: верхняя граница физического диапазона измерения
        :param drift_rate: скорость дрейфа нуля в единицу времени (ед./с)
        :param spike_amplitude: амплитуда импульсного выброса
        :param unit: единица измерения параметра
        :param name: человекочитаемое наименование датчика
        """
        self.nominal = float(nominal)
        self.noise = float(noise)
        self.min_val = float(min_val)
        self.max_val = float(max_val)
        self.drift_rate = float(drift_rate)
        self.spike_amplitude = float(spike_amplitude)
        self.unit = unit
        self.name = name

        self.fault_type = None          # Текущий активный тип отказа
        self.fault_start_time = None    # Время инициации отказа
        self.fault_duration = 0.0       # Длительность отказа (с)
        self.last_value = self.nominal  # Предыдущее значение для режима "залипания"
        self.drift_offset = 0.0         # Накопленное смещение дрейфа
        self.start_time = time.time()   # Системное время старта
        self.rng = np.random.default_rng(42)  # Воспроизводимый генератор псевдослучайных чисел

    def set_seed(self, seed):
        """Фиксация зерна случайности для повторяемости экспериментов."""
        self.rng = np.random.default_rng(seed)

    def set_fault(self, fault_type, duration=10.0, current_time=None):
        """
        Внести регламентированный отказ.
        
        :param fault_type: 'open', 'short', 'drift', 'stuck', 'spike'
        :param duration: длительность отказа в секундах
        :param current_time: явное виртуальное время (если передано), иначе системное time.time()
        """
        self.fault_type = fault_type
        self.fault_duration = float(duration)
        t = current_time if current_time is not None else (time.time() - self.start_time)
        self.fault_start_time = t
        self.drift_offset = 0.0

    def clear_fault(self):
        """Сбросить активный отказ в нормальное состояние."""
        self.fault_type = None
        self.fault_start_time = None
        self.fault_duration = 0.0
        self.drift_offset = 0.0

    def read(self, sim_time=None):
        """
        Прочитать текущее значение датчика.
        
        :param sim_time: относительное время симуляции в секундах (t >= 0).
                         Если не задано, вычисляется относительно системного времени.
        :return: float (текущее считанное значение с учетом шума и внесенного отказа)
        """
        if sim_time is not None:
            t = float(sim_time)
        else:
            t = time.time() - self.start_time

        # Проверка истечения длительности отказа
        if self.fault_type is not None and self.fault_start_time is not None:
            elapsed = t - self.fault_start_time
            if elapsed > self.fault_duration:
                self.clear_fault()

        # Базовый незашумленный физический процесс + гауссов шум измерений
        noise_val = float(self.rng.normal(0.0, self.noise))
        value = self.nominal + noise_val

        # Моделирование типовых видов отказов
        if self.fault_type == 'open':
            # Обрыв цепи: глубокое падение сигнала ниже допустимого диапазона
            # Например, обрыв термосопротивления или обрыв токовой петли 4-20 мА (< 2 мА)
            value = self.min_val - 10.0

        elif self.fault_type == 'short':
            # Короткое замыкание: резкий бросок сигнала выше максимального предела шкалы
            # (например, зашкаливание АЦП до верхнего насыщения)
            value = self.max_val + 50.0

        elif self.fault_type == 'drift':
            # Скрытый отказ: монотонный дрейф нуля измерительного тракта
            elapsed = t - self.fault_start_time if self.fault_start_time is not None else 0.0
            self.drift_offset = self.drift_rate * elapsed
            value += self.drift_offset

        elif self.fault_type == 'stuck':
            # Залипание: датчик "замирает" на последнем зафиксированном значении (шум отсутствует)
            value = self.last_value

        elif self.fault_type == 'spike':
            # Случайный единичный или знакопеременный импульсный выброс большой амплитуды
            spike_sign = self.rng.choice([-1.0, 1.0])
            value += spike_sign * self.spike_amplitude

        # Запоминаем текущее значение для возможного залипания в следующем такте
        self.last_value = value
        return value

    def reset(self):
        """Сброс состояния симулятора к исходному состоянию."""
        self.clear_fault()
        self.last_value = self.nominal
        self.start_time = time.time()
