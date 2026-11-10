# -*- coding: utf-8 -*-
"""
Модуль алгоритмов цифровой фильтрации и многоуровневой диагностики датчиков АС СН.

Включает:
1. Класс SensorDiagnostics:
   - Допусковый контроль (Range Check: min_val <= x <= max_val)
   - Контроль скорости изменения (Rate of Change Check: |dx/dt| <= max_rate)
   - Контроль «залипания» (Stuck Detection: abs(dx) < eps в течение N тактов)
   - Опциональный трендовый анализ медленного дрейфа (Trend / Drift Analysis)
2. Класс MedianFilter:
   - Нелинейный скользящий медианный фильтр для подавления импульсных помех
3. Класс MovingAverageFilter:
   - Линейный фильтр скользящего среднего для подавления высокочастотного шума
4. Класс KalmanFilter1D:
   - Одноканальный скалярный фильтр Калмана
5. Класс Voting2oo3:
   - Модуль мажоритарного голосования «2 из 3» с селекцией и изоляцией дефектного канала
"""

import numpy as np


class SensorDiagnostics:
    """
    Многоуровневый алгоритм оперативной диагностики первичного измерительного канала.
    Выявляет явные отказы (выход за диапазон, аномальная скорость изменения),
    а также скрытые отказы («залипание» и медленный дрейф нуля).
    """

    def __init__(self, min_val=-40.0, max_val=150.0, max_rate=8.0,
                 stuck_threshold=10, stuck_eps=0.001, enable_trend=False,
                 trend_window=30, trend_slope_thresh=0.35):
        """
        Инициализация диагностического блока.
        
        :param min_val: нижняя граница технологического диапазона (ед.)
        :param max_val: верхняя граница технологического диапазона (ед.)
        :param max_rate: максимально допустимая физическая скорость изменения (|dx/dt|, ед./с)
        :param stuck_threshold: количество последовательных тактов постоянства сигнала для фиксации залипания
        :param stuck_eps: порог вариации сигнала (ниже порога значение считается постоянным)
        :param enable_trend: включение расширенного режима трендового анализа скрытого дрейфа
        :param trend_window: глубина буфера для МНК-анализа наклона тренда
        :param trend_slope_thresh: порог угла наклона тренда (ед./с)
        """
        self.min_val = float(min_val)
        self.max_val = float(max_val)
        self.max_rate = float(max_rate)
        self.stuck_threshold = int(stuck_threshold)
        self.stuck_eps = float(stuck_eps)
        self.enable_trend = bool(enable_trend)
        self.trend_window = int(trend_window)
        self.trend_slope_thresh = float(trend_slope_thresh)

        # Состояние алгоритма
        self.prev_value = None
        self.prev_time = None
        self.stuck_counter = 0
        self.faults = []  # Журнал: [(timestamp, value, [problems])]
        self.trend_buffer = []

    def reset(self):
        """Полный сброс памяти алгоритма."""
        self.prev_value = None
        self.prev_time = None
        self.stuck_counter = 0
        self.faults.clear()
        self.trend_buffer.clear()

    def check(self, value, timestamp):
        """
        Проверка одного отсчета технологического сигнала.
        
        :param value: значение отсчета (обычно после предварительной фильтрации)
        :param timestamp: метка времени получения отсчета (с)
        :return: список кодов выявленных дефектов
        """
        problems = []

        # 1. Проверка диапазона (Range Check)
        if value < self.min_val or value > self.max_val:
            problems.append('out_of_range')

        # 2. Проверка скорости изменения (Rate of Change Check)
        if self.prev_value is not None and self.prev_time is not None:
            dt = timestamp - self.prev_time
            if dt > 1e-5:
                rate = abs(value - self.prev_value) / dt
                if rate > self.max_rate:
                    problems.append('rate_exceeded')

        # 3. Проверка «залипания» (Stuck Detection)
        if self.prev_value is not None:
            if abs(value - self.prev_value) < self.stuck_eps:
                self.stuck_counter += 1
                if self.stuck_counter >= self.stuck_threshold:
                    problems.append('stuck')
            else:
                self.stuck_counter = 0

        # 4. Расширенный трендовый анализ дрейфа (если активирован)
        if self.enable_trend:
            self.trend_buffer.append((timestamp, value))
            if len(self.trend_buffer) > self.trend_window:
                self.trend_buffer.pop(0)

            if len(self.trend_buffer) >= self.trend_window:
                ts = np.array([pt[0] for pt in self.trend_buffer])
                ys = np.array([pt[1] for pt in self.trend_buffer])
                dt_span = ts[-1] - ts[0]
                if dt_span > 1.0:
                    slope, _ = np.polyfit(ts, ys, 1)
                    if abs(slope) >= self.trend_slope_thresh and ('stuck' not in problems):
                        problems.append('drift_suspected')

        # Обновление предыдущих значений
        self.prev_value = value
        self.prev_time = timestamp

        # Журналирование обнаруженных аномалий
        if problems:
            self.faults.append((timestamp, value, list(problems)))

        return problems


class MedianFilter:
    """
    Нелинейный скользящий медианный фильтр.
    Эффективно подавляет одиночные импульсные выбросы (spikes),
    сохраняя перепады уровней и исключая фазовое запаздывание на монотонных участках.
    """

    def __init__(self, window_size=5):
        self.window_size = int(window_size)
        self.buffer = []

    def reset(self):
        self.buffer.clear()

    def filter(self, value):
        self.buffer.append(float(value))
        if len(self.buffer) > self.window_size:
            self.buffer.pop(0)
        return float(np.median(self.buffer))


class MovingAverageFilter:
    """
    Линейный фильтр скользящего среднего (Moving Average Filter).
    """

    def __init__(self, window_size=5):
        self.window_size = int(window_size)
        self.buffer = []

    def reset(self):
        self.buffer.clear()

    def filter(self, value):
        self.buffer.append(float(value))
        if len(self.buffer) > self.window_size:
            self.buffer.pop(0)
        return float(np.mean(self.buffer))


class KalmanFilter1D:
    """
    Одномерный дискретный скалярный фильтр Калмана (1D Kalman Filter).
    Обеспечивает оптимальную рекуррентную фильтрацию измерений.
    """

    def __init__(self, initial_state=50.0, initial_variance=1.0, process_variance=1e-4, measurement_variance=0.25):
        self.x = float(initial_state)
        self.p = float(initial_variance)
        self.q = float(process_variance)
        self.r = float(measurement_variance)

    def reset(self, state=50.0, variance=1.0):
        self.x = float(state)
        self.p = float(variance)

    def filter(self, measurement):
        # Экстраполяция (прогноз)
        x_pred = self.x
        p_pred = self.p + self.q

        # Коэффициент усиления Калмана
        k_gain = p_pred / (p_pred + self.r)

        # Коррекция
        self.x = x_pred + k_gain * (float(measurement) - x_pred)
        self.p = (1.0 - k_gain) * p_pred

        return float(self.x)


class Voting2oo3:
    """
    Модуль мажоритарного резервирования «2 из 3» (Two-out-of-three voting).
    Определяет консенсусное значение и диагностирует поврежденный измерительный канал.
    """

    def __init__(self, discrepancy_threshold=3.0):
        self.threshold = float(discrepancy_threshold)

    def vote(self, ch1, ch2, ch3):
        channels = [float(ch1), float(ch2), float(ch3)]
        consensus = float(np.median(channels))

        faulty = []
        for idx, val in enumerate(channels, start=1):
            if abs(val - consensus) > self.threshold:
                faulty.append(f"Channel_{idx}")

        return consensus, faulty
