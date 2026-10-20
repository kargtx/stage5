# -*- coding: utf-8 -*-
"""
Практическое занятие № 3
Дисциплина: «Автоматизированные системы специального назначения» (АССН)
Тема: Разработка алгоритма работы системы в нештатной ситуации (отказ датчика давления: обрыв, КЗ)
Вариант 2: Отказ датчика давления в мажоритарной системе ПАЗ НПС (обрыв, короткое замыкание, деградация 2oo3)
Студенты: Гайнанов Д.И., Гарифуллин К.Р., группа ЭАС-514С
Преподаватель: Антонов В.В.
Уфа, УУНиТ, 2026 г.
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Настройка шрифтов для качественного отображения ГОСТ
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.titlesize'] = 11
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['xtick.labelsize'] = 9
plt.rcParams['ytick.labelsize'] = 9
plt.rcParams['legend.fontsize'] = 8.5
plt.rcParams['figure.titlesize'] = 12

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

def run_sensor_simulation():
    print("=" * 75)
    print("ИМИТАЦИОННОЕ МОДЕЛИРОВАНИЕ АЛГОРИТМА ДИАГНОСТИКИ И ПАРИРОВАНИЯ ОТКАЗОВ ДАТЧИКОВ")
    print("Вариант № 2: Отказ датчика давления (обрыв цепи, короткое замыкание, деградация 2oo3)")
    print("=" * 75)

    # Параметры времени: 0 .. 60 секунд с шагом 50 мс (соответствует циклу контроллера ПЛК)
    dt = 0.05
    t = np.arange(0, 60.0 + dt, dt)
    n_points = len(t)

    # Базовое давление процесса на выходе МНА: 6.5 МПа с плавным технологическим колебанием
    p_true = 6.5 + 0.3 * np.sin(0.15 * t) + 0.1 * np.cos(0.5 * t)
    # Диапазон датчиков: 0 .. 10 МПа -> ток 4 .. 20 мА
    # I = 4.0 + (P / 10.0) * 16.0
    def p_to_curr(p):
        return 4.0 + (p / 10.0) * 16.0

    def curr_to_p(i):
        return (i - 4.0) / 16.0 * 10.0

    # Генерируем нормальные показания с небольшим шумом для трех каналов (PT-01A, PT-01B, PT-01C)
    np.random.seed(42)
    noise_a = np.random.normal(0, 0.03, n_points)
    noise_b = np.random.normal(0, 0.03, n_points)
    noise_c = np.random.normal(0, 0.03, n_points)

    i_a = p_to_curr(p_true) + noise_a
    i_b = p_to_curr(p_true) + noise_b
    i_c = p_to_curr(p_true) + noise_c

    # Сценарий отказов:
    # 1. На t = 14.0 c датчик PT-01A обрывается (ток падает до 0.5 мА < 3.6 мА NAMUR)
    #    На t = 45.0 c происходит замена датчика PT-01A и возврат тока к норме
    for idx, ti in enumerate(t):
        if 14.0 <= ti < 45.0:
            i_a[idx] = 0.4 + np.random.normal(0, 0.05)
        elif ti >= 45.0:
            i_a[idx] = p_to_curr(p_true[idx]) + noise_a[idx]

    # 2. На t = 28.0 c датчик PT-01B замыкается накоротко (ток подскакивает до 23.5 мА > 21.0 мА NAMUR)
    for idx, ti in enumerate(t):
        if ti >= 28.0:
            i_b[idx] = 23.4 + np.random.normal(0, 0.08)

    # 3. Датчик PT-01C остается физически исправным на протяжении всех 60 секунд.

    # 4. Аналитическая (расчетная) модель давления P_model
    # По расходу нефти Q и напору насоса H(Q):
    p_model = p_true + np.random.normal(0, 0.08, n_points)

    # Алгоритм диагностики по стандарту NAMUR NE 43:
    # I < 3.6 мА -> ОБРЫВ (Under-range fault)
    # 3.6 <= I < 3.8 -> Предупреждение (Low warning)
    # 3.8 <= I <= 20.5 -> НОРМАЛЬНЫЙ ДИАПАЗОН
    # 20.5 < I <= 21.0 -> Предупреждение (High warning)
    # I > 21.0 мА -> КОРОТКОЕ ЗАМЫКАНИЕ (Over-range fault)

    status_a = np.zeros(n_points) # 0 - OK, 1 - ОБРЫВ, 2 - КЗ
    status_b = np.zeros(n_points)
    status_c = np.zeros(n_points)

    # Массив итогового измеренного давления P_system
    p_system = np.zeros(n_points)
    redundancy_mode = [] # Режим схемы: "2oo3", "1oo2D", "1oo1 + Model"

    for idx in range(n_points):
        # Диагностика канала A
        if i_a[idx] < 3.6:
            status_a[idx] = 1 # Обрыв
        elif i_a[idx] > 21.0:
            status_a[idx] = 2 # КЗ
        else:
            status_a[idx] = 0 # Норма

        # Диагностика канала B
        if i_b[idx] < 3.6:
            status_b[idx] = 1
        elif i_b[idx] > 21.0:
            status_b[idx] = 2
        else:
            status_b[idx] = 0

        # Диагностика канала C
        if i_c[idx] < 3.6:
            status_c[idx] = 1
        elif i_c[idx] > 21.0:
            status_c[idx] = 2
        else:
            status_c[idx] = 0

        # Логика деградации мажоритарной схемы:
        healthy_vals = []
        if status_a[idx] == 0:
            healthy_vals.append(curr_to_p(i_a[idx]))
        if status_b[idx] == 0:
            healthy_vals.append(curr_to_p(i_b[idx]))
        if status_c[idx] == 0:
            healthy_vals.append(curr_to_p(i_c[idx]))

        n_healthy = len(healthy_vals)

        if n_healthy == 3:
            # 2oo3 (медианное значение трех каналов)
            p_system[idx] = np.median(healthy_vals)
            redundancy_mode.append(3) # Код 3: схема 2oo3
        elif n_healthy == 2:
            # Деградация в 1oo2D: усреднение двух исправных каналов с контролем расхождения
            p_system[idx] = np.mean(healthy_vals)
            redundancy_mode.append(2) # Код 2: схема 1oo2D
        elif n_healthy == 1:
            # Остался 1 физический датчик: привлекаем расчетную модель P_model
            # Если расхождение между PT-01C и P_model < 0.5 МПа -> доверяем датчику C
            p_cand = healthy_vals[0]
            if abs(p_cand - p_model[idx]) < 0.8:
                p_system[idx] = p_cand
            else:
                p_system[idx] = p_model[idx]
            redundancy_mode.append(1) # Код 1: схема 1oo1 + Модель
        else:
            # Все физические датчики отказали: переход на расчетную модель
            p_system[idx] = p_model[idx]
            redundancy_mode.append(0) # Код 0: Аварийный режим по модели

    redundancy_mode = np.array(redundancy_mode)

    print("Симуляция успешно выполнена:")
    print(f"- Всего расчетных точек: {n_points}")
    print(f"- t = 0..14 c: режим 2oo3, все 3 канала исправны")
    print(f"- t = 14..28 c: обрыв PT-01A (ток = {i_a[int(20/dt)]:.2f} мА < 3.6 мА), деградация в 1oo2D")
    print(f"- t = 28..45 c: КЗ PT-01B (ток = {i_b[int(32/dt)]:.2f} мА > 21.0 мА), деградация в 1oo1 + Расчетная модель")
    print(f"- t = 45..60 c: физическая замена PT-01A, ре-арминг и восстановление схемы в 1oo2D")

    # Построение графика результатов симуляции
    fig, axes = plt.subplots(4, 1, figsize=(11, 10), sharex=True)
    plt.subplots_adjust(hspace=0.28)

    # График 1: Токи датчиков и пороги NAMUR NE 43
    ax1 = axes[0]
    ax1.plot(t, i_a, label='PT-01A (Ток канала А)', color='#e11d48', lw=1.6)
    ax1.plot(t, i_b, label='PT-01B (Ток канала B)', color='#f59e0b', lw=1.6)
    ax1.plot(t, i_c, label='PT-01C (Ток канала C)', color='#10b981', lw=1.6)
    ax1.axhline(3.6, color='darkred', ls='--', lw=1.2, label='Порог обрыва NAMUR (< 3.6 мА)')
    ax1.axhline(21.0, color='darkorange', ls='--', lw=1.2, label='Порог КЗ NAMUR (> 21.0 мА)')
    ax1.axhline(4.0, color='gray', ls=':', lw=0.8)
    ax1.axhline(20.0, color='gray', ls=':', lw=0.8)
    ax1.set_ylabel('Ток, мА')
    ax1.set_title('а) Токовые сигналы каналов 4–20 мА и диагностические пороги NAMUR NE 43', fontweight='bold')
    ax1.grid(True, linestyle=':', alpha=0.6)
    ax1.legend(loc='center left', bbox_to_anchor=(1.01, 0.5), frameon=True)
    ax1.set_ylim(-1, 26)

    # Аннотации к событиям
    ax1.annotate('Обрыв PT-01A\n(ток < 3.6 мА)', xy=(14, 0.5), xytext=(15, 6),
                 arrowprops=dict(facecolor='#e11d48', shrink=0.08, width=1, headwidth=5),
                 fontsize=8.5, fontweight='bold', color='#e11d48')
    ax1.annotate('КЗ PT-01B\n(ток > 21 мА)', xy=(28, 23.4), xytext=(22, 17),
                 arrowprops=dict(facecolor='#f59e0b', shrink=0.08, width=1, headwidth=5),
                 fontsize=8.5, fontweight='bold', color='#b45309')
    ax1.annotate('Замена и ввод\nPT-01A в работу', xy=(45, 14.5), xytext=(46, 21),
                 arrowprops=dict(facecolor='#10b981', shrink=0.08, width=1, headwidth=5),
                 fontsize=8.5, fontweight='bold', color='#047857')

    # График 2: Давление в каналах и расчетная модель
    ax2 = axes[1]
    ax2.plot(t, p_true, label='Истинное давление в трубопроводе', color='black', lw=2.0, ls='-')
    ax2.plot(t, curr_to_p(i_c), label='Измерение PT-01C (исправен)', color='#10b981', lw=1.4, alpha=0.85)
    ax2.plot(t, p_model, label='Расчетная модель P_model (по напору насоса)', color='#6366f1', lw=1.3, ls='-.')
    ax2.plot(t, p_system, label='Итоговое давление контура ПАЗ (P_сист)', color='#0284c7', lw=2.2, ls='--')
    ax2.set_ylabel('Давление, МПа')
    ax2.set_title('б) Восстановленные значения давления технологического процесса и расчётная модель', fontweight='bold')
    ax2.grid(True, linestyle=':', alpha=0.6)
    ax2.legend(loc='center left', bbox_to_anchor=(1.01, 0.5), frameon=True)
    ax2.set_ylim(4.5, 8.5)

    # График 3: Статусы отказов каналов
    ax3 = axes[2]
    ax3.step(t, status_a, where='post', label='Статус PT-01A (1: Обрыв)', color='#e11d48', lw=1.7)
    ax3.step(t, status_b + 0.05, where='post', label='Статус PT-01B (2: КЗ)', color='#f59e0b', lw=1.7)
    ax3.step(t, status_c - 0.05, where='post', label='Статус PT-01C (0: Норма)', color='#10b981', lw=1.7)
    ax3.set_yticks([0, 1, 2])
    ax3.set_yticklabels(['0: Норма', '1: Обрыв цепи', '2: Короткое замыкание'])
    ax3.set_ylabel('Статус канала')
    ax3.set_title('в) Дискретные флаги аварийной самодиагностики модулей ввода AI ПЛК', fontweight='bold')
    ax3.grid(True, linestyle=':', alpha=0.6)
    ax3.legend(loc='center left', bbox_to_anchor=(1.01, 0.5), frameon=True)
    ax3.set_ylim(-0.3, 2.5)

    # График 4: Режим мажорирования и конфигурация системы
    ax4 = axes[3]
    ax4.step(t, redundancy_mode, where='post', label='Конфигурация схемы безопасности', color='#0284c7', lw=2.2)
    ax4.set_yticks([0, 1, 2, 3])
    ax4.set_yticklabels(['0: Модель', '1: 1oo1 + Модель', '2: Деградация 1oo2D', '3: Штатный 2oo3'])
    ax4.set_ylabel('Конфигурация')
    ax4.set_xlabel('Время технологического процесса, с')
    ax4.set_title('г) Динамическая реконфигурация структуры безопасности контура ПАЗ во времени', fontweight='bold')
    ax4.grid(True, linestyle=':', alpha=0.6)
    ax4.legend(loc='center left', bbox_to_anchor=(1.01, 0.5), frameon=True)
    ax4.set_ylim(-0.4, 3.5)

    plot_file = os.path.join(OUTPUT_DIR, 'pr3_sensor_fault_diagnostics.png')
    plt.savefig(plot_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"График симуляции успешно сохранён в: {plot_file}")

def generate_flowchart():
    print("\nГенерация блок-схемы алгоритма по ГОСТ 19.701-90...")

    # Создаем аккуратную векторную блок-схему алгоритма обработки отказа датчика
    fig, ax = plt.subplots(figsize=(10, 14), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 140)
    ax.axis('off')

    # Цветовая палитра для уровней:
    # Уровень ПЛК (локальное обнаружение и парирование) - светло-голубой / бирюзовый
    # Уровень SCADA (оперативное управление, сигнализация) - светло-желтый / оранжевый
    # Уровень Диспетчера / КИПиА (регламент, замена) - фиолетовый / зеленый
    c_start = '#3b82f6'
    c_process = '#f1f5f9'
    c_condition = '#fef3c7'
    c_edge = '#1e293b'
    c_alarm = '#fee2e2'
    c_repair = '#dcfce7'

    def draw_box(x, y, w, h, text, shape='rect', bg_color=c_process, border_color=c_edge):
        if shape == 'oval':
            box = patches.FancyBboxPatch((x - w/2, y - h/2), w, h,
                                         boxstyle="round,pad=0.2,rounding_size=1.8",
                                         facecolor=bg_color, edgecolor=border_color, linewidth=1.5)
            ax.add_patch(box)
        elif shape == 'rect':
            box = patches.FancyBboxPatch((x - w/2, y - h/2), w, h,
                                         boxstyle="square,pad=0.0",
                                         facecolor=bg_color, edgecolor=border_color, linewidth=1.3)
            ax.add_patch(box)
        elif shape == 'rhombus':
            # Ромб (условие)
            pts = [[x, y + h/2], [x + w/2, y], [x, y - h/2], [x - w/2, y]]
            poly = patches.Polygon(pts, closed=True, facecolor=bg_color, edgecolor=border_color, linewidth=1.4)
            ax.add_patch(poly)
        elif shape == 'io':
            # Параллелограмм (ввод-вывод)
            offset = w * 0.12
            pts = [[x - w/2 + offset, y + h/2], [x + w/2, y + h/2],
                   [x + w/2 - offset, y - h/2], [x - w/2, y - h/2]]
            poly = patches.Polygon(pts, closed=True, facecolor=bg_color, edgecolor=border_color, linewidth=1.3)
            ax.add_patch(poly)

        ax.text(x, y, text, ha='center', va='center', fontsize=8.2, color='#0f172a',
                fontweight='normal', linespacing=1.2)

    def draw_arrow(x1, y1, x2, y2, label=''):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(facecolor=c_edge, edgecolor=c_edge, width=1.1, headwidth=5, headlength=6))
        if label:
            mx = (x1 + x2) / 2
            my = (y1 + y2) / 2
            ax.text(mx + 1.2, my, label, fontsize=8, color='#0f172a', fontweight='bold', va='center')

    # Заголовок схемы
    ax.text(50, 137, "БЛОК-СХЕМА АЛГОРИТМА ДИАГНОСТИКИ И ПАРИРОВАНИЯ ОТКАЗА ДАТЧИКА (ГОСТ 19.701-90)",
            ha='center', va='center', fontsize=11, fontweight='bold', color='#0f172a')
    ax.text(50, 134, "Контур противоаварийной защиты давления насосного агрегата НПС (мажоритарная схема 2oo3)",
            ha='center', va='center', fontsize=9, fontstyle='italic', color='#475569')

    # Уровни ответственности (пунктирные рамки с подписями слева)
    ax.add_patch(patches.Rectangle((2, 60), 96, 70, fill=True, facecolor='#eff6ff', edgecolor='#93c5fd', ls='--', lw=1, alpha=0.5))
    ax.text(4, 128, "УРОВЕНЬ ПЛК ПАЗ (цикл 50 мс)", fontsize=8, fontweight='bold', color='#1e40af')

    ax.add_patch(patches.Rectangle((2, 28), 96, 30, fill=True, facecolor='#fffbeb', edgecolor='#fde68a', ls='--', lw=1, alpha=0.5))
    ax.text(4, 56, "УРОВЕНЬ SCADA / HMI ОПЕРАТОРА", fontsize=8, fontweight='bold', color='#92400e')

    ax.add_patch(patches.Rectangle((2, 2), 96, 24, fill=True, facecolor='#f0fdf4', edgecolor='#bbf7d0', ls='--', lw=1, alpha=0.5))
    ax.text(4, 24, "УРОВЕНЬ РЕГЛАМЕНТА И СЛУЖБЫ КИПиА", fontsize=8, fontweight='bold', color='#166534')

    # Блоки алгоритма ПЛК:
    # 1. Начало
    draw_box(50, 126, 24, 4.5, "НАЧАЛО ЦИКЛА ПЛК", shape='oval', bg_color='#bfdbfe', border_color='#2563eb')
    draw_arrow(50, 123.75, 50, 119)

    # 2. Опрос АЦП
    draw_box(50, 116.5, 38, 5, "Опрос токовых каналов АЦП:\nI_A, I_B, I_C (4–20 мА)", shape='io', bg_color='#ffffff')
    draw_arrow(50, 114, 50, 108.5)

    # 3. Ромб: Диагностика NAMUR NE 43
    draw_box(50, 104, 44, 9, "Диагностика токовой петли:\nI < 3.6 мА (Обрыв) ИЛИ\nI > 21.0 мА (КЗ)?",
             shape='rhombus', bg_color='#fef08a')

    # Ветка "НЕТ": Проверка скорости и расхождения
    draw_arrow(50, 99.5, 50, 93, label='Нет')
    draw_box(50, 89.5, 42, 7, "Контроль расхождения в тройке:\n|P_i - P_med| > ΔP_tol (0.4 МПа)\nИЛИ |dP/dt| > V_max?",
             shape='rhombus', bg_color='#fef08a')

    # Ветка "НЕТ" для расхождения: Все исправны -> расчет 2oo3
    draw_arrow(50, 86, 50, 80, label='Нет')
    draw_box(50, 77.5, 38, 5, "Штатный режим 2oo3:\nP_расч = Median(P_A, P_B, P_C);\nАварий нет", shape='rect', bg_color='#dbeafe')

    # Ветка "ДА" из ромба NAMUR или ромба расхождения -> Отказ канала
    draw_arrow(72, 104, 82, 104, label='Да')
    draw_arrow(82, 104, 82, 85)
    draw_arrow(71, 89.5, 82, 89.5, label='Да')

    # Блок локализации и деградации
    draw_box(82, 81.5, 32, 7, "Изоляция неисправного канала:\nБлокировка P_i (BAD_QUALITY);\nИсключение из мажорирования",
             shape='rect', bg_color=c_alarm, border_color='#dc2626')
    draw_arrow(82, 78, 82, 73)

    # Ромб: Количество оставшихся исправных каналов
    draw_box(82, 69, 32, 8, "Число исправных\nканалов N_ok?", shape='rhombus', bg_color='#fed7aa')

    # Ветка N_ok = 2: Режим 1oo2D
    draw_arrow(66, 69, 36, 69, label='N=2')
    draw_box(24, 69, 32, 6, "Реконфигурация в 1oo2D:\nP_сист = (P_ok1 + P_ok2) / 2;\nКонтроль расхождения", shape='rect', bg_color='#fed7aa')

    # Ветка N_ok = 1 или 0: Подключение расчетной модели
    draw_arrow(82, 65, 82, 59, label='N<=1')
    draw_box(82, 55, 32, 8, "Аналитическая избыточность:\nP_сист = P_вход + ΔP(Q, n);\nСравнение с P_model", shape='rect', bg_color='#fecdd3')

    # Переход на уровень SCADA
    draw_arrow(24, 66, 24, 48)
    draw_arrow(82, 51, 82, 48)

    draw_box(50, 44, 76, 7, "УРОВЕНЬ SCADA / HMI:\n• Активация светозвуковой сигнализации (ALARM_SENSOR_FAULT);\n• Индикация неисправного датчика на мнемосхеме (желтый/красный мигающий);\n• Регистрация события в архиве аварий SOE с меткой времени 1 мс",
             shape='rect', bg_color='#fef3c7', border_color='#d97706')

    draw_arrow(50, 40.5, 50, 35)

    # Действия оператора
    draw_box(50, 31.5, 76, 7, "ДЕЙСТВИЯ ДЕЖУРНОГО ОПЕРАТОРА:\n• Квитирование аварийного звукового сигнала;\n• Контроль параметров процесса по оставшимся каналам и расчетной модели;\n• Выдача заявки дежурному инженеру КИПиА на ревизию/замену датчика",
             shape='rect', bg_color='#fef9c3', border_color='#ca8a04')

    draw_arrow(50, 28, 50, 21)

    # Регламент КИПиА
    draw_box(50, 16.5, 76, 8, "ПОРЯДОК ЗАМЕНЫ И ВОССТАНОВЛЕНИЯ ДАТЧИКА (СЛУЖБА КИПиА):\n1. Установка программного шунта/блокировки в SCADA;\n2. Перекрытие коренных вентилей импульсной линии, дренаж, замена прибора;\n3. Опрессовка, подача давления, проверка токовой петли (4..20 мА);\n4. Валидация показаний в течение 5 с, снятие блокировки, безударный возврат в 2oo3",
             shape='rect', bg_color=c_repair, border_color='#16a34a')

    draw_arrow(50, 12.5, 50, 7.5)

    # Конец / Возврат
    draw_box(50, 5.5, 34, 4.5, "КОНЕЦ ЦИКЛА / ВОЗВРАТ", shape='oval', bg_color='#bfdbfe', border_color='#2563eb')

    flowchart_file = os.path.join(OUTPUT_DIR, 'pr3_flowchart_algorithm.png')
    plt.savefig(flowchart_file, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Блок-схема успешно сохранена в: {flowchart_file}")

if __name__ == '__main__':
    run_sensor_simulation()
    generate_flowchart()
