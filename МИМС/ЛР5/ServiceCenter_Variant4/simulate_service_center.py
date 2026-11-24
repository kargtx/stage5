import numpy as np
import matplotlib.pyplot as plt
import os

def simulate_advanced(n_masters, num_shifts=1000):
    np.random.seed(42)
    
    profits = []
    repaired_counts = []
    express_counts = []
    standard_counts = []
    rework_counts = []
    balked_counts = []
    wait_times_express = []
    wait_times_standard = []
    utilizations = []
    
    # 8-часовая смена = 480 минут
    SHIFT_DURATION = 480.0
    
    for _ in range(num_shifts):
        t = np.random.exponential(15.0) # средний интервал прибытия 15 мин (4 заказа в час)
        
        # Очередь: кортежи (arrival_time, is_express, priority)
        # Приоритет express = 10, standard = 1
        queue = []
        
        # Мастера: свободное время каждого мастера
        masters_free = [0.0] * n_masters
        master_busy_time = [0.0] * n_masters
        
        shift_repaired = 0
        shift_express = 0
        shift_standard = 0
        shift_rework = 0
        shift_balked = 0
        shift_wait_exp = []
        shift_wait_std = []
        
        # Список событий или шагов
        # Генерируем все заявки за смену
        arrivals = []
        curr_t = t
        while curr_t < SHIFT_DURATION:
            is_express = (np.random.rand() < 0.25) # 25% срочные заказы
            priority = 10 if is_express else 1
            arrivals.append((curr_t, is_express, priority))
            curr_t += np.random.exponential(15.0)
            
        # Моделируем обслуживание с приоритетной очередью
        # Каждая заявка проходит:
        # 1. Приемка: 3 минуты
        # 2. Очередь на ремонт (с учетом приоритета)
        # 3. Ремонт: треугольное(25, 45, 75)
        # 4. Контроль качества (10% на доработку + 12 минут)
        
        # Для моделирования по событиям отслеживаем доступность мастеров
        # Простая и строгая модель дискретных событий:
        events = [] # (time, event_type, data)
        # event_types: ARRIVAL, REPAIR_FINISH, REWORK_FINISH
        
        waiting_orders = [] # (arrival_time, is_express, priority)
        
        # Имитируем продвижение модельного времени
        cur_time = 0.0
        arr_idx = 0
        
        while arr_idx < len(arrivals) or waiting_orders or any(mf > cur_time for mf in masters_free):
            # Ближайшее время: либо следующее прибытие, либо освобождение мастера
            next_times = []
            if arr_idx < len(arrivals):
                next_times.append(arrivals[arr_idx][0])
            for mf in masters_free:
                if mf > cur_time:
                    next_times.append(mf)
                    
            if not next_times:
                break
                
            cur_time = min(next_times)
            
            # Обработка прибывших заявок к этому моменту
            while arr_idx < len(arrivals) and arrivals[arr_idx][0] <= cur_time:
                arr_t, is_exp, prio = arrivals[arr_idx]
                arr_idx += 1
                
                # Добавляем в очередь после приемки (arr_t + 3.0)
                reception_done_t = arr_t + 3.0
                
                # Проверка ожидания: если в очереди уже слишком много (> 10) или очередь мастеров велика
                # клиент уходит если ожидание > 60 мин (или для срочных > 30 мин)
                max_wait_patience = 30.0 if is_exp else 60.0
                
                # Оценим минимальное время освобождения мастера к моменту reception_done_t
                earliest_master = min(masters_free)
                est_wait = max(0.0, earliest_master - reception_done_t)
                
                if est_wait > max_wait_patience:
                    shift_balked += 1
                else:
                    waiting_orders.append((reception_done_t, is_exp, prio))
                    
            # Сортируем очередь: сначала приоритет (по убыванию: 10 выше 1), затем время прибытия (FIFO)
            waiting_orders.sort(key=lambda x: (-x[2], x[0]))
            
            # Назначение свободных мастеров
            for m_i in range(n_masters):
                if masters_free[m_i] <= cur_time and waiting_orders:
                    ord_arr_t, is_exp, prio = waiting_orders.pop(0)
                    
                    wait_dur = max(0.0, cur_time - ord_arr_t)
                    if is_exp:
                        shift_wait_exp.append(wait_dur)
                        shift_express += 1
                    else:
                        shift_wait_std.append(wait_dur)
                        shift_standard += 1
                        
                    # Время ремонта
                    rep_dur = np.random.triangular(25.0, 45.0, 75.0)
                    # 10% вероятность контроля качества с доработкой (+12 мин)
                    if np.random.rand() < 0.10:
                        rep_dur += np.random.uniform(10.0, 15.0)
                        shift_rework += 1
                        
                    masters_free[m_i] = cur_time + rep_dur
                    master_busy_time[m_i] += rep_dur
                    shift_repaired += 1
                    
        # Итоги смены
        # Доходы: Экспресс = 5000 руб, Стандарт = 3000 руб
        rev = shift_express * 5000.0 + shift_standard * 3000.0
        # Затраты на персонал: мастер 600 руб/час * 8ч = 4800 руб
        # Плюс администратор приемки (1 чел) = 300 руб/час * 8ч = 2400 руб
        labor_cost = n_masters * 4800.0 + 2400.0
        # Штрафы за отказы: упущенная выгода 1500 руб за заказ
        penalty = shift_balked * 1500.0
        net_profit = rev - labor_cost - penalty
        
        profits.append(net_profit)
        repaired_counts.append(shift_repaired)
        express_counts.append(shift_express)
        standard_counts.append(shift_standard)
        rework_counts.append(shift_rework)
        balked_counts.append(shift_balked)
        
        wait_times_express.append(np.mean(shift_wait_exp) if shift_wait_exp else 0.0)
        wait_times_standard.append(np.mean(shift_wait_std) if shift_wait_std else 0.0)
        
        avg_busy = np.mean(master_busy_time) / SHIFT_DURATION
        utilizations.append(min(1.0, avg_busy))
        
    return {
        'profit': np.mean(profits),
        'repaired': np.mean(repaired_counts),
        'express': np.mean(express_counts),
        'standard': np.mean(standard_counts),
        'rework': np.mean(rework_counts),
        'balked': np.mean(balked_counts),
        'wait_exp': np.mean(wait_times_express),
        'wait_std': np.mean(wait_times_standard),
        'utilization': np.mean(utilizations) * 100.0,
        'revenue': np.mean(express_counts) * 5000.0 + np.mean(standard_counts) * 3000.0,
        'cost': n_masters * 4800.0 + 2400.0,
        'penalties': np.mean(balked_counts) * 1500.0
    }

print("=== РАСШИРЕННАЯ ИМИТАЦИОННАЯ МОДЕЛЬ СЕРВИСНОГО ЦЕНТРА ===")
print("Включает: Приемку, Срочные/Обычные заказы с приоритетами, Контроль качества (доработку), Отказы")
print("="*115)
print(f"{'Мастера':<8} | {'Обслужено':<10} | {'Экспресс':<9} | {'Стандарт':<9} | {'Отказов':<8} | {'Загрузка,%':<11} | {'Ожид.Станд':<11} | {'Выручка,руб':<12} | {'Прибыль,руб':<12}")
print("="*115)

results = []
for n in range(1, 9):
    r = simulate_advanced(n)
    results.append(r)
    print(f"{n:<8} | {r['repaired']:<10.1f} | {r['express']:<9.1f} | {r['standard']:<9.1f} | {r['balked']:<8.1f} | {r['utilization']:<11.1f} | {r['wait_std']:<11.1f} | {r['revenue']:<12.0f} | {r['profit']:<12.0f}")

print("="*115)

# Построение расширенного графика оптимизации и метрик (2x2 панели)
fig, axs = plt.subplots(2, 2, figsize=(14, 10))
masters = list(range(1, 9))
profits = [r['profit'] for r in results]
utils = [r['utilization'] for r in results]
repaired = [r['repaired'] for r in results]
balked = [r['balked'] for r in results]
wait_std = [r['wait_std'] for r in results]
wait_exp = [r['wait_exp'] for r in results]

opt_idx = np.argmax(profits)
opt_n = masters[opt_idx]
opt_profit = profits[opt_idx]

# Панель 1: Прибыль
axs[0, 0].plot(masters, profits, 'b-o', linewidth=2.5, markersize=8, color='#1f77b4', label='Чистая прибыль')
axs[0, 0].axvline(x=opt_n, color='red', linestyle='--', linewidth=1.5, label=f'Оптимум N={opt_n} ({opt_profit:.0f} руб)')
axs[0, 0].scatter([opt_n], [opt_profit], color='red', s=120, zorder=5)
axs[0, 0].set_title('1. Целевая функция: Чистая прибыль за смену', fontsize=12, fontweight='bold')
axs[0, 0].set_xlabel('Количество мастеров (чел)')
axs[0, 0].set_ylabel('Прибыль, руб')
axs[0, 0].grid(True, linestyle=':', alpha=0.6)
axs[0, 0].legend()

# Панель 2: Загрузка мастеров
axs[0, 1].plot(masters, utils, 'g-s', linewidth=2.5, markersize=8, color='#2ca02c', label='Коэффициент загрузки')
axs[0, 1].axhline(y=80, color='orange', linestyle=':', label='Рациональная зона (~75-80%)')
axs[0, 1].set_title('2. Загрузка персонала (Утилизация)', fontsize=12, fontweight='bold')
axs[0, 1].set_xlabel('Количество мастеров (чел)')
axs[0, 1].set_ylabel('Загрузка, %')
axs[0, 1].grid(True, linestyle=':', alpha=0.6)
axs[0, 1].legend()

# Панель 3: Обслуженные и отказы
axs[1, 0].bar([m - 0.2 for m in masters], repaired, width=0.4, label='Обслужено заказов', color='#3498db')
axs[1, 0].bar([m + 0.2 for m in masters], balked, width=0.4, label='Отказов (потеряно)', color='#e74c3c')
axs[1, 0].set_title('3. Пропускная способность и отказы', fontsize=12, fontweight='bold')
axs[1, 0].set_xlabel('Количество мастеров (чел)')
axs[1, 0].set_ylabel('Количество заявок за смену, шт')
axs[1, 0].grid(True, linestyle=':', alpha=0.6)
axs[1, 0].legend()

# Панель 4: Время ожидания
axs[1, 1].plot(masters, wait_std, 'o-', linewidth=2, color='#e67e22', label='Обычные заказы')
axs[1, 1].plot(masters, wait_exp, 's--', linewidth=2, color='#9b59b6', label='Срочные заказы (Express)')
axs[1, 1].set_title('4. Среднее время ожидания в очереди', fontsize=12, fontweight='bold')
axs[1, 1].set_xlabel('Количество мастеров (чел)')
axs[1, 1].set_ylabel('Время ожидания, мин')
axs[1, 1].grid(True, linestyle=':', alpha=0.6)
axs[1, 1].legend()

plt.suptitle('Комплексный анализ имитационной модели сервисного центра (Вариант № 4)', fontsize=15, fontweight='bold')
plt.tight_layout()

out_img = 'МИМС/ЛР5/ServiceCenter_Variant4/advanced_analysis_dashboard.png'
plt.savefig(out_img, dpi=300)
print(f"Графика дашборда сохранена в {out_img}")
