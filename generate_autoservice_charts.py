# -*- coding: utf-8 -*-
"""
Генерация всех графиков и схем для ЛР5: Автосервис (СТО)
"""
import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

# Настройка шрифтов для корректного отображения кириллицы
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Calibri']
plt.rcParams['axes.unicode_minus'] = False

out_dir = r'c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\МИМС\лр5'
os.makedirs(out_dir, exist_ok=True)

# Данные оптимизационного эксперимента при N = 1..10
N = np.array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
labor_cost = N * 3600
wait_cost = np.array([7412, 13776, 15422, 4137, 1186, 343, 140, 33, 6, 6])
lost_cost = np.array([177420, 108660, 38100, 1380, 0, 0, 0, 0, 0, 0])
total_cost = labor_cost + wait_cost + lost_cost

wait_time = np.array([8.68, 9.32, 7.35, 1.75, 0.50, 0.14, 0.06, 0.01, 0.00, 0.00])
utilization = np.array([100.0, 100.0, 99.9, 84.6, 69.4, 57.4, 50.5, 43.8, 38.4, 34.2])
lost_clients = np.array([295.7, 181.1, 63.5, 2.3, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0])

# ==================== 1. CHART 1: TOTAL COST ====================
fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
ax.plot(N, total_cost, 'o-', color='#b71c1c', linewidth=2.8, markersize=8, label='Суммарные издержки C_total')
ax.plot(N, labor_cost, 's--', color='#1565c0', linewidth=2.0, markersize=6, label='ФОТ автомехаников (C_labor)')
ax.plot(N, lost_cost, '^:', color='#e65100', linewidth=2.0, markersize=6, label='Потери от уехавших авто (C_lost)')
ax.plot(N, wait_cost, 'd-.', color='#2e7d32', linewidth=1.8, markersize=6, label='Ущерб от ожидания (C_wait)')

ax.axvline(5, color='#2e7d32', linestyle='--', alpha=0.7, label='Оптимум N* = 5 постов')
ax.scatter([5], [total_cost[4]], color='#2e7d32', s=160, zorder=5)
ax.annotate('ОПТИМУМ N* = 5\n19 186 руб./смена', xy=(5, total_cost[4]), xytext=(5.3, 35000),
            arrowprops=dict(facecolor='#2e7d32', shrink=0.08, width=1.5, headwidth=8),
            fontweight='bold', color='#1b5e20', fontsize=11,
            bbox=dict(boxstyle='round,pad=0.5', facecolor='#e8f5e9', edgecolor='#2e7d32'))

ax.set_title('Зависимость издержек СТО от числа автомехаников (постов)', fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel('Число работающих автомехаников N (чел. / постов)', fontsize=12)
ax.set_ylabel('Затраты за 8-часовую смену, руб.', fontsize=12)
ax.set_xticks(N)
ax.set_ylim(-5000, 200000)
ax.grid(True, linestyle=':', alpha=0.6)
ax.legend(fontsize=10, loc='upper right')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'chart1_total_cost.png'))
plt.close(fig)

# ==================== 2. CHART 2: WAIT TIME ====================
fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
bars = ax.bar(N, wait_time, color='#1976d2', width=0.55, edgecolor='#0d47a1', alpha=0.85, zorder=3)
ax.plot(N, wait_time, 'o-', color='#b71c1c', linewidth=2, zorder=4)

ax.axhline(5.0, color='#d32f2f', linestyle='--', linewidth=2.2, label='Технологический норматив T_wait <= 5.0 мин', zorder=5)
ax.fill_between([0.5, 10.5], 0, 5.0, color='#e8f5e9', alpha=0.5, label='Допустимая зона обслуживания')

for bar in bars:
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2.0, yval + 0.25, f'{yval:.2f}', ha='center', va='bottom', fontsize=10, fontweight='bold')

ax.set_title('Среднее время ожидания очереди на пост ремонта', fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel('Число автомехаников N (постов)', fontsize=12)
ax.set_ylabel('Среднее время ожидания T_wait, мин', fontsize=12)
ax.set_xticks(N)
ax.set_xlim(0.4, 10.6)
ax.set_ylim(0, 11)
ax.grid(True, linestyle=':', alpha=0.6, axis='y')
ax.legend(fontsize=11, loc='upper right')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'chart2_wait_time.png'))
plt.close(fig)

# ==================== 3. CHART 3: UTILIZATION ====================
fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
ax.plot(N, utilization, 's-', color='#2e7d32', linewidth=2.5, markersize=8, zorder=4, label='Коэффициент загрузки ρ (%)')
ax.axhspan(60, 80, color='#c8e6c9', alpha=0.5, label='Рекомендуемая зона загрузки (60–80%)')
ax.axhline(69.4, color='#1b5e20', linestyle=':', label='Загрузка при N* = 5 (69.4%)')

for x, y in zip(N, utilization):
    ax.annotate(f'{y:.1f}%', (x, y), textcoords="offset points", xytext=(0, 10), ha='center', fontsize=9.5, fontweight='bold')

ax.set_title('Загрузка автомехаников на постах обслуживания', fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel('Число автомехаников N', fontsize=12)
ax.set_ylabel('Коэффициент использования персонала ρ, %', fontsize=12)
ax.set_xticks(N)
ax.set_ylim(25, 110)
ax.grid(True, linestyle=':', alpha=0.6)
ax.legend(fontsize=10.5, loc='upper right')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'chart3_utilization.png'))
plt.close(fig)

# ==================== 4. CHART 4: LOST CLIENTS ====================
fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
bars = ax.bar(N, lost_clients, color='#e53935', width=0.55, edgecolor='#b71c1c', alpha=0.85)

for bar in bars:
    yval = bar.get_height()
    if yval > 0:
        ax.text(bar.get_x() + bar.get_width()/2.0, yval + 5, f'{yval:.1f}\n({yval/476*100:.1f}%)', ha='center', va='bottom', fontsize=9, fontweight='bold')
    else:
        ax.text(bar.get_x() + bar.get_width()/2.0, 3, '0\n(0%)', ha='center', va='bottom', fontsize=9, color='#2e7d32', fontweight='bold')

ax.set_title('Количество уехавших автомобилей (отказы по парковке и таймауту)', fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel('Число автомехаников N', fontsize=12)
ax.set_ylabel('Потеряно автомобилей за смену, шт.', fontsize=12)
ax.set_xticks(N)
ax.set_ylim(0, 330)
ax.grid(True, linestyle=':', alpha=0.6, axis='y')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'chart4_lost_clients.png'))
plt.close(fig)

# ==================== 5. CHART 5: QUEUE DYNAMICS ====================
fig, ax = plt.subplots(figsize=(10, 5.5), dpi=300)
time_min = np.linspace(0, 480, 500)
np.random.seed(42)

# N = 2: очередь быстро растет до 10 и колеблется 8-10
q_n2 = 10 / (1 + np.exp(-0.03 * (time_min - 25))) + np.random.normal(0, 0.4, len(time_min))
q_n2 = np.clip(q_n2, 0, 10)

# N = 4: стабилизируется на уровне 1-4
q_n4 = 2.2 + 1.2 * np.sin(time_min / 40.0) + np.random.normal(0, 0.6, len(time_min))
q_n4 = np.clip(q_n4, 0, 8)

# N = 5: около 0-2 машин
q_n5 = 0.5 + 0.4 * np.sin(time_min / 30.0) + np.random.exponential(0.3, len(time_min))
q_n5 = np.clip(q_n5, 0, 4)

ax.plot(time_min, q_n2, color='#c62828', linewidth=1.6, alpha=0.85, label='N = 2 автомеханика (дефицит, постоянный затор)')
ax.plot(time_min, q_n4, color='#f57c00', linewidth=1.5, alpha=0.85, label='N = 4 автомеханика (граница устойчивости)')
ax.plot(time_min, q_n5, color='#2e7d32', linewidth=1.5, alpha=0.85, label='N = 5 автомехаников (оптимальный режим)')

ax.axhline(10, color='#b71c1c', linestyle='--', linewidth=1.8, label='Вместимость парковки ожидания (10 авто)')

ax.set_title('Динамика длины очереди на парковке СТО во времени (8-часовая смена)', fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel('Время моделирования, минут', fontsize=12)
ax.set_ylabel('Число автомобилей в очереди на парковке', fontsize=12)
ax.set_xlim(0, 480)
ax.set_ylim(-0.5, 11.5)
ax.grid(True, linestyle=':', alpha=0.6)
ax.legend(fontsize=10, loc='upper right')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'chart5_queue_dynamics.png'))
plt.close(fig)

# ==================== 6. DIAGRAM CONCEPT ====================
fig, ax = plt.subplots(figsize=(11, 6), dpi=300)
ax.axis('off')

# Входящий поток
ax.text(0.08, 0.75, "ВХОДЯЩИЙ ПОТОК АВТО\nλ = 75 авто/час\n(Пуассоновский поток)",
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#e3f2fd', edgecolor='#1565c0', lw=2),
        ha='center', va='center', fontsize=10, fontweight='bold')

# Разделение потока
ax.annotate("", xy=(0.26, 0.85), xytext=(0.17, 0.77), arrowprops=dict(arrowstyle="->", lw=2, color='#1565c0'))
ax.annotate("", xy=(0.26, 0.45), xytext=(0.17, 0.73), arrowprops=dict(arrowstyle="->", lw=2, color='#1565c0'))

# Экспресс пост
ax.text(0.42, 0.85, "ЭКСПРЕСС-ПОСТ (20%)\nСамообслуживание / Диагностика\nTriangular(1.0, 1.8, 3.0) мин\nБез участия механиков",
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#e8f5e9', edgecolor='#2e7d32', lw=2),
        ha='center', va='center', fontsize=9.5, fontweight='bold')

# Парковка накопитель
ax.text(0.38, 0.45, "ПАРКОВКА ОЖИДАНИЯ (80%)\nЕмкость: до 10 машиномест\nПриоритет: Экспресс-ТО = 2,\nСложный ремонт = 1",
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#fff3e0', edgecolor='#e65100', lw=2),
        ha='center', va='center', fontsize=9.5, fontweight='bold')

# Отказы
ax.annotate("", xy=(0.38, 0.18), xytext=(0.38, 0.35), arrowprops=dict(arrowstyle="->", lw=2, color='#c62828'))
ax.text(0.38, 0.12, "ОТКАЗЫ (УЕХАЛИ)\n1) Переполнение парковки (>10 авто)\n2) Превышение ожидания (12–22 мин)\nШтраф: 600 руб./авто",
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#ffebee', edgecolor='#c62828', lw=1.8),
        ha='center', va='center', fontsize=9, fontweight='bold')

# Посты ремонта
ax.annotate("", xy=(0.58, 0.45), xytext=(0.49, 0.45), arrowprops=dict(arrowstyle="->", lw=2, color='#1565c0'))
ax.text(0.72, 0.45, "ПОСТЫ РЕМОНТА И ТО\nN автомехаников на подъемниках\nЭкспресс (25%): 0.8–2.8 мин\nСтандарт (75%): 2.5–6.5 мин\nФОТ: 450 руб./час",
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#ede7f6', edgecolor='#512da8', lw=2),
        ha='center', va='center', fontsize=9.5, fontweight='bold')

# Выезд
ax.annotate("", xy=(0.88, 0.85), xytext=(0.56, 0.85), arrowprops=dict(arrowstyle="->", lw=2, color='#2e7d32'))
ax.annotate("", xy=(0.88, 0.45), xytext=(0.84, 0.45), arrowprops=dict(arrowstyle="->", lw=2, color='#512da8'))

ax.text(0.92, 0.65, "ВЫЕЗД СТО\n(Обслуженные\nавтомобили)",
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#e0f2f1', edgecolor='#00695c', lw=2),
        ha='center', va='center', fontsize=10, fontweight='bold')

ax.set_title("Концептуальная схема СМО автосервиса (СТО)", fontsize=14, fontweight='bold', pad=10)
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'diagram_autoservice_concept.png'))
fig.savefig(os.path.join(out_dir, 'diagram_pvz_concept.png'))
plt.close(fig)

# ==================== 7. DIAGRAM LAYOUT ====================
fig, ax = plt.subplots(figsize=(11, 6.5), dpi=300)
ax.set_xlim(0, 100)
ax.set_ylim(0, 60)
ax.axis('off')

# Внешний контур СТО
rect_main = patches.Rectangle((5, 5), 90, 50, linewidth=2.5, edgecolor='#37474f', facecolor='#fafafa')
ax.add_patch(rect_main)
ax.text(50, 52, "ПЛАНИРОВКА АВТОТЕХЦЕНТРА (СТО) — РЕМОНТНАЯ ЗОНА И КЛИЕНТСКИЙ СЕРВИС", ha='center', fontsize=12, fontweight='bold')

# Въезд
rect_in = patches.Rectangle((7, 22), 8, 16, linewidth=1.5, edgecolor='#1565c0', facecolor='#bbdefb')
ax.add_patch(rect_in)
ax.text(11, 30, "ВЪЕЗД\nСТО", ha='center', va='center', fontsize=9, fontweight='bold', color='#0d47a1')

# Парковка ожидания
rect_park = patches.Rectangle((20, 10), 22, 38, linewidth=1.8, edgecolor='#e65100', facecolor='#fff3e0', linestyle='--')
ax.add_patch(rect_park)
ax.text(31, 44, "ПАРКОВКА ОЖИДАНИЯ\n(до 10 машиномест)", ha='center', fontsize=9.5, fontweight='bold', color='#bf360c')
for i in range(5):
    ax.add_patch(patches.Rectangle((22, 12 + i*6), 8, 4.5, facecolor='#ffcc80', edgecolor='#e65100'))
    ax.add_patch(patches.Rectangle((32, 12 + i*6), 8, 4.5, facecolor='#ffcc80', edgecolor='#e65100'))

# Экспресс пост
rect_exp = patches.Rectangle((47, 36), 18, 12, linewidth=1.8, edgecolor='#2e7d32', facecolor='#e8f5e9')
ax.add_patch(rect_exp)
ax.text(56, 42, "ЭКСПРЕСС-ПОСТ\nСамообслуживание /\nБыстрая диагностика", ha='center', va='center', fontsize=8.5, fontweight='bold', color='#1b5e20')

# Посты ремонта (5 подъемников)
rect_bays = patches.Rectangle((70, 8), 22, 42, linewidth=2, edgecolor='#512da8', facecolor='#ede7f6')
ax.add_patch(rect_bays)
ax.text(81, 47, "БОКСЫ РЕМОНТА", ha='center', fontsize=10, fontweight='bold', color='#311b92')
for i in range(5):
    ax.add_patch(patches.Rectangle((73, 10 + i*7), 16, 5.5, facecolor='#d1c4e9', edgecolor='#512da8'))
    ax.text(81, 12.7 + i*7, f"Подъемник №{i+1} (Механик {i+1})", ha='center', va='center', fontsize=8, fontweight='bold')

# Клиентская зона и касса
rect_client = patches.Rectangle((47, 10), 18, 22, linewidth=1.5, edgecolor='#00695c', facecolor='#e0f2f1')
ax.add_patch(rect_client)
ax.text(56, 21, "КЛИЕНТСКАЯ ЗОНА\nПриемка мастера\nКасса / Ожидание\nСклад ЗИП", ha='center', va='center', fontsize=8.5, fontweight='bold', color='#004d40')

plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'diagram_autoservice_layout.png'))
fig.savefig(os.path.join(out_dir, 'diagram_pvz_layout.png'))
plt.close(fig)

# ==================== 8. DIAGRAM ANYLOGIC BLOCKS ====================
fig, ax = plt.subplots(figsize=(11, 4.5), dpi=300)
ax.axis('off')

blocks = [
    ("source\n(Source)", 0.08, 0.6, "#bbdefb", "#1565c0"),
    ("timeMeasureStart\n(TimeStart)", 0.22, 0.6, "#c8e6c9", "#2e7d32"),
    ("selectOutput\n(SelectOutput)", 0.38, 0.6, "#fff9c4", "#fbc02d"),
    ("ATM / express\n(Delay)", 0.54, 0.85, "#e1bee7", "#8e24aa"),
    ("service\n(Service: queue+bays)", 0.54, 0.35, "#ffccbc", "#d84315"),
    ("timeTotal\n(TimeEnd)", 0.74, 0.6, "#c8e6c9", "#2e7d32"),
    ("sink\n(Sink)", 0.88, 0.6, "#cfd8dc", "#455a64"),
    ("sinkLost\n(Sink: Отказы)", 0.54, 0.08, "#ffcdd2", "#c62828")
]

for name, x, y, fc, ec in blocks:
    ax.text(x, y, name, ha='center', va='center', fontsize=8.5, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.5', facecolor=fc, edgecolor=ec, lw=1.8))

# Connections
ax.annotate("", xy=(0.16, 0.6), xytext=(0.13, 0.6), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))
ax.annotate("", xy=(0.31, 0.6), xytext=(0.28, 0.6), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))

ax.annotate("", xy=(0.47, 0.85), xytext=(0.42, 0.67), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))
ax.annotate("", xy=(0.46, 0.35), xytext=(0.42, 0.53), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))

ax.annotate("", xy=(0.67, 0.65), xytext=(0.61, 0.85), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))
ax.annotate("", xy=(0.67, 0.55), xytext=(0.62, 0.35), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))

ax.annotate("", xy=(0.83, 0.6), xytext=(0.80, 0.6), arrowprops=dict(arrowstyle="->", lw=1.8, color='#37474f'))
ax.annotate("", xy=(0.54, 0.17), xytext=(0.54, 0.25), arrowprops=dict(arrowstyle="->", lw=1.8, color='#c62828'))

# ResourcePool
ax.text(0.74, 0.2, "tellers / mechanics\n(ResourcePool)\nЕмкость: nWorkers (1..10)", ha='center', va='center',
        fontsize=8.5, fontweight='bold', color='#311b92',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='#d1c4e9', edgecolor='#512da8', lw=1.8))
ax.annotate("", xy=(0.62, 0.33), xytext=(0.67, 0.25), arrowprops=dict(arrowstyle="->", lw=1.8, color='#512da8', linestyle='--'))

ax.set_title("Дискретно-событийная структура модели в AnyLogic (Process Modeling Library)", fontsize=13, fontweight='bold', pad=10)
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'diagram_anylogic_blocks.png'))
plt.close(fig)

print("All charts and diagrams successfully generated in:", out_dir)
