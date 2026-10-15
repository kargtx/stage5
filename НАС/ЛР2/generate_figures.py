import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Calibri']
plt.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР2"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. GENERATE IDEF0 A0 DIAGRAM (PERFECTED LAYOUT)
# -------------------------------------------------------------
def generate_idef0_diagram():
    fig, ax = plt.subplots(figsize=(15, 8.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Outer frame
    frame = patches.Rectangle((2, 3), 96, 94, linewidth=1.5, edgecolor='#2c3e50', facecolor='white')
    ax.add_patch(frame)
    
    # Title block header
    ax.text(5, 93.5, "МОДЕЛЬ: Процесс АВПКО АСУ ТП и ПАЗ", fontsize=11, fontweight='bold', color='#1a252f')
    ax.text(50, 93.5, "НОТАЦИЯ: IDEF0 (Функциональная модель)", fontsize=10, color='#34495e')
    ax.text(85, 93.5, "ДИАГРАММА: А0", fontsize=11, fontweight='bold', color='#2980b9')
    ax.plot([2, 98], [91, 91], color='#bdc3c7', linewidth=1)
    
    # Function Blocks
    # Block 1
    b1 = patches.Rectangle((15, 60), 22, 17, linewidth=1.8, edgecolor='#2980b9', facecolor='#ebf5fb')
    ax.add_patch(b1)
    ax.text(26, 69.5, "Структурирование\nи кодирование\nсистемы", fontsize=9.5, fontweight='bold', ha='center', va='center', color='#1b4f72')
    ax.text(35, 62, "А1", fontsize=9, fontweight='bold', color='#2980b9')
    
    # Block 2
    b2 = patches.Rectangle((44, 36), 23, 17, linewidth=1.8, edgecolor='#e67e22', facecolor='#fef5e7')
    ax.add_patch(b2)
    ax.text(55.5, 45.5, "Определение\nпоказателей\nкритичности отказов\n(B1, B2, B3, C)", fontsize=9.5, fontweight='bold', ha='center', va='center', color='#7e5109')
    ax.text(65, 38, "А2", fontsize=9, fontweight='bold', color='#e67e22')
    
    # Block 3
    b3 = patches.Rectangle((74, 14), 21, 16, linewidth=1.8, edgecolor='#27ae60', facecolor='#eafaf1')
    ax.add_patch(b3)
    ax.text(84.5, 23, "Формирование\nперечня критичных\nэлементов и мер", fontsize=9.5, fontweight='bold', ha='center', va='center', color='#145a32')
    ax.text(93, 16, "А3", fontsize=9, fontweight='bold', color='#27ae60')
    
    # Arrows style helper
    arrow_props = dict(arrowstyle="->", lw=1.3, color='#2c3e50')
    
    # Inputs (Left)
    ax.annotate("", xy=(15, 72), xytext=(3, 72), arrowprops=arrow_props)
    ax.text(4, 73.5, "Структура и состав КТС", fontsize=8, color='#2c3e50', fontweight='bold')
    
    ax.annotate("", xy=(15, 64), xytext=(3, 64), arrowprops=arrow_props)
    ax.text(4, 65.5, "Параметры надежности ЛР1", fontsize=8, color='#2c3e50', fontweight='bold')
    
    # Arrow B1 -> B2
    ax.plot([37, 40.5, 40.5], [68.5, 68.5, 48], color='#2c3e50', lw=1.3)
    ax.annotate("", xy=(44, 48), xytext=(40.5, 48), arrowprops=arrow_props)
    ax.text(41, 57, "Иерархический\nсписок LCN", fontsize=8, ha='center', color='#1a252f',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='white', edgecolor='#bdc3c7', lw=0.6))
    
    # Arrow B2 -> B3
    ax.plot([67, 70.5, 70.5], [44.5, 44.5, 24], color='#2c3e50', lw=1.3)
    ax.annotate("", xy=(74, 24), xytext=(70.5, 24), arrowprops=arrow_props)
    ax.text(71, 33, "Показатели\nкритичности C", fontsize=8, ha='center', color='#1a252f',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='white', edgecolor='#bdc3c7', lw=0.6))
    
    # Controls (Top)
    # Control to B1
    ax.plot([26, 26], [91, 77], color='#c0392b', lw=1.3)
    ax.annotate("", xy=(26, 77), xytext=(26, 83), arrowprops=dict(arrowstyle="->", lw=1.3, color='#c0392b'))
    ax.text(26, 84, "ГОСТ 27.310-95\nГОСТ Р МЭК 60812", fontsize=7.8, ha='center', color='#922b21', fontweight='bold')
    
    # Control to B2
    ax.plot([55.5, 55.5], [91, 53], color='#c0392b', lw=1.3)
    ax.annotate("", xy=(55.5, 53), xytext=(55.5, 60), arrowprops=dict(arrowstyle="->", lw=1.3, color='#c0392b'))
    ax.text(55.5, 84, "Классификаторы\nотказов и последствий", fontsize=7.8, ha='center', color='#922b21', fontweight='bold')
    
    # Control to B3 (route around text)
    ax.plot([84.5, 84.5], [91, 30], color='#c0392b', lw=1.3)
    ax.annotate("", xy=(84.5, 30), xytext=(84.5, 37), arrowprops=dict(arrowstyle="->", lw=1.3, color='#c0392b'))
    ax.text(84.5, 84, "Критерии допустимого\nриска (SIL 3, RPN>=100)", fontsize=7.8, ha='center', color='#922b21', fontweight='bold')
    
    # Mechanisms (Bottom)
    ax.plot([26, 26], [9, 60], color='#27ae60', lw=1.3)
    ax.annotate("", xy=(26, 60), xytext=(26, 52), arrowprops=dict(arrowstyle="->", lw=1.3, color='#27ae60'))
    ax.text(26, 12, "Конструктор АСУ ТП", fontsize=8, ha='center', color='#145a32', fontweight='bold',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))
    
    ax.plot([55.5, 55.5], [9, 36], color='#27ae60', lw=1.3)
    ax.annotate("", xy=(55.5, 36), xytext=(55.5, 29), arrowprops=dict(arrowstyle="->", lw=1.3, color='#27ae60'))
    ax.text(55.5, 12, "Инженер по надежности", fontsize=8, ha='center', color='#145a32', fontweight='bold',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))
    
    ax.plot([84.5, 84.5], [9, 14], color='#27ae60', lw=1.3)
    ax.annotate("", xy=(84.5, 14), xytext=(84.5, 10), arrowprops=dict(arrowstyle="->", lw=1.3, color='#27ae60'))
    ax.text(84.5, 9.5, "Эксперт по безопасности ПАЗ", fontsize=7.5, ha='center', color='#145a32', fontweight='bold',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))
    
    # Outputs (Right)
    ax.annotate("", xy=(97.5, 22), xytext=(95, 22), arrowprops=arrow_props)
    ax.text(96, 17, "Перечень критичных элементов\nи проектные рекомендации", fontsize=7.5, ha='right', color='#2c3e50', fontweight='bold')
    
    ax.plot([67, 95], [39, 39], color='#2c3e50', lw=1.3)
    ax.annotate("", xy=(97.5, 39), xytext=(95, 39), arrowprops=arrow_props)
    ax.text(96, 41, "Таблица АВПКО (FMECA)", fontsize=7.5, ha='right', color='#2c3e50', fontweight='bold')
    
    # Footer separator
    ax.plot([2, 98], [8, 8], color='#bdc3c7', linewidth=1)
    ax.text(5, 5, "РАЗРАБОТАЛИ: гр. ЭАС-414С Гайнанов Д.И., Гарифуллин К.Р.", fontsize=8.5, color='#2c3e50')
    ax.text(60, 5, "ПРОВЕРИЛ: доц. Ризванов К.А.", fontsize=8.5, color='#2c3e50')
    
    plt.tight_layout()
    output_path = os.path.join(OUTPUT_DIR, "idef0_a0_fmeca.png")
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Generated:", output_path)

# -------------------------------------------------------------
# 2. GENERATE FAULT TREE (LOGICAL STRUCTURE - PERFECT SPACING)
# -------------------------------------------------------------
def generate_fault_tree_logical():
    fig, ax = plt.subplots(figsize=(20, 12), dpi=300)
    ax.set_xlim(0, 190)
    ax.set_ylim(0, 115)
    ax.axis('off')
    
    def draw_event_box(x, y, w, h, text, code, color='#2c3e50', bg='#f8f9f9', is_top=False):
        lw = 2.4 if is_top else 1.4
        box = patches.FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.5,rounding_size=1.5",
                                     linewidth=lw, edgecolor=color, facecolor=bg)
        ax.add_patch(box)
        if code:
            ax.text(x, y + h/2 - 2.2, f"[{code}]", fontsize=8.5, fontweight='bold', ha='center', color=color)
            ax.text(x, y - 1.2, text, fontsize=7.8, ha='center', va='center', color='#1a252f')
        else:
            ax.text(x, y, text, fontsize=8, ha='center', va='center', color='#1a252f')
            
    def draw_basic_event(x, y, r, code, text, color='#2980b9', bg='#ebf5fb'):
        circle = patches.Circle((x, y), r, linewidth=1.4, edgecolor=color, facecolor=bg)
        ax.add_patch(circle)
        ax.text(x, y, code, fontsize=8, fontweight='bold', ha='center', va='center', color=color)
        ax.text(x, y - r - 2.2, text, fontsize=7.2, ha='center', va='top', color='#2c3e50')
        
    def draw_undeveloped_event(x, y, size, code, text, color='#d35400', bg='#fef5e7'):
        pts = [[x, y + size], [x + size*1.2, y], [x, y - size], [x - size*1.2, y]]
        poly = patches.Polygon(pts, linewidth=1.4, edgecolor=color, facecolor=bg)
        ax.add_patch(poly)
        ax.text(x, y, code, fontsize=7.5, fontweight='bold', ha='center', va='center', color=color)
        ax.text(x, y - size - 2.2, text, fontsize=7.2, ha='center', va='top', color='#2c3e50')

    def draw_gate(x, y, gate_type="AND"):
        if gate_type == "AND":
            gate = patches.FancyBboxPatch((x - 2.5, y - 2.5), 5, 5, boxstyle="round,pad=0.2,rounding_size=2",
                                         linewidth=1.5, edgecolor='#c0392b', facecolor='#fadbd8')
            ax.add_patch(gate)
            ax.text(x, y - 0.2, "&", fontsize=11, fontweight='bold', ha='center', va='center', color='#922b21')
        elif gate_type == "OR":
            gate = patches.FancyBboxPatch((x - 2.5, y - 2.5), 5, 5, boxstyle="round,pad=0.2,rounding_size=1",
                                         linewidth=1.5, edgecolor='#2980b9', facecolor='#d4e6f1')
            ax.add_patch(gate)
            ax.text(x, y - 0.2, "≥1", fontsize=10, fontweight='bold', ha='center', va='center', color='#1b4f72')
        elif gate_type == "2oo3":
            gate = patches.FancyBboxPatch((x - 3.2, y - 2.5), 6.4, 5, boxstyle="round,pad=0.2,rounding_size=1",
                                         linewidth=1.5, edgecolor='#8e44ad', facecolor='#ebdef0')
            ax.add_patch(gate)
            ax.text(x, y - 0.2, "2 / 3", fontsize=9, fontweight='bold', ha='center', va='center', color='#512e5f')

    def draw_vline(x, y1, y2, color='#34495e'):
        ax.plot([x, x], [y1, y2], color=color, lw=1.2)
    def draw_hline(x1, x2, y, color='#34495e'):
        ax.plot([x1, x2], [y, y], color=color, lw=1.2)

    # 1. TOP EVENT
    draw_event_box(95, 102, 54, 13, "АВАРИЙНАЯ РАЗГЕРМЕТИЗАЦИЯ КОЛОННЫ\nИ ВЫБРОС В УСЛОВИЯХ ПРЕВЫШЕНИЯ ДАВЛЕНИЯ",
                   "ГОЛОВНОЕ СОБЫТИЕ Т", color='#c0392b', bg='#fdedec', is_top=True)
    
    draw_vline(95, 95.5, 89.5)
    draw_gate(95, 87, "AND")
    draw_vline(95, 84.5, 78)
    draw_hline(42, 142, 78)

    # 2. LEVEL 2: INITIATING EVENT & SAFETY FAILURE
    # Left Branch: Initiating Overpressure (РСУ failure)
    draw_vline(42, 78, 72)
    draw_event_box(42, 66, 40, 10, "Возникновение аварийного роста\nдавления в колонне (P > 3,0 МПа)", "E_иниц", color='#d35400', bg='#fef5e7')
    
    # Right Branch: Protection Failure (ПАЗ failure)
    draw_vline(142, 78, 72)
    draw_event_box(142, 66, 40, 10, "Отказ защитных действий системы ПАЗ\n(непрохождение защитного останова)", "E_защита", color='#8e44ad', bg='#f5eef8')

    # Sub-gate for Initiating Event: OR
    draw_vline(42, 61, 56.5)
    draw_gate(42, 54, "OR")
    draw_vline(42, 51.5, 46)
    draw_hline(18, 66, 46)

    # Sub-events for E_иниц:
    # A1: Отказ контура регулирования давления РСУ
    draw_vline(18, 46, 41)
    draw_event_box(18, 36, 26, 9, "Отказ контура регулирования\nдавления РСУ (сдувка)", "A1", color='#2980b9', bg='#ebf5fb')
    
    # A2: Прекращение орошения колонны
    draw_vline(42, 46, 41)
    draw_event_box(42, 36, 18, 9, "Прекращение\nподачи орошения", "A2", color='#2980b9', bg='#ebf5fb')

    # A3: Неконтролируемый перегрев куба
    draw_vline(66, 46, 41)
    draw_event_box(66, 36, 20, 9, "Перегрев куба\n(отказ подогрева)", "A3", color='#2980b9', bg='#ebf5fb')

    # A1 breakdown: OR gate
    draw_vline(18, 31.5, 26.5)
    draw_gate(18, 24, "OR")
    draw_vline(18, 21.5, 17)
    draw_hline(4.5, 31.5, 17)

    # Basic events under A1:
    draw_vline(4.5, 17, 12)
    draw_basic_event(4.5, 9.5, 2.5, "X1", "Дрейф датчика\nдавления РСУ")

    draw_vline(13.5, 17, 12)
    draw_basic_event(13.5, 9.5, 2.5, "X2", "Сбой ЦПУ\nПЛК РСУ")

    draw_vline(22.5, 17, 12)
    draw_basic_event(22.5, 9.5, 2.5, "X3", "Заклинивание\nрег. клапана")

    draw_vline(31.5, 17, 12)
    draw_basic_event(31.5, 9.5, 2.5, "X4", "Обрыв шины\nсвязи РСУ")

    # Undeveloped under A2 and A3:
    draw_vline(42, 31.5, 12)
    draw_undeveloped_event(42, 9.5, 2.5, "U1", "Отказ насоса\nфлегмы")

    draw_vline(66, 31.5, 12)
    draw_undeveloped_event(66, 9.5, 2.5, "U2", "Сбой подачи\nтеплоносителя")

    # ----------------------------------------------------
    # Sub-gate for Safety Failure E_защита: OR
    draw_vline(142, 61, 56.5)
    draw_gate(142, 54, "OR")
    draw_vline(142, 51.5, 46)
    draw_hline(95, 177, 46)

    # 4 Sub-systems of ПАЗ:
    # B1: Измерительный контур ПАЗ (2oo3)
    draw_vline(95, 46, 41)
    draw_event_box(95, 36, 22, 9, "Отказ измерения\nдавления ПАЗ (2oo3)", "B1", color='#8e44ad', bg='#f5eef8')

    # B2: Контроллер ПАЗ SIL3
    draw_vline(123, 46, 41)
    draw_event_box(123, 36, 22, 9, "Отказ логического\nконтроллера ПАЗ", "B2", color='#8e44ad', bg='#f5eef8')

    # B3: Отсечной клапан ПАЗ
    draw_vline(151, 46, 41)
    draw_event_box(151, 36, 24, 9, "Несрабатывание\nотсечного клапана", "B3", color='#8e44ad', bg='#f5eef8')

    # B4: Электропитание
    draw_vline(177, 46, 41)
    draw_event_box(177, 36, 20, 9, "Полное обесточ.\nцепей ПАЗ", "B4", color='#8e44ad', bg='#f5eef8')

    # B1 (2oo3 gate)
    draw_vline(95, 31.5, 26.5)
    draw_gate(95, 24, "2oo3")
    draw_vline(95, 21.5, 17)
    draw_hline(85, 105, 17)
    draw_vline(85, 17, 12)
    draw_basic_event(85, 9.5, 2.5, "X5", "Отказ датчика\nPT-102A")
    draw_vline(95, 17, 12)
    draw_basic_event(95, 9.5, 2.5, "X6", "Отказ датчика\nPT-102B")
    draw_vline(105, 17, 12)
    draw_basic_event(105, 9.5, 2.5, "X7", "Отказ датчика\nPT-102C")

    # B2 (Контроллер ПАЗ: OR)
    draw_vline(123, 31.5, 26.5)
    draw_gate(123, 24, "OR")
    draw_vline(123, 21.5, 17)
    draw_hline(115, 131, 17)
    draw_vline(115, 17, 12)
    draw_undeveloped_event(115, 9.5, 2.5, "X8", "Общий сбой ПО\nПАЗ (CCF)")
    draw_vline(131, 17, 14)
    draw_gate(131, 11.5, "AND")
    draw_vline(131, 9, 6)
    draw_hline(126, 136, 6)
    draw_vline(126, 6, 4.5)
    draw_basic_event(126, 2.5, 1.8, "X9", "ЦПУ-А")
    draw_vline(136, 6, 4.5)
    draw_basic_event(136, 2.5, 1.8, "X10", "ЦПУ-B")

    # B3 (Отсечной клапан: OR)
    draw_vline(151, 31.5, 26.5)
    draw_gate(151, 24, "OR")
    draw_vline(151, 21.5, 17)
    draw_hline(142, 160, 17)
    draw_vline(142, 17, 12)
    draw_basic_event(142, 9.5, 2.5, "X11", "Заклинивание\nотсекателя")
    draw_vline(151, 17, 12)
    draw_basic_event(151, 9.5, 2.5, "X12", "Отказ соленоида\nпневмопривода")
    draw_vline(160, 17, 12)
    draw_undeveloped_event(160, 9.5, 2.5, "X13", "Засорение линии\nвоздуха КИП")

    # B4 (Питание: AND)
    draw_vline(177, 31.5, 26.5)
    draw_gate(177, 24, "AND")
    draw_vline(177, 21.5, 17)
    draw_hline(171, 183, 17)
    draw_vline(171, 17, 12)
    draw_basic_event(171, 9.5, 2.5, "X14", "Отказ ввода\nсети 220В")
    draw_vline(183, 17, 12)
    draw_basic_event(183, 9.5, 2.5, "X15", "Отказ ИБП /\nразряд АКБ")

    # Title & Legend Box
    title_box = patches.Rectangle((4, 98), 50, 12, linewidth=1, edgecolor='#7f8c8d', facecolor='#f4f6f6')
    ax.add_patch(title_box)
    ax.text(6, 106, "ДЕРЕВО ОТКАЗОВ (FTA)", fontsize=11, fontweight='bold', color='#1a252f')
    ax.text(6, 102.5, "Объект: Блок ректификации ЭЛОУ-АВТ", fontsize=8.5, color='#34495e')
    ax.text(6, 99.5, "Функция безопасности: Защита от избыточного давления", fontsize=8, color='#7f8c8d')

    plt.tight_layout()
    output_path = os.path.join(OUTPUT_DIR, "fault_tree_logical.png")
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Generated:", output_path)

# -------------------------------------------------------------
# 3. GENERATE FAULT TREE WITH PROBABILITIES (PERFECT SPACING)
# -------------------------------------------------------------
def generate_fault_tree_probabilities():
    fig, ax = plt.subplots(figsize=(20, 12), dpi=300)
    ax.set_xlim(0, 190)
    ax.set_ylim(0, 115)
    ax.axis('off')
    
    def draw_prob_box(x, y, w, h, code, text, prob_str, color='#2c3e50', bg='#f8f9f9', is_top=False):
        lw = 2.4 if is_top else 1.4
        box = patches.FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.5,rounding_size=1.5",
                                     linewidth=lw, edgecolor=color, facecolor=bg)
        ax.add_patch(box)
        ax.text(x, y + h/2 - 2.2, f"[{code}]", fontsize=8.5, fontweight='bold', ha='center', color=color)
        ax.text(x, y - 0.2, text, fontsize=7.5, ha='center', va='center', color='#1a252f')
        ax.text(x, y - h/2 + 2, f"P = {prob_str}", fontsize=8.5, fontweight='bold', ha='center', color='#c0392b' if is_top else '#1b4f72')

    def draw_prob_circle(x, y, r, code, text, prob_str, color='#2980b9', bg='#ebf5fb'):
        circle = patches.Circle((x, y), r, linewidth=1.4, edgecolor=color, facecolor=bg)
        ax.add_patch(circle)
        ax.text(x, y + 0.9, code, fontsize=8, fontweight='bold', ha='center', color=color)
        ax.text(x, y - 1.1, f"P={prob_str}", fontsize=6.8, fontweight='bold', ha='center', color='#922b21')
        ax.text(x, y - r - 2.2, text, fontsize=7.2, ha='center', va='top', color='#2c3e50')
        
    def draw_prob_diamond(x, y, size, code, text, prob_str, color='#d35400', bg='#fef5e7'):
        pts = [[x, y + size], [x + size*1.2, y], [x, y - size], [x - size*1.2, y]]
        poly = patches.Polygon(pts, linewidth=1.4, edgecolor=color, facecolor=bg)
        ax.add_patch(poly)
        ax.text(x, y + 0.9, code, fontsize=7.5, fontweight='bold', ha='center', color=color)
        ax.text(x, y - 1.1, f"P={prob_str}", fontsize=6.5, fontweight='bold', ha='center', color='#922b21')
        ax.text(x, y - size - 2.2, text, fontsize=7.2, ha='center', va='top', color='#2c3e50')

    def draw_gate(x, y, gate_type="AND"):
        if gate_type == "AND":
            gate = patches.FancyBboxPatch((x - 2.5, y - 2.5), 5, 5, boxstyle="round,pad=0.2,rounding_size=2",
                                         linewidth=1.5, edgecolor='#c0392b', facecolor='#fadbd8')
            ax.add_patch(gate)
            ax.text(x, y - 0.2, "&", fontsize=11, fontweight='bold', ha='center', va='center', color='#922b21')
        elif gate_type == "OR":
            gate = patches.FancyBboxPatch((x - 2.5, y - 2.5), 5, 5, boxstyle="round,pad=0.2,rounding_size=1",
                                         linewidth=1.5, edgecolor='#2980b9', facecolor='#d4e6f1')
            ax.add_patch(gate)
            ax.text(x, y - 0.2, "≥1", fontsize=10, fontweight='bold', ha='center', va='center', color='#1b4f72')
        elif gate_type == "2oo3":
            gate = patches.FancyBboxPatch((x - 3.2, y - 2.5), 6.4, 5, boxstyle="round,pad=0.2,rounding_size=1",
                                         linewidth=1.5, edgecolor='#8e44ad', facecolor='#ebdef0')
            ax.add_patch(gate)
            ax.text(x, y - 0.2, "2 / 3", fontsize=9, fontweight='bold', ha='center', va='center', color='#512e5f')

    def draw_vline(x, y1, y2, color='#34495e'):
        ax.plot([x, x], [y1, y2], color=color, lw=1.2)
    def draw_hline(x1, x2, y, color='#34495e'):
        ax.plot([x1, x2], [y, y], color=color, lw=1.2)

    # 1. TOP EVENT
    draw_prob_box(95, 102, 54, 14, "ГОЛОВНОЕ СОБЫТИЕ Т",
                  "Аварийный выброс углеводородов из колонны\nпри превышении критического давления",
                  "1,33 · 10⁻³ (0,00133)", color='#c0392b', bg='#fdedec', is_top=True)
    
    draw_vline(95, 95, 89.5)
    draw_gate(95, 87, "AND")
    draw_vline(95, 84.5, 78)
    draw_hline(42, 142, 78)

    # Level 2
    # Initiating Event: E_иниц
    draw_vline(42, 78, 72)
    draw_prob_box(42, 66, 40, 11, "E_иниц", "Аварийный рост давления\nв колонне (P > 3,0 МПа)", "0,0638 (6,38%)", color='#d35400', bg='#fef5e7')

    # Safety Failure: E_защита
    draw_vline(142, 78, 72)
    draw_prob_box(142, 66, 40, 11, "E_защита", "Отказ защитных действий\nсистемы ПАЗ", "0,02087 (2,09%)", color='#8e44ad', bg='#f5eef8')

    # Initiating Event OR Gate
    draw_vline(42, 60.5, 56.5)
    draw_gate(42, 54, "OR")
    draw_vline(42, 51.5, 46)
    draw_hline(18, 66, 46)

    # A1, A2, A3
    draw_vline(18, 46, 41)
    draw_prob_box(18, 36, 26, 10, "A1", "Отказ контура рег.\nдавления РСУ", "0,0638", color='#2980b9', bg='#ebf5fb')
    
    draw_vline(42, 46, 41)
    draw_prob_box(42, 36, 18, 10, "A2", "Отказ подачи\nорошения", "0,018", color='#2980b9', bg='#ebf5fb')

    draw_vline(66, 46, 41)
    draw_prob_box(66, 36, 20, 10, "A3", "Перегрев куба\nколонны", "0,012", color='#2980b9', bg='#ebf5fb')

    # A1 breakdown: OR
    draw_vline(18, 31, 26.5)
    draw_gate(18, 24, "OR")
    draw_vline(18, 21.5, 17)
    draw_hline(4.5, 31.5, 17)

    draw_vline(4.5, 17, 12)
    draw_prob_circle(4.5, 9.5, 2.5, "X1", "Дрейф датчика\nРСУ", "0,025")

    draw_vline(13.5, 17, 12)
    draw_prob_circle(13.5, 9.5, 2.5, "X2", "Сбой ЦПУ\nРСУ", "0,001")

    draw_vline(22.5, 17, 12)
    draw_prob_circle(22.5, 9.5, 2.5, "X3", "Заклинивание\nклапана РСУ", "0,035")

    draw_vline(31.5, 17, 12)
    draw_prob_circle(31.5, 9.5, 2.5, "X4", "Обрыв шины\nРСУ", "0,004")

    draw_vline(42, 31, 12)
    draw_prob_diamond(42, 9.5, 2.5, "U1", "Отказ насоса\nфлегмы", "0,018")

    draw_vline(66, 31, 12)
    draw_prob_diamond(66, 9.5, 2.5, "U2", "Сбой подогрева\nкуба", "0,012")

    # Safety Failure Breakdown: OR
    draw_vline(142, 60.5, 56.5)
    draw_gate(142, 54, "OR")
    draw_vline(142, 51.5, 46)
    draw_hline(95, 177, 46)

    # B1, B2, B3, B4
    draw_vline(95, 46, 41)
    draw_prob_box(95, 36, 22, 10, "B1", "Отказ датчиков\nПАЗ (2oo3)", "6,68 · 10⁻⁴", color='#8e44ad', bg='#f5eef8')

    draw_vline(123, 46, 41)
    draw_prob_box(123, 36, 22, 10, "B2", "Отказ контроллера\nПАЗ SIL3", "1,04 · 10⁻⁴", color='#8e44ad', bg='#f5eef8')

    draw_vline(151, 46, 41)
    draw_prob_box(151, 36, 24, 10, "B3", "Отказ отсекателя\nПАЗ Fail-Safe", "0,01988", color='#8e44ad', bg='#f5eef8')

    draw_vline(177, 46, 41)
    draw_prob_box(177, 36, 20, 10, "B4", "Отказ питания\n(сеть + ИБП)", "2,50 · 10⁻⁴", color='#8e44ad', bg='#f5eef8')

    # B1 2oo3 gate
    draw_vline(95, 31, 26.5)
    draw_gate(95, 24, "2oo3")
    draw_vline(95, 21.5, 17)
    draw_hline(85, 105, 17)
    draw_vline(85, 17, 12)
    draw_prob_circle(85, 9.5, 2.5, "X5", "PT-102A", "0,015")
    draw_vline(95, 17, 12)
    draw_prob_circle(95, 9.5, 2.5, "X6", "PT-102B", "0,015")
    draw_vline(105, 17, 12)
    draw_prob_circle(105, 9.5, 2.5, "X7", "PT-102C", "0,015")

    # B2 OR
    draw_vline(123, 31, 26.5)
    draw_gate(123, 24, "OR")
    draw_vline(123, 21.5, 17)
    draw_hline(115, 131, 17)
    draw_vline(115, 17, 12)
    draw_prob_diamond(115, 9.5, 2.5, "X8", "Сбой ПО (CCF)", "1,0 · 10⁻⁴")
    draw_vline(131, 17, 14)
    draw_gate(131, 11.5, "AND")
    draw_vline(131, 9, 6)
    draw_hline(126, 136, 6)
    draw_vline(126, 6, 4.5)
    draw_prob_circle(126, 2.5, 1.8, "X9", "ЦПУ-А", "0,002")
    draw_vline(136, 6, 4.5)
    draw_prob_circle(136, 2.5, 1.8, "X10", "ЦПУ-B", "0,002")

    # B3 OR
    draw_vline(151, 31, 26.5)
    draw_gate(151, 24, "OR")
    draw_vline(151, 21.5, 17)
    draw_hline(142, 160, 17)
    draw_vline(142, 17, 12)
    draw_prob_circle(142, 9.5, 2.5, "X11", "Заклин. клапана", "0,012")
    draw_vline(151, 17, 12)
    draw_prob_circle(151, 9.5, 2.5, "X12", "Отказ соленоида", "0,005")
    draw_vline(160, 17, 12)
    draw_prob_diamond(160, 9.5, 2.5, "X13", "Линия воздуха", "0,003")

    # B4 AND
    draw_vline(177, 31, 26.5)
    draw_gate(177, 24, "AND")
    draw_vline(177, 21.5, 17)
    draw_hline(171, 183, 17)
    draw_vline(171, 17, 12)
    draw_prob_circle(171, 9.5, 2.5, "X14", "Сбой сети 220В", "0,050")
    draw_vline(183, 17, 12)
    draw_prob_circle(183, 9.5, 2.5, "X15", "Отказ ИБП", "0,005")

    # Formula Calculation Legend Box
    calc_box = patches.Rectangle((4, 95), 52, 14, linewidth=1, edgecolor='#2980b9', facecolor='#ebf5fb')
    ax.add_patch(calc_box)
    ax.text(6, 106, "КОЛИЧЕСТВЕННЫЙ РАСЧЕТ ВЕРОЯТНОСТИ:", fontsize=9, fontweight='bold', color='#1b4f72')
    ax.text(6, 103, "• P(T) = P(E_иниц) · P(E_защита)  [операция И]", fontsize=8.2, color='#2c3e50')
    ax.text(6, 100.5, "• P(E_иниц) ≈ 0,0638 / год (~6,38 %)", fontsize=8.2, color='#2c3e50')
    ax.text(6, 98, "• P(E_защита) ≈ 0,02087 / требование (~2,09 %)", fontsize=8.2, color='#2c3e50')
    ax.text(6, 95.5, "• P(T) = 0,0638 · 0,02087 = 1,33 · 10⁻³ / год (SIL 3)", fontsize=8.5, fontweight='bold', color='#c0392b')

    plt.tight_layout()
    output_path = os.path.join(OUTPUT_DIR, "fault_tree_probabilities.png")
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Generated:", output_path)

# -------------------------------------------------------------
# 4. GENERATE FMECA PARETO & RISK MATRIX CHARTS
# -------------------------------------------------------------
def generate_fmeca_charts():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7), dpi=300)
    
    # 1. RISK MATRIX (B1 vs B2)
    x = np.linspace(1, 10, 100)
    y = np.linspace(1, 10, 100)
    X, Y = np.meshgrid(x, y)
    Z = X * Y # Severity * Occurrence
    
    cmap = matplotlib.colors.LinearSegmentedColormap.from_list("risk_cmap", ["#2ecc71", "#f1c40f", "#e67e22", "#e74c3c"])
    ax1.imshow(Z, origin='lower', extent=[1, 10, 1, 10], aspect='auto', cmap=cmap, alpha=0.55)
    
    ax1.set_xticks(range(1, 11))
    ax1.set_yticks(range(1, 11))
    ax1.grid(color='white', linestyle='--', linewidth=0.8, alpha=0.7)
    
    elements = [
        ("LCN 1.1", "ЦПУ РСУ (1oo2D)", 3, 7, 2, 42),
        ("LCN 1.2", "ЦПУ ПАЗ (SIL3)", 2, 9, 1, 18),
        ("LCN 2.1", "Датчик давления PT-101", 4, 8, 4, 128),
        ("LCN 2.2", "Датчики ПАЗ (2oo3)", 2, 10, 2, 40),
        ("LCN 3.1", "Отсечной клапан XV-101", 4, 10, 5, 200),
        ("LCN 3.2", "Соленоид пневмопривода", 3, 4, 1, 12),
        ("LCN 4.1", "Регулирующий клапан FV", 3, 7, 2, 42),
        ("LCN 5.1", "Коммутатор кольца Ethernet", 3, 3, 1, 9),
        ("LCN 6.1", "Сервер SCADA", 3, 4, 1, 12),
        ("LCN 7.1", "ИБП 24В / Батареи", 2, 6, 2, 24),
        ("LCN 8.1", "Модуль AI аналог. ввода", 2, 5, 1, 10),
        ("LCN 9.1", "Оператор АРМ (ошибка)", 3, 6, 2, 36)
    ]
    
    for lcn, name, b1, b2, b3, c in elements:
        pt_color = '#900c3f' if c >= 100 else ('#d35400' if c >= 40 else '#1e8449')
        ax1.scatter(b1, b2, s=140, color=pt_color, edgecolors='black', linewidth=1.2, zorder=5)
        offset_x = 0.15 if b1 <= 8 else -0.8
        offset_y = 0.15 if b2 <= 8 else -0.4
        ax1.annotate(f"{lcn}\n(C={c})", xy=(b1, b2), xytext=(b1 + offset_x, b2 + offset_y),
                     fontsize=8, fontweight='bold', color='#1a252f', zorder=6,
                     bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.8, edgecolor='#bdc3c7', lw=0.6))
        
    ax1.set_xlabel("Вероятность возникновения отказа, B1 (1 .. 10)", fontsize=10, fontweight='bold', color='#2c3e50')
    ax1.set_ylabel("Тяжесть последствий отказа, B2 (1 .. 10)", fontsize=10, fontweight='bold', color='#2c3e50')
    ax1.set_title("Матрица критичности отказов (FMECA Risk Matrix)", fontsize=12, fontweight='bold', color='#1a252f', pad=12)
    
    ax1.text(2, 2, "ЗЕЛЕНАЯ ЗОНА\n(Приемлемый риск)", fontsize=9, fontweight='bold', color='#145a32', ha='center',
             bbox=dict(boxstyle="round,pad=0.3", facecolor='#abebc6', alpha=0.7, edgecolor='none'))
    ax1.text(8.5, 9.2, "КРАСНАЯ ЗОНА\n(Критический риск)", fontsize=9, fontweight='bold', color='#78281f', ha='center',
             bbox=dict(boxstyle="round,pad=0.3", facecolor='#f5b7b1', alpha=0.8, edgecolor='none'))

    # 2. PARETO CHART FOR CRITICALITY C = B1 * B2 * B3
    sorted_elements = sorted(elements, key=lambda x: x[5], reverse=True)
    labels = [f"{x[0]}\n{x[1][:12]}.." for x in sorted_elements]
    c_values = [x[5] for x in sorted_elements]
    total_c = sum(c_values)
    cum_pct = np.cumsum(c_values) / total_c * 100
    
    x_pos = np.arange(len(labels))
    bars = ax2.bar(x_pos, c_values, color=['#c0392b' if c >= 100 else ('#e67e22' if c >= 40 else '#27ae60') for c in c_values],
                   edgecolor='#2c3e50', linewidth=1, width=0.6, zorder=3)
    
    ax2.set_ylabel("Показатель критичности отказа (RPN / C)", fontsize=10, fontweight='bold', color='#2c3e50')
    ax2.set_xlabel("Компоненты КТС и ПО АС", fontsize=10, fontweight='bold', color='#2c3e50')
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
    ax2.grid(axis='y', linestyle='--', alpha=0.6, zorder=0)
    
    ax2.axhline(100, color='#900c3f', linestyle='--', linewidth=1.5, label='Критический порог C = 100')
    
    ax2_twin = ax2.twinx()
    ax2_twin.plot(x_pos, cum_pct, color='#2980b9', marker='o', linewidth=2.2, markersize=5, label='Кумулятивная доля риска (%)')
    ax2_twin.set_ylabel("Кумулятивная доля риска, %", fontsize=10, fontweight='bold', color='#1b4f72')
    ax2_twin.set_ylim(0, 110)
    ax2_twin.axhline(80, color='#7f8c8d', linestyle=':', linewidth=1.2, label='Порог Парето 80%')
    
    for bar in bars:
        h = bar.get_height()
        ax2.annotate(f"{int(h)}", xy=(bar.get_x() + bar.get_width() / 2, h),
                     xytext=(0, 3), textcoords="offset points", ha='center', va='bottom',
                     fontsize=8, fontweight='bold')
        
    ax2.set_title("Диаграмма Парето критичности отказов АВПКО", fontsize=12, fontweight='bold', color='#1a252f', pad=12)
    ax2.legend(loc='upper right', bbox_to_anchor=(0.7, 0.98), fontsize=8.5)
    ax2_twin.legend(loc='upper right', bbox_to_anchor=(0.99, 0.98), fontsize=8.5)

    plt.tight_layout()
    output_path = os.path.join(OUTPUT_DIR, "pareto_risk_chart.png")
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Generated:", output_path)

if __name__ == "__main__":
    generate_idef0_diagram()
    generate_fault_tree_logical()
    generate_fault_tree_probabilities()
    generate_fmeca_charts()
    print("All figures re-generated successfully!")
