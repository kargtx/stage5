import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

DIR_LR3 = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР3"

def make_figures():
    # Настройка стилей - строго монохромные академические цвета (черный, серый, белый)
    plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
    plt.rcParams['axes.edgecolor'] = 'black'
    plt.rcParams['axes.linewidth'] = 1.0

    # 1. Схема дерева отказов (FTA)
    fig, ax = plt.subplots(figsize=(10, 6.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Вспомогательные функции отрисовки блоков
    def draw_box(x, y, w, h, text, shape='rect'):
        if shape == 'rect':
            box = patches.Rectangle((x - w/2, y - h/2), w, h, edgecolor='black', facecolor='white', linewidth=1.2)
            ax.add_patch(box)
        elif shape == 'gate_and':
            # Вентиль И: прямоугольник снизу, дуга сверху
            path = patches.FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle="round,pad=0.1,rounding_size=2",
                                          edgecolor='black', facecolor='white', linewidth=1.2)
            ax.add_patch(path)
        elif shape == 'circle':
            circ = patches.Circle((x, y), w/2, edgecolor='black', facecolor='white', linewidth=1.2)
            ax.add_patch(circ)
        ax.text(x, y, text, ha='center', va='center', fontsize=7.5, color='black', wrap=True)

    def draw_gate(x, y, gtype, label=""):
        # Отрисовка вентиля И / ИЛИ
        gw, gh = 7, 5
        box = patches.Rectangle((x - gw/2, y - gh/2), gw, gh, edgecolor='black', facecolor='#F0F0F0', linewidth=1.2)
        ax.add_patch(box)
        ax.text(x, y, gtype, ha='center', va='center', fontsize=8, fontweight='bold', color='black')
        if label:
            ax.text(x, y - gh/2 - 2, label, ha='center', va='top', fontsize=6.5, color='black')

    def draw_line(x1, y1, x2, y2):
        ax.plot([x1, x2], [y1, y2], color='black', linewidth=1.0)

    # Головное событие T
    draw_box(50, 93, 38, 7, "T: Аварийная разгерметизация колонны\n(превышение давления P > Pкрит)", 'rect')
    draw_line(50, 89.5, 50, 86)
    
    # Главный вентиль И
    draw_gate(50, 83.5, "И", "G0")
    draw_line(50, 81, 26, 75)
    draw_line(50, 81, 74, 75)
    draw_line(26, 75, 26, 71.5)
    draw_line(74, 75, 74, 71.5)

    # Промежуточное 1: Инициирующий отказ контура регулирования E_иниц
    draw_box(26, 68, 30, 7, "E_иниц: Бросок давления в колонне\n(отказ контура РСУ)", 'rect')
    draw_line(26, 64.5, 26, 61)
    draw_gate(26, 58.5, "ИЛИ", "G1")
    
    # Входы в E_иниц
    draw_line(26, 56, 12, 50)
    draw_line(26, 56, 26, 50)
    draw_line(26, 56, 40, 50)
    draw_line(12, 50, 12, 46.5)
    draw_line(26, 50, 26, 46.5)
    draw_line(40, 50, 40, 46.5)

    draw_box(12, 43, 16, 7, "A1: Отказ измерителя\nдавления РСУ (X1-X4)", 'rect')
    draw_box(26, 43, 15, 7, "A2: Отказ контроллера\nРСУ (ПЛК / ПО)", 'rect')
    draw_box(40, 43, 15, 7, "A3: Отказ клапана\nPV-101 (РСУ)", 'rect')

    # Промежуточное 2: Несрабатывание системы ПАЗ E_защита
    draw_box(74, 68, 34, 7, "E_защита: Несрабатывание системы ПАЗ\nпри аварийном давлении", 'rect')
    draw_line(74, 64.5, 74, 61)
    draw_gate(74, 58.5, "ИЛИ", "G2")

    # 4 ветви защиты ПАЗ: B1 (датчики 2oo3), B2 (ПЛК ПАЗ), B3 (отсекатель XV-101), B4 (питание)
    xs_b = [56, 68, 80, 92]
    for x in xs_b:
        draw_line(74, 56, x, 50)
        draw_line(x, 50, x, 46.5)

    draw_box(56, 43, 12, 7, "B1: Отказ\nдатчиков ПАЗ\n(2oo3)", 'rect')
    draw_box(68, 43, 12, 7, "B2: Отказ\nконтроллера\nПАЗ 1oo2D", 'rect')
    draw_box(80, 43, 12, 7, "B3: Отказ\nотсекателя\nXV-101", 'rect')
    draw_box(92, 43, 12, 7, "B4: Потеря\nпитания\n(Ввод + ИБП)", 'rect')

    # Подветвь B1: 2oo3
    draw_line(56, 39.5, 56, 35.5)
    draw_gate(56, 33, "2 из 3", "G3")
    draw_line(56, 30.5, 50, 25)
    draw_line(56, 30.5, 56, 25)
    draw_line(56, 30.5, 62, 25)
    draw_box(50, 21.5, 5.5, 5.5, "X5\nДат.1", 'circle')
    draw_box(56, 21.5, 5.5, 5.5, "X6\nДат.2", 'circle')
    draw_box(62, 21.5, 5.5, 5.5, "X7\nДат.3", 'circle')

    # Подветвь B2: ПЛК 1oo2
    draw_line(68, 39.5, 68, 35.5)
    draw_gate(68, 33, "И", "G4")
    draw_line(68, 30.5, 65.5, 25)
    draw_line(68, 30.5, 70.5, 25)
    draw_box(65.5, 21.5, 5.5, 5.5, "X9\nЦПУ1", 'circle')
    draw_box(70.5, 21.5, 5.5, 5.5, "X10\nЦПУ2", 'circle')

    # Подветвь B3: Отсекатель
    draw_line(80, 39.5, 80, 35.5)
    draw_gate(80, 33, "ИЛИ", "G5")
    draw_line(80, 30.5, 77, 25)
    draw_line(80, 30.5, 83, 25)
    draw_box(77, 21.5, 5.5, 5.5, "X11\nКран", 'circle')
    draw_box(83, 21.5, 5.5, 5.5, "X12\nСолен.", 'circle')

    # Подветвь B4: Питание
    draw_line(92, 39.5, 92, 35.5)
    draw_gate(92, 33, "И", "G6")
    draw_line(92, 30.5, 89, 25)
    draw_line(92, 30.5, 95, 25)
    draw_box(89, 21.5, 5.5, 5.5, "X14\nСеть", 'circle')
    draw_box(95, 21.5, 5.5, 5.5, "X15\nИБП", 'circle')

    # Базовые события контура РСУ
    draw_line(12, 39.5, 8, 34)
    draw_line(12, 39.5, 16, 34)
    draw_box(8, 29.5, 5.5, 5.5, "X1\nДатч.", 'circle')
    draw_box(16, 29.5, 5.5, 5.5, "X2\nБарь.", 'circle')

    draw_line(40, 39.5, 36, 34)
    draw_line(40, 39.5, 44, 34)
    draw_box(36, 29.5, 5.5, 5.5, "X3\nКлап.", 'circle')
    draw_box(44, 29.5, 5.5, 5.5, "X4\nПозиц.", 'circle')

    # Подпись легенды
    ax.text(50, 7, "Примечание: Квадраты – промежуточные и головные события; Круги – исходные базовые события;\nПрямоугольники с надписью – логические вентили И, ИЛИ, 2 из 3.",
            ha='center', va='center', fontsize=7, style='italic', color='black')

    plt.tight_layout()
    fig_path1 = os.path.join(DIR_LR3, "fig_fta_structure.png")
    fig.savefig(fig_path1, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"Created: {fig_path1}")

    # 2. График анализа важности компонентов (Индексы Фусселя-Везели и Бирнбаума)
    fig2, ax2 = plt.subplots(figsize=(8, 4.5), dpi=300)
    components = [
        "X11: Кран-отсекатель XV-101",
        "X12: Соленоидный клапан ПАЗ",
        "X3: Регулирующий клапан PV-101",
        "X1: Датчик давления РСУ PT-101",
        "X13: Потеря воздуха КИПиА",
        "X8: Общий отказ ПО ПАЗ",
        "X5/X6: Датчик ПАЗ (2oo3)",
        "X14/X15: Отказ питания",
        "X9/X10: ЦПУ ПЛК ПАЗ 1oo2"
    ]
    ifv_vals = [0.572, 0.238, 0.165, 0.118, 0.095, 0.048, 0.032, 0.012, 0.005]
    components.reverse()
    ifv_vals.reverse()

    y_pos = range(len(components))
    bars = ax2.barh(y_pos, ifv_vals, color='#707070', edgecolor='black', height=0.6)
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(components, fontsize=8.5, color='black')
    ax2.set_xlabel("Структурная значимость Фусселя–Везели (I_FV)", fontsize=9.5, fontweight='bold', color='black')
    ax2.set_title("Ранжирование базовых событий по степени влияния на головной отказ", fontsize=10, fontweight='bold', pad=10, color='black')
    ax2.grid(axis='x', linestyle='--', alpha=0.6, color='gray')
    ax2.set_xlim(0, 0.65)

    for bar, val in zip(bars, ifv_vals):
        ax2.text(val + 0.01, bar.get_y() + bar.get_height()/2, f"{val:.3f}", va='center', ha='left', fontsize=8, color='black')

    plt.tight_layout()
    fig_path2 = os.path.join(DIR_LR3, "fig_importance_pareto.png")
    fig2.savefig(fig_path2, dpi=300, bbox_inches='tight')
    plt.close(fig2)
    print(f"Created: {fig_path2}")

if __name__ == '__main__':
    make_figures()
