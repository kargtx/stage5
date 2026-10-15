import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Calibri']
plt.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР2"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def generate_context_diagram():
    fig, ax = plt.subplots(figsize=(15, 9.5), dpi=300)
    ax.set_xlim(0, 150)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # ---------------------------------------------------------
    # 1. OUTER IDEF0 FRAME & HEADER TABLE (BPwin / Ramus style)
    # ---------------------------------------------------------
    # Main outer border
    frame = patches.Rectangle((3, 3), 144, 94, linewidth=1.8, edgecolor='#1c2833', facecolor='white')
    ax.add_patch(frame)
    
    # Header block line
    ax.plot([3, 147], [90, 90], color='#1c2833', lw=1.2)
    # Vertical dividers in header
    ax.plot([30, 30], [90, 97], color='#1c2833', lw=1.0)
    ax.plot([75, 75], [90, 97], color='#1c2833', lw=1.0)
    ax.plot([105, 105], [90, 97], color='#1c2833', lw=1.0)
    ax.plot([128, 128], [90, 97], color='#1c2833', lw=1.0)
    
    # Header fields
    ax.text(4, 95, "USED AT:", fontsize=7, color='#7f8c8d')
    ax.text(4, 92, "УУНиТ, Кафедра АСУ", fontsize=8, fontweight='bold', color='#2c3e50')
    
    ax.text(31, 95, "AUTHOR: Гайнанов Д.И., Гарифуллин К.Р.", fontsize=7.5, color='#2c3e50')
    ax.text(31, 92, "PROJECT: АВПКО АСУ ТП и ПАЗ ЭЛОУ-АВТ", fontsize=8, fontweight='bold', color='#1a5276')
    
    ax.text(76, 95, "DATE: 15.10.2026", fontsize=7.5, color='#2c3e50')
    ax.text(76, 92, "REV: 1.0 (FINAL)", fontsize=7.5, color='#2c3e50')
    
    ax.text(106, 95, "STATUS: [X] RECOMMENDED", fontsize=7.5, color='#27ae60', fontweight='bold')
    ax.text(106, 92, "READER: доц. Ризванов К.А.", fontsize=7.5, color='#2c3e50')
    
    ax.text(129, 95, "CONTEXT:", fontsize=7.5, color='#7f8c8d')
    ax.text(129, 91.5, "TOP (A-0)", fontsize=10, fontweight='bold', color='#c0392b')
    
    # ---------------------------------------------------------
    # 2. CENTRAL FUNCTION BLOCK (NODE A-0)
    # ---------------------------------------------------------
    # Central box
    bx, by, bw, bh = 54, 38, 44, 28
    main_box = patches.Rectangle((bx, by), bw, bh, linewidth=2.2, edgecolor='#1b4f72', facecolor='#ebf5fb')
    ax.add_patch(main_box)
    
    # Box title text
    ax.text(bx + bw/2, by + bh/2 + 3, "Проведение анализа видов,\nпоследствий и критичности\nотказов (АВПКО)\nАСУ ТП и системы ПАЗ",
            fontsize=11.5, fontweight='bold', ha='center', va='center', color='#0b2545', linespacing=1.2)
    
    # Node number in lower right corner of the box
    ax.text(bx + bw - 3, by + 3, "0", fontsize=13, fontweight='bold', ha='right', va='bottom', color='#1b4f72')

    # Arrow helper
    def draw_arrow(x1, y1, x2, y2, color='#1c2833', lw=1.4):
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="->", lw=lw, color=color, mutation_scale=14))

    # ---------------------------------------------------------
    # 3. INPUT ARROWS (ВХОДЫ - СЛЕВА)
    # ---------------------------------------------------------
    # Input 1: Данные о структуре и составе КТС
    draw_arrow(8, 58, bx, 58)
    ax.text(9, 59.5, "Данные о структуре и составе КТС\n(схема РСУ 1oo2D и ПАЗ 2oo3)", fontsize=8, color='#1c2833', fontweight='bold')

    # Input 2: Характеристики надежности ЛР1
    draw_arrow(8, 48, bx, 48)
    ax.text(9, 49.5, "Параметры надежности комплектующих\n(интенсивности отказов λ, MTBF)", fontsize=8, color='#1c2833', fontweight='bold')

    # Input 3: Опыт эксплуатации аналогов
    draw_arrow(8, 41, bx, 41)
    ax.text(9, 42.5, "Статистика эксплуатации ОПО ЭЛОУ-АВТ", fontsize=8, color='#1c2833', fontweight='bold')

    # ---------------------------------------------------------
    # 4. CONTROL ARROWS (УПРАВЛЕНИЕ - СВЕРХУ)
    # ---------------------------------------------------------
    # Control 1: ГОСТ стандарты
    draw_arrow(64, 90, 64, by + bh, color='#922b21')
    ax.text(64.5, 78, "ГОСТ 27.310-95\nГОСТ Р МЭК 60812\nГОСТ Р МЭК 61508", fontsize=8, color='#922b21', fontweight='bold', va='center')

    # Control 2: Классификаторы отказов
    draw_arrow(78, 90, 78, by + bh, color='#922b21')
    ax.text(78.5, 78, "Справочники-классификаторы\nвидов и причин отказов", fontsize=8, color='#922b21', fontweight='bold', va='center')

    # Control 3: Критерии допустимого риска
    draw_arrow(90, 90, 90, by + bh, color='#922b21')
    ax.text(90.5, 78, "Критерии допустимого риска\n(уровень SIL 3, RPN ≥ 100)", fontsize=8, color='#922b21', fontweight='bold', va='center')

    # ---------------------------------------------------------
    # 5. MECHANISM ARROWS (МЕХАНИЗМЫ / РЕСУРСЫ - СНИЗУ)
    # ---------------------------------------------------------
    # Mechanism 1: Конструктор АСУ ТП
    draw_arrow(64, 11, 64, by, color='#145a32')
    ax.text(64, 17, "Конструктор\nАСУ ТП", fontsize=8, color='#145a32', fontweight='bold', ha='center',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))

    # Mechanism 2: Инженер по надежности
    draw_arrow(78, 11, 78, by, color='#145a32')
    ax.text(78, 17, "Инженер\nпо надежности", fontsize=8, color='#145a32', fontweight='bold', ha='center',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))

    # Mechanism 3: Эксперт по безопасности ПАЗ
    draw_arrow(90, 11, 90, by, color='#145a32')
    ax.text(90, 17, "Эксперт по\nбезопасности ПАЗ", fontsize=8, color='#145a32', fontweight='bold', ha='center',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))

    # ---------------------------------------------------------
    # 6. OUTPUT ARROWS (ВЫХОДЫ - СПРАВА)
    # ---------------------------------------------------------
    # Output 1: Таблица АВПКО
    draw_arrow(bx + bw, 58, 142, 58)
    ax.text(141, 59.5, "Таблица АВПКО (FMECA)\nс оценками B1, B2, B3 и C", fontsize=8, color='#1c2833', fontweight='bold', ha='right')

    # Output 2: Перечень критичных элементов
    draw_arrow(bx + bw, 50, 142, 50)
    ax.text(141, 51.5, "Перечень критичных элементов\n(отсекатель XV-101, датчик PT-101)", fontsize=8, color='#1c2833', fontweight='bold', ha='right')

    # Output 3: Рекомендации по обеспечению надежности
    draw_arrow(bx + bw, 42, 142, 42)
    ax.text(141, 43.5, "Проектные рекомендации и меры\n(PST-тестирование, СТД, резервирование)", fontsize=8, color='#1c2833', fontweight='bold', ha='right')

    # ---------------------------------------------------------
    # 7. BOTTOM FOOTER TITLE BLOCK (IDEF0 Standard)
    # ---------------------------------------------------------
    ax.plot([3, 147], [11, 11], color='#1c2833', lw=1.2)
    ax.plot([25, 25], [3, 11], color='#1c2833', lw=1.0)
    ax.plot([120, 120], [3, 11], color='#1c2833', lw=1.0)
    
    ax.text(5, 7, "NODE:", fontsize=7, color='#7f8c8d')
    ax.text(11, 6.2, "A-0", fontsize=11, fontweight='bold', color='#1a5276')
    
    ax.text(27, 7, "TITLE:", fontsize=7, color='#7f8c8d')
    ax.text(37, 6.2, "ПРОЦЕСС ПРОВЕДЕНИЯ АВПКО АСУ ТП И СИСТЕМЫ ПАЗ", fontsize=10.5, fontweight='bold', color='#1c2833')
    
    ax.text(122, 7, "NUMBER:", fontsize=7, color='#7f8c8d')
    ax.text(138, 6.2, "1", fontsize=10, fontweight='bold', color='#1c2833')

    plt.tight_layout()
    output_path = os.path.join(OUTPUT_DIR, "idef0_context_a0.png")
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Context Diagram generated successfully:", output_path)

if __name__ == "__main__":
    generate_context_diagram()
