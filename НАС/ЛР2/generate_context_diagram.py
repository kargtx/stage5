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
    fig, ax = plt.subplots(figsize=(15.5, 9.5), dpi=300)
    ax.set_xlim(0, 155)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # ---------------------------------------------------------
    # 1. OUTER IDEF0 FRAME & HEADER TABLE (BPwin / Ramus style)
    # ---------------------------------------------------------
    frame = patches.Rectangle((3, 3), 149, 94, linewidth=1.8, edgecolor='#1c2833', facecolor='white')
    ax.add_patch(frame)
    
    # Header block line
    ax.plot([3, 152], [90, 90], color='#1c2833', lw=1.2)
    # Vertical dividers in header
    ax.plot([32, 32], [90, 97], color='#1c2833', lw=1.0)
    ax.plot([78, 78], [90, 97], color='#1c2833', lw=1.0)
    ax.plot([110, 110], [90, 97], color='#1c2833', lw=1.0)
    ax.plot([133, 133], [90, 97], color='#1c2833', lw=1.0)
    
    # Header fields
    ax.text(4.5, 94.8, "USED AT:", fontsize=7, color='#7f8c8d')
    ax.text(4.5, 92, "УУНиТ, Кафедра АСУ", fontsize=8, fontweight='bold', color='#2c3e50')
    
    ax.text(33.5, 94.8, "AUTHOR: Гайнанов Д.И., Гарифуллин К.Р.", fontsize=7.5, color='#2c3e50')
    ax.text(33.5, 92, "PROJECT: АВПКО АСУ ТП и ПАЗ ЭЛОУ-АВТ", fontsize=8, fontweight='bold', color='#1a5276')
    
    ax.text(79.5, 94.8, "DATE: 15.10.2026", fontsize=7.5, color='#2c3e50')
    ax.text(79.5, 92, "REV: 1.0 (FINAL)", fontsize=7.5, color='#2c3e50')
    
    ax.text(111.5, 94.8, "STATUS: [X] RECOMMENDED", fontsize=7.5, color='#27ae60', fontweight='bold')
    ax.text(111.5, 92, "READER: доц. Ризванов К.А.", fontsize=7.5, color='#2c3e50')
    
    ax.text(134.5, 94.8, "CONTEXT:", fontsize=7.5, color='#7f8c8d')
    ax.text(134.5, 91.5, "TOP (A-0)", fontsize=10, fontweight='bold', color='#c0392b')
    
    # ---------------------------------------------------------
    # 2. CENTRAL FUNCTION BLOCK (NODE A-0)
    # ---------------------------------------------------------
    bx, by, bw, bh = 54, 37, 48, 29
    main_box = patches.Rectangle((bx, by), bw, bh, linewidth=2.2, edgecolor='#1b4f72', facecolor='#ebf5fb')
    ax.add_patch(main_box)
    
    # Box title text
    ax.text(bx + bw/2, by + bh/2 + 2.5, "Проведение анализа видов,\nпоследствий и критичности\nотказов (АВПКО)\nАСУ ТП и системы ПАЗ",
            fontsize=12, fontweight='bold', ha='center', va='center', color='#0b2545', linespacing=1.25)
    
    # Node number in lower right corner of the box
    ax.text(bx + bw - 3, by + 3, "0", fontsize=14, fontweight='bold', ha='right', va='bottom', color='#1b4f72')

    def draw_arrow(x1, y1, x2, y2, color='#1c2833', lw=1.4):
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="->", lw=lw, color=color, mutation_scale=14))

    # ---------------------------------------------------------
    # 3. INPUT ARROWS (ВХОДЫ - СЛЕВА)
    # ---------------------------------------------------------
    draw_arrow(7, 57, bx, 57)
    ax.text(8.5, 58.5, "Данные о структуре и составе КТС\n(схема РСУ 1oo2D и ПАЗ 2oo3)", fontsize=8, color='#1c2833', fontweight='bold')

    draw_arrow(7, 47, bx, 47)
    ax.text(8.5, 48.5, "Параметры надежности комплектующих\n(интенсивности отказов λ, MTBF)", fontsize=8, color='#1c2833', fontweight='bold')

    draw_arrow(7, 40, bx, 40)
    ax.text(8.5, 41.5, "Статистика эксплуатации ОПО ЭЛОУ-АВТ", fontsize=8, color='#1c2833', fontweight='bold')

    # ---------------------------------------------------------
    # 4. CONTROL ARROWS (УПРАВЛЕНИЕ - СВЕРХУ)
    # ---------------------------------------------------------
    # Control 1: Стандарты надежности
    draw_arrow(66, 90, 66, by + bh, color='#922b21')
    ax.text(67, 82, "ГОСТ 27.310-95\nГОСТ Р МЭК 60812\nГОСТ Р МЭК 61508", fontsize=7.8, color='#922b21', fontweight='bold', va='center')

    # Control 2: Классификаторы и критерии
    draw_arrow(88, 90, 88, by + bh, color='#922b21')
    ax.text(89, 82, "Классификаторы отказов\nи критерии риска\n(SIL 3, RPN ≥ 100)", fontsize=7.8, color='#922b21', fontweight='bold', va='center')

    # ---------------------------------------------------------
    # 5. MECHANISM ARROWS (МЕХАНИЗМЫ / РЕСУРСЫ - СНИЗУ)
    # ---------------------------------------------------------
    draw_arrow(62, 11, 62, by, color='#145a32')
    ax.text(62, 16.5, "Конструктор\nАСУ ТП", fontsize=7.8, color='#145a32', fontweight='bold', ha='center',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))

    draw_arrow(78, 11, 78, by, color='#145a32')
    ax.text(78, 16.5, "Инженер\nпо надежности", fontsize=7.8, color='#145a32', fontweight='bold', ha='center',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))

    draw_arrow(94, 11, 94, by, color='#145a32')
    ax.text(94, 16.5, "Эксперт по\nбезопасности ПАЗ", fontsize=7.8, color='#145a32', fontweight='bold', ha='center',
            bbox=dict(boxstyle="round,pad=0.2", facecolor='#eafaf1', edgecolor='#a9dfbf', lw=0.6))

    # ---------------------------------------------------------
    # 6. OUTPUT ARROWS (ВЫХОДЫ - СПРАВА)
    # ---------------------------------------------------------
    draw_arrow(bx + bw, 57, 147, 57)
    ax.text(146, 58.5, "Таблица АВПКО (FMECA)\nс оценками B1, B2, B3 и C", fontsize=8, color='#1c2833', fontweight='bold', ha='right')

    draw_arrow(bx + bw, 49, 147, 49)
    ax.text(146, 50.5, "Перечень критичных элементов\n(отсекатель XV-101, датчик PT-101)", fontsize=8, color='#1c2833', fontweight='bold', ha='right')

    draw_arrow(bx + bw, 41, 147, 41)
    ax.text(146, 42.5, "Проектные рекомендации и меры\n(PST-тестирование, СТД, резервирование)", fontsize=8, color='#1c2833', fontweight='bold', ha='right')

    # ---------------------------------------------------------
    # 7. BOTTOM FOOTER TITLE BLOCK (IDEF0 Standard)
    # ---------------------------------------------------------
    ax.plot([3, 152], [11, 11], color='#1c2833', lw=1.2)
    ax.plot([26, 26], [3, 11], color='#1c2833', lw=1.0)
    ax.plot([124, 124], [3, 11], color='#1c2833', lw=1.0)
    
    ax.text(5, 7, "NODE:", fontsize=7, color='#7f8c8d')
    ax.text(12, 6.2, "A-0", fontsize=11, fontweight='bold', color='#1a5276')
    
    ax.text(28, 7, "TITLE:", fontsize=7, color='#7f8c8d')
    ax.text(38, 6.2, "ПРОЦЕСС ПРОВЕДЕНИЯ АВПКО АСУ ТП И СИСТЕМЫ ПАЗ", fontsize=10.5, fontweight='bold', color='#1c2833')
    
    ax.text(126, 7, "NUMBER:", fontsize=7, color='#7f8c8d')
    ax.text(142, 6.2, "1", fontsize=10, fontweight='bold', color='#1c2833')

    plt.tight_layout()
    output_path = os.path.join(OUTPUT_DIR, "idef0_context_a0.png")
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Context Diagram regenerated successfully:", output_path)

if __name__ == "__main__":
    generate_context_diagram()
