import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

DIR_LR5 = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР5"

def make_figures():
    plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
    plt.rcParams['axes.edgecolor'] = 'black'
    plt.rcParams['axes.linewidth'] = 1.0

    # Схема гейтовой системы (Stage-Gate) управления качеством и надежностью АСУ ТП
    fig, ax = plt.subplots(figsize=(10.5, 6.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    def draw_stage(x, y, w, h, title, text, fill='#FAFAFA'):
        rect = patches.Rectangle((x - w/2, y - h/2), w, h, edgecolor='black', facecolor=fill, linewidth=1.2)
        ax.add_patch(rect)
        ax.text(x, y + h/2 - 2.8, title, ha='center', va='top', fontsize=8, fontweight='bold', color='black')
        ax.text(x, y - 1.2, text, ha='center', va='center', fontsize=6.8, color='black', wrap=True)

    def draw_gate(x, y, w, h, title, criteria, fill='#EAEAEA'):
        poly = patches.Polygon([[x, y + h/2], [x + w/2, y], [x, y - h/2], [x - w/2, y]],
                               edgecolor='black', facecolor=fill, linewidth=1.2)
        ax.add_patch(poly)
        ax.text(x, y + 2, title, ha='center', va='center', fontsize=7.5, fontweight='bold', color='black')
        ax.text(x, y - 2.5, criteria, ha='center', va='center', fontsize=6, color='black', style='italic')

    def draw_arrow(x1, y1, x2, y2, label=""):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color='black', lw=1.2, mutation_scale=11))
        if label:
            mx, my = (x1 + x2)/2, (y1 + y2)/2
            ax.text(mx, my + 1.2, label, ha='center', va='bottom', fontsize=6.5, color='black')

    # Заголовок
    rect_top = patches.Rectangle((5, 90), 90, 7.5, edgecolor='black', facecolor='#EBEBEB', linewidth=1.2)
    ax.add_patch(rect_top)
    ax.text(50, 93.8, "ГЕЙТОВАЯ СИСТЕМА (STAGE-GATE) УПРАВЛЕНИЯ КАЧЕСТВОМ И НАДЕЖНОСТЬЮ АСУ ТП",
            ha='center', va='center', fontsize=8.5, fontweight='bold', color='black')

    # Стадии и Ворота (по горизонтали)
    # Stage 1: Маркетинг
    draw_stage(14, 62, 18, 44, "СТАДИЯ 1:\nМАРКЕТИНГ",
               "• Анализ рынка и условий ОПО\n"
               "• Требования заказчика\n"
               "• ТЭО и рентабельность\n"
               "• Формирование ТЗ на АСУ\n"
               "• Выбор уровня SIL 3\n"
               "• Оценка TCO системы\n"
               "• Методы: социолог.,\n"
               "  расчетный, экспертный")

    # Gate 1
    draw_gate(27.5, 62, 7.5, 14, "GATE 1", "ТЗ и\nконцепция")

    # Stage 2: Разработка
    draw_stage(41, 62, 18, 44, "СТАДИЯ 2:\nРАЗРАБОТКА",
               "• Проектирование КТС и ПО\n"
               "• Расчет надежности P(t), Kг\n"
               "• FMECA / АВПКО отказов\n"
               "• Моделирование в HIL\n"
               "• Расчет ЗИП по Пуассону\n"
               "• Архитектура 1oo2D / 2oo3\n"
               "• Методы: расчетный,\n"
               "  имитац., аналитич.")

    # Gate 2
    draw_gate(54.5, 62, 7.5, 14, "GATE 2", "Проект\nи РД")

    # Stage 3: Производство
    draw_stage(68, 62, 18, 44, "СТАДИЯ 3:\nПРОИЗВОДСТВО",
               "• 100% входной контроль ПКИ\n"
               "• Сборка и жгутовка шкафов\n"
               "• Учет дефектов монтажа\n"
               "• Прогон 72 ч под нагрузкой\n"
               "• Испытания на вибро и ЭМС\n"
               "• Приемочные испытания FAT\n"
               "• Методы: измерит.,\n"
               "  регистрац., статист.")

    # Gate 3
    draw_gate(81.5, 62, 7.5, 14, "GATE 3", "Заводской\nFAT")

    # Stage 4: Эксплуатация
    draw_stage(93, 62, 14, 44, "СТАДИЯ 4:\nЭКСПЛУАТАЦИЯ",
               "• Шеф-монтаж и ПНР\n"
               "• Испытания SAT\n"
               "• Мониторинг SCADA\n"
               "• Учет Т0 и Тв\n"
               "• ТО по состоянию\n"
               "• Тест PST штока\n"
               "• Продление ресурса")

    # Gate 4
    draw_gate(93, 30, 8.5, 12, "GATE 4", "Ввод в\nпромэкспл.")

    # Соединительные прямые стрелки между стадиями и шлюзами
    draw_arrow(5, 62, 5, 62)
    draw_arrow(23, 62, 23.8, 62)
    draw_arrow(31.2, 62, 32, 62)
    draw_arrow(50, 62, 50.8, 62)
    draw_arrow(58.2, 62, 59, 62)
    draw_arrow(77, 62, 77.8, 62)
    draw_arrow(85.2, 62, 86, 62)
    draw_arrow(93, 40, 93, 36)

    # Ветви решений в гейтах: Go / Recycle / Kill
    # Обозначение обратной связи (Recycle / Доработка)
    ax.plot([27.5, 27.5, 14, 14], [55, 36, 36, 40], color='black', lw=1.1, linestyle='--')
    ax.annotate('', xy=(14, 40), xytext=(14, 38),
                arrowprops=dict(arrowstyle="-|>", color='black', lw=1.1, mutation_scale=9))
    ax.text(20.5, 37.2, "Доработка ТЗ", ha='center', va='bottom', fontsize=5.8, color='black', style='italic')

    ax.plot([54.5, 54.5, 41, 41], [55, 36, 36, 40], color='black', lw=1.1, linestyle='--')
    ax.annotate('', xy=(41, 40), xytext=(41, 38),
                arrowprops=dict(arrowstyle="-|>", color='black', lw=1.1, mutation_scale=9))
    ax.text(47.5, 37.2, "Коррекция РД", ha='center', va='bottom', fontsize=5.8, color='black', style='italic')

    ax.plot([81.5, 81.5, 68, 68], [55, 36, 36, 40], color='black', lw=1.1, linestyle='--')
    ax.annotate('', xy=(68, 40), xytext=(68, 38),
                arrowprops=dict(arrowstyle="-|>", color='black', lw=1.1, mutation_scale=9))
    ax.text(74.5, 37.2, "Устранение брака", ha='center', va='bottom', fontsize=5.8, color='black', style='italic')

    # Блок пояснения решений шлюзов
    rect_bot = patches.Rectangle((5, 8), 90, 16, edgecolor='black', facecolor='#F2F2F2', linewidth=1.2)
    ax.add_patch(rect_bot)
    ax.text(50, 21.5, "КРИТЕРИИ И РЕШЕНИЯ НА КОНТРОЛЬНЫХ ШЛЮЗАХ (GATE REVIEWS):",
            ha='center', va='center', fontsize=7.5, fontweight='bold', color='black')
    ax.text(50, 14.5,
            "1. Go (Проход) – интегральный показатель качества I_stage >= I_порог (0,85...0,90), риски допустимы -> переход на следующий этап;\n"
            "2. Recycle (Доработка) – единичные показатели не соответствуют ТЗ -> возврат на предыдущую стадию для корректировки КД/ТД;\n"
            "3. Hold (Приостановка) – нехватка ресурсов, ожидание комплектующих или согласования экспертов промбезопасности;\n"
            "4. Kill (Прекращение) – экономическая нецелесообразность или непреодолимые несоответствия требованиям надежности.",
            ha='center', va='center', fontsize=6.5, color='black')

    plt.tight_layout()
    fig_path = os.path.join(DIR_LR5, "fig_stage_gate_system.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"Created: {fig_path}")

if __name__ == '__main__':
    make_figures()
