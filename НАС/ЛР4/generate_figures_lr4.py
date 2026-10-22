import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

DIR_LR4 = r"c:\Users\danch\OneDrive\Рабочий стол\stage5\НАС\ЛР4"

def make_figures():
    plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
    plt.rcParams['axes.edgecolor'] = 'black'
    plt.rcParams['axes.linewidth'] = 1.0

    # Схема структуры программы обеспечения надежности (ПОН) на этапах ЖЦ АСУ ТП
    fig, ax = plt.subplots(figsize=(10, 6.2), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    def draw_box(x, y, w, h, text, title="", fill='#FFFFFF'):
        rect = patches.Rectangle((x - w/2, y - h/2), w, h, edgecolor='black', facecolor=fill, linewidth=1.2)
        ax.add_patch(rect)
        if title:
            ax.text(x, y + h/2 - 2.5, title, ha='center', va='top', fontsize=8, fontweight='bold', color='black')
            ax.text(x, y - 1, text, ha='center', va='center', fontsize=6.8, color='black', wrap=True)
        else:
            ax.text(x, y, text, ha='center', va='center', fontsize=7, color='black', wrap=True)

    def draw_arrow(x1, y1, x2, y2, label=""):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-|>", color='black', lw=1.2, mutation_scale=12))
        if label:
            mx, my = (x1 + x2)/2, (y1 + y2)/2
            ax.text(mx, my + 1.5, label, ha='center', va='bottom', fontsize=6.5, color='black', style='italic')

    # Заголовочный блок
    draw_box(50, 94, 90, 7, "КОМПЛЕКСНАЯ СИСТЕМА ОБЕСПЕЧЕНИЯ НАДЕЖНОСТИ (ПОН) АСУ ТП И ПАЗ", fill='#EBEBEB')

    # 3 основных этапа ЖЦ: Разработка, Производство, Эксплуатация
    # 1. Этап разработки
    draw_box(20, 60, 32, 52, 
             "• Сбор условий работы и требований ТЗ\n"
             "• Сравнение и выбор структуры КТС и ПО\n"
             "• Декомпозиция норм надежности (п.3)\n"
             "• Анализ видов и последствий отказов FMECA (п.4)\n"
             "• Оценка технологичности изделий (п.5)\n"
             "• Экспериментальная отработка алгоритмов (п.6)\n"
             "• Расчет состава и норм ЗИП-О / ЗИП-Г (п.7)\n"
             "• Обоснование регламентов ТОиР (п.8)\n"
             "• Заводские приемочные испытания FAT (п.9)\n"
             "• Разработка регламентов и инструкций (п.10)\n"
             "• Экспертиза промбезопасности SIL 3 (п.11)",
             title="СТАДИЯ РАЗРАБОТКИ\n(пп. 1 - 11 ПОН)", fill='#FAFAFA')

    # 2. Этап производства
    draw_box(52, 60, 26, 52,
             "• Учет дефектов сборки шкафов (п.12)\n"
             "• Испытания на надежность (п.13):\n"
             "   - климатические испытания\n"
             "   - вибропрочность, ЭМС\n"
             "   - технологический прогон 72 ч\n"
             "• Анализ дефектов и правка КД/ТД (п.14)\n"
             "• 100% входной контроль ПКИ (п.15)\n"
             "• Контроль техдисциплины монтажа,\n"
             "  обжима и пайки (п.16)",
             title="СТАДИЯ ПРОИЗВОДСТВА\n(пп. 12 - 16 ПОН)", fill='#FAFAFA')

    # 3. Этап эксплуатации
    draw_box(82, 60, 28, 52,
             "• Сбор статистики отказов в SCADA (п.17)\n"
             "• Учет времени восстановления Тв\n"
             "• Анализ надежности и корректировка\n"
             "  периодичности ТО и ЗИП (п.18)\n"
             "• Переход на ТО по состоянию (HART/PST)\n"
             "• Модернизация ПО и продление\n"
             "  назначенного ресурса (п.19)",
             title="СТАДИЯ ЭКСПЛУАТАЦИИ\n(пп. 17 - 19 ПОН)", fill='#FAFAFA')

    # Прямые стрелки процесса
    draw_arrow(50, 90.5, 20, 86.5)
    draw_arrow(36, 60, 39, 60, "КД, ТД, софт")
    draw_arrow(65, 60, 68, 60, "Готовые шкафы")

    # Обратные связи (петли качества и корректировок)
    # От производства к разработке
    ax.annotate('', xy=(36, 40), xytext=(39, 40),
                arrowprops=dict(arrowstyle="-|>", color='black', lw=1.2, mutation_scale=12, linestyle='--'))
    ax.text(37.5, 41.5, "Коррекция КД", ha='center', va='bottom', fontsize=6, color='black', style='italic')

    # От эксплуатации к производству и разработке
    ax.plot([82, 82, 20, 20], [33, 23, 23, 33], color='black', lw=1.1, linestyle=':')
    ax.annotate('', xy=(20, 33.5), xytext=(20, 30),
                arrowprops=dict(arrowstyle="-|>", color='black', lw=1.1, mutation_scale=10))
    ax.text(51, 24.5, "Обратная связь по эксплуатационной надежности (рекламации, статистика отказов)", 
            ha='center', va='bottom', fontsize=6.8, color='black', style='italic')

    # Нижний блок нормативного обеспечения
    draw_box(50, 11, 90, 13, 
             "Нормативно-методический базис ПОН: ГОСТ 24.701-86 (Надежность АСУ), ГОСТ 27.002-2015, ГОСТ Р 27.310-95 (FMECA),\n"
             "ГОСТ Р МЭК 61508 / ГОСТ Р МЭК 61511 (Функциональная безопасность ПАЗ, уровни SIL), ГОСТ 34.601-90 (Стадии создания АС).",
             title="НОРМАТИВНАЯ И МЕТОДИЧЕСКАЯ БАЗА ПЛАНА ОБЕСПЕЧЕНИЯ НАДЕЖНОСТИ", fill='#EBEBEB')

    plt.tight_layout()
    fig_path = os.path.join(DIR_LR4, "fig_pon_lifecycle.png")
    fig.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"Created: {fig_path}")

if __name__ == '__main__':
    make_figures()
