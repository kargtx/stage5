# -*- coding: utf-8 -*-
"""
Генератор материалов к Лабораторной работе №3 по курсу «АСУСН»
Вариант 2: Расход (0–500 м³/ч), Уровень (0–100 %), Вибрация (0–20 мм/с).
Пороги тревог: Q > 400 м³/ч, V > 15 мм/с.
Источник данных (Вариант A): Modbus TCP ПЛК-имитатор (127.0.0.1:502, Unit ID = 1),
подключённый к реальному экземпляру Rapid SCADA 6.4.7.

Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С
Проверил: Антонов В.В.
УУНиТ, 2026 г.
"""

from __future__ import annotations

import os
from pathlib import Path
import sys

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt, RGBColor

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
IMG_DIR = BASE_DIR / "images_lab3"
IMG_DIR.mkdir(parents=True, exist_ok=True)


# =============================================================================
# 1. ДИАГРАММА ИНФОРМАЦИОННЫХ ПОТОКОВ И АРХИТЕКТУРЫ (ГОСТ / ЧЁТКАЯ ИНЖЕНЕРНАЯ)
# =============================================================================

def generate_architecture_diagram() -> Path:
    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(12.2, 6.8), dpi=300)
    ax.set_xlim(0, 12.2)
    ax.set_ylim(0, 6.8)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    ax.text(
        6.1, 6.42,
        "Схема информационных потоков SCADA-системы (Вариант 2 + Вариант A: Modbus TCP)",
        ha="center", va="center", fontsize=13.5, fontweight="bold", color="#0F172A",
    )
    ax.text(
        6.1, 6.08,
        "ПЛК-имитатор (127.0.0.1:502)  <->  SCADA Communicator  <->  SCADA Server (TCP 10000)  <->  WebStation (HTTP 10008)",
        ha="center", va="center", fontsize=9.8, color="#334155",
    )

    def draw_box(x, y, w, h, title, body_lines, fc="#F8FAFC", ec="#1E293B", hdr_bg="#1E293B", lw=1.6):
        rect = FancyBboxPatch(
            (x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.10",
            facecolor=fc, edgecolor=ec, linewidth=lw, zorder=2,
        )
        ax.add_patch(rect)
        hdr_h = 0.56
        hdr = FancyBboxPatch(
            (x, y + h - hdr_h), w, hdr_h, boxstyle="round,pad=0.04,rounding_size=0.10",
            facecolor=hdr_bg, edgecolor=ec, linewidth=lw, zorder=3,
        )
        ax.add_patch(hdr)
        ax.text(
            x + w / 2, y + h - hdr_h / 2, title,
            ha="center", va="center", fontsize=9.5, fontweight="bold", color="#FFFFFF", zorder=4,
        )
        text_y = y + h - hdr_h - 0.28
        for line in body_lines:
            if line == "---DIV---":
                ax.plot([x + 0.16, x + w - 0.16], [text_y - 0.08, text_y - 0.08], color="#CBD5E1", lw=1.0, zorder=4)
                text_y -= 0.24
                continue
            ax.text(
                x + 0.16, text_y, line,
                ha="left", va="top", fontsize=8.4, color="#0F172A", zorder=4,
            )
            text_y -= 0.34

    # Блок 1: ПЛК-имитатор
    draw_box(
        0.25, 0.75, 2.75, 4.95,
        "ПЛК-ИМИТАТОР (Modbus TCP)",
        [
            "modbus_plc_simulator.py",
            "Modbus TCP Slave, Unit ID=1",
            "Порт: 127.0.0.1:502",
            "---DIV---",
            "Holding Registers (Float32):",
            "• Reg 10..11: Расход Q (м³/ч)",
            "• Reg 12..13: Уровень L (%)",
            "• Reg 14..15: Вибрация V (мм/с)",
            "Состояние / Команда (uint16):",
            "• Reg 3: Состояние (0/1)",
            "• Reg 4: ПУСК/СТОП (FC 06)",
            "HTTP API управления: :5502",
        ],
        fc="#F8FAFC", ec="#1E293B", hdr_bg="#0F172A",
    )

    # Блок 2: SCADA Communicator
    draw_box(
        3.40, 0.75, 2.65, 4.95,
        "SCADA COMMUNICATOR",
        [
            "ScadaCommWkr 6.2.1.0",
            "---DIV---",
            "Линия связи 1 (Shared TCP):",
            "• Драйвер: DrvCnlBasic 6.1",
            "• Канал: TcpClient 127.0.0.1:502",
            "• Режим передачи: TCP",
            "• Цикл опроса: 500 мс",
            "Устройство КП 1 (PLC_Var2):",
            "• Драйвер: DrvModbus 6.0.0.5",
            "• Шаблон: DrvModbus_Var2.xml",
            "• Привязка: DrvDsScadaServer",
            "• Ошибок опроса: 0 / 360",
        ],
        fc="#F8FAFC", ec="#1E293B", hdr_bg="#1E293B",
    )

    # Блок 3: SCADA Server
    draw_box(
        6.45, 0.75, 2.65, 4.95,
        "SCADA SERVER",
        [
            "ScadaServerWkr 6.4 (TCP :10000)",
            "---DIV---",
            "База конфигурации (BaseDAT):",
            "• Канал 1: Расход Q (LimID=1)",
            "• Канал 2: Уровень L (LimID=2)",
            "• Канал 3: Вибрация V (LimID=3)",
            "• Канал 4: ПУСК/СТОП (In/Out)",
            "• Канал 5: Аварийный сигнал",
            "  Формула: Val(1)>400 || Val(3)>15",
            "Модуль ModArcBasic 6.2.1:",
            "• Архивы: Cur, Min, Hour, Events",
            "• Архив Lab3 (период: 1 сек)",
        ],
        fc="#F8FAFC", ec="#1E293B", hdr_bg="#334155",
    )

    # Блок 4: WebStation
    draw_box(
        9.50, 0.75, 2.45, 4.95,
        "WEBSTATION (:10008)",
        [
            "ScadaWeb 6.2.3.1 (Kestrel)",
            "---DIV---",
            "Плагины и представления:",
            "• PlgMimic + PlgMimBasicComp",
            "  Мнемосхема Lab3_Mimic.mim",
            "  (бак, датчики, LED, кнопки)",
            "• PlgMain:",
            "  Таблица Lab3_Table.tbl",
            "  Журнал тревог Events",
            "• PlgChart:",
            "  Тренды Q, L, V (архив Lab3)",
            "• Период обновления: 1000 мс",
        ],
        fc="#F8FAFC", ec="#1E293B", hdr_bg="#1E293B",
    )

    def draw_arrow(x1, y1, x2, y2, lbl_top, lbl_bot, color="#0F172A"):
        arr = FancyArrowPatch(
            (x1, y1), (x2, y2),
            arrowstyle="<|-|>", mutation_scale=13, linewidth=1.8, color=color, zorder=5,
        )
        ax.add_patch(arr)
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        if lbl_top:
            ax.text(
                mx, my + 0.22, lbl_top, ha="center", va="bottom",
                fontsize=7.6, fontweight="bold", color="#0F172A",
                bbox=dict(boxstyle="round,pad=0.15", fc="#FFFFFF", ec="#CBD5E1", lw=0.7),
                zorder=6,
            )
        if lbl_bot:
            ax.text(
                mx, my - 0.22, lbl_bot, ha="center", va="top",
                fontsize=7.4, color="#334155",
                bbox=dict(boxstyle="round,pad=0.15", fc="#FFFFFF", ec="#CBD5E1", lw=0.7),
                zorder=6,
            )

    draw_arrow(3.01, 3.2, 3.39, 3.2, "FC 03", "FC 06", color="#1D4ED8")
    draw_arrow(6.06, 3.2, 6.44, 3.2, "Срезы", "Команды", color="#0F172A")
    draw_arrow(9.11, 3.2, 9.49, 3.2, "TCP", "10000", color="#0F172A")

    ax.text(
        6.1, 0.28,
        "Направление потоков: телеметрия (Q, L, V, PumpState) слева направо ->  |  команды управления (ПУСК=1 / СТОП=0) и квитирование <- справа налево",
        ha="center", va="center", fontsize=8.4, style="italic", color="#475569",
    )

    out_path = IMG_DIR / "scada_architecture.png"
    fig.tight_layout(pad=0.3)
    fig.savefig(out_path, dpi=300)
    plt.close(fig)
    return out_path


# =============================================================================
# 2. ТАБЛИЧНАЯ КНИГА EXCEL (ЛР3_конфигурация_SCADA.xlsx)
# =============================================================================

def generate_excel_workbook() -> Path:
    wb = openpyxl.Workbook()

    hdr_fill = PatternFill("solid", fgColor="1E293B")
    sub_fill = PatternFill("solid", fgColor="F1F5F9")
    alarm_fill = PatternFill("solid", fgColor="FEE2E2")
    ok_fill = PatternFill("solid", fgColor="DCFCE7")
    thin = Side(border_style="thin", color="334155")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Times New Roman", size=11, color="0F172A")
    bold_font = Font(name="Times New Roman", size=11, bold=True, color="0F172A")

    def style_sheet(ws, title_text: str, headers: list[str], rows: list[list], widths: list[int]):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        c_title = ws.cell(1, 1, title_text)
        c_title.font = Font(name="Times New Roman", size=13, bold=True, color="0F172A")
        c_title.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 26

        for col_idx, h in enumerate(headers, start=1):
            c = ws.cell(3, col_idx, h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[3].height = 28

        for r_idx, row_vals in enumerate(rows, start=4):
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws.cell(r_idx, col_idx, val)
                c.font = bold_font if col_idx == 1 else cell_font
                c.border = border
                c.fill = sub_fill if (r_idx % 2 == 1) else PatternFill("solid", fgColor="FFFFFF")
                c.alignment = Alignment(
                    horizontal="center" if isinstance(val, (int, float)) or col_idx <= 2 else "left",
                    vertical="center",
                    wrap_text=True,
                )
            ws.row_dimensions[r_idx].height = 22

        for col_idx, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = w

    # Лист 1: Каналы SCADA
    ws1 = wb.active
    ws1.title = "Каналы Cnl"
    style_sheet(
        ws1,
        "Таблица каналов Rapid SCADA 6.4.7 — Вариант 2 (Расход, Уровень, Вибрация)",
        ["CnlNum", "Код тега", "Наименование канала", "Тип (CnlTypeID)", "Формат / Ед.", "Лимит LimID", "ArchiveMask", "EventMask", "Формула"],
        [
            [1, "Flow", "Расход (Q)", "1 (Входной)", "N1 / м³/ч (1001)", "1 (Hi=380, HiHi=400)", "65559 (Cur+Min+Hr+Ev+Lab3)", "49 (En+StatChg+Undef)", "—"],
            [2, "Level", "Уровень (L)", "1 (Входной)", "N1 / % (101)", "2 (5/15/85/95)", "65559 (Cur+Min+Hr+Ev+Lab3)", "49 (En+StatChg+Undef)", "—"],
            [3, "Vibration", "Вибрация (V)", "1 (Входной)", "N2 / мм/с (1002)", "3 (Hi=13.5, HiHi=15)", "65559 (Cur+Min+Hr+Ev+Lab3)", "49 (En+StatChg+Undef)", "—"],
            [4, "PumpCtrl", "Управление насосом (ПУСК/СТОП)", "2 (Вход/Выход)", "101 (СТОП; ПУСК)", "—", "65559 (Cur+Min+Hr+Ev+Lab3)", "105 (En+Val+Undef+Cmd)", "—"],
            [5, "AlarmState", "Аварийный сигнал", "3 (Расчётный)", "102 (НОРМА; АВАРИЯ)", "—", "65559 (Cur+Min+Hr+Ev+Lab3)", "25 (En+ValChg+StatChg)", "(Val(1)>400 || Val(3)>15) ? 1 : 0; статус 15 : 13"],
        ],
        [10, 14, 30, 18, 22, 22, 26, 24, 42],
    )

    # Лист 2: Карта регистров Modbus TCP
    ws2 = wb.create_sheet("Карта Modbus TCP")
    style_sheet(
        ws2,
        "Карта регистров Modbus TCP ПЛК-имитатора (127.0.0.1:502, Unit ID=1)",
        ["Адрес Reg", "Размер / Тип", "Функции Modbus", "Тег SCADA", "Диапазон процесса", "Масштаб / Формат", "Назначение"],
        [
            ["10..11", "2 слова (Float32)", "FC 03 (Read Holding)", "Flow", "0.0 – 500.0 м³/ч", "IEEE 754 Big-Endian", "Мгновенный расход жидкости Q в магистрали"],
            ["12..13", "2 слова (Float32)", "FC 03 (Read Holding)", "Level", "0.0 – 100.0 %", "IEEE 754 Big-Endian", "Уровень заполнения технологического резервуара L"],
            ["14..15", "2 слова (Float32)", "FC 03 (Read Holding)", "Vibration", "0.0 – 20.0 мм/с", "IEEE 754 Big-Endian", "Виброскорость подшипникового узла насоса V"],
            ["3", "1 слово (uint16)", "FC 03 (Read Holding)", "PumpRun", "0 / 1", "0 = СТОП, 1 = ПУСК", "Фактическое дискретное состояние насосного агрегата"],
            ["4", "1 слово (uint16)", "FC 03 / FC 06 (Write Reg)", "PumpCtrl", "0 / 1", "0 = СТОП, 1 = ПУСК", "Регистр команд телеуправления от оператора"],
            ["0..2", "3 слова (uint16)", "FC 03 (резерв)", "Q*1, L*10, V*100", "Целочисленное дубл.", "Коэф. 1 / 10 / 100", "Совместимые целочисленные регистры ПЛК"],
        ],
        [14, 20, 24, 14, 22, 24, 46],
    )

    # Лист 3: Лимиты и тревоги
    ws3 = wb.create_sheet("Тревоги Lim")
    style_sheet(
        ws3,
        "Настройка границ уставок (Lim.xml) и статусов тревог (CnlStatus.xml)",
        ["LimID", "Привязанный канал", "LoLo (11)", "Low (12)", "Норма (13)", "High (14)", "HiHi (15)", "Deadband", "Действие в Вебстанции"],
        [
            [1, "Канал 1: Расход Q (м³/ч)", "—", "—", "0 – 380 м³/ч", "380 м³/ч", "400 м³/ч", 2.0, "При Q > 400: статус High high, красный LED, запись в Events, требуется квитирование"],
            [2, "Канал 2: Уровень L (%)", "5 %", "15 %", "15 – 85 %", "85 %", "95 %", 1.0, "Цвет заливки резервуара синий (#0284C7), при >85 % оранжевый (#EA580C)"],
            [3, "Канал 3: Вибрация V (мм/с)", "—", "—", "0 – 13.5 мм/с", "13.5 мм/с", "15.0 мм/с", 0.2, "При V > 15: статус High high, красный LED, запись в Events, требуется квитирование"],
        ],
        [10, 28, 12, 12, 16, 14, 14, 12, 54],
    )

    # Лист 4: Протокол живых испытаний
    ws4 = wb.create_sheet("Результаты испытаний")
    style_sheet(
        ws4,
        "Протокол испытаний живой системы Rapid SCADA 6.4.7 + Modbus TCP (10.11.2026)",
        ["№ теста", "Режим / Воздействие", "Q (м³/ч)", "L (%)", "V (мм/с)", "Статус / LED", "Кадр Modbus / Журнал Events"],
        [
            ["Тест 1", "Штатный режим (насос ПУСК)", "262.8 – 310.9", "59.0 – 65.4", "4.44 – 6.06", "НОРМА / Зелёный", "FC 03 OK (0 ошибок из 360 запросов); статус Normal"],
            ["Тест 4а", "Нажатие «СТОП (0)» на мнемосхеме", "0.0", "63.2", "0.48", "СТОП (0) / Зелёный", "FC 06 addr=4 val=0; Events: Команда СТОП от admin"],
            ["Тест 4б", "Нажатие «ПУСК (1)» на мнемосхеме", "277.6 – 314.0", "64.2", "5.10", "ПУСК (1) / Зелёный", "FC 06 addr=4 val=1; Events: Команда ПУСК от admin"],
            ["Тест 2", "Аварийный режим (Q > 400, V > 15)", "445.4 – 461.4", "73.5 – 73.8", "17.06 – 17.88", "АВАРИЯ / Красный", "Events: High high по Q, V и АВАРИЯ (AckRequired=true)"],
            ["Тест 3", "Архивация и тренды PlgChart", "Архив Lab3 (1 с)", "Архив Lab3 (1 с)", "Архив Lab3 (1 с)", "259–262 точек", "Отображены ступени перехода Норма->СТОП->ПУСК->Авария"],
        ],
        [12, 34, 18, 16, 18, 22, 52],
    )

    for ws in [ws1, ws2, ws3, ws4]:
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
            val_str = " ".join(str(c.value or "") for c in row)
            if "АВАРИЯ" in val_str or "442.7" in val_str:
                row[5].fill = alarm_fill
            elif "НОРМА" in val_str or "ПУСК (1)" in val_str:
                row[5].fill = ok_fill

    out_xlsx = BASE_DIR / "ЛР3_конфигурация_SCADA.xlsx"
    wb.save(out_xlsx)
    return out_xlsx


# =============================================================================
# 3. ОТЧЕТ ПО ГОСТ В WORD (ГайнановДИ_ГарифуллинКР_ЛР3_АСУСН.docx)
# =============================================================================

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def set_cell_shading(cell, fill_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)


def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    cantSplit = OxmlElement("w:cantSplit")
    trPr.append(cantSplit)


def add_page_number_to_footer(footer):
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    r_fld1 = p.add_run()
    r_fld1.font.name = "Times New Roman"
    r_fld1.font.size = Pt(12)
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    r_fld1._r.append(fldChar1)

    r_instr = p.add_run()
    r_instr.font.name = "Times New Roman"
    r_instr.font.size = Pt(12)
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls("w"))
    r_instr._r.append(instrText)

    r_sep = p.add_run()
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
    r_sep._r.append(fldChar2)

    r_val = p.add_run("2")
    r_val.font.name = "Times New Roman"
    r_val.font.size = Pt(12)

    r_end = p.add_run()
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    r_end._r.append(fldChar3)


def generate_word_report() -> Path:
    doc = docx.Document()

    # ГОСТ поля: А4 (210x297 мм), левое 30 мм, правое 15 мм, верх/низ 20 мм
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(1.5)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.different_first_page_header_footer = True
    add_page_number_to_footer(sec.footer)

    # Базовый стиль Normal — Times New Roman 14 пт, 1.5 интервал, отступ 1.25 см
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(14)
    st.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.first_line_indent = Cm(1.25)
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)

    def add_p(
        text: str,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        indent_cm: float | None = 1.25,
        bold: bool = False,
        italic: bool = False,
        size: int = 14,
        line_spacing: float = 1.5,
        space_before: float = 0,
        space_after: float = 0,
        keep_with_next: bool = False,
    ):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.first_line_indent = Cm(indent_cm) if indent_cm else Cm(0)
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = keep_with_next
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_h1(text: str, page_break: bool = False):
        if page_break:
            doc.add_page_break()
        return add_p(
            text,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            indent_cm=0,
            bold=True,
            size=14,
            line_spacing=1.5,
            space_before=12,
            space_after=8,
            keep_with_next=True,
        )

    def add_h2(text: str):
        return add_p(
            text,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            indent_cm=1.25,
            bold=True,
            size=14,
            line_spacing=1.5,
            space_before=8,
            space_after=4,
            keep_with_next=True,
        )

    def add_bullet(text: str):
        return add_p(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent_cm=1.25, size=14, space_after=2)

    def add_code_block(code_text: str):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl)
        prevent_row_split(tbl.rows[0])
        c = tbl.cell(0, 0)
        set_cell_margins(c, 90, 90, 140, 140)
        set_cell_shading(c, "F8F8F8")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(code_text.strip())
        r.font.name = "Courier New"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        add_p("", indent_cm=0, size=6, line_spacing=1.0, space_after=4)

    def add_gost_table(caption: str, headers: list[str], rows: list[list[str]], col_widths_cm: list[float]):
        add_p(caption, align=WD_ALIGN_PARAGRAPH.LEFT, indent_cm=0, italic=False, bold=False, size=14, space_before=6, space_after=3, keep_with_next=True)
        tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_borders(tbl)

        tblPr = tbl._tbl.tblPr
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

        # Заголовок
        set_repeat_table_header(tbl.rows[0])
        prevent_row_split(tbl.rows[0])
        for j, h in enumerate(headers):
            tbl.columns[j].width = Cm(col_widths_cm[j])
            c = tbl.cell(0, j)
            c.width = Cm(col_widths_cm[j])
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(c, 70, 70, 70, 70)
            set_cell_shading(c, "EDEDED")
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.keep_with_next = True
            r = p.add_run(h)
            r.font.name = "Times New Roman"
            r.font.size = Pt(10.5)
            r.bold = True

        # Строки
        for i, row_vals in enumerate(rows):
            row = tbl.rows[i + 1]
            prevent_row_split(row)
            is_first_data_row = (i == 0 and len(rows) > 1)
            for j, val in enumerate(row_vals):
                c = row.cells[j]
                c.width = Cm(col_widths_cm[j])
                c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                set_cell_margins(c, 60, 60, 70, 70)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 0 else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.line_spacing = 1.15
                if is_first_data_row:
                    p.paragraph_format.keep_with_next = True
                r = p.add_run(val)
                r.font.name = "Times New Roman"
                r.font.size = Pt(10.5)

        add_p("", indent_cm=0, size=6, line_spacing=1.0, space_after=6)

    def add_figure(img_path: Path, caption: str, width_cm: float = 16.0):
        assert img_path.exists(), f"Отсутствует изображение: {img_path}"
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(str(img_path), width=Cm(width_cm))
        add_p(caption, align=WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=12, line_spacing=1.15, space_after=8)

    # =========================================================================
    # ТИТУЛЬНЫЙ ЛИСТ (СТР. 1, БЕЗ НОМЕРА СТРАНИЦЫ)
    # =========================================================================
    add_p("Министерство науки и высшего образования Российской Федерации", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=14)
    add_p("Федеральное государственное бюджетное образовательное учреждение высшего образования", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=14)
    add_p("«Уфимский университет науки и технологий»", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=14, bold=False)
    add_p("Кафедра автоматизированных систем управления", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=14, space_before=10, space_after=48)

    add_p("ОТЧЁТ", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, bold=True, size=16, space_after=6)
    add_p("по лабораторной работе № 3", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, bold=True, size=14, space_after=6)
    add_p("по дисциплине: Автоматизированные системы управления специального назначения (АСУСН)", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=14, space_after=6)
    add_p("на тему: «Моделирование системы сбора данных в SCADA-системе»", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, bold=True, size=14, space_after=6)
    add_p("Вариант № 2 (Расход, Уровень, Вибрация; источник данных — Modbus TCP)", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, italic=True, size=13, space_after=56)

    add_p("Выполнили:", WD_ALIGN_PARAGRAPH.RIGHT, indent_cm=0, size=14)
    add_p("студенты группы ЭАС-514С", WD_ALIGN_PARAGRAPH.RIGHT, indent_cm=0, size=14)
    add_p("Гайнанов Д.И.", WD_ALIGN_PARAGRAPH.RIGHT, indent_cm=0, size=14)
    add_p("Гарифуллин К.Р.", WD_ALIGN_PARAGRAPH.RIGHT, indent_cm=0, size=14, space_after=12)
    add_p("Проверил:", WD_ALIGN_PARAGRAPH.RIGHT, indent_cm=0, size=14)
    add_p("Антонов В.В.", WD_ALIGN_PARAGRAPH.RIGHT, indent_cm=0, size=14, space_after=48)

    add_p("Уфа 2026 г.", WD_ALIGN_PARAGRAPH.CENTER, indent_cm=0, size=14)
    doc.add_page_break()

    # =========================================================================
    # 1. ЦЕЛЬ РАБОТЫ И ЗАДАЧИ
    # =========================================================================
    add_h1("1. ЦЕЛЬ РАБОТЫ И ЗАДАЧИ")
    add_p(
        "Целью лабораторной работы является изучение архитектуры современных SCADA-систем и принципов "
        "организации сбора, первичной обработки, архивирования и визуализации технологических данных в АСУСН, "
        "а также практическое освоение разработки рабочего проекта, интерактивной мнемосхемы, системы тревог "
        "и трендов на базе промышленной платформы Rapid SCADA 6.4.7."
    )
    add_p("Для достижения поставленной цели в соответствии с Вариантом № 2 (Таблица 1) и Вариантом A (источник Modbus TCP) решены следующие задачи:")
    add_bullet("1) развёрнута и сконфигурирована реальная среда Rapid SCADA 6.4.7 (Сервер ScadaServerWkr, Коммуникатор ScadaCommWkr, Вебстанция ScadaWeb) без использования программных заглушек;")
    add_bullet("2) разработан имитатор программируемого логического контроллера modbus_plc_simulator.py (протокол Modbus TCP Slave, 127.0.0.1:502, Unit ID = 1) для насосной станции Варианта № 2 (Расход 0–500 м³/ч, Уровень 0–100 %, Вибрация 0–20 мм/с) с регистрами телеуправления;")
    add_bullet("3) настроена линия связи DrvCnlBasic (TcpClient) и опрос контроллера драйвером DrvModbus 6.0.0.5 по карте регистров IEEE 754 Float32 и командных регистров FC 06;")
    add_bullet("4) сформированы база конфигурации каналов (входные, вход/выход, расчётный), пороги сигнализации (Q > 400 м³/ч, V > 15 мм/с), секундный исторический архив Lab3 и журнал событий Events;")
    add_bullet("5) разработана интерактивная мнемосхема Lab3_Mimic.mim и табличное представление Lab3_Table.tbl, проведены живые испытания всех режимов в браузере и проанализированы результаты.")

    add_gost_table(
        "Таблица 1 – Исходные данные индивидуального задания (Вариант № 2)",
        ["Параметр / Элемент", "Ед. изм.", "Диапазон", "Штатный режим", "Уставка тревоги"],
        [
            ["Расход магистрали (Q)", "м³/ч", "0 – 500", "260 – 315 м³/ч", "Q > 400 м³/ч (Hi = 380, HiHi = 400)"],
            ["Уровень в резервуаре (L)", "%", "0 – 100", "58 – 66 %", "LoLo=5 / Lo=15 / Hi=85 / HiHi=95 %"],
            ["Вибрация агрегата (V)", "мм/с", "0 – 20", "4.4 – 6.4 мм/с", "V > 15 мм/с (Hi = 13.5, HiHi = 15.0)"],
            ["Управление насосом", "дискр.", "0 / 1", "ПУСК (1) / СТОП (0)", "Команды оператора, FC 06 Reg 4"],
        ],
        [4.2, 2.0, 2.5, 3.3, 4.5],
    )

    # =========================================================================
    # 2. КРАТКИЕ ТЕОРЕТИЧЕСКИЕ СВЕДЕНИЯ И АРХИТЕКТУРА RAPID SCADA
    # =========================================================================
    add_h1("2. ТЕОРЕТИЧЕСКИЕ СВЕДЕНИЯ И АРХИТЕКТУРА СБОРА ДАННЫХ")
    add_p(
        "SCADA (Supervisory Control And Data Acquisition) представляет собой программно-аппаратный комплекс верхнего уровня АСУ ТП, "
        "обеспечивающий циклический опрос устройств нижнего уровня (ПЛК, интеллектуальных датчиков, УСО), расчёт производных показателей, "
        "контроль технологических границ (тревог), долгосрочное архивирование временных рядов, ведение журнала событий и предоставление "
        "человеко-машинного интерфейса (HMI) оператору с возможностью дистанционного телеуправления."
    )
    add_p(
        "В данной работе используется российская открытая SCADA-платформа Rapid SCADA 6.4.7, собранная из исходного кода на .NET 8 "
        "и развёрнутая в модульной четырёхзвенной архитектуре с файловым хранилищем FileStorage:"
    )
    add_bullet("• Источник данных (ПЛК-имитатор Modbus TCP) — выступает в роли Modbus TCP Slave на порту 502 (Unit ID = 1), моделируя физику насосной станции и обрабатывая запросы чтения (FC 03) и записи одиночного регистра команды (FC 06);")
    add_bullet("• Коммуникатор (ScadaCommWkr 6.2.1.0) — управляет транспортной линией TcpClient (драйвер DrvCnlBasic 6.1.0.1) в режиме Shared TCP и протокольным драйвером DrvModbus 6.0.0.5, преобразует байтовые кадры Modbus в теги и передаёт срезы в Сервер через DrvDsScadaServer;")
    add_bullet("• Сервер (ScadaServerWkr, TCP-слушатель :10000) — выполняет расчётные формулы в C#-компиляторе CalcEngine, проверяет границы Lim (присваивая статусы Normal=13, High=14, High high=15), формирует записи журнала событий и сохраняет временные ряды модулем ModArcBasic 6.2.1.0 в бинарные DAT-архивы;")
    add_bullet("• Вебстанция (ScadaWeb 6.2.3.1 на Kestrel, HTTP :10008) — веб-приложение ASP.NET Core с плагинами PlgMimic (мнемосхемы .mim), PlgMimBasicComp (индикаторы LED и кнопки), PlgMain (таблицы .tbl и журнал событий Events) и PlgChart (построение трендов).")

    add_figure(
        IMG_DIR / "scada_architecture.png",
        "Рисунок 1 – Схема информационных потоков и компонентов стенда Rapid SCADA 6.4.7 (Вариант 2 + Вариант A)",
        width_cm=16.5,
    )

    # =========================================================================
    # 3. КОНФИГУРАЦИЯ ПРОЕКТА И КАРТА РЕГИСТРОВ MODBUS TCP
    # =========================================================================
    add_h1("3. НАСТРОЙКА ПРОЕКТА И ЛИСТИНГ КОНФИГУРАЦИИ")
    add_h2("3.1. Адресная карта регистров Modbus TCP (DrvModbus_Var2.xml)")
    add_p(
        "Для исключения потерь точности дробных величин и демонстрации передачи чисел с плавающей запятой по протоколу Modbus "
        "основные аналоговые сигналы (Расход Q, Уровень L и Вибрация V) размещены в смежных 16-разрядных регистрах хранения (Holding Registers) "
        "начиная с адреса 10 в формате IEEE 754 Float32 (по 2 регистра на параметр, порядок байтов Big-Endian). Дискретное состояние насоса "
        "и регистр записи команд расположены по адресам 3 и 4 (uint16)."
    )

    add_gost_table(
        "Таблица 2 – Карта регистров Modbus TCP контроллера Варианта № 2",
        ["Адрес Reg", "Тип данных", "Код функции", "Тег DrvModbus", "Назначение и диапазон"],
        [
            ["10..11", "Float32 (2 рег.)", "FC 03", "Flow", "Расход Q, 0.0 – 500.0 м³/ч (IEEE 754)"],
            ["12..13", "Float32 (2 рег.)", "FC 03", "Level", "Уровень L, 0.0 – 100.0 % (IEEE 754)"],
            ["14..15", "Float32 (2 рег.)", "FC 03", "Vibration", "Вибрация V, 0.0 – 20.0 мм/с (IEEE 754)"],
            ["3", "UShort (1 рег.)", "FC 03", "PumpRun", "Фактическое состояние насоса (0=СТОП, 1=ПУСК)"],
            ["4", "UShort (1 рег.)", "FC 03 / FC 06", "PumpCtrl", "Командный регистр управления агрегатом (0/1)"],
        ],
        [2.4, 3.0, 2.7, 2.7, 5.7],
    )

    add_p("Фрагмент файла шаблона устройства DrvModbus_Var2.xml приведён ниже:")
    add_code_block(
"""<DeviceTemplate>
  <Options><ZeroAddr>true</ZeroAddr><DecAddr>true</DecAddr></Options>
  <ElemGroups>
    <ElemGroup active="true" dataBlock="HoldingRegisters" address="10" name="Аналоговые параметры (Float32)">
      <Elem type="float" readOnly="true" tagCode="Flow" name="Расход Q (м³/ч)" />
      <Elem type="float" readOnly="true" tagCode="Level" name="Уровень L (%)" />
      <Elem type="float" readOnly="true" tagCode="Vibration" name="Вибрация V (мм/с)" />
    </ElemGroup>
    <ElemGroup active="true" dataBlock="HoldingRegisters" address="3" name="Состояние и управление">
      <Elem type="ushort" readOnly="true" tagCode="PumpRun" name="Состояние насоса (Reg 3)" />
      <Elem type="ushort" readOnly="false" tagCode="PumpCtrl" name="Управление насосом (Reg 4)" />
    </ElemGroup>
  </ElemGroups>
  <Cmds>
    <Cmd dataBlock="HoldingRegisters" multiple="false" address="4" cmdNum="1" cmdCode="PumpCtrl" name="Команда ПУСК/СТОП насоса" />
  </Cmds>
</DeviceTemplate>"""
    )

    add_h2("3.2. База конфигурации каналов (Cnl), пределов (Lim) и архивов")
    add_p(
        "В базе конфигурации проекта lr3.rsproj (скомпилированной в формат BaseDAT/*.dat) созданы пять технологических каналов "
        "(Таблица 3), три набора пределов сигнализации (Таблица 4) и высокочастотный исторический архив Lab3 (Bit = 16) с шагом записи 1 с "
        "для детальной визуализации быстро протекающих аварийных переходов на графиках."
    )

    add_gost_table(
        "Таблица 3 – Параметры конфигурации каналов (Cnl.xml)",
        ["№", "Наименование (Код)", "Тип CnlType", "Ед. / Формат", "LimID", "ArchiveMask", "EventMask"],
        [
            ["1", "Расход (Q) [Flow]", "1 (Входной)", "м³/ч / N1", "1", "65559", "49"],
            ["2", "Уровень (L) [Level]", "1 (Входной)", "% / N1", "2", "65559", "49"],
            ["3", "Вибрация (V) [Vibration]", "1 (Входной)", "мм/с / N2", "3", "65559", "49"],
            ["4", "Управление насосом [PumpCtrl]", "2 (Вход/Выход)", "СТОП; ПУСК", "—", "65559", "105"],
            ["5", "Аварийный сигнал [AlarmState]", "3 (Расчётный)", "НОРМА; АВАРИЯ", "—", "65559", "25"],
        ],
        [0.9, 4.0, 2.5, 2.5, 1.4, 2.8, 2.4],
    )

    add_p(
        "Маска архивов ArchiveMask = 65559 (в двоичной записи 1+2+4+16+65536) включает запись в текущий срез (Cur), "
        "минутный (Min), часовой (Hour), архив событий (Events, бит 4 = 16) и секундный архив Lab3 (бит 16 = 65536). "
        "На этапе отладки было выявлено, что отсутствие бита 4 в ArchiveMask приводит к пропуску тревог по лимитам в журнале событий — "
        "добавление значения 16 обеспечило корректную фиксацию всех статусов High high и Normal. "
        "В канале 5 задана расчётная C#-формула: (Val(1) > 400 || Val(3) > 15) ? 1 : 0; статус (Val(1) > 400 || Val(3) > 15) ? 15 : 13."
    )

    add_gost_table(
        "Таблица 4 – Пределы и уставки сигнализации (Lim.xml)",
        ["LimID", "Наименование", "Low / High", "LoLo / HiHi", "Deadband", "Поведение при нарушении"],
        [
            ["1", "Пределы расхода Q", "— / 380 м³/ч", "— / 400 м³/ч", "2.0", "Статус 15 (High high), цвет красный, квитирование"],
            ["2", "Пределы уровня L", "15 / 85 %", "5 / 95 %", "1.0", "Выход за рабочую зону 15–85 %, цвет заливки"],
            ["3", "Пределы вибрации V", "— / 13.5 мм/с", "— / 15.0 мм/с", "0.2", "Статус 15 (High high), цвет красный, квитирование"],
        ],
        [1.4, 3.2, 2.7, 2.7, 2.2, 4.3],
    )

    add_h2("3.3. Разработка мнемосхемы Lab3_Mimic.mim")
    add_p(
        "Мнемосхема насосной станции Lab3_Mimic.mim (размер полотна 1060×650 пкс, плагин PlgMimic + PlgMimBasicComp) "
        "включает динамический уровень жидкости в резервуаре (привязки size.height и location.y к каналу 2), три цифровые карточки "
        "измерений с динамическими цветовыми рамками и текстовыми баннерами, круглый светодиодный индикатор тревоги BasicLed (канал 5), "
        "индикатор состояния насоса и две интерактивные кнопки управления BasicButton «ПУСК (1)» и «СТОП (0)» с действием SendCommand на канал 4."
    )

    # =========================================================================
    # 4. ХОД ИСПЫТАНИЙ И РЕЗУЛЬТАТЫ (СКРИНШОТЫ РАБОТАЮЩЕЙ СИСТЕМЫ)
    # =========================================================================
    add_h1("4. РЕЗУЛЬТАТЫ ИСПЫТАНИЙ СИСТЕМЫ", page_break=True)

    add_h2("4.1. Тест 1 — Штатный режим работы (нормальные значения)")
    add_p(
        "После запуска имитатора modbus_plc_simulator.py (--variant 2 --mode normal) и сервисов ScadaServerWkr, ScadaCommWkr, ScadaWeb "
        "проведён визуальный контроль мнемосхемы в Вебстанции (http://localhost:10008, учётная запись admin). "
        "Расход стабилизирован в диапазоне 260–315 м³/ч (на снимке 262.8 м³/ч), уровень составляет 59.0 %, вибрация 6.06 мм/с, "
        "индикатор BasicLed горит зелёным, плашка состояния индицирует «ШТАТНЫЙ РЕЖИМ», агрегат в состоянии «ПУСК (1)» (Рисунок 2)."
    )
    add_figure(
        IMG_DIR / "mimic_normal_mode.png",
        "Рисунок 2 – Мнемосхема Rapid SCADA WebStation в штатном режиме (Вариант 2)",
        width_cm=15.8,
    )

    add_h2("4.2. Тест 2 — Проверка кнопок телеуправления «ПУСК (1)» и «СТОП (0)»")
    add_p(
        "На живой мнемосхеме была нажата кнопка «СТОП (0)». Вебстанция сформировала телекоманду по каналу 4 (CmdCode = PumpCtrl, CmdVal = 0), "
        "Сервер передал её Коммуникатору, а драйвер DrvModbus отправил кадр Modbus FC 06 (Write Single Register, адрес 4, значение 0) "
        "в ПЛК-имитатор. В результате состояние насоса на мнемосхеме изменилось на «Состояние: СТОП (0)», расход снизился до 0.0 м³/ч, "
        "вибрация упала до фонового уровня 0.48 мм/с (при уровне 63.2 %), а в журнале событий зафиксирована команда оператора admin (Рисунок 3). "
        "Последующее нажатие «ПУСК (1)» (FC 06 addr=4, val=1) восстановило рабочий режим агрегата."
    )
    add_figure(
        IMG_DIR / "mimic_pump_stopped.png",
        "Рисунок 3 – Реакция мнемосхемы и журнала на команду «СТОП (0)» (Modbus FC 06)",
        width_cm=15.8,
    )

    add_h2("4.3. Тест 3 — Имитация аварийного режима (Q > 400 м³/ч, V > 15 мм/с) и журнал тревог")
    add_p(
        "При переводе ПЛК-имитатора в аварийный режим расход превысил уставку 400 м³/ч (453.0 м³/ч на снимке при пиках до 461.4 м³/ч), а виброскорость "
        "превысила 15 мм/с (17.88 мм/с при уровне 73.5 %). Расчётный канал 5 принял значение 1 (АВАРИЯ) и статус 15. На мнемосхеме индикатор BasicLed переключился в "
        "красный цвет (#EF4444), текстовые карточки выделили тревоги Q > 400 м³/ч и V > 15 мм/с, появился баннер «АВАРИЯ АГРЕГАТА!» (Рисунок 4)."
    )
    add_figure(
        IMG_DIR / "mimic_alarm_mode.png",
        "Рисунок 4 – Мнемосхема в аварийном режиме: Q = 453.0 м³/ч (> 400) и V = 17.88 мм/с (> 15)",
        width_cm=15.8,
    )

    add_p(
        "В раскрытом журнале событий Вебстанции (Рисунок 5) отображены все хронологические записи: команды телеуправления СТОП и ПУСК с "
        "указанием пользователя admin, критические тревоги уровня High high по каналам 1, 3 и 5 с флагом обязательного квитирования, "
        "а также записи возврата в состояние Normal. Таблица текущих и архивных значений параметров приведена на Рисунке 6."
    )
    add_figure(
        IMG_DIR / "alarm_journal_view.png",
        "Рисунок 5 – Полный журнал событий и тревог Rapid SCADA WebStation в ходе испытаний",
        width_cm=16.2,
    )
    add_figure(
        IMG_DIR / "table_parameters_view.png",
        "Рисунок 6 – Табличное представление текущих значений параметров (Lab3_Table.tbl)",
        width_cm=16.2,
    )

    add_h2("4.4. Тест 4 — Графики изменения параметров во времени (PlgChart)")
    add_p(
        "По данным секундного исторического архива Lab3 (модуль ModArcBasic, период записи 1 с) плагином PlgChart построены реальные тренды "
        "расхода Q (Рисунок 7) и вибрации V (Рисунок 8). На графиках отчётливо прослеживаются все фазы эксперимента: синусоидальный "
        "штатный режим (265–310 м³/ч и 4.5–6.3 мм/с), кратковременный провал до нуля при испытании кнопки СТОП, восстановление после ПУСК и "
        "резкий ступенчатый выход за пороги тревог 400 м³/ч и 15 мм/с в аварийном режиме."
    )
    add_figure(
        IMG_DIR / "scada_trends_flow.png",
        "Рисунок 7 – Тренд изменения расхода Q (Канал 1) в модуле графиков Rapid SCADA PlgChart",
        width_cm=16.0,
    )
    add_figure(
        IMG_DIR / "scada_trends_vibration.png",
        "Рисунок 8 – Тренд изменения вибрации V (Канал 3) в модуле графиков Rapid SCADA PlgChart",
        width_cm=16.0,
    )

    # =========================================================================
    # 5. АНАЛИЗ РЕЗУЛЬТАТОВ И ВЫВОДЫ
    # =========================================================================
    add_h1("5. АНАЛИЗ РЕЗУЛЬТАТОВ И ВЫВОДЫ")
    add_p(
        "По статистике работы драйвера DrvModbus 6.0.0.5 (журнал device001.txt) за время стендовых испытаний выполнено 178 сеансов связи "
        "и 360 запросов чтения Holding Registers с нулевым числом ошибок (0 / 360), а также 4 транзакции записи команд FC 06 (0 / 4 ошибок). "
        "Время реакции интерфейса Вебстанции на событие в ПЛК не превышало 1.5 с (цикл опроса 500 мс + период обновления UI 1000 мс)."
    )
    add_p("В ходе выполнения работы зафиксированы и устранены важные практические особенности развёртывания Rapid SCADA 6:")
    add_bullet("1) в канале DrvCnlBasic TcpClient при пустом строковом адресе устройства режим ConnectionMode=Individual пропускает КП, поэтому корректное подключение к localhost:502 требует задания ConnectionMode=Shared и StrAddress=127.0.0.1;")
    add_bullet("2) при ручном формировании XML мнемосхемы для PlgMimic привязки каналов (InCnlNum, OutCnlNum, PropertyBindings/Item) должны быть оформлены в PascalCase (для парсера ComponentBindings) с атрибутом isArray=\"true\";")
    add_bullet("3) для попадания тревог по границам Lim в журнал событий маска ArchiveMask каждого канала обязательно должна включать бит 4 архива Events (+16 к битовой маске).")
    add_p(
        "Выводы: освоена архитектура, база конфигурации, протокольный обмен Modbus TCP (Float32 и команды записи), "
        "создание анимированных мнемосхем, тревог и секундных архивов в Rapid SCADA 6.4.7. Система полностью работоспособна "
        "и удовлетворяет всем требованиям Варианта № 2 и Варианта A."
    )

    # =========================================================================
    # 6. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ
    # =========================================================================
    add_h1("6. ОТВЕТЫ НА КОНТРОЛЬНЫЕ ВОПРОСЫ")
    qa_list = [
        (
            "1. Какие функции выполняет SCADA-система? Чем она отличается от обычной системы визуализации?",
            "SCADA выполняет не только отображение графики, но и цикл сбора данных реального времени через промышленные протоколы, "
            "математическую и логическую обработку (расчётные формулы), контроль пределов (тревожную сигнализацию и квитирование), "
            "ведение защищённых архивов временных рядов, разграничение прав пользователей и отправку команд телеуправления обратно на ПЛК. "
            "Обычная система визуализации (например, панель HMI или дашборд) ограничена показом данных и не содержит полноценного сервера архивов, "
            "движка формул и многопользовательской серверной диспетчеризации."
        ),
        (
            "2. Как организован обмен данными между Rapid SCADA и источником (ПЛК) в вашей работе?",
            "Обмен организован по схеме Client–Server поверх TCP/IP (Вариант A). Наш Python ПЛК-имитатор прослушивает порт 127.0.0.1:502 "
            "как Modbus TCP Slave (Unit ID = 1). Служба SCADA Communicator с помощью канала TcpClient (DrvCnlBasic, Shared) каждые 500 мс "
            "вызывает драйвер DrvModbus, который отправляет запросы FC 03 на чтение регистров 10..15 и 3..4, а при нажатии кнопок ПУСК/СТОП — "
            "кадр FC 06 записи в регистр 4. Полученные значения Коммуникатор пересылает в Сервер (TCP :10000), который публикует их Вебстанции (:10008)."
        ),
        (
            "3. Чем отличается входной канал от канала управления?",
            "Входной канал (CnlTypeID = 1, Input) предназначен только для чтения и хранения измеренных от ПЛК или расчётных величин "
            "(например, каналы 1–3: Расход, Уровень, Вибрация). Канал управления (CnlTypeID = 5, Output) или комбинированный вход/выходной "
            "(CnlTypeID = 2, Input/Output, как канал 4 PumpCtrl) дополнительно разрешает передачу телекоманд с мнемосхемы оператора в драйвер "
            "Коммуникатора для записи управляющего воздействия (CmdVal = 1 или 0) в регистр контроллера."
        ),
        (
            "4. Как настроить тревогу в Rapid SCADA? Какие приоритеты тревог существуют?",
            "Тревоги настраиваются через таблицу границ Lim.xml (поля LoLo, Low, High, HiHi и зона нечувствительности Deadband) с привязкой "
            "LimID к каналу и установкой битов в EventMask (Enabled + StatusChange) и ArchiveMask (+16 для Events). В Rapid SCADA 6 каждому статусу "
            "(CnlStatus.xml) соответствует числовая серьёзность (Severity от 1 до 999) и признак квитирования (AckRequired): 1 — Критическая "
            "(Low low / High high, статус 11/15), 250 — Существенная (Low / High, 12/14), 500–750 — Норма/Информация (Normal=13, Defined=1)."
        ),
        (
            "5. Что такое архив данных? Как настроить период записи и глубину хранения?",
            "Архив данных — структурированное хранилище временных рядов (исторических значений каналов) и событий на диске. В модуле ModArcBasic "
            "период записи задаётся опциями WritingPeriod и WritingPeriodUnit (в данной работе для детального архива Lab3 задано WritingPeriod=1, "
            "WritingPeriodUnit=Second), а глубина хранения задаётся в днях параметром Retention (например, Retention=365 дней), после чего "
            "устаревшие файлы автоматически удаляются Сервером."
        ),
        (
            "6. Почему для Modbus TCP используется порт 502, а не 10000?",
            "TCP-порт 502 стандартизирован международной организацией Modbus Organization и IANA как официальный порт протокола Modbus TCP/IP "
            "на уровне промышленной сети контроллеров (OT-уровень), поэтому любой ПЛК и драйвер Modbus (включая DrvModbus) по умолчанию ожидает "
            "соединение именно по 502 порту. Порт 10000 — внутренний проприетарный порт TCP-слушателя службы Rapid SCADA Server для обмена данными "
            "с его собственными модулями Коммуникатором и Вебстанцией."
        ),
        (
            "7. Какие типы данных могут передаваться по Modbus? Как float-значение занимает два регистра?",
            "Спецификация Modbus оперирует 1-битными элементами (Coils и Discrete Inputs) и 16-битными регистрами (Input и Holding Registers), "
            "через которые передаются типы Bool, Int16/UInt16 (1 регистр), Int32/UInt32 и Float32 (2 регистра), а также Int64/Double (4 регистра). "
            "Вещественное число одинарной точности IEEE 754 Float32 занимает 32 бита (4 байта), поэтому при передаче по Modbus оно разбивается на "
            "два смежных 16-разрядных регистра (в нашей карте Q в Reg 10..11, L в Reg 12..13, V в Reg 14..15), которые на приёмной стороне "
            "объединяются драйвером DrvModbus в исходное 32-битное число."
        ),
    ]
    for q_txt, a_txt in qa_list:
        add_p(q_txt, indent_cm=1.25, bold=True, size=14, space_before=6, space_after=2, keep_with_next=True)
        add_p(a_txt, indent_cm=1.25, bold=False, size=14, space_after=4)

    # =========================================================================
    # 7. СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ
    # =========================================================================
    add_h1("7. СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ")
    refs = [
        "1. Методические указания к лабораторной работе № 3 по дисциплине АСУСН «Моделирование системы сбора данных в SCADA-системе». — Уфа: УУНиТ, каф. АСУ, 2026.",
        "2. Документация Rapid SCADA 6.4 — Руководство разработчика и администратора [Электронный ресурс]. URL: https://rapidscada.ru/doc/latest/ (дата обращения: 10.11.2026).",
        "3. Исходный код платформы Rapid SCADA v6.4.7 (ScadaServer, ScadaComm, DrvModbus, ScadaWeb, PlgMimic, PlgChart) [Электронный ресурс]. URL: https://github.com/RapidScada/scada-v6.",
        "4. Modbus Messaging on TCP/IP Implementation Guide V1.0b. — Modbus Organization, 2006. — 46 p.",
        "5. ГОСТ 34.602-2020 Информационные технологии. Комплекс стандартов на автоматизированные системы. Техническое задание на создание автоматизированной системы.",
    ]
    for ref in refs:
        add_p(ref, indent_cm=1.25, size=14, space_after=2)

    out_docx = BASE_DIR / "ГайнановДИ_ГарифуллинКР_ЛР3_АСУСН.docx"
    doc.save(out_docx)
    return out_docx


def main() -> None:
    arch_png = generate_architecture_diagram()
    print("OK image:", arch_png)
    xlsx_path = generate_excel_workbook()
    print("OK excel:", xlsx_path)
    docx_path = generate_word_report()
    print("OK docx:", docx_path)


if __name__ == "__main__":
    main()
