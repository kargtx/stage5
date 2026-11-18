# -*- coding: utf-8 -*-
"""
maintenance_viewer.py
Интерактивный графический интерфейс (GUI) диспетчера и инженера по техническому
обслуживанию автоматизированной системы специального назначения (АС СН НПС).

Авторы: Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С
УУНиТ, Кафедра АСУ, 2026 г.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import openpyxl

# Настройки путей
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(SCRIPT_DIR, "ПР5_план_и_регламент_ТО_НПС.xlsx")

class MaintenanceViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("АС СН НПС — Автоматизированное рабочее место планирования ТО и ЗИП (ПР №5)")
        self.root.geometry("1100x720")
        self.root.minsize(950, 600)

        # Стилизация ttk
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        self.style.configure(".", font=("Segoe UI", 10))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e2e8f0")
        self.style.configure("Treeview", rowheight=24)
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[12, 6])

        # Загрузка данных из Excel
        self.wb = None
        self.load_data()

        # Построение интерфейса
        self.create_header()
        self.create_tabs()
        self.create_statusbar()

    def load_data(self):
        if os.path.exists(EXCEL_PATH):
            try:
                self.wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
            except Exception as e:
                messagebox.showerror("Ошибка загрузки", f"Не удалось открыть расчетную книгу Excel:\n{e}")

    def create_header(self):
        header_frame = tk.Frame(self.root, bg="#1e3a8a", height=65)
        header_frame.pack(side=tk.TOP, fill=tk.X)
        header_frame.pack_propagate(False)

        lbl_title = tk.Label(header_frame, text="ПЛАН ТЕХНИЧЕСКОГО ОБСЛУЖИВАНИЯ И РЕГЛАМЕНТ АС СН НПС",
                             font=("Segoe UI", 13, "bold"), fg="#ffffff", bg="#1e3a8a")
        lbl_title.pack(anchor="w", padx=20, pady=(8, 0))

        lbl_sub = tk.Label(header_frame,
                           text="Практическая работа №5 | УУНиТ, Кафедра АСУ, 2026 г. | Гайнанов Д.И., Гарифуллин К.Р., гр. ЭАС-514С",
                           font=("Segoe UI", 9), fg="#93c5fd", bg="#1e3a8a")
        lbl_sub.pack(anchor="w", padx=20, pady=(2, 8))

    def create_tabs(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_equip = ttk.Frame(self.notebook)
        self.tab_matrix = ttk.Frame(self.notebook)
        self.tab_recs = ttk.Frame(self.notebook)
        self.tab_sop = ttk.Frame(self.notebook)
        self.tab_zip = ttk.Frame(self.notebook)
        self.tab_log = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_equip, text=" 📋 Оборудование НПС ")
        self.notebook.add(self.tab_matrix, text=" ⚙️ Матрица видов ТО ")
        self.notebook.add(self.tab_recs, text=" 📜 Технологические регламенты ")
        self.notebook.add(self.tab_sop, text=" 🚨 Нештатные ситуации ")
        self.notebook.add(self.tab_zip, text=" 📦 ЗИП и метрология ")
        self.notebook.add(self.tab_log, text=" ✍️ Журнал ТО и аварий ")

        self.build_tab_equip()
        self.build_tab_matrix()
        self.build_tab_recs()
        self.build_tab_sop()
        self.build_tab_zip()
        self.build_tab_log()

    def create_statusbar(self):
        status_frame = tk.Frame(self.root, bg="#f1f5f9", height=26, relief=tk.RIDGE, bd=1)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        status_frame.pack_propagate(False)

        self.status_lbl = tk.Label(status_frame, text="Готов к работе. Комплекс технических средств АС СН НПС в штатном режиме.",
                                   font=("Segoe UI", 9), fg="#334155", bg="#f1f5f9")
        self.status_lbl.pack(side=tk.LEFT, padx=10)

    # -------------------------------------------------------------------------
    # Вкладка 1: Оборудование
    # -------------------------------------------------------------------------
    def build_tab_equip(self):
        top_bar = ttk.Frame(self.tab_equip)
        top_bar.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(top_bar, text="Поиск / Фильтр:").pack(side=tk.LEFT, padx=(0, 6))
        self.ent_search_equip = ttk.Entry(top_bar, width=30)
        self.ent_search_equip.pack(side=tk.LEFT, padx=(0, 10))
        self.ent_search_equip.bind("<KeyRelease>", lambda e: self.filter_equip())

        btn_refresh = ttk.Button(top_bar, text="Сбросить фильтр", command=self.reset_equip_filter)
        btn_refresh.pack(side=tk.LEFT)

        # Таблица Treeview
        cols = ("id", "name", "model", "cat", "qty", "loc", "crit")
        self.tree_equip = ttk.Treeview(self.tab_equip, columns=cols, show="headings", selectmode="browse")
        self.tree_equip.heading("id", text="Поз. ID")
        self.tree_equip.heading("name", text="Наименование оборудования")
        self.tree_equip.heading("model", text="Тип / Модель")
        self.tree_equip.heading("cat", text="Категория КТС")
        self.tree_equip.heading("qty", text="Кол-во")
        self.tree_equip.heading("loc", text="Место установки")
        self.tree_equip.heading("crit", text="Критичность")

        self.tree_equip.column("id", width=80, anchor="center")
        self.tree_equip.column("name", width=240, anchor="w")
        self.tree_equip.column("model", width=190, anchor="w")
        self.tree_equip.column("cat", width=140, anchor="w")
        self.tree_equip.column("qty", width=60, anchor="center")
        self.tree_equip.column("loc", width=220, anchor="w")
        self.tree_equip.column("crit", width=130, anchor="center")

        sb = ttk.Scrollbar(self.tab_equip, orient=tk.VERTICAL, command=self.tree_equip.yview)
        self.tree_equip.configure(yscrollcommand=sb.set)
        self.tree_equip.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(0, 10))
        sb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=(0, 10))

        self.populate_equip_table()

    def populate_equip_table(self):
        for item in self.tree_equip.get_children():
            self.tree_equip.delete(item)

        if not self.wb or "Перечень_Оборудования" not in self.wb.sheetnames:
            return

        ws = self.wb["Перечень_Оборудования"]
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[0]:
                self.tree_equip.insert("", tk.END, values=row)

    def filter_equip(self):
        query = self.ent_search_equip.get().lower().strip()
        if not self.wb or "Перечень_Оборудования" not in self.wb.sheetnames:
            return
        ws = self.wb["Перечень_Оборудования"]

        for item in self.tree_equip.get_children():
            self.tree_equip.delete(item)

        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[0]:
                match = any(query in str(v).lower() for v in row if v)
                if match:
                    self.tree_equip.insert("", tk.END, values=row)

    def reset_equip_filter(self):
        self.ent_search_equip.delete(0, tk.END)
        self.populate_equip_table()

    # -------------------------------------------------------------------------
    # Вкладка 2: Матрица ТО
    # -------------------------------------------------------------------------
    def build_tab_matrix(self):
        cols = ("id", "name", "m_type", "freq", "qty", "h_single", "h_ann", "role")
        self.tree_matrix = ttk.Treeview(self.tab_matrix, columns=cols, show="headings", selectmode="browse")
        self.tree_matrix.heading("id", text="Поз. ID")
        self.tree_matrix.heading("name", text="Оборудование")
        self.tree_matrix.heading("m_type", text="Вид ТО")
        self.tree_matrix.heading("freq", text="Периодичность")
        self.tree_matrix.heading("qty", text="Ед.")
        self.tree_matrix.heading("h_single", text="Трудоемкость ед. (ч)")
        self.tree_matrix.heading("h_ann", text="Годовая (ч)")
        self.tree_matrix.heading("role", text="Исполнитель")

        self.tree_matrix.column("id", width=80, anchor="center")
        self.tree_matrix.column("name", width=230, anchor="w")
        self.tree_matrix.column("m_type", width=80, anchor="center")
        self.tree_matrix.column("freq", width=120, anchor="center")
        self.tree_matrix.column("qty", width=50, anchor="center")
        self.tree_matrix.column("h_single", width=110, anchor="e")
        self.tree_matrix.column("h_ann", width=90, anchor="e")
        self.tree_matrix.column("role", width=150, anchor="w")

        sb = ttk.Scrollbar(self.tab_matrix, orient=tk.VERTICAL, command=self.tree_matrix.yview)
        self.tree_matrix.configure(yscrollcommand=sb.set)
        self.tree_matrix.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        sb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)

        if self.wb and "Матрица_ТО" in self.wb.sheetnames:
            ws = self.wb["Матрица_ТО"]
            for row in ws.iter_rows(min_row=4, values_only=True):
                if row[0] and "ИТОГО" not in str(row[1]):
                    self.tree_matrix.insert("", tk.END, values=row)

    # -------------------------------------------------------------------------
    # Вкладка 3: Технологические регламенты
    # -------------------------------------------------------------------------
    def build_tab_recs(self):
        paned = ttk.PanedWindow(self.tab_recs, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(paned, width=280)
        right_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        paned.add(right_frame, weight=3)

        ttk.Label(left_frame, text="Выберите оборудование:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        self.lb_recs = tk.Listbox(left_frame, font=("Segoe UI", 10), selectmode=tk.SINGLE, height=12)
        self.lb_recs.pack(fill=tk.BOTH, expand=True)
        self.lb_recs.insert(tk.END, "1. Датчик давления Метран-150-TG")
        self.lb_recs.insert(tk.END, "2. ПЛК РСУ/ПАЗ Regul R500")
        self.lb_recs.insert(tk.END, "3. Электропривод задвижки Auma SA")
        self.lb_recs.insert(tk.END, "4. ИБП APC Smart-UPS RT 10kVA")
        self.lb_recs.select_set(0)
        self.lb_recs.bind("<<ListboxSelect>>", self.on_rec_selected)

        self.txt_rec_detail = scrolledtext.ScrolledText(right_frame, font=("Segoe UI", 10), wrap=tk.WORD)
        self.txt_rec_detail.pack(fill=tk.BOTH, expand=True)

        self.display_rec_metran()

    def on_rec_selected(self, event):
        sel = self.lb_recs.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == 0:
            self.display_rec_metran()
        elif idx == 1:
            self.display_rec_regul()
        elif idx == 2:
            self.display_rec_auma()
        elif idx == 3:
            self.display_rec_ups()

    def display_rec_metran(self):
        text = """ТЕХНОЛОГИЧЕСКИЙ РЕГЛАМЕНТ ТО: ДАТЧИК ДАВЛЕНИЯ МЕТРАН-150-TG
(Позиции: PT-101, PT-102 | Назначение: контроль входного и выкидного давления МНА)

1. ЦЕЛЬ РАБОТ:
   Обеспечение безотказного измерения избыточного давления перекачиваемой нефти,
   исключение утечек через импульсные линии, своевременная калибровка нуля и поверка.

2. СОСТАВ РАБОТ:
   • ТО-1 (2 раза в месяц, 0.5 чел·ч): Внешний осмотр корпуса (Exd/IP66), проверка целостности кабельного ввода,
     визуальный контроль импульсных линий на отсутствие утечек нефти, проверка индикации ЖК-дисплея.
   • ТО-2 (1 раз в месяц, 1.0 чел·ч): Продувка импульсных трубок, проверка и корректировка нулевой точки (Zero Trim)
     через HART-коммуникатор, сверка с контрольным техническим манометром классом 0.4.
   • ТО-4 / Поверка (1 раз в год, 2.5 чел·ч): Демонтаж, поверка на образцовом калибраторе давления Метран-517
     в аккредитованной лаборатории метрологии (5 точек шкалы прямого и обратного хода), оформление свидетельства.

3. ПРИБОРЫ И ИНСТРУМЕНТЫ:
   Портативный HART-коммуникатор Метран-650 / Rosemount 475, эталонный калибратор давления Метран-517,
   мультиметр Fluke 87V, искробезопасный набор слесарно-монтажного инструмента.

4. МЕРЫ БЕЗОПАСНОСТИ:
   Выполнять работы по наряду-допуску на взрывопожароопасных объектах. Закрыть коренной вентиль,
   сбросить остаточное давление через дренажный клапан в емкость сбора утечек, убедиться в отсутствии давления.

5. КРИТЕРИИ РАБОТОСПОСОБНОСТИ:
   Основная приведенная погрешность не более ±0.075%, сопротивление изоляции цепей питания > 20 МОм,
   герметичность импульсных линий при испытании рабочим давлением 6.3 МПа в течение 10 мин.
"""
        self.txt_rec_detail.delete("1.0", tk.END)
        self.txt_rec_detail.insert(tk.END, text)

    def display_rec_regul(self):
        text = """ТЕХНОЛОГИЧЕСКИЙ РЕГЛАМЕНТ ТО: ПРОГРАММИРУЕМЫЙ ЛОГИЧЕСКИЙ КОНТРОЛЛЕР REGUL R500
(Позиции: PLC-01, PLC-02 | Назначение: автоматическое управление РСУ и ПАЗ НПС)

1. ЦЕЛЬ РАБОТ:
   Предупреждение отказов микропроцессорных компонентов, диагностика дублированной шины данных,
   контроль температурного режима шкафов автоматики, резервное копирование прикладного ПО.

2. СОСТАВ РАБОТ:
   • ТО-2 (1 раз в месяц, 1.0 чел·ч): Считывание системных журналов и диагностических буферов CPU,
     контроль состояния светодиодной индикации RUN/FAIL/SYNC, проверка отсутствия перегрева (+45 °C макс).
   • ТО-3 (ежеквартально, 3.0 чел·ч): Обеспыливание шкафа и модулей сухим сжатым воздухом, проверка натяжения
     и протяжка винтовых/пружинных клеммников, тестирование безударного перехода CPU Master -> Slave.
   • ТО-4 (1 раз в год, 6.0-8.0 чел·ч): Выгрузка и сверка контрольных сумм рабочей программы (CRC),
     полное резервное копирование архивов SCADA и образов ПЛК, калибровка аналоговых модулей ввода AI-08-H,
     проверка сквозного прохождения аварийных сигналов ПАЗ на останов МНА.

3. ПРИБОРЫ И ИНСТРУМЕНТЫ:
   Инженерный ноутбук с ПО Epsilon LD / Regul Suite, тепловизор Testo 875, прецизионный калибратор токовой петли
   Fluke 709H, антистатический браслет ESD, промышленный пылесос с фильтром HEPA.

4. МЕРЫ БЕЗОПАСНОСТИ:
   Все операции в шкафу ПЛК производить с обязательным использованием антистатического браслета.
   Переключения и загрузку модифицированных программ согласовывать с главным инженером НПС.
"""
        self.txt_rec_detail.delete("1.0", tk.END)
        self.txt_rec_detail.insert(tk.END, text)

    def display_rec_auma(self):
        text = """ТЕХНОЛОГИЧЕСКИЙ РЕГЛАМЕНТ ТО: ЭЛЕКТРОПРИВОДНАЯ ЗАДВИЖКА AUMA SA 14.2 / AC 01.2
(Позиции: ACT-101 | Назначение: аварийная отсечка и переключение потоков нефти)

1. ЦЕЛЬ РАБОТ:
   Обеспечение гарантированного полного открытия/закрытия задвижки за паспортное время (< 30 с),
   контроль герметичности редуктора, проверка муфт ограничения крутящего момента.

2. СОСТАВ РАБОТ:
   • ТО-1 (2 раза в месяц, 0.5 чел·ч): Внешний осмотр корпуса привода, проверка индикатора положения,
     контроль отсутствия подтеков масла через сальниковые уплотнения штока.
   • ТО-2 (1 раз в месяц, 1.5 чел·ч): Опробование ручного маховика, проверка автоматического перехода
     из ручного в моторный режим, контроль токов электродвигателя при пробном ходе.
   • ТО-3 (ежеквартально, 3.0 чел·ч): Настройка и проверка концевых и моментных микровыключателей,
     проверка состояния клеммных соединений блока управления AC 01.2, смазка штока и подшипников.
   • ТО-4 (1 раз в год, 5.0 чел·ч): Замена редукторного масла, измерение сопротивления изоляции обмоток
     электродвигателя (R > 5 МОм), калибровка датчика непрерывного положения 4-20 мА.

3. ПРИБОРЫ И ИНСТРУМЕНТЫ:
   Токоизмерительные клещи Fluke 376, мегомметр MIC-2500, специальный ключ Auma, смазочный шприц,
   синтетическое редукторное масло Mobil SHC Gear 220.
"""
        self.txt_rec_detail.delete("1.0", tk.END)
        self.txt_rec_detail.insert(tk.END, text)

    def display_rec_ups(self):
        text = """ТЕХНОЛОГИЧЕСКИЙ РЕГЛАМЕНТ ТО: ИСТОЧНИК БЕСПЕРЕБОЙНОГО ПИТАНИЯ APC SMART-UPS 10kVA
(Позиции: UPS-01, BAT-01 | Назначение: гарантированное электроснабжение шкафов АС СН НПС)

1. ЦЕЛЬ РАБОТ:
   Обеспечение непрерывного электропитания КТС АС СН (не менее 60 минут автономной работы) при потере сети 380/220В.

2. СОСТАВ РАБОТ:
   • ТО-1 (2 раза в месяц, 0.3 чел·ч): Контроль панели индикации ИБП, напряжений шины DC и тока заряда АКБ.
   • ТО-2 (1 раз в месяц, 1.0 чел·ч): Визуальный осмотр аккумуляторных батарей на отсутствие вздутий и подтеков,
     замер напряжений на отдельных банках 12В (номинал 13.6-13.8 В в буферном режиме).
   • ТО-3 (ежеквартально, 2.5 чел·ч): Протяжка силовых болтовых соединений батарейной линейки динамическим ключом,
     очистка вентиляторов охлаждения, проверка логов встроенной карты SNMP.
   • ТО-4 (1 раз в год, 4.0 чел·ч): Контрольный разряд батарейной линейки на активную нагрузку для определения
     остаточной фактической емкости (C > 80% от номинала 55 А·ч), тестирование автоматического байпаса.
"""
        self.txt_rec_detail.delete("1.0", tk.END)
        self.txt_rec_detail.insert(tk.END, text)

    # -------------------------------------------------------------------------
    # Вкладка 4: Нештатные ситуации
    # -------------------------------------------------------------------------
    def build_tab_sop(self):
        paned = ttk.PanedWindow(self.tab_sop, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(paned, width=300)
        right_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        paned.add(right_frame, weight=3)

        ttk.Label(left_frame, text="Типовые нештатные ситуации:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        self.lb_sop = tk.Listbox(left_frame, font=("Segoe UI", 10), selectmode=tk.SINGLE, height=10)
        self.lb_sop.pack(fill=tk.BOTH, expand=True)
        self.lb_sop.insert(tk.END, "1. Отказ датчика давления выкида МНА")
        self.lb_sop.insert(tk.END, "2. Обрыв оптоволоконного кольца Ethernet")
        self.lb_sop.insert(tk.END, "3. Отказ основного ввода 380В и пуск ИБП")
        self.lb_sop.insert(tk.END, "4. Отказ ведущего процессора ПЛК Regul")
        self.lb_sop.insert(tk.END, "5. Аварийный загар в насосном зале")
        self.lb_sop.select_set(0)
        self.lb_sop.bind("<<ListboxSelect>>", self.on_sop_selected)

        self.txt_sop_detail = scrolledtext.ScrolledText(right_frame, font=("Segoe UI", 10), wrap=tk.WORD)
        self.txt_sop_detail.pack(fill=tk.BOTH, expand=True)

        self.display_sop_sensor()

    def on_sop_selected(self, event):
        sel = self.lb_sop.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == 0:
            self.display_sop_sensor()
        elif idx == 1:
            self.display_sop_network()
        elif idx == 2:
            self.display_sop_power()
        elif idx == 3:
            self.display_sop_cpu()
        elif idx == 4:
            self.display_sop_gas()

    def display_sop_sensor(self):
        t = """ИНСТРУКЦИЯ: ДЕЙСТВИЯ ПРИ ОТКАЗЕ ДАТЧИКА ДАВЛЕНИЯ НА ВЫКИДЕ МНА (PT-102)

1. ПРИЗНАКИ АВАРИИ:
   • На мнемосхеме SCADA параметр давления подсвечен красным цветом со статусом «BAD QUALITY / ОБРЫВ ЛИНИИ (I < 3.6 мА)».
   • Звуковая сигнализация в операторной, в журнале тревог запись: «PT-102: Недостоверность измерения давления выкида МНА-1».

2. ДЕЙСТВИЯ ДЕЖУРНОГО ОПЕРАТОРА:
   1. Немедленно квитировать аварийную сигнализацию нажатием кнопки «КВИТИРОВАНИЕ» (в течение 10 с).
   2. Проверить текущее состояние алгоритма мажоритирования 2oo3 (2 из 3):
      - При схеме 2oo3 отказ одного датчика НЕ вызывает останов агрегата, система автоматически исключает его из расчета.
      - Проверить давление по дублирующим датчикам PT-102B и PT-102C.
   3. Зафиксировать точное время и код ошибки в оперативном журнале.
   4. Вызвать дежурного инженера АСУ ТП и слесаря КИПиА (нормативное время прибытия не более 15 минут).

3. ДЕЙСТВИЯ РЕМОНТНОЙ БРИГАДЫ:
   1. Получить наряд-допуск на проведение ремонтных работ.
   2. Подключить HART-коммуникатор к датчику: проверить наличие питания 24В и внутренние коды ошибок сенсора.
   3. При аппаратной неисправности электронного блока или мембраны — закрыть отсечной вентиль, сбросить давление
      и произвести замену датчика на резервный экземпляр из фонда неснижаемого ЗИП.
   4. Сконфигурировать диапазон шкалы 0..6.0 МПа через HART-коммуникатор, провести калибровку нуля.
   5. Проверить соответствие показаний в SCADA и перевести контур защиты ПАЗ в штатный режим 2oo3.
"""
        self.txt_sop_detail.delete("1.0", tk.END)
        self.txt_sop_detail.insert(tk.END, t)

    def display_sop_network(self):
        t = """ИНСТРУКЦИЯ: ДЕЙСТВИЯ ПРИ ОБРЫВЕ ОПТОВОЛОКОННОГО КОЛЬЦА INDUSTRIAL ETHERNET

1. ПРИЗНАКИ АВАРИИ:
   • Аварийное сообщение SCADA: «MOXA Turbo Ring: Разрыв кольцевой топологии на сегменте ШУ-01 <-> ШПАЗ-01».
   • Включение индикатора FAULT на лицевой панели коммутатора MOXA EDS-510E.

2. ДЕЙСТВИЯ ОПЕРАТОРА И ИНЖЕНЕРА АСУ ТП:
   1. Квитировать предупреждение. Благодаря технологии быстрого восстановления MOXA Turbo Ring (время < 20 мс)
      потери пакетов управления технологическим процессом не происходит, связь перемаршрутизирована по резервной дуге.
   2. Вызвать инженера АСУ ТП.
   3. Инженер АСУ ТП подключается к Web-интерфейсу коммутатора и локализует номер аварийного оптического порта.
   4. Осмотреть оптический кросс и патч-корды в шкафу автоматики, проверить уровень оптического сигнала.
   5. При физическом повреждении кабеля переключить связь на резервные волокна оптического кабеля.
"""
        self.txt_sop_detail.delete("1.0", tk.END)
        self.txt_sop_detail.insert(tk.END, t)

    def display_sop_power(self):
        t = """ИНСТРУКЦИЯ: ДЕЙСТВИЯ ПРИ ОТКАЗЕ ВВОДА ПИТАНИЯ 380В И ПЕРЕХОДЕ НА ИБП

1. ПРИЗНАКИ АВАРИИ:
   • Сообщение SCADA: «ШГП-01: Пропадание основного ввода сети 380В! Работа от АКБ ИБП».
   • Загорание сигнального табло «РАБОТА ОТ БАТАРЕИ» на щите гарантированного питания.

2. ДЕЙСТВИЯ ПЕРСОНАЛА:
   1. Проконтролировать запуск системы автоматического ввода резерва (АВР) дизель-генераторной установки (ДГУ).
   2. Контролировать расчетное время автономной работы АКБ (по нормативу не менее 60 мин, критический порог — 15 мин).
   3. При неуспешном пуске ДГУ в течение 30 минут подготовить технологический регламент плавного безаварийного
      останова технологических агрегатов НПС.
"""
        self.txt_sop_detail.delete("1.0", tk.END)
        self.txt_sop_detail.insert(tk.END, t)

    def display_sop_cpu(self):
        t = """ИНСТРУКЦИЯ: ДЕЙСТВИЯ ПРИ ОТКАЗЕ ПРОЦЕССОРНОГО МОДУЛЯ ПЛК REGUL R500

1. ПРИЗНАКИ АВАРИИ:
   • Сообщение SCADA: «Контроллер РСУ: Отказ CPU Master! Автоматический безударный переход на CPU Standby».
   • На основном процессоре горит красный светодиод FAIL, на ведомом — зеленый RUN.

2. ДЕЙСТВИЯ ПЕРСОНАЛА:
   1. Технологический процесс продолжается без остановки благодаря горячему резервированию (Hot Standby).
   2. Инженер АСУ ТП выполняет диагностику неисправного CPU через сервисное ПО Epsilon LD.
   3. При аппаратном сбое модуля произвести горячую замену модуля CPU на запасной из комплекта ЗИП.
   4. Программа и конфигурация синхронизируются автоматически по межпроцессорной оптической шине в течение 30 секунд.
"""
        self.txt_sop_detail.delete("1.0", tk.END)
        self.txt_sop_detail.insert(tk.END, t)

    def display_sop_gas(self):
        t = """ИНСТРУКЦИЯ: ДЕЙСТВИЯ ПРИ СРАБАТЫВАНИИ ДАТЧИКА ЗАГАЗОВАННОСТИ (СГОЭС-М11)

1. ПРИЗНАКИ АВАРИИ:
   • 1-й порог (10% НКПР): Предупредительная светозвуковая сигнализация, автоматический пуск аварийной вытяжной вентиляции.
   • 2-й порог (20% НКПР): Аварийная сигнализация ПАЗ, автоматический останов МНА насосного цеха, закрытие входных/выходных задвижек.

2. ДЕЙСТВИЯ ОПЕРАТОРА:
   1. Немедленно объявить тревогу по громкой связи насосного цеха, эвакуировать персонал из загазованной зоны.
   2. Запретить включение электрооборудования в незащищенном исполнении.
   3. Направить газоспасательную службу (ВГСЧ) с переносными газоанализаторами для локализации утечки нефти.
"""
        self.txt_sop_detail.delete("1.0", tk.END)
        self.txt_sop_detail.insert(tk.END, t)

    # -------------------------------------------------------------------------
    # Вкладка 5: ЗИП и метрология
    # -------------------------------------------------------------------------
    def build_tab_zip(self):
        cols = ("cat", "name", "model", "purpose", "stock", "lead", "loc")
        self.tree_zip = ttk.Treeview(self.tab_zip, columns=cols, show="headings", selectmode="browse")
        self.tree_zip.heading("cat", text="Тип ресурса")
        self.tree_zip.heading("name", text="Наименование изделия / прибора")
        self.tree_zip.heading("model", text="Модель / Марка")
        self.tree_zip.heading("purpose", text="Назначение")
        self.tree_zip.heading("stock", text="Норма запаса")
        self.tree_zip.heading("lead", text="Срок пополн. (сут)")
        self.tree_zip.heading("loc", text="Место хранения")

        self.tree_zip.column("cat", width=110, anchor="center")
        self.tree_zip.column("name", width=220, anchor="w")
        self.tree_zip.column("model", width=200, anchor="w")
        self.tree_zip.column("purpose", width=240, anchor="w")
        self.tree_zip.column("stock", width=90, anchor="center")
        self.tree_zip.column("lead", width=110, anchor="center")
        self.tree_zip.column("loc", width=150, anchor="w")

        sb = ttk.Scrollbar(self.tab_zip, orient=tk.VERTICAL, command=self.tree_zip.yview)
        self.tree_zip.configure(yscrollcommand=sb.set)
        self.tree_zip.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        sb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)

        if self.wb and "Нормы_ЗИП_и_Приборы" in self.wb.sheetnames:
            ws = self.wb["Нормы_ЗИП_и_Приборы"]
            for row in ws.iter_rows(min_row=4, values_only=True):
                if row[0]:
                    self.tree_zip.insert("", tk.END, values=row)

    # -------------------------------------------------------------------------
    # Вкладка 6: Журнал ТО
    # -------------------------------------------------------------------------
    def build_tab_log(self):
        btn_frame = ttk.Frame(self.tab_log)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)

        btn_add = ttk.Button(btn_frame, text="➕ Добавить запись о ТО / инциденте", command=self.add_log_dialog)
        btn_add.pack(side=tk.LEFT, padx=(0, 10))

        cols = ("num", "date", "id", "type", "work", "res", "author", "status")
        self.tree_log = ttk.Treeview(self.tab_log, columns=cols, show="headings", selectmode="browse")
        self.tree_log.heading("num", text="№")
        self.tree_log.heading("date", text="Дата и время")
        self.tree_log.heading("id", text="Поз. ID")
        self.tree_log.heading("type", text="Вид ТО / Инцидент")
        self.tree_log.heading("work", text="Выполненные работы")
        self.tree_log.heading("res", text="Заключение")
        self.tree_log.heading("author", text="Исполнитель")
        self.tree_log.heading("status", text="Отметка")

        self.tree_log.column("num", width=40, anchor="center")
        self.tree_log.column("date", width=120, anchor="center")
        self.tree_log.column("id", width=80, anchor="center")
        self.tree_log.column("type", width=140, anchor="w")
        self.tree_log.column("work", width=300, anchor="w")
        self.tree_log.column("res", width=150, anchor="w")
        self.tree_log.column("author", width=120, anchor="w")
        self.tree_log.column("status", width=90, anchor="center")

        sb = ttk.Scrollbar(self.tab_log, orient=tk.VERTICAL, command=self.tree_log.yview)
        self.tree_log.configure(yscrollcommand=sb.set)
        self.tree_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(0, 10))
        sb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=(0, 10))

        if self.wb and "Журнал_ТО_и_Аварий" in self.wb.sheetnames:
            ws = self.wb["Журнал_ТО_и_Аварий"]
            for row in ws.iter_rows(min_row=4, values_only=True):
                if row[0]:
                    self.tree_log.insert("", tk.END, values=row)

    def add_log_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Внесение записи в оперативный журнал ТО")
        dlg.geometry("500x380")
        dlg.resizable(False, False)

        ttk.Label(dlg, text="Позиция оборудования (ID):").pack(anchor="w", padx=15, pady=(15, 2))
        ent_id = ttk.Entry(dlg, width=40)
        ent_id.insert(0, "PT-101")
        ent_id.pack(anchor="w", padx=15)

        ttk.Label(dlg, text="Вид ТО / Инцидент:").pack(anchor="w", padx=15, pady=(8, 2))
        cb_type = ttk.Combobox(dlg, values=["ТО-1 (Еженедельное)", "ТО-2 (Ежемесячное)", "ТО-3 (Ежеквартальное)", "ТО-4 (Поверка)", "Нештатная ситуация"], width=37)
        cb_type.current(1)
        cb_type.pack(anchor="w", padx=15)

        ttk.Label(dlg, text="Выполненные работы / Принятые меры:").pack(anchor="w", padx=15, pady=(8, 2))
        ent_work = ttk.Entry(dlg, width=50)
        ent_work.insert(0, "Плановая калибровка нуля, протяжка клеммных соединений")
        ent_work.pack(anchor="w", padx=15)

        ttk.Label(dlg, text="Заключение:").pack(anchor="w", padx=15, pady=(8, 2))
        ent_res = ttk.Entry(dlg, width=50)
        ent_res.insert(0, "Годен к дальнейшей эксплуатации")
        ent_res.pack(anchor="w", padx=15)

        ttk.Label(dlg, text="Исполнитель:").pack(anchor="w", padx=15, pady=(8, 2))
        ent_author = ttk.Entry(dlg, width=50)
        ent_author.insert(0, "Гайнанов Д.И., Гарифуллин К.Р.")
        ent_author.pack(anchor="w", padx=15)

        def save_log():
            import datetime
            now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
            new_id = len(self.tree_log.get_children()) + 1
            vals = (new_id, now_str, ent_id.get(), cb_type.get(), ent_work.get(), ent_res.get(), ent_author.get(), "Выполнено")
            self.tree_log.insert("", tk.END, values=vals)
            messagebox.showinfo("Успешно", "Запись успешно сохранена в оперативном журнале!")
            dlg.destroy()

        btn_save = ttk.Button(dlg, text="💾 Сохранить запись", command=save_log)
        btn_save.pack(pady=20)


def main():
    root = tk.Tk()
    app = MaintenanceViewerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
