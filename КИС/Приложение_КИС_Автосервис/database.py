# -*- coding: utf-8 -*-
"""
Модуль базы данных SQLite для КИС «Автосервис: Чип-тюнинг» (2026 г.)
Строго соответствует 3-й нормальной форме (3NF) и схеме ДБС (8 сущностей):
- Client (Клиенты)
- Car (Автомобили)
- Tuning_Package (Каталог пакетов тюнинга)
- Service_Request (Заявки клиентов — Данил Гайнанов)
- Work_Order (Наряд-заказы — Данил Гайнанов)
- Diagnostics_Card (Диагностические карты — Входной контроль)
- Order_Execution (Исполнение калибровки с тех. комментарием — Карим Гарифуллин)
- Work_Completion_Act (Акты выполненных работ и гарантия — Карим Гарифуллин)
"""

import sqlite3
import os
import json
from datetime import datetime

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "service_station.db")

def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = get_connection()
    c = conn.cursor()

    # 1. Client
    c.execute("""
    CREATE TABLE IF NOT EXISTS Client (
        id_client INTEGER PRIMARY KEY AUTOINCREMENT,
        client_fio TEXT NOT NULL,
        phone_number TEXT NOT NULL,
        email TEXT
    );
    """)

    # 2. Car (3NF: привязан к Client)
    c.execute("""
    CREATE TABLE IF NOT EXISTS Car (
        vin_code TEXT PRIMARY KEY,
        brand_model TEXT NOT NULL,
        engine_model TEXT NOT NULL,
        ecu_type TEXT NOT NULL,
        plate_number TEXT,
        id_client INTEGER NOT NULL,
        FOREIGN KEY (id_client) REFERENCES Client(id_client) ON DELETE CASCADE
    );
    """)

    # 3. Tuning_Package
    c.execute("""
    CREATE TABLE IF NOT EXISTS Tuning_Package (
        id_package INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        stage TEXT NOT NULL,
        power_gain_hp INTEGER NOT NULL,
        torque_gain_nm INTEGER NOT NULL,
        base_price REAL NOT NULL,
        description TEXT
    );
    """)

    # 4. Service_Request (Зона Данила)
    c.execute("""
    CREATE TABLE IF NOT EXISTS Service_Request (
        id_request INTEGER PRIMARY KEY AUTOINCREMENT,
        vin_code TEXT NOT NULL,
        id_package INTEGER NOT NULL,
        request_date TEXT NOT NULL,
        preliminary_cost REAL NOT NULL,
        eco_options TEXT,
        status TEXT NOT NULL DEFAULT 'Новая',
        FOREIGN KEY (vin_code) REFERENCES Car(vin_code),
        FOREIGN KEY (id_package) REFERENCES Tuning_Package(id_package)
    );
    """)

    # 5. Work_Order (Зона Данила — 3NF: vin_code убран, привязка через id_request!)
    c.execute("""
    CREATE TABLE IF NOT EXISTS Work_Order (
        id_order INTEGER PRIMARY KEY AUTOINCREMENT,
        id_request INTEGER NOT NULL UNIQUE,
        order_date TEXT NOT NULL,
        assigned_master TEXT NOT NULL DEFAULT 'Гайнанов Д.И.',
        assigned_tuner TEXT NOT NULL DEFAULT 'Гарифуллин К.Р.',
        total_cost REAL NOT NULL,
        order_status TEXT NOT NULL DEFAULT 'Утвержден',
        client_signature INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY (id_request) REFERENCES Service_Request(id_request)
    );
    """)

    # 6. Diagnostics_Card (Входной Quality Gate)
    c.execute("""
    CREATE TABLE IF NOT EXISTS Diagnostics_Card (
        id_card INTEGER PRIMARY KEY AUTOINCREMENT,
        id_order INTEGER NOT NULL,
        check_date TEXT NOT NULL,
        dtc_errors_found INTEGER NOT NULL DEFAULT 0,
        dtc_codes TEXT,
        boost_pressure_bar REAL NOT NULL,
        fuel_correction_ltft REAL NOT NULL,
        diagnostic_verdict TEXT NOT NULL,
        FOREIGN KEY (id_order) REFERENCES Work_Order(id_order)
    );
    """)

    # 7. Order_Execution (Зона Карима — ОБЯЗАТЕЛЬНЫЙ ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ МАСТЕРА)
    c.execute("""
    CREATE TABLE IF NOT EXISTS Order_Execution (
        id_execution INTEGER PRIMARY KEY AUTOINCREMENT,
        id_order INTEGER NOT NULL,
        stock_file_hash TEXT NOT NULL,
        mod_file_hash TEXT NOT NULL,
        checksum_valid INTEGER NOT NULL DEFAULT 1,
        flashing_tool TEXT NOT NULL,
        execution_date TEXT NOT NULL,
        technical_comment TEXT NOT NULL,
        FOREIGN KEY (id_order) REFERENCES Work_Order(id_order)
    );
    """)

    # 8. Work_Completion_Act (Зона Карима — Гарантия 24 мес, чек 54-ФЗ)
    c.execute("""
    CREATE TABLE IF NOT EXISTS Work_Completion_Act (
        id_act INTEGER PRIMARY KEY AUTOINCREMENT,
        id_order INTEGER NOT NULL UNIQUE,
        act_date TEXT NOT NULL,
        warranty_period_months INTEGER NOT NULL DEFAULT 24,
        payment_type TEXT NOT NULL,
        fiscal_receipt_number TEXT NOT NULL,
        client_signature_confirmed INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY (id_order) REFERENCES Work_Order(id_order)
    );
    """)

    conn.commit()
    conn.close()

def seed_demo_data():
    conn = get_connection()
    c = conn.cursor()

    # Проверяем, есть ли уже данные
    c.execute("SELECT COUNT(*) FROM Tuning_Package")
    if c.fetchone()[0] > 0:
        conn.close()
        return

    # Заполняем пакеты тюнинга
    packages = [
        ("Stage 1 (Dynamic Boost)", "Stage 1", 52, 90, 28000.0, 'Оптимизация штатных калибровок под топливо АИ-98/100 без вмешательства в "железо" двигателя.'),
        ("Stage 2 (Performance Sport)", "Stage 2", 85, 140, 42000.0, "Калибровка под спортивный даунпайп и увеличенный интеркулер. Снятие лимитера скорости."),
        ("Euro-2 & Decat Optimization", "Custom", 15, 25, 16000.0, "Программное отключение контроля катализатора и вторичных лямбда-зондов после удаления катализатора."),
        ("Eco & DPF/EGR Delete", "EcoTuning", 35, 80, 24000.0, "Программное отключение сажевого фильтра и клапана рециркуляции отработавших газов (EGR) для дизелей.")
    ]
    c.executemany("""
    INSERT INTO Tuning_Package (name, stage, power_gain_hp, torque_gain_nm, base_price, description)
    VALUES (?, ?, ?, ?, ?, ?)
    """, packages)

    # Клиенты
    clients = [
        ("Иванов Алексей Петрович", "+7 (917) 123-45-67", "a.ivanov@mail.ru"),
        ("Сидоров Михаил Сергеевич", "+7 (927) 987-65-43", "m.sidorov@yandex.ru"),
        ("Ахметов Руслан Ильдарович", "+7 (937) 456-78-90", "r.akhmetov@corp.ru"),
        ("Зайцев Дмитрий Николаевич", "+7 (987) 654-32-10", "d.zaytsev@gmail.com")
    ]
    c.executemany("""
    INSERT INTO Client (client_fio, phone_number, email) VALUES (?, ?, ?)
    """, clients)

    # Автомобили
    cars = [
        ("WBA5R11090F123456", "BMW 330i G20", "2.0 B48B20 258 л.с.", "Bosch MG1CS024", "А118АА 102", 1),
        ("WAUZZZF48HA098765", "Audi A4 B9 2.0 TFSI", "2.0 TFSI Gen3 249 л.с.", "Continental SIMOS 18.1", "В249ВВ 102", 2),
        ("WVWZZZAUZJP456789", "Volkswagen Golf R", "2.0 TSI CJXG 310 л.с.", "Continental SIMOS 18.10", "С310СС 777", 3),
        ("TMBAC7NE4K0321456", "Skoda Octavia RS", "2.0 TSI CHHA 230 л.с.", "Bosch MED17.5.2", "Е230ЕЕ 102", 4)
    ]
    c.executemany("""
    INSERT INTO Car (vin_code, brand_model, engine_model, ecu_type, plate_number, id_client)
    VALUES (?, ?, ?, ?, ?, ?)
    """, cars)

    # Заявка 1 (BMW - полностью завершена: прошла все шаги!)
    c.execute("""
    INSERT INTO Service_Request (vin_code, id_package, request_date, preliminary_cost, eco_options, status)
    VALUES ('WBA5R11090F123456', 1, '2026-03-20 10:15', 28000.0, 'Stage 1, Euro-2, Pop&Bang', 'Выполнена')
    """)
    req1_id = c.lastrowid

    # Наряд-заказ 1
    c.execute("""
    INSERT INTO Work_Order (id_request, order_date, assigned_master, assigned_tuner, total_cost, order_status, client_signature)
    VALUES (?, '2026-03-20 10:30', 'Гайнанов Д.И.', 'Гарифуллин К.Р.', 28000.0, 'Закрыт', 1)
    """, (req1_id,))
    ord1_id = c.lastrowid

    # Диагностика 1
    c.execute("""
    INSERT INTO Diagnostics_Card (id_order, check_date, dtc_errors_found, dtc_codes, boost_pressure_bar, fuel_correction_ltft, diagnostic_verdict)
    VALUES (?, '2026-03-20 11:00', 0, 'Ошибок DTC нет', 1.15, 0.8, 'ДВС исправен, допущен к чип-тюнингу')
    """, (ord1_id,))

    # Исполнение 1 (с обязательным комментарием Карима!)
    c.execute("""
    INSERT INTO Order_Execution (id_order, stock_file_hash, mod_file_hash, checksum_valid, flashing_tool, execution_date, technical_comment)
    VALUES (?, 'SHA256:7f83b1657ff1fc53b92dc18148a1d65dfc2d4b1fa3d677284addd200126d9069', 
               'SHA256:3a4b9c1e5f8821049ad567bce123f89012345678abcdef0123456789abcdef01',
               1, 'Autotuner OBD-II Master', '2026-03-20 12:45',
               'ПО ЭБУ Bosch MG1CS024 успешно обновлено до версии Stage 1. Ограничитель скорости 250 км/ч снят. Наддув увеличен на +0.22 бар в среднем диапазоне (2500-5500 об/мин). Углы опережения зажигания скорректированы под топливо АИ-98/100. На диностенде зафиксировано 310 л.с. и 490 Нм (прибавка +52 л.с. и +90 Нм). Откаты по детонации Knock Retard = 0.0°. Рекомендована замена свечей зажигания через 20 000 км.')
    """, (ord1_id,))

    # Акт 1 (закрыт с чеком и гарантией 24 мес)
    c.execute("""
    INSERT INTO Work_Completion_Act (id_order, act_date, warranty_period_months, payment_type, fiscal_receipt_number, client_signature_confirmed)
    VALUES (?, '2026-03-20 14:00', 24, 'Банковская карта (Эквайринг)', 'ФД-00049281', 1)
    """, (ord1_id,))

    # Заявка 2 (Audi - в работе у калибровщика, ожидает записи и комментария мастера!)
    c.execute("""
    INSERT INTO Service_Request (vin_code, id_package, request_date, preliminary_cost, eco_options, status)
    VALUES ('WAUZZZF48HA098765', 2, '2026-03-20 11:20', 42000.0, 'Stage 2, Downpipe, Decat Euro-2', 'В работе')
    """)
    req2_id = c.lastrowid

    c.execute("""
    INSERT INTO Work_Order (id_request, order_date, assigned_master, assigned_tuner, total_cost, order_status, client_signature)
    VALUES (?, '2026-03-20 11:45', 'Гайнанов Д.И.', 'Гарифуллин К.Р.', 42000.0, 'В работе', 1)
    """, (req2_id,))
    ord2_id = c.lastrowid

    c.execute("""
    INSERT INTO Diagnostics_Card (id_order, check_date, dtc_errors_found, dtc_codes, boost_pressure_bar, fuel_correction_ltft, diagnostic_verdict)
    VALUES (?, '2026-03-20 12:15', 0, 'Ошибок DTC нет', 1.20, 1.2, 'ДВС исправен, допущен к чип-тюнингу')
    """, (ord2_id,))

    # Заявка 3 (Golf R - на этапе входной диагностики)
    c.execute("""
    INSERT INTO Service_Request (vin_code, id_package, request_date, preliminary_cost, eco_options, status)
    VALUES ('WVWZZZAUZJP456789', 2, '2026-03-20 13:00', 42000.0, 'Stage 2, Pop & Bang', 'Утверждена')
    """)
    req3_id = c.lastrowid

    c.execute("""
    INSERT INTO Work_Order (id_request, order_date, assigned_master, assigned_tuner, total_cost, order_status, client_signature)
    VALUES (?, '2026-03-20 13:15', 'Гайнанов Д.И.', 'Гарифуллин К.Р.', 42000.0, 'На диагностике', 1)
    """, (req3_id,))

    # Заявка 4 (Skoda - новая заявка в CRM)
    c.execute("""
    INSERT INTO Service_Request (vin_code, id_package, request_date, preliminary_cost, eco_options, status)
    VALUES ('TMBAC7NE4K0321456', 1, '2026-03-20 14:10', 28000.0, 'Stage 1', 'Новая')
    """)

    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    seed_demo_data()
    print("База данных SQLite (3NF) успешно инициализирована и заполнена тестовыми данными.")
