# -*- coding: utf-8 -*-
"""
HTTP-сервер и REST API для КИС «Автосервис: Чип-тюнинг» (2026 г.)
Работает на базе встроенной библиотеки Python (без сторонних зависимостей).
Предоставляет полнофункциональный REST API к базе данных SQLite и отдает Single-Page Web App.
"""

import http.server
import socketserver
import os
import json
import urllib.parse
from datetime import datetime
import database

PORT = 8000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

class CISRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def _send_json(self, data, status_code=200):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS, PUT, DELETE")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS, PUT, DELETE")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path.startswith("/api/"):
            self._handle_api_get(path, urllib.parse.parse_qs(parsed.query))
            return

        # Если обращение к корню, отдаем index.html
        if path == "/":
            self.path = "/index.html"
        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path.startswith("/api/"):
            length = int(self.headers.get("Content-Length", 0))
            body_bytes = self.rfile.read(length) if length > 0 else b"{}"
            try:
                body = json.loads(body_bytes.decode("utf-8")) if body_bytes else {}
            except Exception:
                body = {}
            self._handle_api_post(path, body)
            return

        self.send_error(404, "Not Found")

    def _handle_api_get(self, path, query):
        conn = database.get_connection()
        c = conn.cursor()

        try:
            if path == "/api/stats":
                # Сводная аналитика дашборда
                c.execute("SELECT COUNT(*) FROM Client")
                total_clients = c.fetchone()[0]

                c.execute("SELECT COUNT(*) FROM Car")
                total_cars = c.fetchone()[0]

                c.execute("SELECT COUNT(*) FROM Service_Request")
                total_requests = c.fetchone()[0]

                c.execute("SELECT COUNT(*) FROM Work_Order")
                total_orders = c.fetchone()[0]

                c.execute("SELECT COUNT(*) FROM Work_Order WHERE order_status = 'Закрыт'")
                completed_orders = c.fetchone()[0]

                c.execute("SELECT COUNT(*) FROM Work_Order WHERE order_status IN ('В работе', 'На диагностике', 'Утвержден')")
                active_orders = c.fetchone()[0]

                c.execute("SELECT COALESCE(SUM(total_cost), 0) FROM Work_Order WHERE order_status = 'Закрыт'")
                total_revenue = c.fetchone()[0]

                c.execute("SELECT COUNT(*) FROM Work_Completion_Act")
                warranty_certificates = c.fetchone()[0]

                self._send_json({
                    "total_clients": total_clients,
                    "total_cars": total_cars,
                    "total_requests": total_requests,
                    "total_orders": total_orders,
                    "active_orders": active_orders,
                    "completed_orders": completed_orders,
                    "total_revenue": total_revenue,
                    "warranty_certificates": warranty_certificates,
                    "system_year": 2026,
                    "engineers": {
                        "master": "Гайнанов Д.И. (Данил)",
                        "tuner": "Гарифуллин К.Р. (Карим)",
                        "supervisor": "Демченко М.С."
                    }
                })

            elif path == "/api/clients":
                c.execute("""
                SELECT cl.*, 
                       (SELECT COUNT(*) FROM Car cr WHERE cr.id_client = cl.id_client) AS car_count
                FROM Client cl
                ORDER BY cl.id_client DESC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/cars":
                c.execute("""
                SELECT cr.*, cl.client_fio, cl.phone_number
                FROM Car cr
                JOIN Client cl ON cr.id_client = cl.id_client
                ORDER BY cr.brand_model ASC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/packages":
                c.execute("SELECT * FROM Tuning_Package ORDER BY base_price ASC")
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/requests":
                c.execute("""
                SELECT sr.*, 
                       cr.brand_model, cr.engine_model, cr.ecu_type, cr.plate_number,
                       cl.client_fio, cl.phone_number,
                       tp.name AS package_name, tp.stage, tp.power_gain_hp, tp.torque_gain_nm,
                       wo.id_order, wo.order_status
                FROM Service_Request sr
                JOIN Car cr ON sr.vin_code = cr.vin_code
                JOIN Client cl ON cr.id_client = cl.id_client
                JOIN Tuning_Package tp ON sr.id_package = tp.id_package
                LEFT JOIN Work_Order wo ON sr.id_request = wo.id_request
                ORDER BY sr.id_request DESC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/orders":
                c.execute("""
                SELECT wo.*,
                       sr.preliminary_cost, sr.eco_options, sr.request_date,
                       cr.vin_code, cr.brand_model, cr.engine_model, cr.ecu_type, cr.plate_number,
                       cl.client_fio, cl.phone_number, cl.email,
                       tp.name AS package_name, tp.stage, tp.power_gain_hp, tp.torque_gain_nm,
                       dc.id_card, dc.diagnostic_verdict, dc.dtc_errors_found,
                       oe.id_execution, oe.flashing_tool, oe.checksum_valid, oe.technical_comment,
                       wca.id_act, wca.act_date, wca.fiscal_receipt_number, wca.warranty_period_months
                FROM Work_Order wo
                JOIN Service_Request sr ON wo.id_request = sr.id_request
                JOIN Car cr ON sr.vin_code = cr.vin_code
                JOIN Client cl ON cr.id_client = cl.id_client
                JOIN Tuning_Package tp ON sr.id_package = tp.id_package
                LEFT JOIN Diagnostics_Card dc ON wo.id_order = dc.id_order
                LEFT JOIN Order_Execution oe ON wo.id_order = oe.id_order
                LEFT JOIN Work_Completion_Act wca ON wo.id_order = wca.id_order
                ORDER BY wo.id_order DESC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/diagnostics":
                c.execute("""
                SELECT dc.*, 
                       wo.id_request, wo.order_status,
                       cr.vin_code, cr.brand_model, cr.engine_model, cr.ecu_type,
                       cl.client_fio
                FROM Diagnostics_Card dc
                JOIN Work_Order wo ON dc.id_order = wo.id_order
                JOIN Service_Request sr ON wo.id_request = sr.id_request
                JOIN Car cr ON sr.vin_code = cr.vin_code
                JOIN Client cl ON cr.id_client = cl.id_client
                ORDER BY dc.id_card DESC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/executions":
                c.execute("""
                SELECT oe.*,
                       wo.total_cost, wo.order_status,
                       cr.vin_code, cr.brand_model, cr.engine_model, cr.ecu_type,
                       cl.client_fio,
                       tp.name AS package_name
                FROM Order_Execution oe
                JOIN Work_Order wo ON oe.id_order = wo.id_order
                JOIN Service_Request sr ON wo.id_request = sr.id_request
                JOIN Car cr ON sr.vin_code = cr.vin_code
                JOIN Client cl ON cr.id_client = cl.id_client
                JOIN Tuning_Package tp ON sr.id_package = tp.id_package
                ORDER BY oe.id_execution DESC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            elif path == "/api/acts":
                c.execute("""
                SELECT wca.*,
                       wo.total_cost,
                       cr.vin_code, cr.brand_model, cr.plate_number,
                       cl.client_fio, cl.phone_number,
                       tp.name AS package_name
                FROM Work_Completion_Act wca
                JOIN Work_Order wo ON wca.id_order = wo.id_order
                JOIN Service_Request sr ON wo.id_request = sr.id_request
                JOIN Car cr ON sr.vin_code = cr.vin_code
                JOIN Client cl ON cr.id_client = cl.id_client
                JOIN Tuning_Package tp ON sr.id_package = tp.id_package
                ORDER BY wca.id_act DESC
                """)
                rows = [dict(r) for r in c.fetchall()]
                self._send_json(rows)

            else:
                self._send_json({"error": "Unknown API endpoint"}, 404)

        except Exception as e:
            self._send_json({"error": str(e)}, 500)
        finally:
            conn.close()

    def _handle_api_post(self, path, body):
        conn = database.get_connection()
        c = conn.cursor()

        try:
            # 1. Регистрация Клиента и Автомобиля (Сеанс 1.3)
            if path == "/api/clients":
                client_fio = body.get("client_fio", "").strip()
                phone_number = body.get("phone_number", "").strip()
                email = body.get("email", "").strip()

                if not client_fio or not phone_number:
                    self._send_json({"error": "ФИО и телефон обязательны для заполнения"}, 400)
                    return

                c.execute("""
                INSERT INTO Client (client_fio, phone_number, email)
                VALUES (?, ?, ?)
                """, (client_fio, phone_number, email))
                client_id = c.lastrowid

                # Если сразу указан автомобиль
                vin_code = body.get("vin_code", "").strip().upper()
                if vin_code:
                    brand_model = body.get("brand_model", "Не указана").strip()
                    engine_model = body.get("engine_model", "Не указана").strip()
                    ecu_type = body.get("ecu_type", "Bosch Standard").strip()
                    plate_number = body.get("plate_number", "").strip().upper()

                    c.execute("""
                    INSERT OR REPLACE INTO Car (vin_code, brand_model, engine_model, ecu_type, plate_number, id_client)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """, (vin_code, brand_model, engine_model, ecu_type, plate_number, client_id))

                conn.commit()
                self._send_json({"success": True, "id_client": client_id, "message": "Клиент успешно зарегистрирован в КИС"})

            # 2. Добавление автомобиля к клиенту
            elif path == "/api/cars":
                vin_code = body.get("vin_code", "").strip().upper()
                id_client = int(body.get("id_client", 0))
                brand_model = body.get("brand_model", "").strip()
                engine_model = body.get("engine_model", "").strip()
                ecu_type = body.get("ecu_type", "").strip()
                plate_number = body.get("plate_number", "").strip().upper()

                if not vin_code or not id_client or not brand_model:
                    self._send_json({"error": "VIN-код, модель и клиент обязательны"}, 400)
                    return

                c.execute("""
                INSERT OR REPLACE INTO Car (vin_code, brand_model, engine_model, ecu_type, plate_number, id_client)
                VALUES (?, ?, ?, ?, ?, ?)
                """, (vin_code, brand_model, engine_model, ecu_type, plate_number, id_client))
                conn.commit()
                self._send_json({"success": True, "vin_code": vin_code, "message": "Автомобиль добавлен в парк КИС"})

            # 3. Формирование заявки (Сеанс 1.5)
            elif path == "/api/requests":
                vin_code = body.get("vin_code", "").strip().upper()
                id_package = int(body.get("id_package", 0))
                eco_options = body.get("eco_options", "").strip()
                
                # Стоимость пакета
                c.execute("SELECT base_price FROM Tuning_Package WHERE id_package = ?", (id_package,))
                row = c.fetchone()
                if not row:
                    self._send_json({"error": "Пакет тюнинга не найден"}, 404)
                    return
                preliminary_cost = row[0]
                # Доп. опции
                if "Euro-2" in eco_options:
                    preliminary_cost += 5000.0
                if "Pop&Bang" in eco_options:
                    preliminary_cost += 4000.0

                now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                c.execute("""
                INSERT INTO Service_Request (vin_code, id_package, request_date, preliminary_cost, eco_options, status)
                VALUES (?, ?, ?, ?, ?, 'Новая')
                """, (vin_code, id_package, now_str, preliminary_cost, eco_options))
                req_id = c.lastrowid
                conn.commit()
                self._send_json({"success": True, "id_request": req_id, "cost": preliminary_cost, "message": "Заявка на чип-тюнинг успешно создана"})

            # 4. Утверждение Наряд-заказа — Зона Данила (Сеанс 1.8)
            elif path == "/api/orders":
                id_request = int(body.get("id_request", 0))
                assigned_master = body.get("assigned_master", "Гайнанов Д.И.").strip()
                assigned_tuner = body.get("assigned_tuner", "Гарифуллин К.Р.").strip()

                c.execute("SELECT preliminary_cost FROM Service_Request WHERE id_request = ?", (id_request,))
                r = c.fetchone()
                if not r:
                    self._send_json({"error": "Заявка не найдена"}, 404)
                    return
                total_cost = r[0]

                now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                c.execute("""
                INSERT INTO Work_Order (id_request, order_date, assigned_master, assigned_tuner, total_cost, order_status, client_signature)
                VALUES (?, ?, ?, ?, ?, 'На диагностике', 1)
                """, (id_request, now_str, assigned_master, assigned_tuner, total_cost))
                order_id = c.lastrowid

                # Обновляем статус заявки
                c.execute("UPDATE Service_Request SET status = 'В работе' WHERE id_request = ?", (id_request,))
                conn.commit()
                self._send_json({"success": True, "id_order": order_id, "message": f"Наряд-заказ № НЗ-2026/{order_id:04d} утвержден Гайнановым Д.И."})

            # 5. Входная диагностика ДВС (Сеанс 1.10)
            elif path == "/api/diagnostics":
                id_order = int(body.get("id_order", 0))
                dtc_errors_found = int(body.get("dtc_errors_found", 0))
                dtc_codes = body.get("dtc_codes", "Ошибок DTC нет").strip()
                boost_pressure_bar = float(body.get("boost_pressure_bar", 1.15))
                fuel_correction_ltft = float(body.get("fuel_correction_ltft", 0.0))
                
                verdict = "ДВС исправен, допущен к чип-тюнингу" if dtc_errors_found == 0 else "Неисправность ДВС! Чип-тюнинг заблокирован (на ремонт)"
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

                c.execute("""
                INSERT INTO Diagnostics_Card (id_order, check_date, dtc_errors_found, dtc_codes, boost_pressure_bar, fuel_correction_ltft, diagnostic_verdict)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (id_order, now_str, dtc_errors_found, dtc_codes, boost_pressure_bar, fuel_correction_ltft, verdict))
                card_id = c.lastrowid

                # Если ошибок нет — переводим в статус «В работе» у калибровщика, иначе — «Снят с тюнинга»
                new_status = "В работе" if dtc_errors_found == 0 else "Снят с тюнинга"
                c.execute("UPDATE Work_Order SET order_status = ? WHERE id_order = ?", (new_status, id_order))
                conn.commit()

                self._send_json({
                    "success": True, 
                    "id_card": card_id, 
                    "verdict": verdict, 
                    "status": new_status,
                    "message": "Диагностическая карта сохранена. " + ("Автомобиль допущен в тюнинг-цех к Гарифуллину К.Р." if dtc_errors_found == 0 else "Автомобиль направлен на ремонт.")
                })

            # 6. АРМ Калибровщика — Зона Карима (Сеанс 1.13)
            # ОБЯЗАТЕЛЬНОЕ ТРЕБОВАНИЕ: Технический комментарий мастера!
            elif path == "/api/executions":
                id_order = int(body.get("id_order", 0))
                technical_comment = body.get("technical_comment", "").strip()

                if not technical_comment or len(technical_comment) < 20:
                    self._send_json({"error": "Требование КИС: обязателен развернутый технический комментарий мастера-калибровщика Гарифуллина К.Р. (не менее 20 символов)!"}, 400)
                    return

                flashing_tool = body.get("flashing_tool", "Autotuner OBD-II Master").strip()
                stock_hash = body.get("stock_file_hash", "SHA256:7f83b1657ff1fc53b92dc18148a1d65dfc2d4b1fa3d677284addd200126d9069")
                mod_hash = body.get("mod_file_hash", "SHA256:4c2a1e9f8821049ad567bce123f89012345678abcdef0123456789abcdef02")
                checksum_valid = 1
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

                c.execute("""
                INSERT INTO Order_Execution (id_order, stock_file_hash, mod_file_hash, checksum_valid, flashing_tool, execution_date, technical_comment)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (id_order, stock_hash, mod_hash, checksum_valid, flashing_tool, now_str, technical_comment))
                exec_id = c.lastrowid

                # Обновляем статус наряд-заказа на «Выполнен» (готов к сдаче)
                c.execute("UPDATE Work_Order SET order_status = 'Выполнен' WHERE id_order = ?", (id_order,))
                conn.commit()

                self._send_json({
                    "success": True, 
                    "id_execution": exec_id, 
                    "message": "Калибровка ЭБУ успешно записана в контроллер. Технический комментарий зафиксирован. Заказ переведен в статус «Выполнен»."
                })

            # 7. Акт выполненных работ и гарантия — Зона Карима (Сеанс 1.16)
            elif path == "/api/acts":
                id_order = int(body.get("id_order", 0))
                payment_type = body.get("payment_type", "Банковская карта (Эквайринг)").strip()
                warranty_period_months = int(body.get("warranty_period_months", 24))
                fiscal_receipt = f"ФД-{int(datetime.now().timestamp()) % 100000:05d}"
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

                c.execute("""
                INSERT INTO Work_Completion_Act (id_order, act_date, warranty_period_months, payment_type, fiscal_receipt_number, client_signature_confirmed)
                VALUES (?, ?, ?, ?, ?, 1)
                """, (id_order, now_str, warranty_period_months, payment_type, fiscal_receipt))
                act_id = c.lastrowid

                # Закрываем наряд-заказ и заявку
                c.execute("UPDATE Work_Order SET order_status = 'Закрыт' WHERE id_order = ?", (id_order,))
                c.execute("""
                UPDATE Service_Request 
                SET status = 'Выполнена' 
                WHERE id_request = (SELECT id_request FROM Work_Order WHERE id_order = ?)
                """, (id_order,))
                conn.commit()

                self._send_json({
                    "success": True,
                    "id_act": act_id,
                    "fiscal_receipt": fiscal_receipt,
                    "warranty_months": warranty_period_months,
                    "message": f"Акт выполненных работ № АКТ-2026/{act_id:04d} сформирован. Чек {fiscal_receipt} пробит по 54-ФЗ. Гарантия {warranty_period_months} мес. активирована!"
                })

            # 8. Сброс к демо-данным
            elif path == "/api/reset":
                conn.close()
                if os.path.exists(database.DB_FILE):
                    os.remove(database.DB_FILE)
                database.init_db()
                database.seed_demo_data()
                self._send_json({"success": True, "message": "База данных успешно пересоздана и сброшена к исходным демо-данным 2026 года!"})
                return

            else:
                self._send_json({"error": "Unknown POST route"}, 404)

        except Exception as e:
            self._send_json({"error": str(e)}, 500)
        finally:
            conn.close()

def run_server():
    database.init_db()
    database.seed_demo_data()
    
    server_address = ("", PORT)
    httpd = http.server.HTTPServer(server_address, CISRequestHandler)
    print(f"==================================================================")
    print(f" КИС «АВТОСЕРВИС: ЧИП-ТЮНИНГ» ЗАПУЩЕНА (ВЕРСИЯ 2026 Г.)")
    print(f" Локальный адрес сервера: http://localhost:{PORT}")
    print(f" База данных: {database.DB_FILE} (SQLite в 3NF)")
    print(f" Исполнители: Гайнанов Д.И. (Данил) & Гарифуллин К.Р. (Карим)")
    print(f" Руководитель: ст. преп. Демченко М.С.")
    print(f" Нажмите Ctrl+C для остановки сервера.")
    print(f"==================================================================")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nСервер остановлен.")
        httpd.server_close()

if __name__ == "__main__":
    run_server()
