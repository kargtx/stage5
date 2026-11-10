# -*- coding: utf-8 -*-
"""
Локальный интерактивный веб-просмотрщик базы данных мониторинга НПС
Для демонстрации Практической работы № 4 преподавателю (Антонов В.В.)
Студенты: Гайнанов Д.И., Гарифуллин К.Р., группа ЭАС-514С
Не требует сторонних библиотек (работает на встроенных модулях Python)
"""

import os
import sys
import json
import sqlite3
import time
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nps_monitoring.db")
PORT = 8088

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>БД Мониторинга НПС — Практическая работа №4 (АССН)</title>
    <style>
        :root {
            --bg: #0f172a;
            --surface: #1e293b;
            --surface-hover: #334155;
            --border: #475569;
            --text: #f8fafc;
            --text-dim: #94a3b8;
            --primary: #3b82f6;
            --primary-hover: #2563eb;
            --accent: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background: var(--bg);
            color: var(--text);
            display: flex;
            height: 100vh;
            overflow: hidden;
        }
        /* Боковая панель */
        .sidebar {
            width: 290px;
            background: var(--surface);
            border-right: 1px solid var(--border);
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
        }
        .sidebar-header {
            padding: 18px 16px;
            border-bottom: 1px solid var(--border);
            background: rgba(15, 23, 42, 0.4);
        }
        .sidebar-header h2 {
            font-size: 14px;
            font-weight: 700;
            color: var(--primary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .sidebar-header p {
            font-size: 11.5px;
            color: var(--text-dim);
            margin-top: 4px;
            line-height: 1.4;
        }
        .table-list {
            flex: 1;
            overflow-y: auto;
            padding: 10px 8px;
        }
        .table-btn {
            width: 100%;
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 12px;
            border: none;
            background: transparent;
            color: var(--text);
            font-size: 13px;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.15s ease;
            margin-bottom: 3px;
            text-align: left;
        }
        .table-btn:hover { background: var(--surface-hover); }
        .table-btn.active {
            background: var(--primary);
            color: #fff;
            font-weight: 600;
        }
        .table-badge {
            background: rgba(255, 255, 255, 0.12);
            padding: 2px 7px;
            border-radius: 10px;
            font-size: 11px;
        }
        .sidebar-footer {
            padding: 12px 14px;
            border-top: 1px solid var(--border);
            font-size: 11px;
            color: var(--text-dim);
            line-height: 1.4;
            background: rgba(15, 23, 42, 0.2);
        }

        /* Главная область */
        .main-content {
            flex: 1;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            background: var(--bg);
        }
        .topbar {
            padding: 12px 20px;
            background: var(--surface);
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .tabs {
            display: flex;
            gap: 8px;
        }
        .tab-btn {
            background: transparent;
            border: 1px solid var(--border);
            color: var(--text-dim);
            padding: 6px 14px;
            font-size: 12.5px;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .tab-btn.active, .tab-btn:hover {
            background: var(--surface-hover);
            color: #fff;
            border-color: var(--primary);
        }
        .tab-btn.active { background: var(--primary); border-color: var(--primary); }

        /* Вкладки контента */
        .view-panel {
            flex: 1;
            overflow: hidden;
            display: none;
            flex-direction: column;
        }
        .view-panel.active { display: flex; }

        /* Панель таблицы */
        .table-container {
            flex: 1;
            overflow: auto;
            padding: 16px 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12.5px;
            background: var(--surface);
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid var(--border);
        }
        th, td {
            padding: 9px 12px;
            border-bottom: 1px solid var(--border);
            text-align: left;
        }
        th {
            background: #111827;
            color: var(--primary);
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        tr:hover { background: rgba(255, 255, 255, 0.03); }
        .badge-pk { color: var(--danger); font-weight: bold; margin-right: 4px; }
        .badge-fk { color: var(--primary); font-weight: bold; margin-right: 4px; }

        /* Консоль SQL */
        .sql-panel {
            padding: 16px 20px;
            display: flex;
            flex-direction: column;
            gap: 12px;
            flex: 1;
            overflow: hidden;
        }
        .query-presets {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        .preset-btn {
            background: var(--surface);
            border: 1px solid var(--border);
            color: var(--text);
            padding: 6px 12px;
            font-size: 12px;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .preset-btn:hover {
            background: var(--surface-hover);
            border-color: var(--primary);
            color: var(--primary);
        }
        .editor-container {
            display: flex;
            gap: 10px;
        }
        #sqlEditor {
            flex: 1;
            height: 110px;
            background: #090d16;
            color: #38bdf8;
            font-family: Consolas, "Courier New", monospace;
            font-size: 13px;
            padding: 10px;
            border: 1px solid var(--border);
            border-radius: 6px;
            resize: none;
        }
        .run-btn {
            background: var(--accent);
            border: none;
            color: #000;
            font-weight: 700;
            padding: 0 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            transition: background 0.2s;
        }
        .run-btn:hover { background: #059669; color: #fff; }
        .query-status {
            font-size: 12px;
            color: var(--text-dim);
            display: flex;
            justify-content: space-between;
        }
        .query-results {
            flex: 1;
            overflow: auto;
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 6px;
        }

        /* Панель ER-Диаграммы */
        .erd-panel {
            flex: 1;
            padding: 20px;
            overflow: auto;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .erd-img {
            max-width: 95%;
            height: auto;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.5);
            border: 1px solid var(--border);
        }
    </style>
</head>
<body>

    <!-- Сайдбар со списком таблиц -->
    <div class="sidebar">
        <div class="sidebar-header">
            <h2>База данных НПС</h2>
            <p>Практика №4 • АС СН (ЭАС-514С)<br>Гайнанов Д.И., Гарифуллин К.Р.</p>
        </div>
        <div class="table-list" id="tableList">
            <!-- Таблицы загружаются динамически -->
        </div>
        <div class="sidebar-footer">
            СУБД: SQLite 3 (3НФ)<br>
            Файл: <code>nps_monitoring.db</code><br>
            Записей: 11 500+ телеметрии
        </div>
    </div>

    <!-- Основное содержимое -->
    <div class="main-content">
        <div class="topbar">
            <h3 id="currentTitle" style="font-size: 15px; font-weight: 600;">Таблица: tech_objects</h3>
            <div class="tabs">
                <button class="tab-btn active" onclick="switchTab('data')">Данные таблицы</button>
                <button class="tab-btn" onclick="switchTab('sql')">SQL Консоль (Запросы)</button>
                <button class="tab-btn" onclick="switchTab('erd')">ER-Диаграмма (3НФ)</button>
            </div>
        </div>

        <!-- Вкладка 1: Просмотр данных таблицы -->
        <div class="view-panel active" id="panel-data">
            <div class="table-container" id="dataContainer">
                <table id="dataTable">
                    <thead><tr id="tableHead"></tr></thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </div>

        <!-- Вкладка 2: Выполнение SQL запросов -->
        <div class="view-panel" id="panel-sql">
            <div class="sql-panel">
                <div class="query-presets">
                    <span style="font-size: 12px; color: var(--text-dim); align-self: center;">Типовые запросы задания:</span>
                    <button class="preset-btn" onclick="loadPreset(1)">1. Последние значения МНА-01</button>
                    <button class="preset-btn" onclick="loadPreset(2)">2. Агрегация давления за смену</button>
                    <button class="preset-btn" onclick="loadPreset(3)">3. Тревоги и отказы за сутки</button>
                    <button class="preset-btn" onclick="loadPreset(4)">4. Тренд за 1 час (LIMIT 720)</button>
                </div>
                <div class="editor-container">
                    <textarea id="sqlEditor" spellcheck="false">SELECT * FROM tech_objects;</textarea>
                    <button class="run-btn" onclick="runQuery()">Выполнить (F5)</button>
                </div>
                <div class="query-status">
                    <span id="queryStats">Готов к выполнению</span>
                    <span id="queryTime"></span>
                </div>
                <div class="query-results" id="queryResults">
                    <table id="queryTable">
                        <thead><tr id="qHead"></tr></thead>
                        <tbody id="qBody"></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Вкладка 3: Просмотр ER-диаграммы -->
        <div class="view-panel" id="panel-erd">
            <div class="erd-panel">
                <h4 style="margin-bottom: 12px; color: var(--primary);">Концептуальная и логическая ER-диаграмма 3НФ (нотация Crow's Foot без пересечений)</h4>
                <img src="/pr4_erd_diagram.png" class="erd-img" alt="ER-диаграмма 3НФ">
            </div>
        </div>
    </div>

    <script>
        let currentTable = 'tech_objects';

        const PRESET_QUERIES = {
            1: `SELECT 
    o.object_code,
    p.param_code,
    p.name AS param_name,
    m.value AS last_val,
    p.unit,
    m.timestamp,
    m.quality_code
FROM tech_objects o
JOIN object_parameters op ON o.object_id = op.object_id
JOIN parameters p ON op.parameter_id = p.parameter_id
JOIN measurements m ON m.measurement_id = (
    SELECT m2.measurement_id 
    FROM measurements m2 
    WHERE m2.binding_id = op.binding_id 
    ORDER BY m2.timestamp DESC, m2.measurement_id DESC 
    LIMIT 1
)
WHERE o.object_code = 'MNA-01';`,

            2: `SELECT 
    strftime('%H:00', m.timestamp) AS hour_interval,
    ROUND(AVG(m.value), 3) AS avg_pressure,
    ROUND(MIN(m.value), 3) AS min_pressure,
    ROUND(MAX(m.value), 3) AS max_pressure,
    COUNT(*) AS points_count
FROM measurements m
JOIN object_parameters op ON m.binding_id = op.binding_id
JOIN parameters p ON op.parameter_id = p.parameter_id
WHERE p.param_code = 'PT_DISCH'
GROUP BY strftime('%H', m.timestamp)
ORDER BY hour_interval ASC;`,

            3: `SELECT 
    e.timestamp,
    e.severity,
    e.event_type,
    o.object_code,
    e.description,
    u.username AS ack_by_user
FROM events e
JOIN tech_objects o ON e.object_id = o.object_id
LEFT JOIN users u ON e.ack_user_id = u.user_id
WHERE e.severity IN ('WARNING', 'CRITICAL', 'EMERGENCY')
ORDER BY e.timestamp DESC;`,

            4: `SELECT 
    m.timestamp,
    m.value,
    m.raw_current_ma,
    m.quality_code
FROM measurements m
JOIN object_parameters op ON m.binding_id = op.binding_id
WHERE op.parameter_id = 1
ORDER BY m.timestamp DESC
LIMIT 720;`
        };

        async function init() {
            const res = await fetch('/api/tables');
            const data = await res.json();
            const list = document.getElementById('tableList');
            list.innerHTML = '';
            data.forEach((t, idx) => {
                const btn = document.createElement('button');
                btn.className = `table-btn ${t.name === currentTable ? 'active' : ''}`;
                btn.onclick = () => loadTable(t.name);
                btn.innerHTML = `<span>📊 ${t.name}</span> <span class="table-badge">${t.count}</span>`;
                btn.id = `btn-${t.name}`;
                list.appendChild(btn);
            });
            loadTable(currentTable);
        }

        async function loadTable(tableName) {
            currentTable = tableName;
            document.querySelectorAll('.table-btn').forEach(b => b.classList.remove('active'));
            const activeBtn = document.getElementById(`btn-${tableName}`);
            if (activeBtn) activeBtn.classList.add('active');
            document.getElementById('currentTitle').innerText = `Таблица: ${tableName}`;

            const res = await fetch(`/api/data?table=${tableName}&limit=200`);
            const data = await res.json();

            const head = document.getElementById('tableHead');
            const body = document.getElementById('tableBody');
            head.innerHTML = data.columns.map(c => `<th>${c}</th>`).join('');
            body.innerHTML = data.rows.map(row => 
                `<tr>${row.map(val => `<td>${val !== null ? val : '<span style="color:gray;">NULL</span>'}</td>`).join('')}</tr>`
            ).join('');
        }

        function switchTab(tabId) {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            event.target.classList.add('active');
            document.querySelectorAll('.view-panel').forEach(p => p.classList.remove('active'));
            document.getElementById(`panel-${tabId}`).classList.add('active');
        }

        function loadPreset(num) {
            document.getElementById('sqlEditor').value = PRESET_QUERIES[num];
            runQuery();
        }

        async function runQuery() {
            const sql = document.getElementById('sqlEditor').value.trim();
            if (!sql) return;
            const status = document.getElementById('queryStats');
            const timer = document.getElementById('queryTime');
            status.innerText = 'Выполнение...';
            timer.innerText = '';

            try {
                const res = await fetch('/api/query', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ query: sql })
                });
                const data = await res.json();
                if (data.error) {
                    status.innerHTML = `<span style="color:var(--danger)">Ошибка: ${data.error}</span>`;
                    return;
                }
                status.innerText = `Успешно: выбрано ${data.rows.length} строк`;
                timer.innerText = `Время: ${data.time_ms.toFixed(3)} мс`;

                const qHead = document.getElementById('qHead');
                const qBody = document.getElementById('qBody');
                qHead.innerHTML = data.columns.map(c => `<th>${c}</th>`).join('');
                qBody.innerHTML = data.rows.map(row => 
                    `<tr>${row.map(val => `<td>${val !== null ? val : '<span style="color:gray;">NULL</span>'}</td>`).join('')}</tr>`
                ).join('');
            } catch (err) {
                status.innerHTML = `<span style="color:var(--danger)">Ошибка сети/сервера: ${err.message}</span>`;
            }
        }

        // F5 shortcut to run query
        window.addEventListener('keydown', (e) => {
            if (e.key === 'F5' || (e.ctrlKey && e.key === 'Enter')) {
                e.preventDefault();
                runQuery();
            }
        });

        init();
    </script>
</body>
</html>
"""

class DBViewerHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass # отключаем лишний консольный шум

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode('utf-8'))
        elif parsed.path == '/pr4_erd_diagram.png':
            img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pr4_erd_diagram.png')
            if os.path.exists(img_path):
                self.send_response(200)
                self.send_header('Content-Type', 'image/png')
                self.end_headers()
                with open(img_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404)
        elif parsed.path == '/api/tables':
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
            tables = []
            for (tname,) in cur.fetchall():
                cur.execute(f"SELECT COUNT(*) FROM [{tname}]")
                cnt = cur.fetchone()[0]
                tables.append({"name": tname, "count": cnt})
            conn.close()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps(tables).encode('utf-8'))
        elif parsed.path == '/api/data':
            params = parse_qs(parsed.query)
            tname = params.get('table', ['tech_objects'])[0]
            limit = int(params.get('limit', [100])[0])
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute(f"PRAGMA table_info([{tname}])")
            cols = [r[1] for r in cur.fetchall()]
            cur.execute(f"SELECT * FROM [{tname}] LIMIT {limit}")
            rows = cur.fetchall()
            conn.close()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({"columns": cols, "rows": rows}).encode('utf-8'))
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path == '/api/query':
            content_len = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_len)
            req = json.loads(post_body.decode('utf-8'))
            query = req.get('query', '')

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            t0 = time.perf_counter()
            try:
                cur.execute(query)
                rows = cur.fetchall()
                cols = [desc[0] for desc in cur.description] if cur.description else []
                t1 = time.perf_counter()
                res = {
                    "columns": cols,
                    "rows": rows[:500], # ограничим 500 для быстрой передачи
                    "time_ms": (t1 - t0) * 1000
                }
            except Exception as e:
                res = {"error": str(e)}
            finally:
                conn.close()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps(res).encode('utf-8'))

def start_server():
    server_address = ('127.0.0.1', PORT)
    httpd = HTTPServer(server_address, DBViewerHandler)
    url = f"http://localhost:{PORT}"
    print(f"============================================================")
    print(f"ВЕБ-ИНТЕРФЕЙС БАЗЫ ДАННЫХ НПС ЗАПУЩЕН!")
    print(f"Откройте в браузере: {url}")
    print(f"Для остановки сервера нажмите Ctrl + C")
    print(f"============================================================")
    webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nСервер остановлен.")
        httpd.server_close()

if __name__ == '__main__':
    start_server()

