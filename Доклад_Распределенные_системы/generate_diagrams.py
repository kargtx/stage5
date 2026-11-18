# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig_dir = r"c:\Users\kargt\Desktop\УЧЕБА\5 КУРС\stage5\Доклад_Распределенные_системы\figures"
os.makedirs(fig_dir, exist_ok=True)

# 1. Modern Distributed Data Architecture Diagram
def make_architecture_diagram():
    fig, ax = plt.subplots(figsize=(12, 7), dpi=300)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8.5)
    ax.axis('off')

    # Background canvas
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    # Title
    ax.text(6, 8.1, "Современная архитектура распределенной системы обработки данных (2024–2026)",
            fontsize=14, fontweight='bold', ha='center', color='#0F172A', fontfamily='sans-serif')

    layers = [
        ("УРОВЕНЬ ИСТОЧНИКОВ И СБОРА ДАННЫХ (Data Sources & Ingestion)", 6.9, 0.8, "#E2E8F0", "#334155",
         ["IoT / АСУТП (SCADA)", "Микросервисы (CDC, Debezium)", "Мобильные / Web-приложения", "Внешние API & Файлы"]),
        ("УРОВЕНЬ БРОКЕРОВ И ПОТОКОВ (Distributed Messaging & Buffering)", 5.5, 0.8, "#DBEAFE", "#1E40AF",
         ["Apache Kafka (Log)", "Apache Pulsar (Tiered)", "RabbitMQ / MQTT", "Redpanda (C++)"]),
        ("УРОВЕНЬ РАСПРЕДЕЛЕННЫХ ВЫЧИСЛЕНИЙ (Compute Engines)", 4.1, 0.8, "#DCFCE7", "#166534",
         ["Apache Flink (Streaming)", "Apache Spark (Batch & Micro-batch)", "Trino / Presto (Ad-hoc SQL)", "Ray (Distributed AI / LLM)"]),
        ("УРОВЕНЬ ХРАНЕНИЯ (Distributed Storage & Lakehouse)", 2.7, 0.8, "#FEF3C7", "#92400E",
         ["Apache Iceberg / Delta Lake", "Ceph / HDFS / MinIO (S3)", "ClickHouse (OLAP)", "CockroachDB / TiDB (NewSQL)"]),
        ("УРОВЕНЬ ПОТРЕБЛЕНИЯ И АНАЛИТИКИ (Serving, BI & Machine Learning)", 1.3, 0.8, "#F3E8FF", "#6B21A8",
         ["BI-дашборды и отчеты", "Real-time мониторинг", "LLM Inference & RAG", "Автоматические регуляторы"])
    ]

    for title, y, h, bg_col, text_col, items in layers:
        # Container box
        rect = patches.FancyBboxPatch((0.5, y), 11, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                      facecolor=bg_col, edgecolor='#94A3B8', linewidth=1.2)
        ax.add_patch(rect)
        ax.text(0.7, y + h - 0.22, title, fontsize=10, fontweight='bold', color=text_col)

        # 4 Sub-blocks
        col_w = 2.5
        for i, item in enumerate(items):
            box_x = 0.8 + i * 2.65
            sub_rect = patches.FancyBboxPatch((box_x, y + 0.1), col_w, h - 0.38,
                                              boxstyle="round,pad=0.05,rounding_size=0.1",
                                              facecolor="#FFFFFF", edgecolor=text_col, linewidth=0.9)
            ax.add_patch(sub_rect)
            ax.text(box_x + col_w / 2, y + 0.1 + (h - 0.38) / 2, item,
                    fontsize=8.5, ha='center', va='center', color='#1E293B', fontweight='semibold')

    # Draw vertical flow arrows
    for y_arrow in [6.85, 5.45, 4.05, 2.65]:
        ax.annotate("", xy=(6, y_arrow - 0.35), xytext=(6, y_arrow),
                    arrowprops=dict(arrowstyle="->,head_width=0.4,head_length=0.4",
                                    color="#0284C7", lw=2.2))

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "diagram_distributed_architecture.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print("Saved:", out_path)

# 2. CAP and PACELC Diagram
def make_cap_diagram():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6), dpi=300)
    fig.patch.set_facecolor('#F8FAFC')

    # --- AX1: CAP Theorem ---
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 10)
    ax1.axis('off')
    ax1.set_facecolor('#F8FAFC')
    ax1.text(5, 9.4, "Теорема CAP (Брюэр)", fontsize=13, fontweight='bold', ha='center', color='#0F172A')
    ax1.text(5, 8.8, "В распределенной сети возможны только 2 из 3 свойств", fontsize=9, ha='center', color='#475569')

    # Circles
    c_c = plt.Circle((3.6, 6.0), 2.2, facecolor='#60A5FA', alpha=0.4, edgecolor='#2563EB', lw=2)
    c_a = plt.Circle((6.4, 6.0), 2.2, facecolor='#34D399', alpha=0.4, edgecolor='#059669', lw=2)
    c_p = plt.Circle((5.0, 3.6), 2.2, facecolor='#F87171', alpha=0.4, edgecolor='#DC2626', lw=2)
    ax1.add_patch(c_c)
    ax1.add_patch(c_a)
    ax1.add_patch(c_p)

    # Labels
    ax1.text(2.6, 7.3, "Consistency (C)\nСогласованность", fontsize=9.5, fontweight='bold', ha='center', color='#1E40AF')
    ax1.text(7.4, 7.3, "Availability (A)\nДоступность", fontsize=9.5, fontweight='bold', ha='center', color='#065F46')
    ax1.text(5.0, 2.1, "Partition Tolerance (P)\nУстойчивость к разделению", fontsize=9.5, fontweight='bold', ha='center', color='#991B1B')

    # Intersections
    ax1.text(5.0, 6.8, "CA\n(Классические СУБД\nPostgres, Oracle)\n*Не в сети!", fontsize=7.5, ha='center', color='#1E293B')
    ax1.text(3.3, 4.3, "CP\n(HBase, MongoDB,\nRedis, Spanner)", fontsize=8, fontweight='bold', ha='center', color='#1E293B')
    ax1.text(6.7, 4.3, "AP\n(Cassandra, CouchDB,\nDynamoDB, DNS)", fontsize=8, fontweight='bold', ha='center', color='#1E293B')
    ax1.text(5.0, 5.0, "P обязателен\nв реальных сетях!", fontsize=7.5, fontweight='bold', ha='center', color='#B91C1C')

    # --- AX2: PACELC Theorem ---
    ax2.set_xlim(0, 10)
    ax2.set_ylim(0, 10)
    ax2.axis('off')
    ax2.set_facecolor('#F8FAFC')
    ax2.text(5, 9.4, "Теорема PACELC (Абади)", fontsize=13, fontweight='bold', ha='center', color='#0F172A')
    ax2.text(5, 8.8, "Поведение в штатном режиме и при сбоях сети", fontsize=9, ha='center', color='#475569')

    # Formula structure
    rect_if = patches.FancyBboxPatch((0.5, 5.2), 9.0, 3.2, boxstyle="round,pad=0.1",
                                    facecolor="#FEE2E2", edgecolor="#DC2626", lw=1.5)
    ax2.add_patch(rect_if)
    ax2.text(5, 7.8, "ЕСЛИ ЕСТЬ РАЗДЕЛЕНИЕ СЕТИ (If Partition - P):", fontsize=10, fontweight='bold', ha='center', color='#991B1B')
    ax2.text(2.8, 6.6, "Выбор Availability (A)\nВысокая доступность\n(Cassandra, DynamoDB)", fontsize=8.5, ha='center', color='#1E293B')
    ax2.text(5.0, 6.6, "ИЛИ", fontsize=11, fontweight='bold', ha='center', color='#991B1B')
    ax2.text(7.2, 6.6, "Выбор Consistency (C)\nСтрогая согласованность\n(HBase, Bigtable)", fontsize=8.5, ha='center', color='#1E293B')

    rect_else = patches.FancyBboxPatch((0.5, 1.2), 9.0, 3.4, boxstyle="round,pad=0.1",
                                      facecolor="#DCFCE7", edgecolor="#16A34A", lw=1.5)
    ax2.add_patch(rect_else)
    ax2.text(5, 4.0, "ИНАЧЕ, В ШТАТНОМ РЕЖИМЕ (Else - E):", fontsize=10, fontweight='bold', ha='center', color='#166534')
    ax2.text(2.8, 2.7, "Выбор Latency (L)\nНизкая задержка ответа\n(MongoDB, Cassandra)", fontsize=8.5, ha='center', color='#1E293B')
    ax2.text(5.0, 2.7, "ИЛИ", fontsize=11, fontweight='bold', ha='center', color='#166534')
    ax2.text(7.2, 2.7, "Выбор Consistency (C)\nАбсолютная свежесть данных\n(Spanner, Megastore)", fontsize=8.5, ha='center', color='#1E293B')

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "diagram_cap_pacelc.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print("Saved:", out_path)

# 3. Evolution of Processing Models
def make_evolution_diagram():
    fig, ax = plt.subplots(figsize=(12, 6.5), dpi=300)
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 8.5)
    ax.axis('off')

    ax.text(6, 8.0, "Эволюция парадигм обработки данных: от MapReduce до Streaming & Lakehouse",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A')

    steps = [
        ("Поколение 1: Пакетная обработка (2004–2012)\nApache Hadoop (MapReduce + HDFS)",
         "• Чтение/запись на жесткие диски (I/O bottleneck)\n• Задержки: от десятков минут до часов\n• Сложная модель программирования Map/Reduce",
         "#E2E8F0", "#334155", 6.2),
        ("Поколение 2: Вычисления в ОЗУ (2012–2018)\nApache Spark (RDD, Spark SQL, Micro-batch)",
         "• Вычисления в RAM в 10–100 раз быстрее Hadoop\n• Универсальный API (SQL, Graph, MLlib)\n• Микро-пакеты (задержки 100–500 мс)",
         "#DBEAFE", "#1E40AF", 4.3),
        ("Поколение 3: Истинный Stream Processing (2018–2023)\nApache Flink, Kafka Streams",
         "• Потоковая обработка по событиям (задержки < 10 мс)\n• Event Time, Watermarks, Stateful processing\n• Семантика ровно один раз (Exactly-Once)",
         "#DCFCE7", "#166534", 2.4),
        ("Поколение 4: Современный Lakehouse & Unified Compute (2024–2026)\nApache Iceberg, ClickHouse, Ray Distributed AI",
         "• Открытые табличные форматы данных (Iceberg, Delta Lake)\n• Конвергенция потоков, аналитики и моделей ИИ\n• Serverless & Zero-copy архитектуры",
         "#F3E8FF", "#6B21A8", 0.5)
    ]

    for title, desc, bg, edge, y in steps:
        rect = patches.FancyBboxPatch((0.8, y), 10.4, 1.5, boxstyle="round,pad=0.08",
                                      facecolor=bg, edgecolor=edge, lw=1.3)
        ax.add_patch(rect)
        ax.text(1.1, y + 1.15, title, fontsize=10, fontweight='bold', color=edge)
        ax.text(1.1, y + 0.55, desc, fontsize=8.5, color='#1E293B', va='center')

    # Flow arrows
    for y_a in [6.15, 4.25, 2.35]:
        ax.annotate("", xy=(6, y_a - 0.4), xytext=(6, y_a),
                    arrowprops=dict(arrowstyle="->,head_width=0.4,head_length=0.4",
                                    color="#0284C7", lw=2.2))

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "diagram_batch_vs_stream.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print("Saved:", out_path)

# 4. Consensus & Replication (Raft)
def make_consensus_diagram():
    fig, ax = plt.subplots(figsize=(11, 6), dpi=300)
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 7.5)
    ax.axis('off')

    ax.text(5.5, 7.0, "Распределенный консенсус: Архитектура алгоритма Raft",
            fontsize=13, fontweight='bold', ha='center', color='#0F172A')

    # Leader Node
    rect_l = patches.FancyBboxPatch((4.0, 4.5), 3.0, 1.8, boxstyle="round,pad=0.1",
                                   facecolor="#DCFCE7", edgecolor="#16A34A", lw=2)
    ax.add_patch(rect_l)
    ax.text(5.5, 5.8, "ЛИДЕР (Leader)", fontsize=11, fontweight='bold', ha='center', color='#166534')
    ax.text(5.5, 5.2, "• Принимает запросы клиентов\n• Формирует журнал (Log)\n• Отправляет Heartbeats",
            fontsize=8, ha='center', color='#1E293B')

    # Follower 1
    rect_f1 = patches.FancyBboxPatch((0.8, 1.2), 2.8, 1.8, boxstyle="round,pad=0.1",
                                    facecolor="#DBEAFE", edgecolor="#2563EB", lw=1.5)
    ax.add_patch(rect_f1)
    ax.text(2.2, 2.5, "ВЕДОМЫЙ 1 (Follower)", fontsize=10, fontweight='bold', ha='center', color='#1E40AF')
    ax.text(2.2, 1.8, "• Репликация записей\n• Ответ подтверждением\n• Таймер выборов",
            fontsize=7.8, ha='center', color='#1E293B')

    # Follower 2
    rect_f2 = patches.FancyBboxPatch((4.1, 1.2), 2.8, 1.8, boxstyle="round,pad=0.1",
                                    facecolor="#DBEAFE", edgecolor="#2563EB", lw=1.5)
    ax.add_patch(rect_f2)
    ax.text(5.5, 2.5, "ВЕДОМЫЙ 2 (Follower)", fontsize=10, fontweight='bold', ha='center', color='#1E40AF')
    ax.text(5.5, 1.8, "• Репликация записей\n• Ответ подтверждением\n• Таймер выборов",
            fontsize=7.8, ha='center', color='#1E293B')

    # Follower 3
    rect_f3 = patches.FancyBboxPatch((7.4, 1.2), 2.8, 1.8, boxstyle="round,pad=0.1",
                                    facecolor="#DBEAFE", edgecolor="#2563EB", lw=1.5)
    ax.add_patch(rect_f3)
    ax.text(8.8, 2.5, "ВЕДОМЫЙ 3 (Follower)", fontsize=10, fontweight='bold', ha='center', color='#1E40AF')
    ax.text(8.8, 1.8, "• Репликация записей\n• Ответ подтверждением\n• Таймер выборов",
            fontsize=7.8, ha='center', color='#1E293B')

    # Arrows from Leader to Followers
    ax.annotate("", xy=(2.2, 3.1), xytext=(4.5, 4.5),
                arrowprops=dict(arrowstyle="<->", color="#0284C7", lw=1.8, ls="--"))
    ax.annotate("", xy=(5.5, 3.1), xytext=(5.5, 4.5),
                arrowprops=dict(arrowstyle="<->", color="#0284C7", lw=1.8, ls="--"))
    ax.annotate("", xy=(8.8, 3.1), xytext=(6.5, 4.5),
                arrowprops=dict(arrowstyle="<->", color="#0284C7", lw=1.8, ls="--"))

    ax.text(5.5, 0.4, "Кворум = 3 узла из 4 (Большинство N/2 + 1) гарантирует защиту от Split-Brain",
            fontsize=9.5, fontweight='bold', ha='center', color='#0F172A')

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "diagram_consensus_raft.png")
    plt.savefig(out_path, dpi=300)
    plt.close()
    print("Saved:", out_path)

make_architecture_diagram()
make_cap_diagram()
make_evolution_diagram()
make_consensus_diagram()
print("All 4 diagrams generated successfully!")
