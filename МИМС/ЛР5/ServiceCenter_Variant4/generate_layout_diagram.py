import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_layout():
    fig, ax = plt.subplots(figsize=(14, 8), dpi=300)
    
    # Фон здания сервисного центра
    building = patches.FancyBboxPatch((0.5, 0.5), 13.0, 7.0, boxstyle="round,pad=0.1",
                                      facecolor="#F8FAFC", edgecolor="#334155", linewidth=3)
    ax.add_patch(building)
    
    # 1. Зона входа и стойка приемки (Reception)
    reception = patches.Rectangle((1.0, 4.0), 3.5, 3.0, facecolor="#E0F2FE", edgecolor="#0284C7", linewidth=2, linestyle="--")
    ax.add_patch(reception)
    ax.text(2.75, 6.6, "1. ЗОНА ПРИЕМКИ (RECEPTION)", fontsize=11, fontweight="bold", ha="center", color="#0369A1")
    
    # Стойка администратора
    desk = patches.Rectangle((1.5, 4.5), 2.5, 0.8, facecolor="#BAE6FD", edgecolor="#0284C7", linewidth=1.5)
    ax.add_patch(desk)
    ax.text(2.75, 4.9, "Стойка оформления\n(receptionService)", fontsize=9, ha="center", va="center", color="#0C4A6E")
    
    # 2. Зона ожидания клиентов (Waiting Lounge)
    waiting = patches.Rectangle((1.0, 1.0), 3.5, 2.5, facecolor="#FEF3C7", edgecolor="#D97706", linewidth=2, linestyle="--")
    ax.add_patch(waiting)
    ax.text(2.75, 3.1, "2. ЗОНА ОЖИДАНИЯ (LOUNGE)", fontsize=11, fontweight="bold", ha="center", color="#B45309")
    
    # Диваны/очередь
    sofa1 = patches.Rectangle((1.4, 1.4), 1.2, 0.6, facecolor="#FDE68A", edgecolor="#D97706", linewidth=1)
    sofa2 = patches.Rectangle((3.0, 1.4), 1.2, 0.6, facecolor="#FDE68A", edgecolor="#D97706", linewidth=1)
    ax.add_patch(sofa1)
    ax.add_patch(sofa2)
    ax.text(2.0, 1.7, "Диван A", fontsize=8, ha="center", va="center", color="#78350F")
    ax.text(3.6, 1.7, "Диван B", fontsize=8, ha="center", va="center", color="#78350F")
    ax.text(2.75, 2.3, "Очередь клиентов (queueOrders)\nТайм-аут: 60 мин -> отказ", fontsize=9, ha="center", color="#92400E")

    # 3. Ремонтный цех мастеров (Repair Workshop)
    workshop = patches.Rectangle((5.0, 2.2), 5.5, 4.8, facecolor="#F1F5F9", edgecolor="#475569", linewidth=2)
    ax.add_patch(workshop)
    ax.text(7.75, 6.6, "3. РЕМОНТНАЯ ЗОНА (WORKSHOP: N = 4 МАСТЕРА)", fontsize=11, fontweight="bold", ha="center", color="#1E293B")
    
    # Рабочие места 4 мастеров
    wb_coords = [(5.3, 4.5), (7.8, 4.5), (5.3, 2.8), (7.8, 2.8)]
    for idx, (wx, wy) in enumerate(wb_coords):
        wb = patches.Rectangle((wx, wy), 2.2, 1.3, facecolor="#E2E8F0", edgecolor="#64748B", linewidth=1.5)
        ax.add_patch(wb)
        ax.text(wx + 1.1, wy + 0.9, f"Рабочее место #{idx+1}", fontsize=9, fontweight="bold", ha="center", color="#0F172A")
        ax.text(wx + 1.1, wy + 0.45, "Паяльная станция,\nОсциллограф, ПК", fontsize=8, ha="center", color="#475569")

    # 4. Зона контроля качества (Quality Control)
    qc = patches.Rectangle((11.0, 4.2), 2.2, 2.8, facecolor="#EDE9FE", edgecolor="#7C3AED", linewidth=2, linestyle="--")
    ax.add_patch(qc)
    ax.text(12.1, 6.6, "4. ОТК (QC)", fontsize=11, fontweight="bold", ha="center", color="#6D28D9")
    ax.text(12.1, 5.2, "Тестирование\n(90% OK,\n10% доработка)", fontsize=9, ha="center", color="#5B21B6")

    # 5. Выдача заказов и склад (Ready Storage & Pickup)
    pickup = patches.Rectangle((11.0, 1.0), 2.2, 2.8, facecolor="#DCFCE7", edgecolor="#16A34A", linewidth=2, linestyle="--")
    ax.add_patch(pickup)
    ax.text(12.1, 3.4, "5. ВЫДАЧА (SINK)", fontsize=11, fontweight="bold", ha="center", color="#15803D")
    ax.text(12.1, 2.0, "Выдача клиенту\nи оплата\n(Выручка 3-5 тыс.)", fontsize=9, ha="center", color="#166534")

    # Стрелки потока
    arrowprops = dict(arrowstyle="->", color="#2563EB", lw=2, mutation_scale=15)
    ax.annotate("", xy=(1.0, 5.5), xytext=(0.1, 5.5), arrowprops=arrowprops) # вход
    ax.text(0.55, 5.8, "Вход клиентов\n(sourceOrders)", fontsize=8, ha="center", fontweight="bold", color="#1D4ED8")
    
    # от приемки к очереди
    ax.annotate("", xy=(2.75, 3.7), xytext=(2.75, 4.3), arrowprops=arrowprops)
    
    # от очереди к мастерам
    ax.annotate("", xy=(4.9, 4.0), xytext=(3.7, 2.5), arrowprops=arrowprops)
    
    # от мастеров к ОТК
    ax.annotate("", xy=(10.9, 5.5), xytext=(10.6, 5.5), arrowprops=arrowprops)
    
    # от ОТК к выдаче
    ax.annotate("", xy=(12.1, 3.9), xytext=(12.1, 4.2), arrowprops=arrowprops)

    # Выход отказов (balked)
    arrow_lost = dict(arrowstyle="->", color="#DC2626", lw=2, mutation_scale=15)
    ax.annotate("", xy=(0.1, 2.0), xytext=(0.9, 2.0), arrowprops=arrow_lost)
    ax.text(0.5, 2.4, "Отказ (sinkLost)\n(Ожидание > 60 мин)", fontsize=8, ha="center", fontweight="bold", color="#B91C1C")

    # Нижняя плашка с дашбордом
    dash = patches.Rectangle((0.5, 0.05), 13.0, 0.35, facecolor="#0F172A", edgecolor="#0F172A")
    ax.add_patch(dash)
    ax.text(7.0, 0.22, "AnyLogic 8.9 Animation View: Планировка сервисного центра с визуализацией движения заявок и загрузки мастеров",
            fontsize=9.5, fontweight="bold", color="white", ha="center", va="center")

    ax.set_xlim(-0.2, 14.0)
    ax.set_ylim(-0.2, 8.0)
    ax.axis("off")
    plt.tight_layout()
    
    out_path = 'МИМС/ЛР5/ServiceCenter_Variant4/service_center_layout_2d.png'
    plt.savefig(out_path, dpi=300)
    print(f"2D планировка анимации успешно сохранена в: {out_path}")

if __name__ == '__main__':
    draw_layout()
