/**
 * КИС «АВТОСЕРВИС: ЧИП-ТЮНИНГ» — КЛИЕНТСКАЯ ЛОГИКА (2026 Г.)
 * Разработка: студенты гр. ЭАС-514С Гайнанов Д.И. (Данил) & Гарифуллин К.Р. (Карим)
 * Руководитель: ст. преп. Демченко М.С.
 */

// Глобальное состояние
const state = {
  stats: {},
  clients: [],
  cars: [],
  packages: [],
  requests: [],
  orders: [],
  diagnostics: [],
  executions: [],
  acts: [],
  activeArmOrder: null,
  activeDbTable: 'Client',
  isServerMode: true
};

// Инициализация при загрузке
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initDbTabs();
  initModals();
  loadAllData();

  // Обновление даты
  const dateEl = document.getElementById('current-system-date');
  if (dateEl) {
    dateEl.innerText = new Date().toISOString().split('T')[0] + ' (2026)';
  }
});

/* ==========================================================================
   1. НАВИГАЦИЯ ПО ВКЛАДКАМ
   ========================================================================== */
function initTabs() {
  const navBtns = document.querySelectorAll('.nav-btn');
  navBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-tab');
      switchTab(targetId);
    });
  });

  const quickNewBtn = document.getElementById('btn-quick-new-order');
  if (quickNewBtn) {
    quickNewBtn.addEventListener('click', () => {
      openModal('modal-client');
    });
  }

  const resetBtn = document.getElementById('btn-reset-demo');
  if (resetBtn) {
    resetBtn.addEventListener('click', handleResetDemo);
  }
}

function switchTab(tabId) {
  // Переключение активной кнопки в сайдбаре
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const activeBtn = document.querySelector(`.nav-btn[data-tab="${tabId}"]`);
  if (activeBtn) activeBtn.classList.add('active');

  // Переключение панели
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  const targetPanel = document.getElementById(tabId);
  if (targetPanel) targetPanel.classList.add('active');

  // Обновление заголовка
  const titleEl = document.getElementById('page-title');
  const titles = {
    'tab-dashboard': 'Сводный мониторинг и сквозной процесс КИС',
    'tab-clients': 'Сеанс 1.3: Картотека клиентов и автомобилей (Client & Car)',
    'tab-requests': 'Сеанс 1.5: Оформление заявок на чип-тюнинг (Service_Request)',
    'tab-orders': 'Сеанс 1.8: Утверждение Наряд-заказа [Зона Данила] (Work_Order)',
    'tab-diagnostics': 'Сеанс 1.10: Входная диагностика ДВС (Quality Gate)',
    'tab-execution': 'Сеанс 1.13: АРМ Калибровщика [Зона Карима] (Order_Execution)',
    'tab-acts': 'Сеанс 1.16: Закрытие наряда, Акт и Гарантия (Work_Completion_Act)',
    'tab-database': 'Информационная база данных КИС (Реляционная модель 3NF)'
  };
  if (titleEl && titles[tabId]) {
    titleEl.innerText = titles[tabId];
  }

  // Обновляем данные конкретной вкладки
  if (tabId === 'tab-execution') {
    renderArmOrders();
  } else if (tabId === 'tab-database') {
    renderRawDbTable(state.activeDbTable);
  }
}

/* ==========================================================================
   2. ЗАГРУЗКА ДАННЫХ ИЗ REST API
   ========================================================================== */
async function apiGet(endpoint) {
  try {
    const res = await fetch(endpoint);
    if (!res.ok) throw new Error(`HTTP error ${res.status}`);
    return await res.json();
  } catch (err) {
    console.warn(`API GET ${endpoint} недоступен, переключение на локальный режим:`, err);
    state.isServerMode = false;
    return getLocalFallback(endpoint);
  }
}

async function apiPost(endpoint, data) {
  try {
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const json = await res.json();
    if (!res.ok) throw new Error(json.error || `HTTP error ${res.status}`);
    return json;
  } catch (err) {
    console.warn(`API POST ${endpoint} ошибка:`, err);
    throw err;
  }
}

async function loadAllData() {
  try {
    const [stats, clients, cars, packages, requests, orders, diagnostics, executions, acts] = await Promise.all([
      apiGet('/api/stats'),
      apiGet('/api/clients'),
      apiGet('/api/cars'),
      apiGet('/api/packages'),
      apiGet('/api/requests'),
      apiGet('/api/orders'),
      apiGet('/api/diagnostics'),
      apiGet('/api/executions'),
      apiGet('/api/acts')
    ]);

    state.stats = stats || {};
    state.clients = clients || [];
    state.cars = cars || [];
    state.packages = packages || [];
    state.requests = requests || [];
    state.orders = orders || [];
    state.diagnostics = diagnostics || [];
    state.executions = executions || [];
    state.acts = acts || [];

    renderDashboard();
    renderClients();
    renderCars();
    renderRequests();
    renderOrders();
    renderDiagnostics();
    renderArmOrders();
    renderActs();
    updateBadges();

  } catch (err) {
    showToast('Ошибка загрузки данных КИС: ' + err.message, 'danger');
  }
}

/* ==========================================================================
   3. ОТРИСОВКА ПАНЕЛЕЙ И ТАБЛИЦ
   ========================================================================== */
function renderDashboard() {
  document.getElementById('kpi-revenue').innerText = (state.stats.total_revenue || 0).toLocaleString('ru-RU') + ' ₽';
  document.getElementById('kpi-active-orders').innerText = state.stats.active_orders || 0;
  document.getElementById('kpi-cars-count').innerText = state.stats.total_cars || 0;
  document.getElementById('kpi-warranty-count').innerText = state.stats.warranty_certificates || 0;

  // Таблица заказов на дашборде
  const tbody = document.getElementById('tbody-dashboard-orders');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (state.orders.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; color: var(--text-dim);">Нет зарегистрированных наряд-заказов</td></tr>';
    return;
  }

  state.orders.slice(0, 5).forEach(o => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono font-bold">НЗ-2026/${String(o.id_order).padStart(4, '0')}</td>
      <td><strong>${o.brand_model}</strong><br><span class="font-mono text-dim" style="font-size:11px;">${o.vin_code}</span></td>
      <td>${o.client_fio}</td>
      <td><span class="badge badge-accent">${o.package_name || 'Stage 1'}</span></td>
      <td>${o.assigned_master}</td>
      <td>${o.assigned_tuner}</td>
      <td class="font-bold">${Number(o.total_cost).toLocaleString('ru-RU')} ₽</td>
      <td>${getStatusPill(o.order_status)}</td>
      <td>
        <button class="btn btn-outline btn-sm" onclick="openOrderDoc(${o.id_order})">📄 Бланк</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function updateBadges() {
  document.getElementById('badge-clients-count').innerText = state.clients.length;
  document.getElementById('badge-requests-count').innerText = state.requests.length;
  document.getElementById('badge-orders-count').innerText = state.orders.length;

  const tuningCount = state.orders.filter(o => o.order_status === 'В работе').length;
  document.getElementById('badge-tuning-count').innerText = tuningCount;
}

function getStatusPill(status) {
  let cls = 'status-new';
  if (status === 'На диагностике') cls = 'status-diag';
  else if (status === 'В работе') cls = 'status-tuning';
  else if (status === 'Выполнен' || status === 'Закрыт') cls = 'status-done';
  else if (status.includes('Снят') || status.includes('Брак')) cls = 'status-fail';
  return `<span class="status-pill ${cls}">${status}</span>`;
}

// Сеанс 1.3: Клиенты
function renderClients() {
  const tbody = document.getElementById('tbody-clients');
  if (!tbody) return;
  tbody.innerHTML = '';

  state.clients.forEach(cl => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono">${cl.id_client}</td>
      <td><strong>${cl.client_fio}</strong></td>
      <td class="font-mono">${cl.phone_number}</td>
      <td>${cl.email || '—'}</td>
      <td><span class="badge">${cl.car_count || 1} ТС</span></td>
      <td>
        <button class="btn btn-secondary btn-sm" onclick="prefillNewRequest('${cl.client_fio}')">+ Заявка</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// Автомобили
function renderCars() {
  const tbody = document.getElementById('tbody-cars');
  if (!tbody) return;
  tbody.innerHTML = '';

  state.cars.forEach(cr => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono font-bold">${cr.vin_code}</td>
      <td><strong>${cr.brand_model}</strong></td>
      <td>${cr.engine_model}</td>
      <td><span class="tag">${cr.ecu_type}</span></td>
      <td class="font-mono">${cr.plate_number || '—'}</td>
      <td>${cr.client_fio} (${cr.phone_number})</td>
    `;
    tbody.appendChild(tr);
  });
}

// Сеанс 1.5: Заявки
function renderRequests() {
  const tbody = document.getElementById('tbody-requests');
  if (!tbody) return;
  tbody.innerHTML = '';

  state.requests.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono font-bold">ЗЯВ-${String(r.id_request).padStart(4, '0')}</td>
      <td class="font-mono" style="font-size:11px;">${r.request_date}</td>
      <td><strong>${r.brand_model}</strong> (${r.plate_number || r.vin_code})</td>
      <td>${r.client_fio}</td>
      <td><span class="badge badge-accent">${r.package_name}</span></td>
      <td class="text-success">+${r.power_gain_hp} л.с. / +${r.torque_gain_nm} Нм</td>
      <td><span style="font-size:11px;">${r.eco_options || '—'}</span></td>
      <td class="font-bold">${Number(r.preliminary_cost).toLocaleString('ru-RU')} ₽</td>
      <td>${getStatusPill(r.status)}</td>
      <td>
        ${!r.id_order ? `<button class="btn btn-primary btn-sm" onclick="approveRequestToOrder(${r.id_request})">📋 В наряд-заказ</button>` : `<span class="badge">Наряд №${r.id_order}</span>`}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// Сеанс 1.8: Наряд-заказы (Зона Данила)
function renderOrders() {
  const tbody = document.getElementById('tbody-orders');
  if (!tbody) return;
  tbody.innerHTML = '';

  state.orders.forEach(o => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono font-bold">НЗ-2026/${String(o.id_order).padStart(4, '0')}</td>
      <td class="font-mono" style="font-size:11px;">${o.order_date}</td>
      <td><strong>${o.brand_model}</strong><br><span class="font-mono text-dim" style="font-size:11px;">${o.vin_code}</span></td>
      <td>${o.client_fio}</td>
      <td><span class="tag">${o.assigned_master}</span></td>
      <td><span class="tag">${o.assigned_tuner}</span></td>
      <td class="font-bold">${Number(o.total_cost).toLocaleString('ru-RU')} ₽</td>
      <td><span class="text-success">✔ Подписан</span></td>
      <td>${getStatusPill(o.order_status)}</td>
      <td>
        <div style="display:flex; gap:6px;">
          <button class="btn btn-outline btn-sm" onclick="openOrderDoc(${o.id_order})" title="Печать официального договора Наряд-заказа Данила">📄 Бланк</button>
          ${o.order_status === 'Утвержден' ? `<button class="btn btn-primary btn-sm" onclick="openDiagForOrder(${o.id_order})">🔍 На диагностику</button>` : ''}
          ${o.order_status === 'Выполнен' ? `<button class="btn btn-success btn-sm" onclick="openActForOrder(${o.id_order})">📜 Выдать Акт</button>` : ''}
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// Сеанс 1.10: Диагностика
function renderDiagnostics() {
  const tbody = document.getElementById('tbody-diagnostics');
  if (!tbody) return;
  tbody.innerHTML = '';

  state.diagnostics.forEach(d => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono font-bold">ДК-${String(d.id_card).padStart(4, '0')}</td>
      <td class="font-mono">НЗ-${d.id_order}</td>
      <td class="font-mono" style="font-size:11px;">${d.check_date}</td>
      <td><strong>${d.brand_model}</strong> (${d.client_fio})</td>
      <td><span class="${d.dtc_errors_found ? 'text-danger font-bold' : 'text-success'}">${d.dtc_codes}</span></td>
      <td>${d.boost_pressure_bar} бар</td>
      <td>${d.fuel_correction_ltft}%</td>
      <td><span class="status-pill ${d.dtc_errors_found ? 'status-fail' : 'status-done'}">${d.diagnostic_verdict}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

// Сеанс 1.13: АРМ Калибровщика (Зона Карима)
function renderArmOrders() {
  const container = document.getElementById('arm-orders-container');
  if (!container) return;
  container.innerHTML = '';

  const activeForTuning = state.orders.filter(o => o.order_status === 'В работе' || o.order_status === 'Выполнен');

  if (activeForTuning.length === 0) {
    container.innerHTML = '<div style="color:var(--text-dim); font-size:12px; padding:12px;">Нет заказов со статусом «В работе». Проведите входную диагностику!</div>';
    return;
  }

  activeForTuning.forEach(o => {
    const card = document.createElement('div');
    card.className = `arm-order-card ${state.activeArmOrder && state.activeArmOrder.id_order === o.id_order ? 'selected' : ''}`;
    card.onclick = () => selectArmOrder(o);
    card.innerHTML = `
      <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
        <span class="font-mono font-bold">НЗ-${o.id_order}</span>
        ${getStatusPill(o.order_status)}
      </div>
      <div style="font-weight:600; font-size:13px; color:#fff;">${o.brand_model}</div>
      <div style="font-size:11px; color:var(--text-muted); margin:2px 0;">Клиент: ${o.client_fio}</div>
      <div style="display:flex; justify-content:space-between; margin-top:6px; font-size:11px;">
        <span class="tag" style="padding:1px 5px;">${o.ecu_type}</span>
        <span class="badge badge-accent">${o.package_name}</span>
      </div>
    `;
    container.appendChild(card);
  });

  // Если заказ уже выбран, обновляем панель
  if (state.activeArmOrder) {
    const refreshed = state.orders.find(o => o.id_order === state.activeArmOrder.id_order);
    if (refreshed) selectArmOrder(refreshed);
  } else if (activeForTuning.length > 0) {
    selectArmOrder(activeForTuning[0]);
  }
}

function selectArmOrder(order) {
  state.activeArmOrder = order;

  document.querySelectorAll('.arm-order-card').forEach(c => c.classList.remove('selected'));
  renderArmOrdersListSelection();

  document.getElementById('arm-empty-state').style.display = 'none';
  document.getElementById('arm-active-box').style.display = 'block';

  document.getElementById('arm-car-title').innerText = `${order.brand_model} (VIN: ${order.vin_code})`;
  document.getElementById('arm-ecu-tag').innerText = `ECU: ${order.ecu_type}`;
  document.getElementById('arm-pkg-tag').innerText = `Пакет: ${order.package_name} (+${order.power_gain_hp} л.с.)`;

  // Хэши
  document.getElementById('arm-stock-hash').innerText = `SHA256:7f83b1657ff1fc53b92dc18148a1d65dfc2d4b1fa3d677284addd200126d9069`;
  document.getElementById('arm-mod-hash').innerText = `SHA256:4c2a1e9f8821049ad567bce123f89012345678abcdef0123456789abcdef02`;

  const commentInput = document.getElementById('arm-tech-comment');
  if (order.technical_comment) {
    commentInput.value = order.technical_comment;
  } else {
    commentInput.value = `Прошивка ${order.package_name} успешно записана в ЭБУ ${order.ecu_type}. Наддув скорректирован до 1.35 бар. Смесь Lambda 0.82 в режиме полного газа. Откаты по детонации Knock Retard = 0.0°. Требуется использование топлива АИ-98/100. Рекомендовано сократить интервал замены масла до 7 500 км.`;
  }

  // Привязка кнопок
  document.getElementById('btn-arm-read-stock').onclick = () => simulateStockRead();
  document.getElementById('btn-arm-write-flash').onclick = () => simulateEcuFlash();
  document.getElementById('btn-arm-recovery').onclick = () => handleRecoveryMode();
}

function renderArmOrdersListSelection() {
  document.querySelectorAll('.arm-order-card').forEach(c => {
    const text = c.innerText;
    if (state.activeArmOrder && text.includes(`НЗ-${state.activeArmOrder.id_order}`)) {
      c.classList.add('selected');
    }
  });
}

// Имитация вычитывания стокового дампа ROM
function simulateStockRead() {
  showToast('Подключение программатора к разъему OBD-II / Bench... Чтение сток-дампа ROM.', 'primary');
  setTimeout(() => {
    showToast('Стоковый дамп успешно вычитан (размер: 4 194 304 байт) и архивирован в КИС с хэшем SHA-256.', 'success');
  }, 1200);
}

// Аварийный Recovery Mode
function handleRecoveryMode() {
  if (!confirm('ВНИМАНИЕ! Вы собираетесь принудительно запустить Recovery Mode (аварийный откат блока на исходный заводской дамп). Продолжить?')) {
    return;
  }
  showToast('Инициализирован Recovery Mode через сервисный Boot-режим...', 'danger');
  setTimeout(() => {
    showToast('ЭБУ успешно восстановлен до 100% заводского состояния из архива КИС.', 'success');
  }, 1500);
}

// Прошивка ЭБУ и фиксация комментария мастера (Карим)
async function simulateEcuFlash() {
  if (!state.activeArmOrder) return;

  const comment = document.getElementById('arm-tech-comment').value.trim();
  const valMsg = document.getElementById('comment-val-msg');

  // ВАЛИДАЦИЯ: комментарий обязателен!
  if (!comment || comment.length < 20) {
    valMsg.innerText = 'ОШИБКА: Технический комментарий мастера-калибровщика обязателен (не менее 20 символов)!';
    document.getElementById('arm-tech-comment').focus();
    return;
  }
  valMsg.innerText = '';

  const progressWrap = document.getElementById('flash-progress-wrap');
  const progressBar = document.getElementById('flash-progress-bar');
  const percentText = document.getElementById('flash-percent');
  const statusText = document.getElementById('flash-status-text');

  progressWrap.style.display = 'block';
  progressBar.style.width = '0%';
  percentText.innerText = '0%';

  const tool = document.getElementById('arm-tool').value;

  const steps = [
    { p: 15, text: 'Проверка контрольной суммы Checksum: 0x9F4C2A1E... OK' },
    { p: 35, text: 'Снятие защиты от тюнинга TPROT и разблокировка контроллера...' },
    { p: 65, text: 'Запись калибровочных таблиц наддува и УОЗ во флеш-память...' },
    { p: 85, text: 'Верификация целостности микропрограммы и сброс адаптаций...' },
    { p: 100, text: 'Запись успешно завершена! ЭБУ готов к тестовому заезду.' }
  ];

  let currentStep = 0;
  const interval = setInterval(async () => {
    if (currentStep < steps.length) {
      const s = steps[currentStep];
      progressBar.style.width = s.p + '%';
      percentText.innerText = s.p + '%';
      statusText.innerText = s.text;
      currentStep++;
    } else {
      clearInterval(interval);

      try {
        const payload = {
          id_order: state.activeArmOrder.id_order,
          flashing_tool: tool,
          technical_comment: comment
        };
        const res = await apiPost('/api/executions', payload);
        showToast(res.message, 'success');
        await loadAllData();
      } catch (err) {
        showToast('Ошибка сохранения исполнения в БД: ' + err.message, 'danger');
      }
    }
  }, 400);
}

// Сеанс 1.16: Акты
function renderActs() {
  const tbody = document.getElementById('tbody-acts');
  if (!tbody) return;
  tbody.innerHTML = '';

  state.acts.forEach(a => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="font-mono font-bold">АКТ-2026/${String(a.id_act).padStart(4, '0')}</td>
      <td class="font-mono">НЗ-${a.id_order}</td>
      <td class="font-mono" style="font-size:11px;">${a.act_date}</td>
      <td><strong>${a.brand_model}</strong> (${a.plate_number || '—'})</td>
      <td>${a.client_fio} (${a.phone_number})</td>
      <td class="font-bold">${Number(a.total_cost).toLocaleString('ru-RU')} ₽</td>
      <td><span class="tag">${a.payment_type}</span></td>
      <td class="font-mono text-dim">${a.fiscal_receipt_number}</td>
      <td><span class="badge badge-accent">${a.warranty_period_months} мес.</span></td>
      <td>
        <button class="btn btn-outline btn-sm" onclick="openActDoc(${a.id_order})">📜 Сертификат</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/* ==========================================================================
   4. МОДАЛЬНЫЕ ОКНА И ОБРАБОТЧИКИ ФОРМ
   ========================================================================== */
function initModals() {
  // Открытие модалок
  document.getElementById('btn-open-modal-client')?.addEventListener('click', () => openModal('modal-client'));
  document.getElementById('btn-open-modal-request')?.addEventListener('click', () => {
    populateRequestFormSelects();
    openModal('modal-request');
  });
  document.getElementById('btn-open-modal-order')?.addEventListener('click', () => {
    populateOrderFormSelects();
    openModal('modal-order');
  });
  document.getElementById('btn-open-modal-diag')?.addEventListener('click', () => {
    populateDiagFormSelects();
    openModal('modal-diag');
  });
  document.getElementById('btn-open-modal-act')?.addEventListener('click', () => {
    populateActFormSelects();
    openModal('modal-act');
  });
}

function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.showModal();
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.close();
}

// Сеанс 1.3: Сабмит Клиента и ТС
async function handleClientSubmit(e) {
  e.preventDefault();
  const payload = {
    client_fio: document.getElementById('client-fio').value,
    phone_number: document.getElementById('client-phone').value,
    email: document.getElementById('client-email').value,
    vin_code: document.getElementById('car-vin').value,
    brand_model: document.getElementById('car-model').value,
    engine_model: document.getElementById('car-engine').value,
    ecu_type: document.getElementById('car-ecu').value,
    plate_number: document.getElementById('car-plate').value
  };

  try {
    const res = await apiPost('/api/clients', payload);
    showToast(res.message, 'success');
    closeModal('modal-client');
    e.target.reset();
    await loadAllData();
  } catch (err) {
    showToast(err.message, 'danger');
  }
}

// Имитация авто-декодера VIN
function simulateVinDecode() {
  const vin = document.getElementById('car-vin').value.trim();
  if (vin.length < 5) {
    showToast('Введите не менее 5 символов VIN-кода', 'danger');
    return;
  }
  showToast('VIN-декодер: сопоставление контроллера ЭБУ и модели ТС по базе данных AutoData...', 'primary');
  setTimeout(() => {
    if (vin.startsWith('WBA') || vin.startsWith('wba')) {
      document.getElementById('car-model').value = 'BMW 330i G20 xDrive';
      document.getElementById('car-engine').value = '2.0 B48B20 258 л.с.';
      document.getElementById('car-ecu').value = 'Bosch MG1CS024';
    } else if (vin.startsWith('WAU') || vin.startsWith('wau')) {
      document.getElementById('car-model').value = 'Audi A4 B9 2.0 TFSI';
      document.getElementById('car-engine').value = '2.0 TFSI Gen3 249 л.с.';
      document.getElementById('car-ecu').value = 'Continental SIMOS 18.1';
    } else {
      document.getElementById('car-model').value = 'Volkswagen Tiguan 2.0 TSI';
      document.getElementById('car-engine').value = '2.0 TSI EA888 220 л.с.';
      document.getElementById('car-ecu').value = 'Bosch MED17.5.2';
    }
    showToast('Параметры ТС и тип ЭБУ успешно определены!', 'success');
  }, 600);
}

function autoFillCarByVin(val) {
  if (val.length === 17) {
    simulateVinDecode();
  }
}

// Сеанс 1.5: Подготовка селектов заявки
function populateRequestFormSelects() {
  const carSelect = document.getElementById('req-car-vin');
  const pkgSelect = document.getElementById('req-package');

  carSelect.innerHTML = state.cars.map(c => 
    `<option value="${c.vin_code}">${c.brand_model} (${c.vin_code}) — ${c.client_fio}</option>`
  ).join('');

  pkgSelect.innerHTML = state.packages.map(p => 
    `<option value="${p.id_package}" data-price="${p.base_price}" data-hp="${p.power_gain_hp}" data-nm="${p.torque_gain_nm}">
      ${p.name} [${p.stage}] — ${Number(p.base_price).toLocaleString('ru-RU')} ₽ (+${p.power_gain_hp} л.с.)
    </option>`
  ).join('');

  updateRequestCalc();
}

function updateRequestCalc() {
  const pkgSelect = document.getElementById('req-package');
  const opt = pkgSelect.options[pkgSelect.selectedIndex];
  if (!opt) return;

  let basePrice = Number(opt.getAttribute('data-price') || 28000);
  const hp = opt.getAttribute('data-hp') || 52;
  const nm = opt.getAttribute('data-nm') || 90;

  if (document.getElementById('opt-euro2').checked) basePrice += 5000;
  if (document.getElementById('opt-popbang').checked) basePrice += 4000;

  document.getElementById('prev-power').innerText = `+${hp} л.с.`;
  document.getElementById('prev-torque').innerText = `+${nm} Нм`;
  document.getElementById('prev-cost').innerText = `${basePrice.toLocaleString('ru-RU')} ₽`;
}

async function handleRequestSubmit(e) {
  e.preventDefault();
  const eco = [];
  if (document.getElementById('opt-euro2').checked) eco.push('Euro-2 Decat');
  if (document.getElementById('opt-popbang').checked) eco.push('Pop & Bang');
  if (document.getElementById('opt-speedlim').checked) eco.push('V-Max Off');

  const payload = {
    vin_code: document.getElementById('req-car-vin').value,
    id_package: document.getElementById('req-package').value,
    eco_options: eco.join(', ') || 'Стандартная настройка'
  };

  try {
    const res = await apiPost('/api/requests', payload);
    showToast(res.message, 'success');
    closeModal('modal-request');
    await loadAllData();
    switchTab('tab-requests');
  } catch (err) {
    showToast(err.message, 'danger');
  }
}

// Быстрый перевод заявки в наряд-заказ (Зона Данила)
function approveRequestToOrder(reqId) {
  populateOrderFormSelects(reqId);
  openModal('modal-order');
}

function populateOrderFormSelects(preferredId = null) {
  const select = document.getElementById('order-req-id');
  const unlinkedRequests = state.requests.filter(r => !r.id_order);

  select.innerHTML = unlinkedRequests.map(r => 
    `<option value="${r.id_request}" ${preferredId === r.id_request ? 'selected' : ''}>
      Заявка №${r.id_request}: ${r.brand_model} (${r.client_fio}) — ${Number(r.preliminary_cost).toLocaleString('ru-RU')} ₽
    </option>`
  ).join('');

  if (unlinkedRequests.length === 0) {
    select.innerHTML = '<option value="">Нет новых заявок, требующих наряд-заказа</option>';
  }
}

async function handleOrderSubmit(e) {
  e.preventDefault();
  const reqId = document.getElementById('order-req-id').value;
  if (!reqId) {
    showToast('Выберите заявку для утверждения', 'danger');
    return;
  }

  const payload = {
    id_request: reqId,
    assigned_master: document.getElementById('order-master').value,
    assigned_tuner: document.getElementById('order-tuner').value
  };

  try {
    const res = await apiPost('/api/orders', payload);
    showToast(res.message, 'success');
    closeModal('modal-order');
    await loadAllData();
    switchTab('tab-orders');
  } catch (err) {
    showToast(err.message, 'danger');
  }
}

// Сеанс 1.10: Диагностика
function openDiagForOrder(orderId) {
  populateDiagFormSelects(orderId);
  openModal('modal-diag');
}

function populateDiagFormSelects(preferredOrderId = null) {
  const select = document.getElementById('diag-order-id');
  const waitingDiag = state.orders.filter(o => o.order_status === 'На диагностике' || o.order_status === 'Утвержден');

  select.innerHTML = waitingDiag.map(o => 
    `<option value="${o.id_order}" ${preferredOrderId === o.id_order ? 'selected' : ''}>
      НЗ-${o.id_order}: ${o.brand_model} (${o.client_fio})
    </option>`
  ).join('');

  if (waitingDiag.length === 0) {
    select.innerHTML = '<option value="">Нет автомобилей, ожидающих диагностику</option>';
  }
}

function handleDiagDtcChange(val) {
  const wrap = document.getElementById('diag-codes-wrap');
  wrap.style.display = (val === '1') ? 'block' : 'none';
}

async function handleDiagSubmit(e) {
  e.preventDefault();
  const orderId = document.getElementById('diag-order-id').value;
  if (!orderId) {
    showToast('Выберите наряд-заказ', 'danger');
    return;
  }

  const hasDtc = parseInt(document.getElementById('diag-dtc-select').value);
  const payload = {
    id_order: orderId,
    dtc_errors_found: hasDtc,
    dtc_codes: hasDtc ? document.getElementById('diag-codes-text').value : 'Ошибок DTC нет',
    boost_pressure_bar: document.getElementById('diag-boost').value,
    fuel_correction_ltft: document.getElementById('diag-ltft').value
  };

  try {
    const res = await apiPost('/api/diagnostics', payload);
    showToast(res.message, hasDtc ? 'danger' : 'success');
    closeModal('modal-diag');
    await loadAllData();
    switchTab(hasDtc ? 'tab-diagnostics' : 'tab-execution');
  } catch (err) {
    showToast(err.message, 'danger');
  }
}

// Сеанс 1.16: Акты
function openActForOrder(orderId) {
  populateActFormSelects(orderId);
  openModal('modal-act');
}

function populateActFormSelects(preferredOrderId = null) {
  const select = document.getElementById('act-order-id');
  const completed = state.orders.filter(o => o.order_status === 'Выполнен');

  select.innerHTML = completed.map(o => 
    `<option value="${o.id_order}" ${preferredOrderId === o.id_order ? 'selected' : ''}>
      НЗ-${o.id_order}: ${o.brand_model} (${o.client_fio}) — ${Number(o.total_cost).toLocaleString('ru-RU')} ₽
    </option>`
  ).join('');

  if (completed.length === 0) {
    select.innerHTML = '<option value="">Нет выполненных заказов, готовых к закрытию</option>';
  }
}

async function handleActSubmit(e) {
  e.preventDefault();
  const orderId = document.getElementById('act-order-id').value;
  if (!orderId) {
    showToast('Выберите выполненный заказ', 'danger');
    return;
  }

  const payload = {
    id_order: orderId,
    payment_type: document.getElementById('act-payment').value,
    warranty_period_months: document.getElementById('act-warranty').value
  };

  try {
    const res = await apiPost('/api/acts', payload);
    showToast(res.message, 'success');
    closeModal('modal-act');
    await loadAllData();
    openActDoc(orderId);
  } catch (err) {
    showToast(err.message, 'danger');
  }
}

/* ==========================================================================
   5. ОФИЦИАЛЬНЫЕ ПЕЧАТНЫЕ ДОКУМЕНТЫ (НАРЯД-ЗАКАЗ ДАНИЛА И АКТ КАРИМА)
   ========================================================================== */
function openOrderDoc(orderId) {
  const order = state.orders.find(o => o.id_order === orderId);
  if (!order) return;

  const viewer = document.getElementById('doc-viewer-body');
  document.getElementById('doc-viewer-title').innerText = `Наряд-заказ № КИС-2026/${String(order.id_order).padStart(4, '0')}`;

  viewer.innerHTML = `
    <div class="doc-header-table">
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
          <h2 style="margin:0; font-size:18px;">ООО «АВТОСЕРВИС КОРП: ЧИП-ТЮНИНГ»</h2>
          <div style="font-size:12px; color:#555;">г. Уфа, ул. Чернышевского, 145 • Тел: +7 (347) 299-00-26 • ИНН 0274123456</div>
        </div>
        <div class="doc-stamp">КИС «АВТОСЕРВИС»<br>УТВЕРЖДЕНО 2026</div>
      </div>
    </div>

    <h2 style="text-align:center; font-size:18px; margin:16px 0 6px;">НАРЯД-ЗАКАЗ № КИС-2026/${String(order.id_order).padStart(4, '0')}</h2>
    <div style="text-align:center; font-size:12px; margin-bottom:20px;">на выполнение работ по программной модификации ЭБУ (чип-тюнинг) от ${order.order_date} г.</div>

    <table style="width:100%; border-collapse:collapse; font-size:12px; margin-bottom:16px;" border="1" cellpadding="6">
      <tr>
        <td style="width:25%; font-weight:bold; background:#f5f5f5;">Заказчик (Владелец):</td>
        <td>${order.client_fio} (Тел: ${order.phone_number})</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Автомобиль:</td>
        <td><strong>${order.brand_model}</strong> (Госномер: ${order.plate_number || 'б/н'})</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Идентификатор VIN:</td>
        <td style="font-family:monospace; font-weight:bold;">${order.vin_code}</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Двигатель / Блок ЭБУ:</td>
        <td>${order.engine_model} / ${order.ecu_type}</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Мастер-приемщик:</td>
        <td><strong>${order.assigned_master}</strong> (Ведущий специалист сервиса)</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Инженер-калибровщик:</td>
        <td><strong>${order.assigned_tuner}</strong> (Специалист цеха чип-тюнинга)</td>
      </tr>
    </table>

    <h4 style="margin:12px 0 6px; font-size:14px;">Перечень заказанных услуг и калькуляция:</h4>
    <table style="width:100%; border-collapse:collapse; font-size:12px; margin-bottom:20px;" border="1" cellpadding="6">
      <tr style="background:#f0f0f0;">
        <th>№</th>
        <th>Наименование технологической операции</th>
        <th>Нормо-часы</th>
        <th>Стоимость (руб.)</th>
      </tr>
      <tr>
        <td>1</td>
        <td>Входная компьютерная диагностика параметров ДВС и сканирование кодов DTC</td>
        <td>0.5</td>
        <td>1 500.00</td>
      </tr>
      <tr>
        <td>2</td>
        <td>Вычитывание сток-дампа ROM и программная модификация ${order.package_name}</td>
        <td>2.0</td>
        <td>${(order.total_cost - 1500).toFixed(2)}</td>
      </tr>
      <tr style="font-weight:bold; background:#f9f9f9;">
        <td colspan="3" style="text-align:right;">ИТОГО К ОПЛАТЕ:</td>
        <td>${Number(order.total_cost).toLocaleString('ru-RU')} руб.</td>
      </tr>
    </table>

    <div style="font-size:11px; color:#444; margin-bottom:30px;">
      Заказчик подтверждает согласие с правилами эксплуатации форсированного двигателя (заправка бензином не ниже АИ-98).
    </div>

    <div style="display:flex; justify-content:space-between; margin-top:30px; font-size:12px;">
      <div>Мастер-приемщик: _________ / <strong>${order.assigned_master}</strong> /</div>
      <div>Заказчик: _________ / <strong>${order.client_fio}</strong> /</div>
    </div>
  `;

  openModal('modal-doc-viewer');
}

function openActDoc(orderId) {
  const order = state.orders.find(o => o.id_order === orderId);
  const act = state.acts.find(a => a.id_order === orderId);
  if (!order) return;

  const viewer = document.getElementById('doc-viewer-body');
  document.getElementById('doc-viewer-title').innerText = `Акт выполненных работ и Гарантийный талон № АКТ-2026/${order.id_order}`;

  viewer.innerHTML = `
    <div class="doc-header-table">
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
          <h2 style="margin:0; font-size:18px;">ООО «АВТОСЕРВИС КОРП: ЧИП-ТЮНИНГ»</h2>
          <div style="font-size:12px; color:#555;">Официальный сертифицированный центр модификации ПО блоков управления</div>
        </div>
        <div class="doc-stamp" style="border-color:#16a34a; color:#15803d;">ГАРАНТИЯ 24 МЕС.<br>АКТИВИРОВАНО 2026</div>
      </div>
    </div>

    <h2 style="text-align:center; font-size:18px; margin:16px 0 6px;">АКТ ВЫПОЛНЕННЫХ РАБОТ № АКТ-2026/${order.id_order}</h2>
    <div style="text-align:center; font-size:12px; margin-bottom:20px;">к Наряд-заказу № КИС-2026/${order.id_order} от ${act ? act.act_date : '2026-03-20'} г.</div>

    <table style="width:100%; border-collapse:collapse; font-size:12px; margin-bottom:16px;" border="1" cellpadding="6">
      <tr>
        <td style="width:25%; font-weight:bold; background:#f5f5f5;">Заказчик:</td>
        <td>${order.client_fio}</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Автомобиль:</td>
        <td><strong>${order.brand_model}</strong> (VIN: <span style="font-family:monospace;">${order.vin_code}</span>)</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Выполненная калибровка:</td>
        <td><strong>${order.package_name}</strong> (Прибавка: +${order.power_gain_hp} л.с., +${order.torque_gain_nm} Нм)</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Инженер-калибровщик:</td>
        <td><strong>${order.assigned_tuner}</strong> (Карим Гарифуллин)</td>
      </tr>
      <tr>
        <td style="font-weight:bold; background:#f5f5f5;">Фискальный чек (54-ФЗ):</td>
        <td><span style="font-family:monospace;">${act ? act.fiscal_receipt_number : 'ФД-00049281'}</span> (${act ? act.payment_type : 'Оплачено'})</td>
      </tr>
    </table>

    <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:4px; padding:12px; margin-bottom:16px;">
      <h4 style="margin:0 0 4px 0; font-size:12px; color:#1e3a8a;">ТЕХНИЧЕСКИЙ КОММЕНТАРИЙ МАСТЕРА-КАЛИБРОВЩИКА (ГАРИФУЛЛИН К.Р.):</h4>
      <p style="margin:0; font-size:11px; font-style:italic; line-height:1.4;">
        ${order.technical_comment || 'ПО ЭБУ Bosch MG1CS024 успешно обновлено до версии Stage 1. Ограничитель скорости 250 км/ч снят. Наддув увеличен на +0.22 бар в среднем диапазоне. УОЗ оптимизированы под АИ-98/100. Зафиксировано 310 л.с. и 490 Нм. Откаты зажигания отсутствуют.'}
      </p>
    </div>

    <div style="border:2px solid #2563eb; background:#eff6ff; padding:12px; border-radius:6px; margin-bottom:20px;">
      <h4 style="margin:0 0 4px 0; font-size:13px; color:#1d4ed8;">СЕРТИФИКАТ ГАРАНТИЙНЫХ ОБЯЗАТЕЛЬСТВ (24 МЕСЯЦА)</h4>
      <p style="margin:0; font-size:11px; color:#1e40af;">
        Сервисный центр гарантирует безотказное функционирование записанной микропрограммы ЭБУ, стабильность контрольных сумм и целостность микропроцессора на протяжении 24 месяцев со дня подписания настоящего Акта. Исходный заводской дамп сохранен в защищенном архиве сервиса с правом бесплатного отката.
      </p>
    </div>

    <div style="display:flex; justify-content:space-between; margin-top:30px; font-size:12px;">
      <div>Работы сдал (Инженер): _________ / <strong>${order.assigned_tuner}</strong> /</div>
      <div>Автомобиль принял (Клиент): _________ / <strong>${order.client_fio}</strong> /</div>
    </div>
  `;

  openModal('modal-doc-viewer');
}

/* ==========================================================================
   6. ОБОЗРЕВАТЕЛЬ БАЗЫ ДАННЫХ 3NF
   ========================================================================== */
function initDbTabs() {
  document.querySelectorAll('.db-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.db-tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const table = btn.getAttribute('data-dbtable');
      state.activeDbTable = table;
      renderRawDbTable(table);
    });
  });
}

function renderRawDbTable(table) {
  const thead = document.getElementById('thead-raw-db');
  const tbody = document.getElementById('tbody-raw-db');
  if (!thead || !tbody) return;

  let rows = [];
  if (table === 'Client') rows = state.clients;
  else if (table === 'Car') rows = state.cars;
  else if (table === 'Tuning_Package') rows = state.packages;
  else if (table === 'Service_Request') rows = state.requests;
  else if (table === 'Work_Order') rows = state.orders;
  else if (table === 'Diagnostics_Card') rows = state.diagnostics;
  else if (table === 'Order_Execution') rows = state.executions;
  else if (table === 'Work_Completion_Act') rows = state.acts;

  if (rows.length === 0) {
    thead.innerHTML = '<tr><th>Таблица пуста</th></tr>';
    tbody.innerHTML = '';
    return;
  }

  const cols = Object.keys(rows[0]);
  thead.innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';
  tbody.innerHTML = rows.map(r => 
    '<tr>' + cols.map(c => `<td>${r[c] !== null && r[c] !== undefined ? r[c] : 'NULL'}</td>`).join('') + '</tr>'
  ).join('');
}

/* ==========================================================================
   7. СБРОС К ДЕМО-ДАННЫМ
   ========================================================================== */
async function handleResetDemo() {
  if (!confirm('Вы действительно хотите пересоздать базу данных и сбросить ее к исходным демонстрационным данным 2026 года?')) {
    return;
  }
  try {
    const res = await apiPost('/api/reset', {});
    showToast(res.message, 'success');
    await loadAllData();
  } catch (err) {
    showToast('Ошибка сброса демо-данных: ' + err.message, 'danger');
  }
}

/* ==========================================================================
   8. УВЕДОМЛЕНИЯ TOAST
   ========================================================================== */
function showToast(message, type = 'primary') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerText = message;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}
