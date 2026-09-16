const STORAGE_KEY = 'straw-bale-recorder-v1';
const state = { fields: [], stocktakes: [], loads: [], stockMovements: [], customers: [] };
let map = null;
let markerLayer = null;
let mapOpened = false;
let pendingFieldLocation = null;
let serverStateAvailable = false;

const els = {};

window.addEventListener('DOMContentLoaded', async () => {
  collectElements();
  els.seasonYear.textContent = new Date().getFullYear();
  bindEvents();
  await hydrateState();
  render();
  initMap();
});

function collectElements() {
  [
    'seasonYear', 'seasonTotal', 'cropTotals', 'customerTotals', 'dailyFieldsTotal', 'dailyBalesTotal', 'dailyMoistureAverage',
    'recentFields', 'fieldList', 'estimatedStock', 'latestStocktakeTotal', 'removedSinceStocktake', 'mapFallback',
    'boughtInSinceStocktake', 'pendingLoads', 'completedLoads', 'stocktakeHistory', 'stockMovementsList',
    'fieldSearch', 'addFieldButtonFields', 'addStocktakeButton', 'addStockMovementButton', 'fieldDialog',
    'fieldForm', 'fieldDialogTitle', 'closeFieldDialogButton', 'fieldId', 'fieldCustomer', 'fieldFarm',
    'fieldName', 'fieldHectares', 'fieldBales', 'fieldMoisture', 'fieldCrop', 'fieldPhoto', 'photoPreview',
    'deleteFieldButton', 'partCompleteButton', 'completeFieldButton', 'addDeliveryButton', 'recentDeliveries',
    'deliveryDialog', 'deliveryForm', 'deliveryDialogTitle', 'closeDeliveryDialogButton', 'deliveryId',
    'deliveryCustomer', 'deliveryRegistration', 'deliveryDateTime', 'deliveryBales', 'deliveryWeight',
    'deliveryStatus', 'saveDeliveryButton', 'stocktakeDialog', 'stocktakeForm', 'closeStocktakeButton',
    'stocktakeDate', 'stocktakeBales', 'stocktakeNotes', 'stockMovementDialog', 'stockMovementForm',
    'stockMovementDialogTitle', 'closeStockMovementButton', 'stockMovementId', 'stockMovementType',
    'stockMovementDate', 'stockMovementCustomer', 'stockMovementBales', 'stockMovementNotes',
    'deleteStockMovementButton', 'straw-customers'
  ].forEach((id) => { els[id] = document.getElementById(id); });
}

function bindEvents() {
  document.querySelectorAll('[data-nav]').forEach((button) => {
    button.addEventListener('click', () => showView(button.dataset.nav));
  });
  els.addFieldButtonFields.addEventListener('click', () => openFieldDialog());
  els.fieldSearch.addEventListener('input', renderFieldList);
  els.fieldForm.addEventListener('submit', saveFieldFromForm);
  els.closeFieldDialogButton.addEventListener('click', closeFieldDialog);
  els.deleteFieldButton.addEventListener('click', deleteCurrentField);
  els.partCompleteButton.addEventListener('click', () => saveFieldWithStatus('part-complete'));
  els.completeFieldButton.addEventListener('click', () => saveFieldWithStatus('complete'));
  els.fieldPhoto.addEventListener('change', handlePhotoSelection);
  els.fieldDialog.addEventListener('close', () => { pendingFieldLocation = null; });
  els.addDeliveryButton.addEventListener('click', () => openDeliveryDialog());
  els.deliveryForm.addEventListener('submit', saveDeliveryFromForm);
  els.closeDeliveryDialogButton.addEventListener('click', () => els.deliveryDialog.close());
  els.addStocktakeButton.addEventListener('click', openStocktakeDialog);
  els.stocktakeForm.addEventListener('submit', saveStocktake);
  els.closeStocktakeButton.addEventListener('click', () => els.stocktakeDialog.close());
  els.addStockMovementButton.addEventListener('click', () => openStockMovementDialog());
  els.stockMovementForm.addEventListener('submit', saveStockMovement);
  els.closeStockMovementButton.addEventListener('click', () => els.stockMovementDialog.close());
  els.stockMovementType.addEventListener('change', updateStockMovementDefaults);
  els.deleteStockMovementButton.addEventListener('click', deleteCurrentStockMovement);
}

function showView(viewName) {
  document.querySelectorAll('.view').forEach((view) => view.classList.toggle('active', view.id === `${viewName}View`));
  document.querySelectorAll('.nav-button').forEach((button) => button.classList.toggle('active', button.dataset.nav === viewName));
  if (viewName === 'map' && map) {
    setTimeout(() => {
      map.invalidateSize();
      renderMapMarkers();
      if (!mapOpened) fitMapToFields();
      mapOpened = true;
    }, 80);
  }
}

function loadLocalState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (saved && typeof saved === 'object') {
      return {
        fields: Array.isArray(saved.fields) ? saved.fields : [],
        stocktakes: Array.isArray(saved.stocktakes) ? saved.stocktakes : [],
        loads: Array.isArray(saved.loads) ? saved.loads : [],
        stockMovements: Array.isArray(saved.stockMovements) ? saved.stockMovements : [],
        customers: Array.isArray(saved.customers) ? saved.customers : []
      };
    }
  } catch (error) {
    console.warn('Unable to load saved straw records', error);
  }
  return { fields: [], stocktakes: [], loads: [], stockMovements: [], customers: [] };
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  if (serverStateAvailable) saveServerState();
}

async function hydrateState() {
  const localState = loadLocalState();
  Object.assign(state, localState);
  try {
    const response = await fetch('/api/straw/state', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load Straw app records');
    const serverState = await response.json();
    serverStateAvailable = true;
    const serverHasRecords = serverState.fields.length || serverState.stocktakes.length || serverState.stockMovements.length;
    const localHasRecords = state.fields.length || state.stocktakes.length || state.stockMovements.length;
    if (serverHasRecords) {
      state.fields = serverState.fields;
      state.stocktakes = serverState.stocktakes;
      state.stockMovements = serverState.stockMovements;
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } else if (localHasRecords) {
      await saveServerState();
    }
  } catch (error) {
    console.warn('Unable to load central Straw app records', error);
  }
  try {
    const response = await fetch('/api/straw/deliveries', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load deliveries');
    state.loads = (await response.json()).deliveries;
  } catch (error) {
    console.warn('Unable to load delivered loads', error);
    state.loads = Array.isArray(state.loads) ? state.loads : [];
  }
  try {
    const response = await fetch('/api/straw/customers', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load straw customers');
    state.customers = (await response.json()).customers;
  } catch (error) {
    console.warn('Unable to load straw customers', error);
    state.customers = Array.isArray(state.customers) ? state.customers : [];
  }
  for (const field of state.fields) addCustomerLocally(field.customer);
}

async function saveServerState() {
  try {
    const response = await fetch('/api/straw/state', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: state.fields, stocktakes: state.stocktakes, stockMovements: state.stockMovements })
    });
    serverStateAvailable = response.ok;
  } catch (error) {
    serverStateAvailable = false;
    console.warn('Unable to save central Straw app records', error);
  }
}

function render() {
  renderTotals();
  renderFieldList();
  renderRecentFields();
  renderDeliveries();
  renderStock();
  renderStrawCustomers();
  renderMapMarkers();
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character]));
}

function formatDeliveryDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return value || 'No date';
  return new Date(`${value}T12:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
}

function deliveryCards(deliveries) {
  return deliveries.map((delivery) => `
    <div class="field-card delivery-card">
      <div><strong>${escapeHtml(delivery.customer)}</strong><span class="delivery-reg">${escapeHtml(delivery.registration)}</span></div>
      <div>${escapeHtml(formatDeliveryDate(delivery.delivery_date))} at ${escapeHtml(delivery.delivery_time)}</div>
      <div>${Number(delivery.bale_total || 0)} bales · ${delivery.weight_total ? `Weight: ${escapeHtml(delivery.weight_total)}` : '<strong class="awaiting-weight">Awaiting weight</strong>'}</div>
      <div class="dialog-actions"><button class="secondary-action" type="button" data-edit-delivery="${delivery.id}">Edit Load Out</button></div>
    </div>
  `).join('');
}

function bindDeliveryEditButtons(container) {
  container.querySelectorAll('[data-edit-delivery]').forEach((button) => {
    button.addEventListener('click', () => openDeliveryDialog(Number(button.dataset.editDelivery)));
  });
}

function renderDeliveries() {
  const recent = [...state.loads].slice(0, 5);
  els.recentDeliveries.innerHTML = recent.length ? deliveryCards(recent) : '<div class="field-card">No deliveries recorded yet.</div>';
  bindDeliveryEditButtons(els.recentDeliveries);
}

function addCustomerLocally(customer) {
  const name = String(customer || '').trim().replace(/\s+/g, ' ');
  if (!name || state.customers.some((value) => value.toLowerCase() === name.toLowerCase())) return;
  state.customers.push(name);
  state.customers.sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
}

function renderStrawCustomers() {
  els['straw-customers'].replaceChildren(...state.customers.map((customer) => new Option(customer, customer)));
}

function rememberStrawCustomer(customer) {
  const name = String(customer || '').trim().replace(/\s+/g, ' ');
  if (!name) return;
  addCustomerLocally(name);
  renderStrawCustomers();
  fetch('/api/straw/customers', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ customer: name })
  }).catch((error) => console.warn('Unable to remember straw customer', error));
}

function renderTotals() {
  const totalBales = state.fields.reduce((sum, field) => sum + Number(field.bales || 0), 0);
  els.seasonTotal.textContent = String(totalBales);
  const cropTotals = {};
  Array.from(els.fieldCrop.options).forEach((option) => { cropTotals[option.value] = 0; });
  state.fields.forEach((field) => {
    cropTotals[field.crop || 'Other'] = (cropTotals[field.crop || 'Other'] || 0) + Number(field.bales || 0);
  });
  const entries = Object.entries(cropTotals);
  els.cropTotals.innerHTML = entries.map(([crop, total]) => `
    <div class="panel crop-total" data-crop="${crop}">
      <span>${crop}</span>
      <strong>${total}</strong>
    </div>
  `).join('') || '<div class="panel">No crop varieties set up yet</div>';
  const customerTotals = {};
  state.fields.forEach((field) => {
    const customer = String(field.customer || '').trim() || 'No customer entered';
    customerTotals[customer] = (customerTotals[customer] || 0) + Number(field.bales || 0);
  });
  const customers = Object.entries(customerTotals).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  els.customerTotals.innerHTML = customers.map(([customer, total]) => `
    <div class="customer-total-row"><strong>${escapeHtml(customer)}</strong><span>${total} bales</span></div>
  `).join('') || '<div class="field-card">Customer totals will appear when field records are added.</div>';
  const todays = state.fields.filter((field) => field.status !== 'complete' || true);
  els.dailyFieldsTotal.textContent = String(todays.length);
  els.dailyBalesTotal.textContent = String(totalBales);
  const moistureValues = state.fields.map((field) => Number(field.moisture || 0)).filter((value) => value > 0);
  els.dailyMoistureAverage.textContent = moistureValues.length ? `${(moistureValues.reduce((sum, v) => sum + v, 0) / moistureValues.length).toFixed(1)}%` : '-';
}

function renderRecentFields() {
  const recent = [...state.fields].slice(-4).reverse();
  els.recentFields.innerHTML = recent.map((field) => `
    <div class="field-card">
      <strong>${field.name || 'Unnamed field'}</strong>
      <div>${field.customer || 'Unknown customer'} · ${field.farm || 'Unknown farm'}</div>
      <div>${field.bales || 0} bales</div>
    </div>
  `).join('') || '<div class="field-card">No recent fields yet</div>';
}

function renderFieldList() {
  const q = (els.fieldSearch.value || '').trim().toLowerCase();
  const items = state.fields.filter((field) => {
    const haystack = `${field.customer || ''} ${field.farm || ''} ${field.name || ''} ${field.crop || ''}`.toLowerCase();
    return haystack.includes(q);
  });
  els.fieldList.innerHTML = items.map((field) => `
    <div class="field-card">
      <div><strong>${field.name || 'Unnamed field'}</strong></div>
      <div>${field.customer || 'Unknown customer'} · ${field.farm || 'Unknown farm'}</div>
      <div>${field.bales || 0} bales · ${field.crop || 'Unknown'} · ${field.hectares || 0} ha</div>
      <div class="dialog-actions">
        <button class="secondary-action" type="button" data-edit-field="${field.id}">Edit</button>
      </div>
    </div>
  `).join('') || '<div class="field-card">No fields match this search.</div>';
  els.fieldList.querySelectorAll('[data-edit-field]').forEach((button) => {
    button.addEventListener('click', () => openFieldDialog(button.dataset.editField));
  });
}

function renderStock() {
  const latest = [...state.stocktakes].sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0))[0] || null;
  const afterLatest = (value) => !latest || new Date(value) >= new Date(latest.date);
  const loadsAfterCount = state.loads.filter((load) => afterLatest(`${load.delivery_date}T${load.delivery_time}`));
  const movementsAfterCount = state.stockMovements.filter((movement) => afterLatest(movement.date));
  const removedLoads = loadsAfterCount.reduce((sum, load) => sum + Number(load.bale_total || 0), 0);
  const ducksAllocated = movementsAfterCount.filter((movement) => movement.type === 'ducks').reduce((sum, movement) => sum + Number(movement.bales || 0), 0);
  const boughtIn = movementsAfterCount.filter((movement) => movement.type === 'bought-in').reduce((sum, movement) => sum + Number(movement.bales || 0), 0);
  els.latestStocktakeTotal.textContent = latest ? String(Number(latest.bales || 0)) : 'No count';
  els.removedSinceStocktake.textContent = String(removedLoads + ducksAllocated);
  els.boughtInSinceStocktake.textContent = String(boughtIn);
  els.estimatedStock.textContent = latest ? String(Number(latest.bales || 0) + boughtIn - removedLoads - ducksAllocated) : 'No count';

  const pendingLoads = state.loads.filter((load) => !load.weight_total);
  const completedLoads = state.loads.filter((load) => load.weight_total);
  els.pendingLoads.innerHTML = pendingLoads.length ? deliveryCards(pendingLoads) : '<div class="empty-state">No pending loads out.</div>';
  els.completedLoads.innerHTML = completedLoads.length ? deliveryCards(completedLoads) : '<div class="empty-state">No completed loads out.</div>';
  bindDeliveryEditButtons(els.pendingLoads);
  bindDeliveryEditButtons(els.completedLoads);

  const stocktakes = [...state.stocktakes].sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0));
  els.stocktakeHistory.innerHTML = stocktakes.length ? stocktakes.map((stocktake) => `
    <div class="stock-card">
      <span><strong>${escapeHtml(formatDateTime(stocktake.date))}</strong><span class="field-meta">${escapeHtml(stocktake.notes || 'Stocktake')}</span></span>
      <span><span class="bale-count">${Number(stocktake.bales || 0)}</span><button class="text-button" type="button" data-delete-stocktake="${escapeHtml(stocktake.id)}">Delete</button></span>
    </div>
  `).join('') : '<div class="empty-state">No stocktakes yet.</div>';
  els.stocktakeHistory.querySelectorAll('[data-delete-stocktake]').forEach((button) => {
    button.addEventListener('click', () => deleteStocktake(button.dataset.deleteStocktake));
  });

  const movements = [...state.stockMovements].sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0));
  els.stockMovementsList.innerHTML = movements.length ? movements.map((movement) => `
    <button class="stock-card" type="button" data-edit-stock-movement="${escapeHtml(movement.id)}">
      <span><strong>${movement.type === 'bought-in' ? 'Bought in' : 'Ducks at home'} · ${escapeHtml(formatDateTime(movement.date))}</strong><span class="field-meta">${escapeHtml([movement.customer, movement.notes].filter(Boolean).join(' · ') || 'Stock movement')}</span></span>
      <span class="bale-count">${movement.type === 'bought-in' ? '+' : '-'}${Number(movement.bales || 0)}</span>
    </button>
  `).join('') : '<div class="empty-state">No bought in or ducks records.</div>';
  els.stockMovementsList.querySelectorAll('[data-edit-stock-movement]').forEach((button) => {
    button.addEventListener('click', () => openStockMovementDialog(state.stockMovements.find((movement) => movement.id === button.dataset.editStockMovement)));
  });
}

function dateTimeInputValue(value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function formatDateTime(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 'No date' : date.toLocaleString('en-GB', { dateStyle: 'medium', timeStyle: 'short' });
}

function makeRecordId() {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function openStocktakeDialog() {
  els.stocktakeForm.reset();
  els.stocktakeDate.value = dateTimeInputValue();
  els.stocktakeDialog.showModal();
}

function saveStocktake(event) {
  event.preventDefault();
  if (!els.stocktakeForm.reportValidity()) return;
  const now = new Date().toISOString();
  state.stocktakes.push({ id: makeRecordId(), date: new Date(els.stocktakeDate.value).toISOString(), bales: Math.round(Number(els.stocktakeBales.value || 0)), notes: els.stocktakeNotes.value.trim(), createdAt: now, updatedAt: now });
  saveState();
  renderStock();
  els.stocktakeDialog.close();
}

function deleteStocktake(id) {
  if (!window.confirm('Delete this stocktake?')) return;
  state.stocktakes = state.stocktakes.filter((stocktake) => stocktake.id !== id);
  saveState();
  renderStock();
}

function openStockMovementDialog(movement = null) {
  els.stockMovementForm.reset();
  els.stockMovementId.value = movement ? movement.id : '';
  els.stockMovementType.value = movement ? movement.type : 'ducks';
  els.stockMovementDate.value = dateTimeInputValue(movement ? movement.date : new Date());
  els.stockMovementCustomer.value = movement ? movement.customer || '' : '';
  els.stockMovementBales.value = movement ? movement.bales : '';
  els.stockMovementNotes.value = movement ? movement.notes || '' : '';
  els.stockMovementDialogTitle.textContent = movement ? 'Edit Stock Movement' : 'Stock Movement';
  els.deleteStockMovementButton.hidden = !movement;
  updateStockMovementDefaults();
  els.stockMovementDialog.showModal();
}

function updateStockMovementDefaults() {
  if (els.stockMovementType.value === 'ducks' && !els.stockMovementCustomer.value.trim()) els.stockMovementCustomer.value = 'Ducks at home';
  if (els.stockMovementType.value === 'bought-in' && els.stockMovementCustomer.value.trim() === 'Ducks at home') els.stockMovementCustomer.value = '';
}

function saveStockMovement(event) {
  event.preventDefault();
  if (!els.stockMovementForm.reportValidity()) return;
  const existing = state.stockMovements.find((movement) => movement.id === els.stockMovementId.value);
  const now = new Date().toISOString();
  const record = {
    id: existing ? existing.id : makeRecordId(), type: els.stockMovementType.value,
    date: new Date(els.stockMovementDate.value).toISOString(),
    customer: els.stockMovementCustomer.value.trim() || (els.stockMovementType.value === 'ducks' ? 'Ducks at home' : ''),
    bales: Math.round(Number(els.stockMovementBales.value || 0)), notes: els.stockMovementNotes.value.trim(),
    createdAt: existing ? existing.createdAt : now, updatedAt: now
  };
  if (existing) Object.assign(existing, record); else state.stockMovements.push(record);
  if (record.type === 'bought-in') rememberStrawCustomer(record.customer);
  saveState();
  renderStock();
  els.stockMovementDialog.close();
}

function deleteCurrentStockMovement() {
  const id = els.stockMovementId.value;
  if (!id || !window.confirm('Delete this stock movement?')) return;
  state.stockMovements = state.stockMovements.filter((movement) => movement.id !== id);
  saveState();
  renderStock();
  els.stockMovementDialog.close();
}

function localDateTimeValue(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

function openDeliveryDialog(deliveryId = null) {
  const delivery = deliveryId ? state.loads.find((item) => Number(item.id) === Number(deliveryId)) : null;
  els.deliveryForm.reset();
  els.deliveryId.value = delivery ? delivery.id : '';
  els.deliveryCustomer.value = delivery ? delivery.customer || '' : '';
  els.deliveryRegistration.value = delivery ? delivery.registration || '' : '';
  els.deliveryDateTime.value = delivery ? `${delivery.delivery_date || ''}T${delivery.delivery_time || ''}` : localDateTimeValue();
  els.deliveryBales.value = delivery ? delivery.bale_total : '';
  els.deliveryWeight.value = delivery ? delivery.weight_total || '' : '';
  els.deliveryDialogTitle.textContent = delivery ? 'Edit Load Out' : 'Log Load Out';
  els.saveDeliveryButton.textContent = delivery ? 'Save Changes' : 'Save Load Out';
  els.deliveryStatus.textContent = delivery && !delivery.weight_total ? 'Add the weight when it is sent to you.' : 'Weight can be left blank and added later.';
  els.deliveryStatus.dataset.error = 'false';
  els.deliveryDialog.showModal();
}

async function saveDeliveryFromForm(event) {
  event.preventDefault();
  if (!els.deliveryForm.reportValidity()) return;
  const existing = state.loads.find((item) => Number(item.id) === Number(els.deliveryId.value));
  const [deliveryDate, deliveryTime] = els.deliveryDateTime.value.split('T');
  const payload = {
    customer: els.deliveryCustomer.value.trim(),
    registration: els.deliveryRegistration.value.trim(),
    delivery_date: deliveryDate || '',
    delivery_time: deliveryTime || '',
    bale_total: els.deliveryBales.value,
    weight_total: els.deliveryWeight.value
  };
  if (existing) payload.version = existing.version;
  els.saveDeliveryButton.disabled = true;
  els.deliveryStatus.textContent = 'Saving…';
  els.deliveryStatus.dataset.error = 'false';
  try {
    const response = await fetch(existing ? `/api/straw/deliveries/${existing.id}` : '/api/straw/deliveries', {
      method: existing ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || 'Unable to save this delivery.');
    state.loads = state.loads.filter((item) => Number(item.id) !== Number(result.delivery.id));
    state.loads.push(result.delivery);
    state.loads.sort((a, b) => b.delivery_date.localeCompare(a.delivery_date) || b.delivery_time.localeCompare(a.delivery_time) || b.id - a.id);
    saveState();
    addCustomerLocally(result.delivery.customer);
    render();
    els.deliveryDialog.close();
  } catch (error) {
    els.deliveryStatus.textContent = error.message;
    els.deliveryStatus.dataset.error = 'true';
  } finally {
    els.saveDeliveryButton.disabled = false;
  }
}

function openFieldDialog(fieldId = null) {
  const field = fieldId ? state.fields.find((item) => item.id === fieldId) : null;
  els.fieldForm.reset();
  els.fieldId.value = field ? field.id : '';
  els.fieldCustomer.value = field ? field.customer || '' : '';
  els.fieldFarm.value = field ? field.farm || '' : '';
  els.fieldName.value = field ? field.name || '' : '';
  els.fieldHectares.value = field ? field.hectares || '' : '';
  els.fieldBales.value = field ? field.bales || '' : '';
  els.fieldMoisture.value = field ? field.moisture || '' : '';
  const crop = field && field.crop ? field.crop : (els.fieldCrop.options[0]?.value || '');
  if (crop && !Array.from(els.fieldCrop.options).some(option => option.value === crop)) els.fieldCrop.add(new Option(crop, crop));
  els.fieldCrop.value = crop;
  els.fieldPhoto.value = '';
  els.photoPreview.style.display = 'none';
  els.photoPreview.src = field && field.photo ? field.photo : '';
  if (field && field.photo) { els.photoPreview.style.display = 'block'; }
  els.fieldDialog.showModal();
}

function closeFieldDialog() {
  pendingFieldLocation = null;
  els.fieldDialog.close();
}

function saveFieldFromForm(event) {
  event.preventDefault();
  const id = els.fieldId.value || crypto.randomUUID();
  const existing = state.fields.find((item) => item.id === id);
  const payload = {
    ...(existing || {}),
    id,
    customer: els.fieldCustomer.value.trim(),
    farm: els.fieldFarm.value.trim(),
    name: els.fieldName.value.trim(),
    hectares: Number(els.fieldHectares.value || 0),
    bales: Number(els.fieldBales.value || 0),
    moisture: Number(els.fieldMoisture.value || 0),
    crop: els.fieldCrop.value,
    photo: els.photoPreview.src || '',
    status: existing ? existing.status || 'active' : 'active',
    lat: existing && existing.lat !== undefined ? existing.lat : pendingFieldLocation?.lat,
    lng: existing && existing.lng !== undefined ? existing.lng : pendingFieldLocation?.lng,
    updatedAt: new Date().toISOString()
  };
  const existingIndex = state.fields.findIndex((item) => item.id === id);
  if (existingIndex >= 0) state.fields[existingIndex] = payload; else state.fields.push(payload);
  rememberStrawCustomer(payload.customer);
  saveState();
  render();
  closeFieldDialog();
}

function saveFieldWithStatus(status) {
  const id = els.fieldId.value || crypto.randomUUID();
  const existing = state.fields.find((item) => item.id === id);
  const payload = {
    ...(existing || {}),
    id,
    customer: els.fieldCustomer.value.trim(),
    farm: els.fieldFarm.value.trim(),
    name: els.fieldName.value.trim(),
    hectares: Number(els.fieldHectares.value || 0),
    bales: Number(els.fieldBales.value || 0),
    moisture: Number(els.fieldMoisture.value || 0),
    crop: els.fieldCrop.value,
    photo: els.photoPreview.src || '',
    status,
    lat: existing && existing.lat !== undefined ? existing.lat : pendingFieldLocation?.lat,
    lng: existing && existing.lng !== undefined ? existing.lng : pendingFieldLocation?.lng,
    updatedAt: new Date().toISOString()
  };
  const existingIndex = state.fields.findIndex((item) => item.id === id);
  if (existingIndex >= 0) state.fields[existingIndex] = payload; else state.fields.push(payload);
  rememberStrawCustomer(payload.customer);
  saveState();
  render();
  closeFieldDialog();
}

function deleteCurrentField() {
  const id = els.fieldId.value;
  if (!id) return;
  state.fields = state.fields.filter((field) => field.id !== id);
  saveState();
  render();
  closeFieldDialog();
}

function handlePhotoSelection(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    els.photoPreview.src = reader.result;
    els.photoPreview.style.display = 'block';
  };
  reader.readAsDataURL(file);
}

function initMap() {
  if (!window.L) {
    els.mapFallback.classList.add('active');
    return;
  }
  map = L.map('map', { zoomControl: true }).setView([52.569259, 1.406654], 11);
  L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    maxZoom: 19,
    attribution: 'Tiles &copy; Esri'
  }).addTo(map);
  L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}', {
    maxZoom: 19,
    attribution: 'Labels &copy; Esri'
  }).addTo(map);
  markerLayer = L.layerGroup().addTo(map);
  map.on('click', (event) => {
    pendingFieldLocation = {
      lat: Math.round(event.latlng.lat * 1000000) / 1000000,
      lng: Math.round(event.latlng.lng * 1000000) / 1000000
    };
    openFieldDialog();
  });
  renderMapMarkers();
}

function fieldHasLocation(field) {
  return field && Number.isFinite(Number(field.lat)) && Number.isFinite(Number(field.lng));
}

function fieldPinColour(field) {
  if (field.status === 'complete' || field.completed === true) return '#2f8f46';
  if (field.status === 'part-complete' || field.status === 'in-progress') return '#d99a2b';
  return '#c64232';
}

function fieldCropCode(crop) {
  return { Wheat: 'W', Barley: 'B', 'Spring Barley': 'SB', Oats: 'O', Hay: 'H' }[crop] || 'Oth';
}

function makeFieldIcon(field) {
  const bales = Math.round(Number(field.bales || 0));
  const baleLabel = bales >= 10000 ? `${Math.round(bales / 1000)}k` : String(bales);
  return L.divIcon({
    className: 'crop-marker',
    html: `<span style="background:${fieldPinColour(field)}"><b>${escapeHtml(`${baleLabel}-${fieldCropCode(field.crop)}`)}</b></span>`,
    iconSize: [44, 44],
    iconAnchor: [22, 44],
    popupAnchor: [0, -42]
  });
}

function renderMapMarkers() {
  if (!markerLayer) return;
  markerLayer.clearLayers();
  state.fields.filter(fieldHasLocation).forEach((field) => {
    const marker = L.marker([Number(field.lat), Number(field.lng)], { icon: makeFieldIcon(field) }).addTo(markerLayer);
    marker.bindPopup(`<strong>${escapeHtml(field.name || 'Unnamed field')}</strong><br>${escapeHtml(field.customer || 'No customer')} · ${escapeHtml(field.farm || 'No farm')}<br>${escapeHtml(field.crop || 'Other')} · ${Number(field.bales || 0)} bales<br><button class="map-edit-field" type="button">Edit field</button>`);
    marker.on('popupopen', () => {
      marker.getPopup().getElement()?.querySelector('.map-edit-field')?.addEventListener('click', () => openFieldDialog(field.id));
    });
  });
}

function fitMapToFields() {
  if (!map) return;
  const locations = state.fields.filter(fieldHasLocation).map((field) => [Number(field.lat), Number(field.lng)]);
  if (locations.length) {
    map.fitBounds(L.latLngBounds(locations), { padding: [36, 36], maxZoom: 15 });
  } else {
    map.setView([52.569259, 1.406654], 11);
  }
}
