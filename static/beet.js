(() => {
  const form = document.getElementById('beet-form');
  const status = document.getElementById('beet-status');
  const list = document.getElementById('beet-list');
  const search = document.getElementById('beet-search');
  const totals = document.getElementById('beet-totals');
  const fieldOptions = document.getElementById('field-options');
  const fieldMap = window.BEET_FIELD_MAP || {};
  const fieldNames = ['job_date', 'customer', 'field', 'area', 'diesel_used', 'diesel_price'];
  const settings = window.BEET_SETTINGS;
  const settingsForm = document.getElementById('beet-settings-form');
  const settingsStatus = document.getElementById('settings-status');
  const preview = document.getElementById('escalator-preview');
  const money = value => `£${Number(value).toFixed(2)}`;
  let entries = [];
  let current = null;
  let dirty = false;
  let saving = false;

  function today() {
    const date = new Date();
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  }
  function message(text, error = false) {
    status.textContent = text;
    status.dataset.error = String(error);
  }
  function formatDate(value) {
    if (!value) return '';
    return new Date(`${value}T12:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  }
  function updatePreview() {
    const litres = parseFloat(form.elements.diesel_used.value);
    const price = parseFloat(form.elements.diesel_price.value);
    const area = parseFloat(form.elements.area.value);
    const base = current ? Number(current.base_price) : settings.base_price;
    const pct = current ? Number(current.trigger_pct) : settings.trigger_pct;
    const trigger = Math.round(base * (1 + pct / 100) * 10000) / 10000;
    const basis = `base ${(base * 100).toFixed(1)}p, starts at ${(trigger * 100).toFixed(1)}p`;
    if (!(litres >= 0) || !(price >= 0)) {
      preview.textContent = `Diesel escalator (${basis}): enter litres and price.`;
      return;
    }
    if (Math.round(price * 10000) / 10000 < trigger) {
      preview.textContent = `No surcharge: ${(price * 100).toFixed(1)}p is below the ${(trigger * 100).toFixed(1)}p trigger.`;
      return;
    }
    const total = litres * (price - base);
    preview.textContent = area > 0
      ? `${litres} L · surcharge ${money(total / area)} per ha (${money(total)} total on ${area} ha)`
      : `${litres} L · surcharge ${money(total)} total (enter the area for a per-hectare figure)`;
  }
  function updateFieldOptions() {
    const wanted = form.elements.customer.value.trim().toLowerCase();
    const match = Object.keys(fieldMap).find(name => name.toLowerCase() === wanted);
    fieldOptions.replaceChildren(...(match ? fieldMap[match] : []).map(name => {
      const option = document.createElement('option');
      option.value = name;
      return option;
    }));
  }
  function render() {
    list.replaceChildren();
    const query = search.value.trim().toLowerCase();
    const filtered = entries.filter(entry => ['customer', 'field'].some(field => String(entry[field] || '').toLowerCase().includes(query)));
    const area = filtered.reduce((sum, entry) => sum + Number(entry.area), 0);
    const diesel = filtered.reduce((sum, entry) => sum + Number(entry.diesel_used), 0);
    const escalator = filtered.reduce((sum, entry) => sum + Number(entry.escalator), 0);
    totals.textContent = filtered.length ? `${filtered.length} jobs · ${area.toFixed(2)} ha · ${diesel.toFixed(1)} litres · ${money(escalator)} surcharge` + (area > 0 ? ` (${money(escalator / area)} per ha)` : '') : '';
    if (!filtered.length) {
      list.textContent = entries.length ? 'No matching sugar beet jobs.' : 'No sugar beet jobs saved yet.';
      return;
    }
    for (const entry of filtered) {
      const article = document.createElement('article');
      article.className = 'maintenance-record';
      const details = document.createElement('div');
      const heading = document.createElement('h3');
      heading.textContent = `${entry.customer} · ${entry.field}`;
      const info = document.createElement('p');
      const surcharge = Number(entry.escalator) > 0
        ? `surcharge ${entry.escalator_per_ha ? money(entry.escalator_per_ha) + ' per ha' : 'no area'} (${money(entry.escalator)} total)`
        : 'no surcharge';
      info.textContent = `${formatDate(entry.job_date)} · ${entry.area} ha · ${entry.diesel_used} litres · ${surcharge}`;
      details.append(heading, info);
      const open = document.createElement('button');
      open.type = 'button';
      open.className = 'secondary';
      open.textContent = 'Open';
      open.addEventListener('click', () => openEntry(entry));
      article.append(details, open);
      list.append(article);
    }
  }
  async function load() {
    const response = await fetch('/api/beet', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load sugar beet jobs.');
    entries = (await response.json()).entries;
    render();
  }
  function fill(entry) {
    current = entry;
    for (const field of fieldNames) form.elements[field].value = entry ? entry[field] : (field === 'job_date' ? today() : field === 'diesel_price' ? window.BEET_DEFAULT_PRICE : '');
    updateFieldOptions();
    updatePreview();
    dirty = false;
    document.getElementById('record-heading').textContent = entry ? `Edit ${entry.field}` : 'New sugar beet job';
    message(entry ? 'Job opened. Press Save Job after making changes.' : 'Complete the fields and press Save.');
  }
  function openEntry(entry) {
    if (dirty && !window.confirm('Discard your unsaved changes and open this job?')) return;
    fill(entry);
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  form.addEventListener('input', event => {
    if (event.target.name === 'customer') updateFieldOptions();
    updatePreview();
    dirty = true;
    message('Unsaved changes — press Save Job.');
  });
  form.addEventListener('submit', async event => {
    event.preventDefault();
    if (saving || !form.reportValidity()) return;
    saving = true;
    document.getElementById('save-record').disabled = true;
    message('Saving…');
    const payload = Object.fromEntries(fieldNames.map(field => [field, form.elements[field].value]));
    if (current) payload.version = current.version;
    try {
      const response = await fetch(current ? `/api/beet/${current.id}` : '/api/beet', {
        method: current ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(15000),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Unable to save the sugar beet job.');
      current = result.entry;
      entries = entries.filter(entry => entry.id !== current.id);
      entries.push(current);
      entries.sort((a, b) => b.job_date.localeCompare(a.job_date) || b.id - a.id);
      dirty = false;
      render();
      message(`Saved at ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}.`);
    } catch (error) {
      message(`${error.message} Your changes remain in the form.`, true);
    } finally {
      saving = false;
      document.getElementById('save-record').disabled = false;
    }
  });
  document.getElementById('new-record').addEventListener('click', () => {
    if (dirty && !window.confirm('Discard your unsaved changes and start a new job?')) return;
    fill(null);
  });
  settingsForm.elements.base_price.value = settings.base_price;
  settingsForm.elements.trigger_pct.value = settings.trigger_pct;
  settingsForm.addEventListener('input', () => { settingsStatus.textContent = 'Unsaved changes — press Save Escalator Settings.'; settingsStatus.dataset.error = 'false'; });
  settingsForm.addEventListener('submit', async event => {
    event.preventDefault();
    if (!settingsForm.reportValidity()) return;
    try {
      const response = await fetch('/api/beet/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ base_price: settingsForm.elements.base_price.value, trigger_pct: settingsForm.elements.trigger_pct.value }),
        signal: AbortSignal.timeout(15000),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Unable to save the escalator settings.');
      Object.assign(settings, result.settings);
      settingsStatus.textContent = 'Saved. New jobs will use these figures.';
      settingsStatus.dataset.error = 'false';
      updatePreview();
    } catch (error) {
      settingsStatus.textContent = error.message;
      settingsStatus.dataset.error = 'true';
    }
  });
  search.addEventListener('input', render);
  window.addEventListener('beforeunload', event => {
    if (dirty || saving) { event.preventDefault(); event.returnValue = ''; }
  });
  fill(null);
  load().catch(error => { list.textContent = error.message; });
})();
