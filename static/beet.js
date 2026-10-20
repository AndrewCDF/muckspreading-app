(() => {
  const form = document.getElementById('beet-form');
  const status = document.getElementById('beet-status');
  const list = document.getElementById('beet-list');
  const filter = document.getElementById('beet-filter');
  const totals = document.getElementById('beet-totals');
  const fieldOptions = document.getElementById('field-options');
  const fieldMap = window.BEET_FIELD_MAP || {};
  const fieldNames = ['job_date', 'customer', 'area', 'diesel_used', 'tractor_diesel', 'diesel_price'];
  const fieldRows = document.getElementById('field-rows');
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
  const trim = value => String(Math.round(value * 10000) / 10000);
  // A saved job quotes the base and trigger it was recorded with; a new job uses the current settings.
  function termsText() {
    const base = current ? Number(current.base_price) : settings.base_price;
    const pct = current ? Number(current.trigger_pct) : settings.trigger_pct;
    const ppl = trim(base * 100);
    return `All contracting charges are worked at ${ppl}ppl. We absorb any diesel price fluctuation within ${trim(pct)}% (up to ${trim(base * (1 + pct / 100) * 100)}ppl). Above that, a surcharge is added on the full rise over ${ppl}ppl.`;
  }
  function showTerms() {
    document.getElementById('escalator-terms').textContent = termsText();
  }
  function updatePreview() {
    const harvester = parseFloat(form.elements.diesel_used.value);
    const tractor = parseFloat(form.elements.tractor_diesel.value) || 0;
    const litres = harvester + tractor;
    const price = parseFloat(form.elements.diesel_price.value);
    const area = parseFloat(form.elements.area.value);
    const base = current ? Number(current.base_price) : settings.base_price;
    const pct = current ? Number(current.trigger_pct) : settings.trigger_pct;
    const trigger = Math.round(base * (1 + pct / 100) * 10000) / 10000;
    const basis = `base ${(base * 100).toFixed(1)}p, starts at ${(trigger * 100).toFixed(1)}p`;
    if (!(price >= 0)) {
      preview.textContent = `Surcharge not calculated yet — no diesel price entered (${basis}).`;
      return;
    }
    if (!(litres >= 0)) {
      preview.textContent = `Diesel escalator (${basis}): enter litres used.`;
      return;
    }
    if (Math.round(price * 10000) / 10000 < trigger) {
      preview.textContent = `No surcharge: ${(price * 100).toFixed(1)}p is below the ${(trigger * 100).toFixed(1)}p trigger.`;
      return;
    }
    const total = litres * (price - base);
    const litreText = tractor > 0 ? `${harvester} L harvester + ${tractor} L tractor = ${litres} L` : `${litres} L`;
    preview.textContent = area > 0
      ? `${litreText} · surcharge ${money(total / area)} per ha (${money(total)} total on ${area} ha)`
      : `${litreText} · surcharge ${money(total)} total (enter the area for a per-hectare figure)`;
  }
  function addFieldRow(name = '') {
    const row = document.createElement('div');
    row.className = 'field-row';
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'field-input';
    input.setAttribute('list', 'field-options');
    input.maxLength = 160;
    input.autocomplete = 'off';
    input.value = name;
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'secondary';
    remove.textContent = 'Remove';
    remove.addEventListener('click', () => {
      if (fieldRows.children.length === 1) {
        input.value = '';
      } else {
        row.remove();
      }
      numberFieldRows();
      dirty = true;
      message('Unsaved changes — press Save Job.');
    });
    row.append(input, remove);
    fieldRows.append(row);
    numberFieldRows();
    return input;
  }
  function numberFieldRows() {
    [...fieldRows.children].forEach((row, index) => {
      row.querySelector('input').placeholder = `Field ${index + 1}`;
      row.querySelector('input').setAttribute('aria-label', `Field ${index + 1} name`);
    });
  }
  function setFieldRows(names) {
    fieldRows.replaceChildren();
    for (const name of (names.length ? names : [''])) addFieldRow(name);
  }
  // One blank row means no fields were entered; with several rows, blank names are numbered.
  function fieldValues() {
    const names = [...fieldRows.querySelectorAll('input')].map(input => input.value.trim());
    return names.length === 1 && !names[0] ? [] : names;
  }
  document.getElementById('add-field').addEventListener('click', () => {
    addFieldRow().focus();
    dirty = true;
    message('Unsaved changes — press Save Job.');
  });
  function updateFieldOptions() {
    const wanted = form.elements.customer.value.trim().toLowerCase();
    const match = Object.keys(fieldMap).find(name => name.toLowerCase() === wanted);
    fieldOptions.replaceChildren(...(match ? fieldMap[match] : []).map(name => {
      const option = document.createElement('option');
      option.value = name;
      return option;
    }));
  }
  function matchesFilter(entry) {
    switch (filter.value) {
      case 'unpriced': return !entry.priced;
      case 'ready': return entry.priced && !entry.invoiced;
      case 'invoiced': return entry.invoiced;
      default: return true;
    }
  }
  function render() {
    list.replaceChildren();
    const filtered = entries.filter(matchesFilter)
      .sort((a, b) => b.job_date.localeCompare(a.job_date) || b.id - a.id);
    const priced = filtered.filter(entry => entry.priced);
    const unpriced = filtered.length - priced.length;
    const area = filtered.reduce((sum, entry) => sum + Number(entry.area), 0);
    const diesel = filtered.reduce((sum, entry) => sum + Number(entry.total_diesel), 0);
    const escalator = priced.reduce((sum, entry) => sum + Number(entry.escalator), 0);
    totals.textContent = filtered.length
      ? `${filtered.length} jobs · ${area.toFixed(2)} ha · ${diesel.toFixed(1)} litres · ${money(escalator)} surcharge` + (unpriced ? ` on ${priced.length} priced (${unpriced} still need a diesel price)` : '')
      : '';
    if (!filtered.length) {
      list.textContent = entries.length ? 'No sugar beet jobs match this view.' : 'No sugar beet jobs saved yet.';
      return;
    }
    let monthList = null;
    let monthKey = '';
    for (const entry of filtered) {
      if (entry.job_date.slice(0, 7) !== monthKey) {
        monthKey = entry.job_date.slice(0, 7);
        const inMonth = filtered.filter(item => item.job_date.slice(0, 7) === monthKey);
        const monthArea = inMonth.reduce((sum, item) => sum + Number(item.area), 0);
        const monthSurcharge = inMonth.filter(item => item.priced).reduce((sum, item) => sum + Number(item.escalator), 0);
        const section = document.createElement('section');
        section.className = 'beet-month';
        const monthHeading = document.createElement('h3');
        monthHeading.className = 'month-heading';
        monthHeading.textContent = new Date(`${monthKey}-01T12:00:00`).toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
        const monthTotals = document.createElement('p');
        monthTotals.className = 'month-totals';
        monthTotals.textContent = `${inMonth.length} ${inMonth.length === 1 ? 'job' : 'jobs'} · ${monthArea.toFixed(2)} ha · ${money(monthSurcharge)} surcharge`;
        section.append(monthHeading, monthTotals);
        list.append(section);
        monthList = section;
      }
      const article = document.createElement('article');
      article.className = 'maintenance-record';
      const details = document.createElement('div');
      const heading = document.createElement('h3');
      heading.textContent = entry.field ? `${entry.customer} · ${entry.field}` : entry.customer;
      const info = document.createElement('p');
      const priceText = !entry.priced ? 'diesel price needed'
        : Number(entry.escalator) > 0 ? `surcharge ${entry.escalator_per_ha ? money(entry.escalator_per_ha) + ' per ha (' + money(entry.escalator) + ' total)' : money(entry.escalator) + ' total'}`
        : 'no surcharge';
      info.textContent = `${formatDate(entry.job_date)} · ${entry.area ? entry.area + ' ha · ' : ''}${entry.tractor_diesel ? `${entry.diesel_used} L harvester + ${entry.tractor_diesel} L tractor` : `${entry.diesel_used} litres`} · `;
      const state = document.createElement('span');
      if (!entry.priced) state.className = 'price-needed';
      state.textContent = priceText;
      info.append(state);
      if (entry.invoiced) {
        const badge = document.createElement('p');
        badge.className = 'invoiced-badge';
        badge.textContent = `Invoiced ${formatDate(entry.invoiced_date)}`;
        details.append(heading, info, badge);
      } else {
        details.append(heading, info);
      }
      const open = document.createElement('button');
      open.type = 'button';
      open.className = 'secondary';
      open.textContent = 'Open';
      open.addEventListener('click', () => openEntry(entry));
      const actions = document.createElement('div');
      actions.className = 'record-actions';
      actions.append(open);
      if (entry.invoiced || entry.priced) {
        const toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = entry.invoiced ? 'secondary' : 'primary';
        toggle.textContent = entry.invoiced ? 'Undo invoiced' : 'Mark invoiced';
        toggle.addEventListener('click', () => setInvoiced(entry, !entry.invoiced));
        actions.append(toggle);
      }
      article.append(details, actions);
      monthList.append(article);
    }
  }
  async function setInvoiced(entry, invoiced) {
    if (current && current.id === entry.id && dirty && !window.confirm('Discard your unsaved changes to this job?')) return;
    try {
      const payload = { ...Object.fromEntries(fieldNames.map(field => [field, entry[field]])), fields: entry.fields };
      const response = await fetch(`/api/beet/${entry.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...payload, invoiced, version: entry.version }),
        signal: AbortSignal.timeout(15000),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Unable to update the job.');
      replaceEntry(result.entry);
      if (current && current.id === entry.id) fill(result.entry);
      message(invoiced ? 'Marked as invoiced.' : 'Invoice marker removed.');
    } catch (error) {
      message(error.message, true);
    }
  }
  function replaceEntry(saved) {
    entries = entries.filter(item => item.id !== saved.id);
    entries.push(saved);
    entries.sort((a, b) => b.job_date.localeCompare(a.job_date) || b.id - a.id);
    render();
  }
  async function load() {
    const response = await fetch('/api/beet', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load sugar beet jobs.');
    entries = (await response.json()).entries;
    render();
  }
  function fill(entry) {
    current = entry;
    for (const field of fieldNames) form.elements[field].value = entry ? entry[field] : (field === 'job_date' ? today() : '');
    form.elements.invoiced.checked = Boolean(entry && entry.invoiced);
    setFieldRows(entry ? entry.fields : []);
    updateFieldOptions();
    updatePreview();
    showTerms();
    dirty = false;
    document.getElementById('record-heading').textContent = entry ? `Edit ${entry.field || entry.customer}` : 'New sugar beet job';
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
    const payload = { ...Object.fromEntries(fieldNames.map(field => [field, form.elements[field].value])), fields: fieldValues() };
    payload.invoiced = form.elements.invoiced.checked;
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
      dirty = false;
      replaceEntry(result.entry);
      fill(null);
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
      showTerms();
      settingsStatus.textContent = 'Saved. New jobs will use these figures.';
      settingsStatus.dataset.error = 'false';
      updatePreview();
    } catch (error) {
      settingsStatus.textContent = error.message;
      settingsStatus.dataset.error = 'true';
    }
  });
  showTerms();
  const settingsPanel = document.getElementById('escalator-settings');
  const settingsToggle = document.getElementById('toggle-settings');
  settingsToggle.addEventListener('click', () => {
    const open = settingsPanel.hidden;
    settingsPanel.hidden = !open;
    settingsToggle.setAttribute('aria-expanded', String(open));
    settingsToggle.textContent = open ? 'Hide Diesel Escalator Settings' : 'Diesel Escalator Settings';
    if (open) settingsPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  });
  filter.addEventListener('change', render);
  window.addEventListener('beforeunload', event => {
    if (dirty || saving) { event.preventDefault(); event.returnValue = ''; }
  });
  fill(null);
  load().catch(error => { list.textContent = error.message; });
})();
