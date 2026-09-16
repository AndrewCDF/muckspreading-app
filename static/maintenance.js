(() => {
  const form = document.getElementById('maintenance-form');
  const status = document.getElementById('maintenance-status');
  const list = document.getElementById('maintenance-list');
  const search = document.getElementById('maintenance-search');
  const fieldNames = ['maintenance_date', 'machinery', 'machine_hours', 'work_completed'];
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
  function render() {
    list.replaceChildren();
    const query = search.value.trim().toLowerCase();
    const filtered = entries.filter(entry => fieldNames.some(field => String(entry[field] || '').toLowerCase().includes(query)));
    if (!filtered.length) {
      list.textContent = entries.length ? 'No matching maintenance records.' : 'No maintenance records saved yet.';
      return;
    }
    for (const entry of filtered) {
      const article = document.createElement('article');
      article.className = 'maintenance-record';
      const details = document.createElement('div');
      const heading = document.createElement('h3');
      heading.textContent = `${entry.machinery} · ${formatDate(entry.maintenance_date)}`;
      const work = document.createElement('p');
      work.className = 'work';
      work.textContent = entry.work_completed;
      details.append(heading, work);
      if (entry.machine_hours) {
        const hours = document.createElement('p');
        hours.textContent = `Machine hours: ${entry.machine_hours}`;
        details.append(hours);
      }
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
    const response = await fetch('/api/maintenance', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load maintenance records.');
    entries = (await response.json()).entries;
    render();
  }
  function fill(entry) {
    current = entry;
    for (const field of fieldNames) form.elements[field].value = entry ? entry[field] : (field === 'maintenance_date' ? today() : '');
    dirty = false;
    document.getElementById('record-heading').textContent = entry ? `Edit ${entry.machinery}` : 'New maintenance record';
    message(entry ? 'Record opened. Press Save Maintenance Record after making changes.' : 'Complete the required fields and press Save.');
  }
  function openEntry(entry) {
    if (dirty && !window.confirm('Discard your unsaved changes and open this record?')) return;
    fill(entry);
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  form.addEventListener('input', () => {
    dirty = true;
    message('Unsaved changes — press Save Maintenance Record.');
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
      const response = await fetch(current ? `/api/maintenance/${current.id}` : '/api/maintenance', {
        method: current ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(15000),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Unable to save the maintenance record.');
      current = result.entry;
      entries = entries.filter(entry => entry.id !== current.id);
      entries.push(current);
      entries.sort((a, b) => b.maintenance_date.localeCompare(a.maintenance_date) || b.id - a.id);
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
    if (dirty && !window.confirm('Discard your unsaved changes and start a new record?')) return;
    fill(null);
  });
  search.addEventListener('input', render);
  window.addEventListener('beforeunload', event => {
    if (dirty || saving) { event.preventDefault(); event.returnValue = ''; }
  });
  fill(null);
  load().catch(error => { list.textContent = error.message; });
})();
