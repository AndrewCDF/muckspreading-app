(() => {
  const form = document.getElementById('timesheet-form');
  const staffNames = form ? Array.from(form.elements.name.options, option => option.value) : [];
  const status = document.getElementById('save-status');
  const list = document.getElementById('saved-entries');
  const calendar = document.getElementById('month-calendar');
  const editor = document.getElementById('day-editor');
  const calendarStatus = document.getElementById('calendar-status');
  const parameters = new URLSearchParams(window.location.search);
  let displayedMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  if (/^\d{4}-(0[1-9]|1[0-2])$/.test(parameters.get('month') || '')) {
    displayedMonth = new Date(`${parameters.get('month')}-01T12:00:00`);
  }
  const selectedDate = editor ? editor.dataset.date : null;
  const fields = ['name', 'date', 'company', 'start', 'finish', 'notes'];
  let current = null;
  let entries = [];
  let dirty = false;
  let saving = null;
  let revision = 0;
  let conflict = false;
  let switching = false;

  function today() {
    return dateKey(new Date());
  }
  function dateKey(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  }
  function entrySegmentForDate(entry, date) {
    const segments = Array.isArray(entry.daily_segments) ? entry.daily_segments : [];
    const segment = segments.find(item => item.date === date);
    if (segment) return segment;
    if (entry.date === date) {
      return { date, start: entry.start, finish: entry.finish, duration_label: entry.duration_label || '' };
    }
    return null;
  }
  function renderCalendar() {
    if (!calendar) return;
    const monthKey = `${displayedMonth.getFullYear()}-${String(displayedMonth.getMonth() + 1).padStart(2, '0')}`;
    document.getElementById('month-export').href = `/timesheet/month/${monthKey}/export.xlsx`;
    document.getElementById('month-print').href = `/timesheet/month/${monthKey}/print`;
    document.getElementById('complete-month-value').value = monthKey;
    document.getElementById('calendar-month').textContent = displayedMonth.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
    calendar.replaceChildren();
    const year = displayedMonth.getFullYear();
    const month = displayedMonth.getMonth();
    const offset = (displayedMonth.getDay() + 6) % 7;
    const days = new Date(year, month + 1, 0).getDate();
    const cells = Math.ceil((offset + days) / 7) * 7;
    for (let index = 0; index < cells; index++) {
      const day = index - offset + 1;
      if (day < 1 || day > days) {
        const blank = document.createElement('div');
        blank.className = 'calendar-blank';
        blank.setAttribute('aria-hidden', 'true');
        calendar.append(blank);
        continue;
      }
      const date = dateKey(new Date(year, month, day));
      const records = entries.filter(entry => entrySegmentForDate(entry, date));
      const button = document.createElement('a');
      button.href = `/timesheet/day/${date}`;
      button.className = 'calendar-day';
      button.dataset.date = date;
      button.classList.toggle('is-today', date === today());
      button.classList.toggle('is-selected', date === selectedDate);
      if (date === today()) button.setAttribute('aria-current', 'date');
      button.setAttribute('aria-label', `${new Date(year, month, day).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}, ${records.length} saved ${records.length === 1 ? 'entry' : 'entries'}`);
      const number = document.createElement('span');
      number.className = 'day-number';
      number.textContent = day;
      button.append(number);
      for (const company of new Set(records.map(entry => entry.company))) {
        const companyEntries = records.filter(entry => entry.company === company);
        if (!companyEntries.length) continue;
        const badge = document.createElement('span');
        badge.className = company.startsWith('Cherry') ? 'calendar-company cherry' : 'calendar-company farrell';
        const fullLabel = document.createElement('span');
        fullLabel.className = 'company-full';
        fullLabel.textContent = company === 'Cherry Dene Farm Ltd.' ? 'Cherry Dene' : company === 'A. Farrell Contracting Ltd.' ? 'A. Farrell' : company;
        const shortLabel = document.createElement('span');
        shortLabel.className = 'company-short';
        shortLabel.textContent = company === 'Cherry Dene Farm Ltd.' ? 'CD' : company === 'A. Farrell Contracting Ltd.' ? 'AF' : company.split(/\s+/).slice(0, 2).map(word => word[0]).join('');
        shortLabel.setAttribute('aria-hidden', 'true');
        badge.append(fullLabel, shortLabel);
        button.setAttribute('aria-label', `${button.getAttribute('aria-label')}, ${company}`);
        badge.title = companyEntries.map(entry => {
          const segment = entrySegmentForDate(entry, date);
          return `${entry.name}: ${segment.start || '—'}–${segment.finish || '—'}${segment.duration_label ? ` (${segment.duration_label})` : ''}`;
        }).join('\n');
        button.append(badge);
      }
      calendar.append(button);
    }
  }
  function message(text, error = false) {
    status.textContent = text;
    status.dataset.error = String(error);
  }
  function durationLabel(totalMinutes) {
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    const parts = [];
    if (hours) parts.push(`${hours} hour${hours === 1 ? '' : 's'}`);
    if (minutes || !parts.length) parts.push(`${minutes} minute${minutes === 1 ? '' : 's'}`);
    return parts.join(' ');
  }
  function renderShiftSummary() {
    if (!form) return;
    const summary = document.getElementById('shift-summary');
    const start = form.elements.start.value;
    const finish = form.elements.finish.value;
    if (!start || !finish) {
      summary.textContent = '';
      return;
    }
    const [startHour, startMinute] = start.split(':').map(Number);
    const [finishHour, finishMinute] = finish.split(':').map(Number);
    const startTotal = startHour * 60 + startMinute;
    let finishTotal = finishHour * 60 + finishMinute;
    const nextDay = finishTotal < startTotal;
    if (nextDay) finishTotal += 24 * 60;
    const total = durationLabel(finishTotal - startTotal);
    if (nextDay) {
      const finishDate = new Date(`${form.elements.date.value}T12:00:00`);
      finishDate.setDate(finishDate.getDate() + 1);
      summary.textContent = `Finishes the next day, ${finishDate.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}. Total: ${total}.`;
    } else {
      summary.textContent = `Total: ${total}.`;
    }
  }
  function renderEntries() {
    renderCalendar();
    if (!list) return;
    list.replaceChildren();
    const filtered = entries.filter(entry => entrySegmentForDate(entry, selectedDate));
    if (!filtered.length) {
      list.textContent = 'No entries for this day yet. Add the details below.';
    }
    for (const entry of filtered) {
      const row = document.createElement('article');
      row.className = 'saved-day';
      const details = document.createElement('div');
      const title = document.createElement('strong');
      title.textContent = `${selectedDate} · ${entry.name}`;
      const company = document.createElement('p');
      company.textContent = entry.company;
      const segment = entrySegmentForDate(entry, selectedDate);
      const hours = document.createElement('p');
      hours.textContent = segment.start || segment.finish ? `${segment.start || 'Start not entered'} – ${segment.finish || 'In progress'}${segment.duration_label ? ` · ${segment.duration_label}` : ''}` : 'Times not entered';
      if (entry.date !== selectedDate) {
        const carried = document.createElement('p');
        carried.textContent = `Continued from ${new Date(`${entry.date}T12:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'long' })}`;
        details.append(title, company, hours, carried);
      } else {
        details.append(title, company, hours);
      }
      const notes = document.createElement('p');
      notes.textContent = entry.notes.slice(0, 180) + (entry.notes.length > 180 ? '…' : '');
      details.append(notes);
      const open = entry.date === selectedDate ? document.createElement('button') : document.createElement('a');
      open.className = 'secondary';
      open.textContent = entry.date === selectedDate ? 'Open' : 'Open original day';
      open.setAttribute('aria-label', `Open ${entry.date}, ${entry.name}, ${entry.company}`);
      if (entry.date === selectedDate) {
        open.type = 'button';
        open.addEventListener('click', () => switchEntry(entry.id));
      } else {
        open.href = `/timesheet/day/${entry.date}`;
      }
      row.append(details, open);
      list.append(row);
    }
  }
  async function loadEntries() {
    const response = await fetch('/api/timesheets', { cache: 'no-store' });
    if (!response.ok) throw new Error('Unable to load saved days. Refresh to try again.');
    entries = (await response.json()).entries;
    if (calendarStatus) calendarStatus.textContent = 'Company labels show days with saved entries.';
    renderEntries();
  }
  function fill(entry) {
    current = entry;
    const nameSelect = form.elements.name;
    nameSelect.replaceChildren(...staffNames.map(name => new Option(name, name)));
    if (entry && !staffNames.includes(entry.name)) nameSelect.add(new Option(entry.name, entry.name));
    for (const field of fields) {
      const value = entry ? entry[field] : (field === 'date' ? selectedDate : field === 'name' ? staffNames[0] || '' : '');
      if ((field === 'start' || field === 'finish') && value && !Array.from(form.elements[field].options).some(option => option.value === value)) {
        form.elements[field].add(new Option(value, value));
      }
      form.elements[field].value = value;
    }
    renderShiftSummary();
    dirty = false;
    conflict = false;
    revision++;
    message(entry ? 'Saved day opened. You can keep adding jobs.' : 'Enter your details to start.');
  }
  async function save() {
    if (saving) return saving;
    if (switching || document.getElementById('entry-fields').disabled) return false;
    if (!dirty && current) return true;
    if (conflict) return false;
    if (!form.checkValidity()) {
      message('Not saved yet — enter your name, date and company, and check any times.', true);
      return false;
    }
    const sentRevision = revision;
    const data = Object.fromEntries(fields.map(field => [field, form.elements[field].value]));
    if (current) data.version = current.version;
    message('Saving…');
    document.getElementById('save-entry').disabled = true;
    saving = (async () => {
      try {
        const response = await fetch(current ? `/api/timesheets/${current.id}` : '/api/timesheets', {
          method: current ? 'PUT' : 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(data),
          signal: AbortSignal.timeout(15000),
        });
        const payload = await response.json();
        if (!response.ok) {
          conflict = response.status === 409;
          throw new Error(payload.error || 'Unable to save. Please try again.');
        }
        current = payload.entry;
        entries = entries.filter(entry => entry.id !== current.id);
        entries.push(current);
        entries.sort((a, b) => b.date.localeCompare(a.date) || a.name.localeCompare(b.name));
        renderEntries();
        dirty = revision !== sentRevision;
        message(dirty ? 'Earlier changes saved. Press Save timesheet to save your latest changes.' : `Saved at ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}. You can return and add more later.`);
        return true;
      } catch (error) {
        message(`${error.message} Your changes are still in this form.`, true);
        return false;
      }
    })();
    const result = await saving;
    saving = null;
    document.getElementById('save-entry').disabled = false;
    return result;
  }
  async function switchEntry(id, anotherCompany = false) {
    if (switching || saving || document.getElementById('entry-fields').disabled) return;
    switching = true;
    try {
      if (anotherCompany && !form.reportValidity()) return;
      if (dirty && !window.confirm('Your changes have not been saved. Press Cancel to keep editing and use Save timesheet, or OK to discard them.')) return;
      document.getElementById('entry-fields').disabled = true;
      if (id !== null) {
        await loadEntries();
        const entry = entries.find(item => item.id === id);
        if (!entry) throw new Error('This saved day could not be found.');
        fill(entry);
      } else {
        const name = form.elements.name.value;
        const company = form.elements.company.value;
        const date = form.elements.date.value;
        if (anotherCompany) {
          const otherOption = Array.from(form.elements.company.options).find(option => option.value && option.value !== company);
          if (!otherOption) throw new Error('Add another active company in settings.xlsx first.');
          const otherCompany = otherOption.value;
          await loadEntries();
          const existing = entries.find(entry => entry.name.toLowerCase() === name.trim().toLowerCase() && entry.date === date && entry.company === otherCompany);
          if (existing) {
            fill(existing);
            message('Opened the other company’s entry for the same day. You can keep adding jobs.');
          } else {
            fill(null);
            if (!Array.from(form.elements.name.options).some(option => option.value === name)) form.elements.name.add(new Option(name, name));
            form.elements.name.value = name;
            form.elements.date.value = date;
            form.elements.company.value = otherCompany;
            message('Same person and day, other company. Add its times and job notes below.');
          }
        }
      }
      renderEntries();
      document.getElementById('timesheet-heading').focus({ preventScroll: true });
    } catch (error) {
      message(error.message, true);
    } finally {
      document.getElementById('entry-fields').disabled = false;
      switching = false;
    }
  }
  if (form) {
    form.addEventListener('input', () => {
    dirty = true;
    revision++;
    renderShiftSummary();
    if (!conflict) {
      message('Unsaved changes — press Save timesheet to save.');
    }
  });
  form.addEventListener('submit', event => { event.preventDefault(); save(); });
  document.querySelectorAll('.calendar-return').forEach(link => link.addEventListener('click', event => {
    event.preventDefault();
    if (switching || saving) return;
    if (dirty && !window.confirm('Leave without saving your changes? Press Cancel to return and use Save timesheet.')) return;
    dirty = false;
    window.location.assign(link.href);
  }));
  document.getElementById('another-company').addEventListener('click', () => switchEntry(null, true));
  } else {
  document.getElementById('previous-month').addEventListener('click', () => {
    displayedMonth = new Date(displayedMonth.getFullYear(), displayedMonth.getMonth() - 1, 1);
    renderCalendar();
  });
  document.getElementById('next-month').addEventListener('click', () => {
    displayedMonth = new Date(displayedMonth.getFullYear(), displayedMonth.getMonth() + 1, 1);
    renderCalendar();
  });
  document.getElementById('today-month').addEventListener('click', () => {
    displayedMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    renderCalendar();
    window.location.assign(`/timesheet/day/${today()}`);
  });
  document.getElementById('complete-month-form').addEventListener('submit', event => {
    const label = displayedMonth.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
    if (!window.confirm(`Complete ${label} and email the full timesheet to the selected recipients?`)) {
      event.preventDefault();
      return;
    }
    document.getElementById('complete-month').disabled = true;
    document.getElementById('complete-month').textContent = 'Sending…';
  });
  }
  window.addEventListener('beforeunload', event => {
    if (dirty || saving) { event.preventDefault(); event.returnValue = ''; }
  });
  if (form) {
    fill(null);
    document.getElementById('entry-fields').disabled = true;
    message('Loading this day…');
    loadEntries().then(() => {
      const dayEntries = entries.filter(entry => entry.date === selectedDate);
      const names = new Set(dayEntries.map(entry => entry.name.toLowerCase()));
      const existing = names.size === 1 ? dayEntries[0] : null;
      fill(existing || null);
      document.getElementById('entry-fields').disabled = false;
    }).catch(error => message(error.message, true));
  } else {
  renderCalendar();
  loadEntries().catch(error => { calendarStatus.textContent = error.message; });
  }
})();
