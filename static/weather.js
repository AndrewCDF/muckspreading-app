(() => {
  const VERDICT = { grit: 'GRITTING REQUIRED', watch: 'Keep an eye on it', ok: 'NO Gritting' };
  const statusEl = document.getElementById('wx-status');
  let forecast = null;

  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const num = (value, digits = 1) => (value === null || value === undefined || Number.isNaN(Number(value)) ? '–' : Number(value).toFixed(digits).replace(/\.0$/, digits === 1 ? '' : '.0'));
  const deg = value => (value === null || value === undefined ? '–' : `${num(value)}°`);
  const degC = value => (value === null || value === undefined ? '–' : `${num(value)}°C`);
  const mm = value => (value === null || value === undefined ? '–' : `${num(value)} mm`);

  function stat(label, value, extra = '', cls = '') {
    return `<div class="wx-stat ${cls}"><span>${esc(label)}</span><strong>${value}</strong>${extra ? `<small>${extra}</small>` : ''}</div>`;
  }

  function renderNow(now, hours) {
    const road = now.surface === null || now.surface === undefined ? '–' : degC(now.surface);
    const tonight = now.tonight;
    const nextRain = now.next_rain
      ? `${now.next_rain.starts_in_hours === 0 ? 'Raining now' : `Next rain ${esc(now.next_rain.from)}`}`
      : 'Dry for the whole forecast';
    const nextRainDetail = now.next_rain ? `${esc(now.next_rain.from)} to ${esc(now.next_rain.until)}, ${num(now.next_rain.mm)} mm` : '';
    const roadClass = now.level === 'grit' ? 'wx-grit' : now.level === 'watch' ? 'wx-watch' : '';
    document.getElementById('wx-now-body').innerHTML = `
      <div class="wx-now">
        <div class="wx-now-main">
          <div class="wx-now-icon" aria-hidden="true">${esc(now.icon)}</div>
          <div><div class="wx-now-temp">${deg(now.temp)}</div><div class="wx-now-desc">${esc(now.description)} · feels like ${deg(now.feels)}</div>
            <div class="wx-now-sub">High ${deg(now.high)}${now.high_time ? ` at ${esc(now.high_time)}` : ''} · Low ${deg(now.low)}${now.low_time ? ` at ${esc(now.low_time)}` : ''}<br>Sunrise ${esc(now.sunrise)} · Sunset ${esc(now.sunset)}</div></div>
        </div>
        <div class="wx-stats">
          ${stat('Road surface now', road, `${now.level ? esc(VERDICT[now.level] || '') : ''}${now.ground6 !== null && now.ground6 !== undefined ? ` · ground at 6 cm ${degC(now.ground6)}` : ''}`, `wx-road ${roadClass}`)}
          ${tonight ? stat(esc(tonight.label), esc(VERDICT[tonight.level] || ''), `road surface low ${degC(tonight.min_surface)} at ${esc(tonight.min_surface_day)} ${esc(tonight.min_surface_time)}${tonight.rain_advice ? ` · ${esc(tonight.rain_advice)}` : ''}`, `wx-road ${tonight.level === 'grit' ? 'wx-grit' : tonight.level === 'watch' ? 'wx-watch' : ''}`) : ''}
          ${stat('Rain', nextRain, `${nextRainDetail ? `${nextRainDetail} · ` : ''}${mm(now.rain)} now · ${num(now.rain_next_6h)} mm in the next 6 hours (up to ${num(now.prob_next_6h, 0)}% chance) · ${num(now.rain_today)} mm today in total`, 'wx-road')}
          ${stat('Wind', `${num(now.wind)} mph ${esc(now.compass)}`, `gusts ${num(now.gust)} mph`)}
          ${stat('Humidity', `${num(now.humidity, 0)}%`, `dew point ${degC(now.dew)}`)}
          ${stat('Pressure', `${num(now.pressure, 0)} hPa`, now.pressure_trend ? `${esc(now.pressure_trend)} (${now.pressure_change > 0 ? '+' : ''}${num(now.pressure_change)} hPa in 3 hours)` : '')}
          ${stat('Cloud', `${num(now.cloud, 0)}%`, now.visibility_km !== null && now.visibility_km !== undefined ? `visibility ${num(now.visibility_km)} km` : '')}
          ${stat('UV index', num(now.uv), `highest today ${num(now.uv_max)}`)}
        </div>
        <div class="wx-next-hours wx-chart-block">
          <button type="button" class="wx-chart-toggle" id="wx-hours-toggle" aria-expanded="false" aria-controls="wx-hours-body"><span>Next 12 hours</span><span class="wx-chart-chevron" aria-hidden="true">＋</span></button>
          <div class="wx-chart-body" id="wx-hours-body" hidden>
          <div class="wx-table-wrap"><table class="wx-table"><thead><tr><th>Time</th><th>Air</th><th>Road</th><th>Rain mm</th><th>Chance</th><th>Wind mph</th></tr></thead><tbody>${hours.slice(0, 12).map(hour => {
            const cls = hour.level === 'grit' ? 'wx-row-grit' : hour.level === 'watch' ? 'wx-row-watch' : '';
            return `<tr class="${cls}"><td>${esc(hour.hour)} ${esc(hour.icon)}</td><td>${deg(hour.temp)}</td><td>${deg(hour.surface)}</td><td>${num(hour.rain)}</td><td>${num(hour.prob, 0)}%</td><td>${num(hour.wind)} ${esc(hour.compass)}</td></tr>`;
          }).join('')}</tbody></table></div>
          </div>
        </div>
      </div>`;
    const toggle = document.getElementById('wx-hours-toggle');
    toggle.addEventListener('click', () => {
      const body = document.getElementById('wx-hours-body');
      body.hidden = !body.hidden;
      toggle.setAttribute('aria-expanded', String(!body.hidden));
      toggle.querySelector('.wx-chart-chevron').textContent = body.hidden ? '＋' : '－';
    });
  }

  // ---- charts -------------------------------------------------------------------------------------------
  const CHARTS = [
    { title: 'Temperature and road surface (°C)', unit: '°', freeze: true, bar: { key: 'rain', label: 'Rain each hour', unit: ' mm', cls: 's-rainkey', share: 0.3 },
      series: [{ key: 'temp', label: 'Air', cls: 's-air' }, { key: 'feels', label: 'Feels like', cls: 's-feels' }, { key: 'surface', label: 'Road surface', cls: 's-surface' }] },
    { daily: true, title: 'Rain per day (mm)' },
    { title: 'Chance of rain (%)', unit: '%', min: 0, max: 100, series: [{ key: 'prob', label: 'Chance', cls: 's-prob' }] },
    { title: 'Wind (mph)', unit: '', min: 0, series: [{ key: 'wind', label: 'Wind', cls: 's-wind' }, { key: 'gust', label: 'Gusts', cls: 's-gust' }] },
    { title: 'Pressure (hPa)', unit: '', series: [{ key: 'pressure', label: 'Pressure', cls: 's-pressure' }] },
    { title: 'Cloud cover and humidity (%)', unit: '%', min: 0, max: 100, series: [{ key: 'cloud', label: 'Cloud', cls: 's-cloud' }, { key: 'humidity', label: 'Humidity', cls: 's-humidity' }] },
  ];

  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const dateNumber = hour => Number(hour.date.slice(8, 10));
  const monthName = hour => MONTHS[Number(hour.date.slice(5, 7)) - 1];
  const when = hour => `${hour.day} ${dateNumber(hour)} ${monthName(hour)} ${hour.hour}`;
  const isNight = hour => hour.hour >= '18:00' || hour.hour < '06:00';

  function drawChart(config, hours, options = {}) {
    const dayView = Boolean(options.dayView);
    const wrap = document.createElement('figure');
    wrap.className = 'wx-chart';
    const legend = config.series.map(item => `<span class="${item.cls}"><i></i>${esc(item.label)}</span>`).join('') +
      (config.bar ? `<span class="${config.bar.cls}"><i style="border-top-width:8px"></i>${esc(config.bar.label)}</span>` : '') +
      '<span class="wx-night-key"><i></i>Night (18:00 to 06:00)</span>';
    wrap.innerHTML = `<h3>${esc(config.title)}</h3><div class="wx-legend">${legend}</div><p class="wx-readout" aria-live="polite">&nbsp;</p><div class="wx-svg"></div><ul class="wx-extremes"></ul>`;
    const W = 360, H = 200, L = 38, R = 8, T = 10, B = 44, plotH = H - T - B, plotW = W - L - R;
    const keys = config.series.map(item => item.key);
    const values = hours.flatMap(hour => keys.map(key => hour[key])).filter(value => value !== null && value !== undefined);
    let low = config.min !== undefined ? config.min : Math.floor(Math.min(...values, config.freeze ? -1 : Infinity));
    let high = config.max !== undefined ? config.max : Math.ceil(Math.max(...values, config.bar && !values.length ? 1 : -Infinity));
    if (config.bar && config.bar.share === 1) { high = Math.max(2, Math.ceil(Math.max(...hours.map(hour => hour[config.bar.key] || 0)))); }
    if (config.freeze) { low = Math.min(low, -2); high = Math.max(high, 6); }
    if (high <= low) { high = low + 1; }
    // round the axis to tidy values (for example every 5 degrees) unless both ends were fixed
    let step = (high - low) / 4;
    if (config.min === undefined || config.max === undefined) {
      const magnitude = 10 ** Math.floor(Math.log10(step));
      step = [1, 2, 2.5, 5, 10].map(factor => factor * magnitude).find(candidate => candidate >= step) || step;
      low = Math.floor(low / step) * step;
      high = Math.ceil(high / step) * step;
    }
    const x = index => L + (index / Math.max(1, hours.length - 1)) * plotW;
    const y = value => T + (1 - (value - low) / (high - low)) * plotH;
    const base = H - B;
    const barMax = config.bar ? Math.max(config.bar.share === 1 ? high : 2, ...hours.map(hour => hour[config.bar.key] || 0)) : 1;
    const barW = Math.max(0.8, plotW / hours.length * 0.8);
    const parts = [`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(config.title)} for ${dayView ? esc(options.dayLabel || 'the day') : 'the next 14 days'}">`];

    // night bands first, so everything else draws over them
    let runStart = null;
    hours.forEach((hour, index) => {
      const night = isNight(hour);
      if (night && runStart === null) runStart = index;
      if ((!night || index === hours.length - 1) && runStart !== null) {
        const end = night ? index : index - 1;
        const edge = night ? end : Math.min(end + 1, hours.length - 1);  // run on to the first hour after the night
        parts.push(`<rect x="${x(runStart).toFixed(1)}" y="${T}" width="${Math.max(1, x(edge) - x(runStart)).toFixed(1)}" height="${plotH}" class="wx-night-band"/>`);
        runStart = null;
      }
    });
    if (config.freeze && low < 0) parts.push(`<rect x="${L}" y="${y(0).toFixed(1)}" width="${plotW}" height="${(base - y(0)).toFixed(1)}" class="wx-freeze"/>`);

    // value axis with five labelled lines
    for (let tick = low; tick <= high + 1e-9; tick += (high - low) / Math.max(1, Math.round((high - low) / step))) {
      parts.push(`<line x1="${L}" x2="${W - R}" y1="${y(tick).toFixed(1)}" y2="${y(tick).toFixed(1)}" class="wx-grid"/><text x="${L - 4}" y="${(y(tick) + 3).toFixed(1)}" text-anchor="end" class="wx-axis">${Math.round(tick * 10) / 10}${config.unit}</text>`);
    }
    if (dayView) {
      // one day: a line and a label every three hours
      hours.forEach((hour, index) => {
        if (hour.hour.endsWith(':00') && Number(hour.hour.slice(0, 2)) % 3 === 0) {
          parts.push(`<line x1="${x(index).toFixed(1)}" x2="${x(index).toFixed(1)}" y1="${T}" y2="${base}" class="wx-grid-day"/><text x="${x(index).toFixed(1)}" y="${base + 14}" text-anchor="middle" class="wx-axis wx-axis-day">${esc(hour.hour)}</text>`);
        }
      });
    } else {
    // time grid: midnight lines, lighter lines at 06:00, 12:00 and 18:00
    hours.forEach((hour, index) => {
      if (hour.hour === '00:00') parts.push(`<line x1="${x(index).toFixed(1)}" x2="${x(index).toFixed(1)}" y1="${T}" y2="${base}" class="wx-grid-day"/>`);
      else if (['06:00', '12:00', '18:00'].includes(hour.hour)) parts.push(`<line x1="${x(index).toFixed(1)}" x2="${x(index).toFixed(1)}" y1="${T}" y2="${base}" class="wx-grid-minor"/>`);
    });
    // day labels centred under each day: weekday, date number and, at the start or on the 1st, the month
    let groupStart = 0;
    let labelled = 0;
    hours.forEach((hour, index) => {
      const last = index === hours.length - 1;
      if (last || hours[index + 1].date !== hour.date) {
        if (index - groupStart >= 7) {
          const centre = ((x(groupStart) + x(index)) / 2).toFixed(1);
          const first = hours[groupStart];
          parts.push(`<text x="${centre}" y="${base + 12}" text-anchor="middle" class="wx-axis wx-axis-day">${esc(first.day)}</text><text x="${centre}" y="${base + 23}" text-anchor="middle" class="wx-axis wx-axis-date">${dateNumber(first)}</text>`);
          if (labelled === 0 || dateNumber(first) === 1) parts.push(`<text x="${centre}" y="${base + 33}" text-anchor="middle" class="wx-axis">${monthName(first)}</text>`);
          labelled += 1;
        }
        groupStart = index + 1;
      }
    });
    }
    if (config.bar) hours.forEach((hour, index) => {
      if (hour[config.bar.key] > 0) {
        const share = config.bar.share === 1 ? plotH : plotH * config.bar.share;
        const barH = Math.max(1.5, (hour[config.bar.key] / barMax) * share);
        parts.push(`<rect x="${(x(index) - barW / 2).toFixed(1)}" y="${(base - barH).toFixed(1)}" width="${barW.toFixed(1)}" height="${barH.toFixed(1)}" class="wx-bar-rain"/>`);
      }
    });
    if (config.freeze && low < 0) parts.push(`<line x1="${L}" x2="${W - R}" y1="${y(0).toFixed(1)}" y2="${y(0).toFixed(1)}" class="wx-zero"/>`);
    config.series.forEach(item => {
      let path = '';
      let pen = false;
      hours.forEach((hour, index) => {
        const value = hour[item.key];
        if (value === null || value === undefined) { pen = false; return; }
        path += `${pen ? 'L' : 'M'}${x(index).toFixed(1)},${y(value).toFixed(1)}`;
        pen = true;
      });
      parts.push(`<path d="${path}" class="wx-line ${item.cls}"/>`);
    });
    parts.push(`<line class="wx-cursor" x1="0" x2="0" y1="${T}" y2="${base}" visibility="hidden"/></svg>`);
    wrap.querySelector('.wx-svg').innerHTML = parts.join('');

    // lowest and highest of each line, with when
    const extremes = [];
    config.series.forEach(item => {
      const known = hours.filter(hour => hour[item.key] !== null && hour[item.key] !== undefined);
      if (!known.length) return;
      const lowest = known.reduce((a, b) => (b[item.key] < a[item.key] ? b : a));
      const highest = known.reduce((a, b) => (b[item.key] > a[item.key] ? b : a));
      extremes.push(`<li class="${item.cls}"><strong>${esc(item.label)}</strong> lowest ${num(lowest[item.key])}${config.unit} (${esc(when(lowest))}) · highest ${num(highest[item.key])}${config.unit} (${esc(when(highest))})</li>`);
    });
    if (config.bar) {
      const total = hours.reduce((sum, hour) => sum + (hour[config.bar.key] || 0), 0);
      const wettest = hours.reduce((a, b) => ((b[config.bar.key] || 0) > (a[config.bar.key] || 0) ? b : a));
      const wetHours = hours.filter(hour => (hour[config.bar.key] || 0) > 0).length;
      extremes.push(`<li class="${config.bar.cls}"><strong>${esc(config.bar.label)}</strong> ${num(total)}${config.bar.unit} in total over ${wetHours} wet hours${total > 0 ? ` · heaviest ${num(wettest[config.bar.key])}${config.bar.unit} (${esc(when(wettest))})` : ''}</li>`);
    }
    wrap.querySelector('.wx-extremes').innerHTML = extremes.join('');

    const svg = wrap.querySelector('svg');
    const cursor = svg.querySelector('.wx-cursor');
    const readout = wrap.querySelector('.wx-readout');
    const show = event => {
      const box = svg.getBoundingClientRect();
      const position = ((event.clientX - box.left) / box.width) * W;
      const index = Math.max(0, Math.min(hours.length - 1, Math.round(((position - L) / plotW) * (hours.length - 1))));
      const hour = hours[index];
      cursor.setAttribute('x1', x(index).toFixed(1));
      cursor.setAttribute('x2', x(index).toFixed(1));
      cursor.setAttribute('visibility', 'visible');
      const bits = config.series.map(item => `${item.label} ${hour[item.key] === null ? '–' : num(hour[item.key])}${config.unit}`);
      if (config.bar) bits.push(`${config.bar.label} ${num(hour[config.bar.key] || 0)}${config.bar.unit}`);
      readout.textContent = `${when(hour)} · ${bits.join(' · ')}`;
    };
    svg.addEventListener('pointermove', show);
    svg.addEventListener('pointerdown', show);
    return wrap;
  }

  // One bar per day showing the day's total rain, with the amount written above it.
  function drawDailyRain(config, days) {
    const wrap = document.createElement('figure');
    wrap.className = 'wx-chart';
    wrap.innerHTML = `<h3>${esc(config.title)}</h3><div class="wx-legend"><span class="s-rainkey"><i style="border-top-width:8px"></i>Total rain for the day</span></div><div class="wx-svg"></div><ul class="wx-extremes"></ul>`;
    const W = 360, H = 200, L = 38, R = 8, T = 18, B = 44, plotH = H - T - B, plotW = W - L - R;
    const totals = days.map(day => day.rain_mm || 0);
    let high = Math.max(2, ...totals);
    const magnitude = 10 ** Math.floor(Math.log10(high / 4));
    const step = [1, 2, 2.5, 5, 10].map(factor => factor * magnitude).find(candidate => candidate >= high / 4) || high / 4;
    high = Math.ceil(high / step) * step;
    const slot = plotW / days.length;
    const y = value => T + (1 - value / high) * plotH;
    const base = H - B;
    const parts = [`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Total rain for each of the next 14 days">`];
    for (let tick = 0; tick <= high + 1e-9; tick += step) {
      parts.push(`<line x1="${L}" x2="${W - R}" y1="${y(tick).toFixed(1)}" y2="${y(tick).toFixed(1)}" class="wx-grid"/><text x="${L - 4}" y="${(y(tick) + 3).toFixed(1)}" text-anchor="end" class="wx-axis">${Math.round(tick * 10) / 10}</text>`);
    }
    days.forEach((day, index) => {
      const centre = L + slot * index + slot / 2;
      const height = totals[index] > 0 ? Math.max(1.5, (totals[index] / high) * plotH) : 0;
      if (height) parts.push(`<rect x="${(centre - slot * 0.32).toFixed(1)}" y="${(base - height).toFixed(1)}" width="${(slot * 0.64).toFixed(1)}" height="${height.toFixed(1)}" rx="2" class="wx-bar-rain"/><text x="${centre.toFixed(1)}" y="${(base - height - 3).toFixed(1)}" text-anchor="middle" class="wx-axis wx-axis-date">${num(totals[index])}</text>`);
      parts.push(`<text x="${centre.toFixed(1)}" y="${base + 12}" text-anchor="middle" class="wx-axis wx-axis-day">${esc(day.short)}</text><text x="${centre.toFixed(1)}" y="${base + 23}" text-anchor="middle" class="wx-axis wx-axis-date">${Number(day.date.slice(8, 10))}</text>`);
      if (index === 0 || Number(day.date.slice(8, 10)) === 1) parts.push(`<text x="${centre.toFixed(1)}" y="${base + 33}" text-anchor="middle" class="wx-axis">${MONTHS[Number(day.date.slice(5, 7)) - 1]}</text>`);
    });
    parts.push('</svg>');
    wrap.querySelector('.wx-svg').innerHTML = parts.join('');
    const total = totals.reduce((sum, value) => sum + value, 0);
    const wettest = days.reduce((a, b) => ((b.rain_mm || 0) > (a.rain_mm || 0) ? b : a));
    const dry = totals.filter(value => value === 0).length;
    wrap.querySelector('.wx-extremes').innerHTML = `<li class="s-rainkey"><strong>Rain</strong> ${num(total)} mm over the 14 days · ${dry} dry ${dry === 1 ? 'day' : 'days'}${total > 0 ? ` · wettest ${esc(wettest.label)} with ${num(wettest.rain_mm)} mm` : ''}</li>`;
    return wrap;
  }

  // Each 14-day chart sits behind its own button and is drawn the first time it is opened.
  function renderCharts(hours) {
    const host = document.getElementById('wx-charts');
    const blocks = CHARTS.map(config => {
      const block = document.createElement('div');
      block.className = 'wx-chart-block';
      block.innerHTML = `<button type="button" class="wx-chart-toggle" aria-expanded="false"><span>${esc(config.title)}</span><span class="wx-chart-chevron" aria-hidden="true">＋</span></button><div class="wx-chart-body" hidden></div>`;
      const toggle = block.querySelector('.wx-chart-toggle');
      const body = block.querySelector('.wx-chart-body');
      block.setOpen = open => {
        if (open && !body.firstChild) body.append(config.daily ? drawDailyRain(config, forecast.days) : drawChart(config, hours));
        body.hidden = !open;
        toggle.setAttribute('aria-expanded', String(open));
        toggle.querySelector('.wx-chart-chevron').textContent = open ? '－' : '＋';
      };
      toggle.addEventListener('click', () => block.setOpen(body.hidden));
      return block;
    });
    const all = document.createElement('button');
    all.type = 'button';
    all.className = 'wx-chart-all';
    all.textContent = 'Show all charts';
    all.addEventListener('click', () => {
      const open = all.textContent === 'Show all charts';
      blocks.forEach(block => block.setOpen(open));
      all.textContent = open ? 'Hide all charts' : 'Show all charts';
    });
    host.replaceChildren(all, ...blocks);
  }

  // ---- days ---------------------------------------------------------------------------------------------
  function hourlyTable(hours) {
    const rows = hours.map(hour => {
      const cls = hour.level === 'grit' ? 'wx-row-grit' : hour.level === 'watch' ? 'wx-row-watch' : '';
      const cold = hour.surface !== null && hour.surface <= 0.5 ? ' class="wx-cold"' : '';
      return `<tr class="${cls}"><td>${esc(hour.hour)} ${esc(hour.icon)}</td><td>${deg(hour.temp)}</td><td>${deg(hour.feels)}</td><td${cold}>${deg(hour.surface)}</td>
        <td>${num(hour.rain)}</td><td>${num(hour.prob, 0)}%</td><td>${num(hour.wind)}</td><td>${num(hour.gust)}</td><td>${esc(hour.compass)}</td>
        <td>${num(hour.cloud, 0)}%</td><td>${num(hour.humidity, 0)}%</td><td>${num(hour.pressure, 0)}</td><td>${num(hour.visibility_km)}</td><td>${deg(hour.dew)}</td><td>${num(hour.uv)}</td></tr>`;
    }).join('');
    return `<div class="wx-table-wrap"><table class="wx-table"><thead><tr><th>Time</th><th>Air</th><th>Feels</th><th>Road</th><th>Rain mm</th><th>Chance</th><th>Wind mph</th><th>Gust</th><th>Dir</th><th>Cloud</th><th>Humid.</th><th>hPa</th><th>Vis km</th><th>Dew</th><th>UV</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  const DAY_CHARTS = [
    { title: 'Temperature and road surface (°C)', unit: '°', freeze: true, bar: { key: 'rain', label: 'Rain', unit: ' mm', cls: 's-rainkey', share: 0.3 },
      series: [{ key: 'temp', label: 'Air', cls: 's-air' }, { key: 'feels', label: 'Feels like', cls: 's-feels' }, { key: 'surface', label: 'Road surface', cls: 's-surface' }] },
    { title: 'Rain through the day (running total, mm)', unit: '', min: 0, series: [{ key: 'cum', label: 'Rain so far', cls: 's-prob' }] },
    { title: 'Chance of rain (%)', unit: '%', min: 0, max: 100, series: [{ key: 'prob', label: 'Chance of rain', cls: 's-prob' }] },
    { title: 'Wind (mph)', unit: '', min: 0, series: [{ key: 'wind', label: 'Wind', cls: 's-wind' }, { key: 'gust', label: 'Gusts', cls: 's-gust' }] },
    { title: 'Cloud cover and humidity (%)', unit: '%', min: 0, max: 100, series: [{ key: 'cloud', label: 'Cloud', cls: 's-cloud' }, { key: 'humidity', label: 'Humidity', cls: 's-humidity' }] },
  ];
  let selectedDate = null;

  function dayStats(day) {
    const night = day.night;
    const roadClass = night && night.level === 'grit' ? 'wx-grit' : night && night.level === 'watch' ? 'wx-watch' : '';
    // the whole night, 18:00 to 06:00 the next morning, because gritting is overnight
    const overnight = night
      ? stat('Lowest road surface overnight, 18:00 to 06:00', degC(night.min_surface), `at ${esc(night.min_surface_day)} ${esc(night.min_surface_time)} · air low ${degC(night.min_air)}${night.hours_covered < 12 ? ' · the forecast only covers part of this night' : ''}`, roadClass)
      : stat('Lowest road surface', degC(day.min_surface), day.min_surface_time ? `at ${esc(day.min_surface_time)}` : '');
    return `<div class="wx-stats">
      ${overnight}
      ${stat('Air high / low', `${deg(day.temp_max)} / ${deg(day.temp_min)}`, `high at ${esc(day.temp_max_time)} · low at ${esc(day.temp_min_time)}<br>feels ${deg(day.feels_max)} at ${esc(day.feels_max_time)} / ${deg(day.feels_min)} at ${esc(day.feels_min_time)}`)}
      ${stat('Rain, total for the day', mm(day.rain_mm), `${num(day.rain_hours, 0)} wet hours · up to ${num(day.prob_max, 0)}% chance`)}
      ${stat('Snow', `${num(day.snow_cm)} cm`)}
      ${stat('Wind', `${num(day.wind_max)} mph ${esc(day.wind_compass)}`, `gusts to ${num(day.gust_max)} mph`)}
      ${stat('Sunrise / sunset', `${esc(day.sunrise)} / ${esc(day.sunset)}`, `${num(day.daylight_hours)} h of daylight`)}
      ${stat('Sunshine', `${num(day.sunshine_hours)} h`)}
      ${stat('UV index (max)', num(day.uv_max))}
    </div>`;
  }

  function renderDayDetail(host, day, index, hours) {
    const dayHours = hours.slice(day.start ?? 0, day.end ?? 0);
    const night = day.night;
    const title = index === 0 ? `Today · ${day.label}` : index === 1 ? `Tomorrow · ${day.label}` : day.label;
    const chip = night ? `<span class="wx-chip wx-${night.level}">${esc(VERDICT[night.level])}</span>` : '';
    let nightHtml = '';
    if (night) {
      const window = night.grit_from ? ` Below the grit threshold ${esc(night.grit_from)} to ${esc(night.grit_until)}.` : '';
      nightHtml = `<p class="wx-advice">Night (18:00 to 06:00): ${esc(VERDICT[night.level])}. Road surface low ${degC(night.min_surface)} at ${esc(night.min_surface_day)} ${esc(night.min_surface_time)}.${window}${night.rain_advice ? ` ${esc(night.rain_advice)}` : ''}</p>`;
    }
    host.innerHTML = `<div class="wx-detail-head"><span class="wx-day-icon" aria-hidden="true">${esc(day.icon)}</span>
      <div><h3>${esc(title)}</h3><p class="hint">${esc(day.description)}${index === 0 && dayHours.length ? `, hours from ${esc(dayHours[0].hour)}` : ', 24 hours'}</p></div>${chip}</div>
      ${nightHtml}${dayStats(day)}<div class="wx-day-charts"></div><h4>Hour by hour</h4><div data-hourly></div>`;
    let running = 0;
    const withTotals = dayHours.map(hour => ({ ...hour, cum: Math.round((running += hour.rain || 0) * 10) / 10 }));
    host.querySelector('.wx-day-charts').replaceChildren(...DAY_CHARTS.map(config => drawChart(config, withTotals, { dayView: true, dayLabel: day.label })));
    host.querySelector('[data-hourly]').innerHTML = hourlyTable(dayHours);
  }

  // Open one day (and close the others); opening an open day closes it.
  function selectDay(date, scroll) {
    const days = forecast.days;
    document.querySelectorAll('.wx-dayrow').forEach(row => {
      const open = row.dataset.date === date && selectedDate !== date;
      row.classList.toggle('is-open', open);
      row.querySelector('.wx-row').setAttribute('aria-expanded', String(open));
      const detail = row.querySelector('.wx-row-detail');
      detail.hidden = !open;
      if (open) {
        const index = days.findIndex(day => day.date === date);
        renderDayDetail(detail, days[index], index, forecast.hours);
        if (scroll) row.scrollIntoView({ block: 'start', behavior: 'smooth' });
      } else {
        detail.replaceChildren();
      }
    });
    selectedDate = selectedDate === date ? null : date;
  }

  // The daily forecast: a single row per day with the road surface low and the gritting verdict for the night.
  function renderForecast(days) {
    const list = document.getElementById('wx-strip');
    list.innerHTML = days.map((day, index) => {
      const night = day.night;
      const level = night ? night.level : '';
      const name = index === 0 ? 'Today' : day.short;
      const road = night ? night.min_surface : day.min_surface;
      const roadTime = night ? `${night.min_surface_day} ${night.min_surface_time}` : day.min_surface_time;
      return `<div class="wx-dayrow ${level ? `wx-${level}` : ''}" data-date="${esc(day.date)}">
        <button type="button" class="wx-row" aria-expanded="false">
          <span class="wx-row-day"><strong>${esc(name)}</strong><small>${Number(day.date.slice(8, 10))} ${MONTHS[Number(day.date.slice(5, 7)) - 1]}</small></span>
          <span class="wx-row-icon" aria-hidden="true">${esc(day.icon)}</span>
          <span class="wx-row-temp"><span>▲ ${deg(day.temp_max)} <small>${esc(day.temp_max_time)}</small></span><span>▼ ${deg(day.temp_min)} <small>${esc(day.temp_min_time)}</small></span></span>
          <span class="wx-row-verdict">${night ? esc(VERDICT[level]) : ''}</span>
          <span class="wx-row-facts"><span>${day.rain_mm > 0 ? `💧 ${num(day.rain_mm)} mm` : 'Dry'}</span><span>💨 ${num(day.wind_max, 0)} mph ${esc(day.wind_compass)}</span><span>Road low ${degC(road)}${roadTime ? ` at ${esc(roadTime)}` : ''}</span></span>
        </button>
        <div class="wx-row-detail" hidden></div>
      </div>`;
    }).join('');
    list.querySelectorAll('.wx-dayrow').forEach(row => row.querySelector('.wx-row').addEventListener('click', () => selectDay(row.dataset.date, true)));
  }

  async function load() {
    try {
      const response = await fetch(window.WEATHER_URL, { cache: 'no-store' });
      const data = await response.json();
      if (!data.ok) throw new Error(data.message || 'The forecast could not be loaded.');
      statusEl.textContent = `Updated ${data.updated}${data.stale ? ' (the forecast service did not answer, so this may be a little out of date)' : ''}.`;
      renderNow(data.current, data.hours);
      const keep = selectedDate !== null || forecast === null ? selectedDate : null;
      forecast = data;
      renderForecast(data.days);
      selectedDate = null;
      selectDay(keep || data.days[0].date, false);
      renderCharts(data.hours);
      for (const id of ['wx-forecast-panel', 'wx-charts-panel']) document.getElementById(id).hidden = false;
    } catch (error) {
      statusEl.textContent = `${error.message} Check the internet connection and try again.`;
    }
  }

  load();
  window.setInterval(load, 15 * 60 * 1000);
})();
