(() => {
  'use strict';
  const $ = (s) => document.querySelector(s);

  if (typeof XLSX === 'undefined' || typeof Chart === 'undefined') {
    const n = $('#notice');
    n.hidden = false;
    n.textContent = 'Bibliotekene for Excel og grafer ble ikke lastet. Siden trenger internettilgang første gang den åpnes.';
    $('#source').textContent = 'Mangler biblioteker';
    return;
  }

  // ---------- Formatering ----------
  const nf = new Intl.NumberFormat('nb-NO', { maximumFractionDigits: 2 });
  const nf0 = new Intl.NumberFormat('nb-NO', { maximumFractionDigits: 0 });
  const nfc = new Intl.NumberFormat('nb-NO', { notation: 'compact', maximumFractionDigits: 1 });
  const df = new Intl.DateTimeFormat('nb-NO', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const mf = new Intl.DateTimeFormat('nb-NO', { month: 'short', year: 'numeric' });
  const fmtNum = (v) => (v == null || !isFinite(v)) ? '–' : (Math.abs(v) >= 1000 ? nf0 : nf).format(v);
  const cellText = (v) => v == null ? '' : v instanceof Date ? df.format(v) : typeof v === 'number' ? nf.format(v) : String(v);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const isEmpty = (v) => v == null || (typeof v === 'string' && v.trim() === '');
  const AGG = { sum: 'Sum', avg: 'Gjennomsnitt', count: 'Antall rader', min: 'Minste', max: 'Største' };

  // ---------- Tolking av celler ----------
  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v !== 'string') return null;
    let s = v.trim().replace(/[\s ]/g, '').replace(/(kr|nok|%|,-)$/i, '').replace(/^(kr|nok)/i, '');
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
    else if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');
    return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : null;
  }
  function toDate(v) {
    if (v instanceof Date) {
      if (isNaN(v)) return null;
      // Excel-datoer kan komme ut noen minutter eller timer fra midnatt; rund av til nærmeste dag da.
      const mins = v.getHours() * 60 + v.getMinutes();
      if (mins < 120 || mins > 1320) {
        const t = new Date(v.getTime() + 12 * 3600e3);
        return new Date(t.getFullYear(), t.getMonth(), t.getDate());
      }
      return v;
    }
    if (typeof v !== 'string') return null;
    const s = v.trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
    if (m) return new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
    m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
    if (m) return new Date(+m[3], m[2] - 1, +m[1]);
    return null;
  }

  function readSheet(ws) {
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    // Overskriftsraden er den første raden (blant de 20 første) som er omtrent like bred som tabellen,
    // slik at en tittel eller merknad over tabellen hoppes over.
    const filled = (r) => (r ? r.reduce((a, v) => a + (isEmpty(v) ? 0 : 1), 0) : 0);
    let widest = 0;
    for (const r of aoa.slice(0, 200)) widest = Math.max(widest, filled(r));
    if (!widest) return { cols: [], rows: [] };
    let h = aoa.findIndex((r, i) => i < 20 && filled(r) >= Math.max(1, Math.ceil(widest * 0.6)));
    if (h < 0) h = aoa.findIndex((r) => filled(r) > 0);
    const body = aoa.slice(h + 1).filter((r) => r && r.some((v) => !isEmpty(v)));
    let width = aoa[h].length;
    for (const r of body) if (r.length > width) width = r.length;

    const seen = {};
    const names = [];
    for (let i = 0; i < width; i++) {
      let n = aoa[h][i];
      n = isEmpty(n) ? `Kolonne ${XLSX.utils.encode_col(i)}` : (n instanceof Date ? df.format(n) : String(n).trim());
      if (seen[n]) { seen[n]++; n = `${n} (${seen[n]})`; } else seen[n] = 1;
      names.push(n);
    }
    const cols = names.map((name, i) => {
      let nNum = 0, nDate = 0, n = 0;
      for (const r of body) {
        const v = r[i];
        if (isEmpty(v)) continue;
        n++;
        if (toDate(v)) nDate++;
        else if (toNum(v) !== null) nNum++;
      }
      const type = n === 0 ? 'text' : nDate / n >= 0.8 ? 'date' : nNum / n >= 0.8 ? 'number' : 'text';
      return { name, i, type, filled: n };
    }).filter((c) => c.filled > 0 || !c.name.startsWith('Kolonne '));

    const rows = body.map((r) => cols.map((c) => {
      const v = r[c.i];
      if (isEmpty(v)) return null;
      if (c.type === 'number') return toNum(v);
      if (c.type === 'date') return toDate(v);
      return v instanceof Date ? df.format(v) : String(v).trim();
    }));

    cols.forEach((c, j) => {
      const set = new Set();
      for (const r of rows) {
        const v = r[j];
        if (v != null) set.add(v instanceof Date ? v.getTime() : v);
        if (set.size > 1000) break;
      }
      c.unique = set.size;
      c.values = c.type !== 'date' && set.size <= 60
        ? [...set].sort((a, b) => typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), 'nb'))
        : null;
    });
    return { cols, rows };
  }

  // ---------- Tilstand ----------
  const S = {
    book: null, file: null, fileName: '', loadedAt: null, sample: true,
    sheet: null, parsed: {}, cols: [], rows: [],
    filters: {}, dateCol: null, dateFrom: null, dateTo: null, search: '',
    sort: { j: -1, dir: 1 }, limit: 100,
    main: {}, second: {}, charts: {},
  };

  function loadBook(wb, label, isSample, file) {
    S.book = wb;
    S.sample = isSample;
    S.fileName = label;
    S.file = file || null;
    S.loadedAt = new Date();
    S.parsed = {};
    const prev = S.sheet;
    renderTabs();
    const names = wb.SheetNames;
    const first = names.includes(prev) && !isSample ? prev : (names.find((n) => getSheet(n).rows.length) || names[0]);
    selectSheet(first);
    renderSource();
  }
  function getSheet(name) {
    if (!S.parsed[name]) S.parsed[name] = readSheet(S.book.Sheets[name]);
    return S.parsed[name];
  }

  function selectSheet(name) {
    S.sheet = name;
    const { cols, rows } = getSheet(name);
    S.cols = cols; S.rows = rows;
    S.filters = {}; S.dateFrom = null; S.dateTo = null; S.search = '';
    S.sort = { j: -1, dir: 1 }; S.limit = 100;
    S.dateCol = cols.findIndex((c) => c.type === 'date');
    if (S.dateCol < 0) S.dateCol = null;
    pickDefaults();
    renderTabs();
    renderFilters();
    renderAll();
  }

  function pickDefaults() {
    const cols = S.cols;
    const idx = (c) => cols.indexOf(c);
    const texts = cols.filter((c) => c.type === 'text' && c.unique >= 2 && c.unique <= 60);
    const dates = cols.filter((c) => c.type === 'date');
    const nums = cols.filter((c) => c.type === 'number' && c.unique > 1);
    const prefs = ['omsetning', 'beløp', 'belop', 'salg', 'inntekt', 'sum', 'total', 'verdi', 'budsjett', 'kostnad', 'pris', 'amount', 'revenue', 'antall'];
    let value = null;
    for (const p of prefs) { value = nums.find((c) => c.name.toLowerCase().includes(p)); if (value) break; }
    value = value || nums.find((c) => c.unique > 12) || nums[0] || null;
    const group = texts[0] || dates[0] || cols.find((c) => c.type === 'number' && c.unique <= 30 && c !== value) || null;
    S.main = {
      group: group ? idx(group) : null,
      gran: 'month',
      value: value ? idx(value) : null,
      agg: value ? 'sum' : 'count',
      type: 'auto',
    };
    // Sekundærgraf: utvikling over tid hvis hovedgrafen ikke allerede er tidsbasert, ellers fordeling.
    const groupIsDate = group && group.type === 'date';
    if (dates.length && !groupIsDate) {
      S.second = { group: idx(dates[0]), gran: 'month', type: 'line' };
    } else {
      const other = texts.find((c) => c !== group && c.unique <= 12) || texts.find((c) => c !== group);
      S.second = other ? { group: idx(other), gran: 'month', type: 'doughnut' } : { group: null };
    }
  }

  // ---------- Filtrering og gruppering ----------
  function filteredRows() {
    const fs = Object.entries(S.filters).filter(([, v]) => v !== '');
    const q = S.search.trim().toLowerCase();
    const dc = S.dateCol;
    return S.rows.filter((r) => {
      for (const [j, v] of fs) if (r[j] == null || String(r[j]) !== v) return false;
      if (dc != null && (S.dateFrom || S.dateTo)) {
        const d = r[dc];
        if (!d) return false;
        if (S.dateFrom && d < S.dateFrom) return false;
        if (S.dateTo && d > S.dateTo) return false;
      }
      if (q && !r.some((v) => v != null && cellText(v).toLowerCase().includes(q))) return false;
      return true;
    });
  }

  function keyOf(v, col, gran) {
    if (v == null) return null;
    if (col.type === 'date') {
      const y = v.getFullYear(), m = v.getMonth();
      if (gran === 'year') return { k: `${y}`, s: y * 1e4, l: String(y) };
      if (gran === 'quarter') { const q = Math.floor(m / 3) + 1; return { k: `${y}K${q}`, s: y * 1e4 + q, l: `K${q} ${y}` }; }
      if (gran === 'day') return { k: `${y}-${m}-${v.getDate()}`, s: y * 1e4 + (m + 1) * 100 + v.getDate(), l: df.format(new Date(y, m, v.getDate())) };
      return { k: `${y}-${m}`, s: y * 1e4 + (m + 1) * 100, l: mf.format(new Date(y, m, 1)) };
    }
    if (col.type === 'number') return { k: String(v), s: v, l: fmtNum(v) };
    return { k: v, s: 0, l: v };
  }

  function aggValue(b, agg) {
    if (agg === 'count') return b.cnt;
    if (!b.n) return null;
    if (agg === 'sum') return b.sum;
    if (agg === 'avg') return b.sum / b.n;
    if (agg === 'min') return b.min;
    return b.max;
  }

  function groupBy(rows, gj, gran, vj, agg, limit) {
    const col = S.cols[gj];
    const m = new Map();
    for (const r of rows) {
      const key = keyOf(r[gj], col, gran) || { k: '\u0000tom', s: Infinity, l: '(tom)' };
      let b = m.get(key.k);
      if (!b) { b = { l: key.l, s: key.s, sum: 0, n: 0, cnt: 0, min: Infinity, max: -Infinity }; m.set(key.k, b); }
      b.cnt++;
      if (vj != null) {
        const v = r[vj];
        if (v != null) { b.sum += v; b.n++; if (v < b.min) b.min = v; if (v > b.max) b.max = v; }
      }
    }
    let arr = [...m.values()];
    arr.forEach((b) => (b.v = aggValue(b, agg)));
    const ordered = col.type !== 'text';
    if (ordered) arr.sort((a, b) => a.s - b.s);
    else arr.sort((a, b) => (b.v ?? -Infinity) - (a.v ?? -Infinity));
    if (!ordered && limit && arr.length > limit) {
      const rest = arr.slice(limit - 1);
      const o = { l: `Øvrige (${rest.length})`, sum: 0, n: 0, cnt: 0, min: Infinity, max: -Infinity };
      for (const b of rest) { o.sum += b.sum; o.n += b.n; o.cnt += b.cnt; o.min = Math.min(o.min, b.min); o.max = Math.max(o.max, b.max); }
      o.v = aggValue(o, agg);
      arr = arr.slice(0, limit - 1).concat(o);
    }
    return { arr, ordered };
  }

  // ---------- Grafer ----------
  function tokens() {
    const cs = getComputedStyle(document.documentElement);
    const g = (n) => cs.getPropertyValue(n).trim();
    return {
      fg: g('--fg'), muted: g('--muted'), grid: g('--grid'), surface: g('--surface'), line: g('--line'),
      cat: ['--c1', '--c2', '--c3', '--c4', '--c5', '--c6', '--c7', '--c8'].map(g),
    };
  }
  function alpha(hex, a) {
    const h = hex.replace('#', '');
    const n = parseInt(h.length === 3 ? h.split('').map((c) => c + c).join('') : h, 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
  }
  const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;

  function drawChart(key, canvas, type, labels, values, seriesLabel) {
    if (S.charts[key]) S.charts[key].destroy();
    const t = tokens();
    const pie = type === 'doughnut';
    const horiz = type === 'hbar';
    const line = type === 'line';
    const box = canvas.parentElement;
    box.style.height = horiz ? `${Math.max(300, labels.length * 26 + 60)}px` : '';
    const ds = {
      label: seriesLabel,
      data: values,
      backgroundColor: pie ? labels.map((_, i) => t.cat[i % t.cat.length]) : line ? alpha(t.cat[0], 0.14) : t.cat[0],
      borderColor: pie ? t.surface : t.cat[0],
      borderWidth: pie ? 2 : line ? 2 : 0,
      borderRadius: pie || line ? 0 : 4,
      maxBarThickness: 44,
      fill: line,
      tension: 0.25,
      pointRadius: line ? (values.length > 40 ? 0 : 3) : 0,
      pointBackgroundColor: t.cat[0],
      spanGaps: true,
    };
    const valueAxis = {
      beginAtZero: true,
      grid: { color: t.grid }, border: { display: false },
      ticks: { color: t.muted, callback: (v) => nfc.format(v) },
    };
    const catAxis = {
      grid: { display: false }, border: { color: t.line },
      ticks: { color: t.muted, autoSkip: !horiz, maxRotation: 0, callback(v) { const l = this.getLabelForValue(v); return String(l).length > 22 ? String(l).slice(0, 21) + '…' : l; } },
    };
    S.charts[key] = new Chart(canvas, {
      type: pie ? 'doughnut' : line ? 'line' : 'bar',
      data: { labels, datasets: [ds] },
      options: {
        responsive: true, maintainAspectRatio: false,
        animation: reduced ? false : { duration: 300 },
        indexAxis: horiz ? 'y' : 'x',
        cutout: pie ? '58%' : undefined,
        plugins: {
          legend: { display: pie, position: 'bottom', labels: { color: t.muted, boxWidth: 10, boxHeight: 10, padding: 12 } },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const v = pie ? ctx.parsed : horiz ? ctx.parsed.x : ctx.parsed.y;
                if (pie) {
                  const tot = ctx.dataset.data.reduce((a, b) => a + (b || 0), 0);
                  return ` ${fmtNum(v)} (${tot ? nf.format((v / tot) * 100) : 0} %)`;
                }
                return ` ${seriesLabel}: ${fmtNum(v)}`;
              },
            },
          },
        },
        scales: pie ? {} : horiz ? { x: valueAxis, y: catAxis } : { x: catAxis, y: valueAxis },
      },
    });
  }

  function seriesName(cfg) {
    return cfg.agg === 'count' || cfg.value == null ? 'Antall rader' : `${AGG[cfg.agg]} av ${S.cols[cfg.value].name}`;
  }

  function renderMain(rows) {
    const cfg = S.main;
    renderMainControls();
    const box = $('#main-chart').parentElement;
    if (cfg.group == null || !rows.length) {
      if (S.charts.main) { S.charts.main.destroy(); S.charts.main = null; }
      $('#main-title').textContent = 'Hovedgraf';
      $('#main-sub').textContent = rows.length ? 'Fant ingen kolonne å gruppere etter.' : 'Ingen rader passer med filtrene.';
      return;
    }
    const gcol = S.cols[cfg.group];
    let type = cfg.type;
    const peek = groupBy(rows, cfg.group, cfg.gran, cfg.value, cfg.agg, 0);
    if (type === 'auto') type = gcol.type === 'date' ? 'line' : peek.arr.length > 8 ? 'hbar' : 'bar';
    const limit = type === 'doughnut' ? 8 : type === 'hbar' ? 25 : 40;
    const { arr } = groupBy(rows, cfg.group, cfg.gran, cfg.value, cfg.agg, gcol.type === 'text' ? limit : 0);
    const name = seriesName(cfg);
    $('#main-title').textContent = `${name} per ${gcol.name.toLowerCase()}`;
    $('#main-sub').textContent = `${arr.length} grupper · ${nf0.format(rows.length)} rader`;
    box.hidden = false;
    drawChart('main', $('#main-chart'), type, arr.map((b) => b.l), arr.map((b) => b.v), name);
  }

  function renderSecond(rows) {
    const cfg = S.second;
    const panel = $('#second-panel');
    if (cfg.group == null || !rows.length) {
      panel.hidden = true;
      if (S.charts.second) { S.charts.second.destroy(); S.charts.second = null; }
      return;
    }
    panel.hidden = false;
    const gcol = S.cols[cfg.group];
    const base = { value: S.main.value, agg: S.main.agg };
    const name = seriesName(base);
    const { arr } = groupBy(rows, cfg.group, cfg.gran, base.value, base.agg, gcol.type === 'text' ? 8 : 0);
    const type = gcol.type === 'date' ? 'line' : cfg.type;
    $('#second-title').textContent = gcol.type === 'date' ? `Utvikling over tid` : `Fordeling på ${gcol.name.toLowerCase()}`;
    $('#second-sub').textContent = `${name} · ${gcol.name}`;
    renderSecondControls();
    drawChart('second', $('#second-chart'), type, arr.map((b) => b.l), arr.map((b) => b.v), name);
  }

  // ---------- Kontroller ----------
  function selectHtml(id, label, options, current) {
    return `<div class="field"><label for="${id}">${esc(label)}</label><select id="${id}">${
      options.map((o) => `<option value="${esc(o.v)}"${String(o.v) === String(current) ? ' selected' : ''}>${esc(o.l)}</option>`).join('')
    }</select></div>`;
  }
  const granOpts = [{ v: 'day', l: 'Dag' }, { v: 'month', l: 'Måned' }, { v: 'quarter', l: 'Kvartal' }, { v: 'year', l: 'År' }];

  function renderMainControls() {
    const cols = S.cols;
    const groupOpts = cols.map((c, j) => ({ c, j }))
      .filter(({ c }) => c.type === 'text' || c.type === 'date' || (c.type === 'number' && c.unique <= 40))
      .map(({ c, j }) => ({ v: j, l: c.name }));
    const valueOpts = [{ v: '', l: 'Ingen (tell rader)' }].concat(
      cols.map((c, j) => ({ c, j })).filter(({ c }) => c.type === 'number').map(({ c, j }) => ({ v: j, l: c.name })));
    const cfg = S.main;
    const isDate = cfg.group != null && cols[cfg.group].type === 'date';
    let html = '';
    if (groupOpts.length) html += selectHtml('m-group', 'Grupper etter', groupOpts, cfg.group);
    if (isDate) html += selectHtml('m-gran', 'Periode', granOpts, cfg.gran);
    html += selectHtml('m-value', 'Verdi', valueOpts, cfg.value ?? '');
    if (cfg.value != null) html += selectHtml('m-agg', 'Beregning', ['sum', 'avg', 'count', 'min', 'max'].map((v) => ({ v, l: AGG[v] })), cfg.agg);
    html += selectHtml('m-type', 'Graf', [
      { v: 'auto', l: 'Automatisk' }, { v: 'bar', l: 'Stolper' }, { v: 'hbar', l: 'Liggende stolper' },
      { v: 'line', l: 'Linje' }, { v: 'doughnut', l: 'Smultring' },
    ], cfg.type);
    $('#main-controls').innerHTML = html;
  }

  function renderSecondControls() {
    const cfg = S.second;
    const opts = S.cols.map((c, j) => ({ c, j }))
      .filter(({ c }) => c.type === 'date' || (c.type === 'text' && c.unique >= 2 && c.unique <= 200))
      .map(({ c, j }) => ({ v: j, l: c.name }));
    let html = selectHtml('s-group', 'Vis etter', opts, cfg.group);
    if (S.cols[cfg.group].type === 'date') html += selectHtml('s-gran', 'Periode', granOpts, cfg.gran);
    else html += selectHtml('s-type', 'Graf', [{ v: 'doughnut', l: 'Smultring' }, { v: 'hbar', l: 'Liggende stolper' }, { v: 'bar', l: 'Stolper' }], cfg.type);
    $('#second-controls').innerHTML = html;
  }

  document.addEventListener('change', (e) => {
    const id = e.target.id, v = e.target.value;
    const num = (x) => (x === '' ? null : +x);
    if (id === 'm-group') { S.main.group = num(v); if (S.cols[S.main.group].type === 'date' && S.main.type === 'doughnut') S.main.type = 'auto'; }
    else if (id === 'm-gran') S.main.gran = v;
    else if (id === 'm-value') { S.main.value = num(v); if (S.main.value == null) S.main.agg = 'count'; else if (S.main.agg === 'count') S.main.agg = 'sum'; }
    else if (id === 'm-agg') S.main.agg = v;
    else if (id === 'm-type') S.main.type = v;
    else if (id === 's-group') { S.second.group = num(v); if (S.second.type === 'line') S.second.type = 'doughnut'; }
    else if (id === 's-gran') S.second.gran = v;
    else if (id === 's-type') S.second.type = v;
    else if (id.startsWith('f-')) { S.filters[+id.slice(2)] = v; e.target.classList.toggle('active', v !== ''); S.limit = 100; }
    else if (id === 'd-from') S.dateFrom = v ? new Date(v + 'T00:00:00') : null;
    else if (id === 'd-to') S.dateTo = v ? new Date(v + 'T23:59:59') : null;
    else return;
    renderAll();
  });

  // ---------- Filtre ----------
  function renderFilters() {
    const el = $('#filters');
    const parts = [];
    S.cols.forEach((c, j) => {
      if (!c.values || c.unique < 2 || c.unique > 60) return;
      if (c.type === 'number' && (c.unique > 12 || !c.values.every(Number.isInteger))) return;
      parts.push(selectHtml(`f-${j}`, c.name, [{ v: '', l: 'Alle' }].concat(c.values.map((x) => ({ v: String(x), l: cellText(x) }))), ''));
    });
    if (S.dateCol != null) {
      const ds = S.rows.map((r) => r[S.dateCol]).filter(Boolean);
      const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      let mn = null, mx = null;
      for (const d of ds) { if (!mn || d < mn) mn = d; if (!mx || d > mx) mx = d; }
      const name = esc(S.cols[S.dateCol].name);
      const lim = mn ? ` min="${iso(mn)}" max="${iso(mx)}"` : '';
      parts.push(`<div class="field"><label for="d-from">${name} fra</label><input type="date" id="d-from"${lim}></div>`);
      parts.push(`<div class="field"><label for="d-to">${name} til</label><input type="date" id="d-to"${lim}></div>`);
    }
    parts.push(`<div class="field search"><label for="search">Søk i rader</label><input type="search" id="search" placeholder="Fritekst …"></div>`);
    parts.push(`<button class="btn link" id="reset" type="button">Nullstill filtre</button>`);
    el.innerHTML = parts.join('');
    $('#search').addEventListener('input', (e) => { S.search = e.target.value; S.limit = 100; renderAll(); });
    $('#reset').addEventListener('click', () => { renderFilters(); S.filters = {}; S.dateFrom = S.dateTo = null; S.search = ''; renderAll(); });
  }

  // ---------- Nøkkeltall ----------
  function renderKpis(rows) {
    const tiles = [];
    const total = S.rows.length;
    tiles.push({ l: 'Rader', v: nf0.format(rows.length), s: rows.length === total ? `i arket «${S.sheet}»` : `av ${nf0.format(total)} totalt etter filter` });
    const vj = S.main.value;
    if (vj != null) {
      let sum = 0, n = 0, mn = Infinity, mx = -Infinity;
      for (const r of rows) { const v = r[vj]; if (v != null) { sum += v; n++; if (v < mn) mn = v; if (v > mx) mx = v; } }
      const nm = S.cols[vj].name;
      tiles.push({ l: `Sum ${nm}`, v: fmtNum(n ? sum : null), s: `${nf0.format(n)} verdier` });
      tiles.push({ l: `Snitt ${nm}`, v: fmtNum(n ? sum / n : null), s: n ? `spenn ${fmtNum(mn)} til ${fmtNum(mx)}` : 'ingen verdier' });
    }
    const gj = S.main.group;
    if (gj != null && S.cols[gj].type === 'text' && rows.length) {
      const { arr } = groupBy(rows, gj, 'month', vj, S.main.agg, 0);
      const top = arr[0];
      if (top) tiles.push({ l: `Størst ${S.cols[gj].name.toLowerCase()}`, v: top.l, s: `${seriesName(S.main)}: ${fmtNum(top.v)}` });
    }
    if (S.dateCol != null && rows.length) {
      let mn = null, mx = null;
      for (const r of rows) { const d = r[S.dateCol]; if (d) { if (!mn || d < mn) mn = d; if (!mx || d > mx) mx = d; } }
      if (mn) tiles.push({ l: 'Periode', v: `${mf.format(mn)} – ${mf.format(mx)}`, s: `${df.format(mn)} til ${df.format(mx)}` });
    }
    $('#kpis').innerHTML = tiles.map((t) =>
      `<div class="kpi"><span class="k-label">${esc(t.l)}</span><span class="k-value">${esc(t.v)}</span><span class="k-sub">${esc(t.s)}</span></div>`).join('');
  }

  // ---------- Tabell ----------
  const TYPE_LABEL = { number: 'tall', date: 'dato', text: 'tekst' };
  function renderTable(rows) {
    const cols = S.cols;
    $('#table-title').textContent = `Data fra «${S.sheet}»`;
    $('#columns').innerHTML = cols.map((c) => `<span class="pill">${esc(c.name)}<b>${TYPE_LABEL[c.type]}</b></span>`).join('');
    let data = rows;
    if (S.sort.j >= 0) {
      const j = S.sort.j, dir = S.sort.dir;
      data = rows.slice().sort((a, b) => {
        const x = a[j], y = b[j];
        if (x == null) return 1; if (y == null) return -1;
        if (x instanceof Date || typeof x === 'number') return (x - y) * dir;
        return String(x).localeCompare(String(y), 'nb', { numeric: true }) * dir;
      });
    }
    const thead = `<tr>${cols.map((c, j) => {
      const num = c.type === 'number' ? ' class="num"' : '';
      const arrow = S.sort.j === j ? (S.sort.dir > 0 ? ' ▲' : ' ▼') : '';
      const aria = S.sort.j === j ? ` aria-sort="${S.sort.dir > 0 ? 'ascending' : 'descending'}"` : '';
      return `<th${num}${aria}><button type="button" data-sort="${j}">${esc(c.name)}${arrow}</button></th>`;
    }).join('')}</tr>`;
    const shown = data.slice(0, S.limit);
    const tbody = shown.length
      ? shown.map((r) => `<tr>${r.map((v, j) => `<td class="${cols[j].type === 'number' ? 'num' : ''}${v == null ? ' empty' : ''}">${v == null ? '' : esc(cellText(v))}</td>`).join('')}</tr>`).join('')
      : `<tr><td colspan="${cols.length || 1}" class="empty-state">Ingen rader passer med filtrene.</td></tr>`;
    $('#table thead').innerHTML = thead;
    $('#table tbody').innerHTML = tbody;
    $('#table-count').textContent = `Viser ${nf0.format(shown.length)} av ${nf0.format(data.length)} rader`;
    $('#more').hidden = data.length <= S.limit;
  }
  $('#table').addEventListener('click', (e) => {
    const b = e.target.closest('[data-sort]');
    if (!b) return;
    const j = +b.dataset.sort;
    S.sort = S.sort.j === j ? (S.sort.dir > 0 ? { j, dir: -1 } : { j: -1, dir: 1 }) : { j, dir: 1 };
    renderTable(filteredRows());
  });
  $('#more').addEventListener('click', () => { S.limit += 100; renderTable(filteredRows()); });

  // ---------- Oversikt ----------
  function renderTabs() {
    if (!S.book) return;
    const names = S.book.SheetNames;
    const el = $('#tabs');
    el.hidden = names.length < 2;
    el.innerHTML = names.map((n) => {
      const cnt = S.parsed[n] ? S.parsed[n].rows.length : null;
      return `<button class="tab" type="button" role="tab" data-sheet="${esc(n)}" aria-selected="${n === S.sheet}">${esc(n)}${cnt != null ? `<span class="count">${nf0.format(cnt)}</span>` : ''}</button>`;
    }).join('');
  }
  $('#tabs').addEventListener('click', (e) => {
    const b = e.target.closest('[data-sheet]');
    if (b && b.dataset.sheet !== S.sheet) selectSheet(b.dataset.sheet);
  });

  function renderSource() {
    const n = $('#notice');
    const t = S.loadedAt ? S.loadedAt.toLocaleTimeString('nb-NO', { hour: '2-digit', minute: '2-digit' }) : '';
    const sheets = S.book.SheetNames.length;
    if (S.sample) {
      $('#source').textContent = 'Eksempeldata';
      n.hidden = false;
      n.className = 'notice';
      n.innerHTML = '<span><strong>Du ser eksempeldata.</strong> Velg en Excel-fil, eller slipp den hvor som helst på siden. Første rad med innhold i hvert ark brukes som kolonneoverskrifter.</span>';
    } else {
      $('#source').textContent = `${S.fileName} · ${sheets} ${sheets === 1 ? 'ark' : 'ark'} · lest kl. ${t}`;
      n.hidden = true;
    }
    $('#reload').hidden = S.sample || !S.file;
  }

  function renderAll() {
    if (!S.cols.length) {
      $('#kpis').innerHTML = '';
      $('#charts').hidden = true;
      $('#table thead').innerHTML = '';
      $('#table tbody').innerHTML = `<tr><td class="empty-state">Arket «${esc(S.sheet)}» er tomt.</td></tr>`;
      $('#table-count').textContent = '';
      $('#columns').innerHTML = '';
      return;
    }
    $('#charts').hidden = false;
    const rows = filteredRows();
    renderKpis(rows);
    renderMain(rows);
    renderSecond(rows);
    renderTable(rows);
  }

  // ---------- Filinnlesing ----------
  function showError(msg) {
    const n = $('#notice');
    n.hidden = false;
    n.className = 'notice error';
    n.textContent = msg;
  }
  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xls|xlsb|ods|csv)$/i.test(file.name)) {
      showError(`«${file.name}» ser ikke ut som et regneark. Velg en .xlsx-, .xls-, .ods- eller .csv-fil.`);
      return;
    }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      if (!wb.SheetNames.length) throw new Error('ingen ark');
      loadBook(wb, file.name, false, file);
    } catch (err) {
      showError(`Kunne ikke lese «${file.name}». Sjekk at filen ikke er passordbeskyttet eller skadet, og prøv igjen.`);
      console.error(err);
    }
  }
  $('#pick').addEventListener('click', () => $('#file').click());
  $('#file').addEventListener('change', (e) => { handleFile(e.target.files[0]); e.target.value = ''; });
  $('#reload').addEventListener('click', () => {
    // Nettleseren holder på filreferansen, så en lagret endring i Excel leses inn på nytt her.
    if (S.file) handleFile(S.file).catch(() => showError('Filen kunne ikke leses på nytt. Velg den igjen.'));
  });

  let dragDepth = 0;
  const hasFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
  window.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; $('#drop').hidden = false; });
  window.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) $('#drop').hidden = true; });
  window.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dragDepth = 0; $('#drop').hidden = true;
    handleFile(e.dataTransfer.files[0]);
  });

  // ---------- Eksempeldata ----------
  function mulberry32(a) {
    return () => { a |= 0; a = (a + 0x6D2B79F5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  }
  function sampleBook() {
    const rnd = mulberry32(2025);
    const pick = (items, w) => { let x = rnd() * w.reduce((a, b) => a + b, 0); for (let i = 0; i < items.length; i++) { x -= w[i]; if (x <= 0) return items[i]; } return items[items.length - 1]; };
    const regions = ['Oslo', 'Bergen', 'Trondheim', 'Stavanger', 'Tromsø'];
    const rw = [1.7, 1.15, 0.9, 0.8, 0.45];
    const products = [['Kontorstol', 2490], ['Skrivebord', 4990], ['Skjerm 27"', 3290], ['Hodetelefoner', 1290], ['Tastatur', 890], ['Dokkingstasjon', 1890]];
    const pw = [1.2, 0.7, 1, 1.3, 1.5, 0.8];
    const channels = ['Nettbutikk', 'Butikk', 'Bedriftssalg'];
    const cw = [1.4, 1, 0.6];
    const season = [0.85, 0.8, 0.9, 0.95, 1, 1.05, 0.65, 0.95, 1.15, 1.1, 1.3, 1.55];
    const rows = [];
    for (let i = 0; i < 900; i++) {
      let m;
      do { m = Math.floor(rnd() * 12); } while (rnd() > season[m] / 1.55);
      const d = new Date(2025, m, 1 + Math.floor(rnd() * 28));
      const [prod, price] = pick(products, pw);
      const ch = pick(channels, cw);
      const qty = 1 + Math.floor(rnd() * rnd() * (ch === 'Bedriftssalg' ? 24 : 5));
      const oms = Math.round(qty * price * (0.85 + rnd() * 0.2));
      const cost = Math.round(oms * (0.52 + rnd() * 0.16));
      rows.push([d, pick(regions, rw), prod, ch, qty, oms, cost]);
    }
    rows.sort((a, b) => a[0] - b[0]);
    const salg = [['Dato', 'Region', 'Produkt', 'Kanal', 'Antall', 'Omsetning', 'Kostnad']].concat(rows);
    const budsjett = [['Måned', 'Region', 'Budsjett']];
    for (let m = 0; m < 12; m++) regions.forEach((r, i) => budsjett.push([new Date(2025, m, 1), r, Math.round(rw[i] * season[m] * 118000 / 1000) * 1000]));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(salg, { cellDates: true, dateNF: 'yyyy-mm-dd' }), 'Salg');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(budsjett, { cellDates: true, dateNF: 'mmm yyyy' }), 'Budsjett');
    return wb;
  }
  $('#download-sample').addEventListener('click', () => {
    try { XLSX.writeFile(sampleBook(), 'eksempel-dashbord.xlsx', { cellDates: true }); }
    catch (e) { showError('Nettleseren tillot ikke nedlasting her. Åpne HTML-filen direkte fra disken og prøv igjen.'); }
  });

  // Tegn grafene på nytt når fargetema endres.
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', renderAll);
  new MutationObserver(renderAll).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

  loadBook(sampleBook(), 'Eksempeldata', true);
})();
