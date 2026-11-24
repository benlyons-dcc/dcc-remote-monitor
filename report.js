// DCC Environmental Monitor: Word report for a job.
// Uses the job page's data and helpers (job, devices, deviceSettings, rowsOf, valueOf,
// TYPES, CODES, parseTs, targetOf, DRY_GPK, DRY_RH, SETTLE_HOURS, TYPE_LABEL, PALETTE).
// The Word file is built in the browser with the docx library; nothing leaves the page.

(function () {
  const DOCX_URL = 'https://cdn.jsdelivr.net/npm/docx@9.8.1/dist/index.iife.js';
  const JSZIP_URL = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
  const LOGO_URL = 'dcc-emblem.png';
  const BRAND = { navy: '0B4F8A', orange: 'C43C08', flame: 'E0701A', sky: '0A93D2', text: '1F2D3D', mid: '5E6E80', rule: 'C9D4E0', band: 'EEF3F8' };

  // A4, 2 cm margins → 9638 twentieths of a point of usable width
  const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN;
  const CHART_PX_W = 1600, CHART_PX_H = 600;

  // ── Small helpers ────────────────────────────────────────
  const pad = n => String(n).padStart(2, '0');
  const ukDate = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
  const ukDateTime = d => `${ukDate(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const isoDay = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const fmt = (v, dp) => (v == null || isNaN(v) ? '—' : Number(v).toFixed(dp));
  const sensorLabel = src => {
    const loc = deviceSettings[src]?.location;
    return loc ? `${src} ${loc}` : src;
  };
  const sensorType = src => (deviceSettings[src]?.type === 'ambient' ? 'ambient' : 'enclosed');

  let docxLoading = null;
  function loadDocx() {
    if (window.docx) return Promise.resolve(window.docx);
    if (!docxLoading) {
      docxLoading = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = DOCX_URL;
        s.onload = () => (window.docx ? resolve(window.docx) : reject(new Error('Word library did not load')));
        s.onerror = () => { docxLoading = null; reject(new Error('Could not load the Word library. Check your connection.')); };
        document.head.appendChild(s);
      });
    }
    return docxLoading;
  }

  function loadScript(url, globalName) {
    if (window[globalName]) return Promise.resolve(window[globalName]);
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = url;
      s.onload = () => (window[globalName] ? resolve(window[globalName]) : reject(new Error(globalName + ' did not load')));
      s.onerror = () => reject(new Error('Could not load ' + globalName));
      document.head.appendChild(s);
    });
  }

  // The docx library writes empty folder entries into the zip. Word desktop ignores them,
  // but files with them have failed to open when co-edited in OneDrive/SharePoint. Re-zip without them.
  async function stripFolderEntries(blob) {
    try {
      const JSZip = await loadScript(JSZIP_URL, 'JSZip');
      const src = await JSZip.loadAsync(blob);
      const out = new JSZip();
      const files = [];
      src.forEach((path, f) => { if (!f.dir) files.push(path); });
      for (const path of files) out.file(path, await src.file(path).async('uint8array'), { createFolders: false });
      return await out.generateAsync({ type: 'blob', compression: 'DEFLATE',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    } catch (e) {
      console.warn('Could not tidy the Word file, downloading as built', e);
      return blob;
    }
  }

  async function fetchBytes(url) {
    try {
      const r = await fetch(url);
      if (!r.ok) return null;
      return new Uint8Array(await r.arrayBuffer());
    } catch (e) { return null; }
  }

  // ── Modal ────────────────────────────────────────────────
  window.openReportModal = function () {
    if (!job) return;
    const sensors = devices.filter(d => rowsOf(d).length);
    if (!sensors.length) { alert('There are no readings on this job yet, so there is nothing to report.'); return; }

    // Period: the chart's current period
    const s = document.getElementById('startDate').value;
    const e = document.getElementById('endDate').value;
    const hours = parseInt(document.getElementById('rangeSelect').value) || 0;
    const allMs = sensors.flatMap(d => rowsOf(d).map(r => parseTs(r.recorded_at).getTime()));
    const firstDay = new Date(Math.min(...allMs));
    const lastDay = new Date(Math.max(...allMs));
    let from = s ? new Date(s + 'T00:00:00') : hours ? new Date(Date.now() - hours * 3600000) : firstDay;
    let to = e ? new Date(e + 'T00:00:00') : lastDay;
    if (from < firstDay) from = firstDay;
    document.getElementById('rptFrom').value = isoDay(from);
    document.getElementById('rptTo').value = isoDay(to);

    // Sensors: all with readings. Readings: what's ticked on the chart, else humidity.
    document.getElementById('rptSensors').innerHTML = sensors.map(d =>
      `<label class="rpt-check"><input type="checkbox" value="${esc(d)}" checked> ${esc(sensorLabel(d))}</label>`).join('');
    const ticked = selectedTypes();
    const chosen = ticked.length ? ticked : ['HUM'];
    document.getElementById('rptTypes').innerHTML = CODES.map(c =>
      `<label class="rpt-check"><input type="checkbox" value="${c}" ${chosen.includes(c) ? 'checked' : ''}> ${c} <span style="color:var(--text-dim)">${TYPES[c].name}</span></label>`).join('');

    const msg = document.getElementById('rptMsg'); msg.textContent = ''; msg.className = 'form-msg';
    document.getElementById('reportModal').classList.add('open');
  };

  window.closeReportModal = function (ev) {
    if (ev && ev.target !== document.getElementById('reportModal')) return;
    document.getElementById('reportModal').classList.remove('open');
  };

  window.generateReport = async function () {
    const msg = document.getElementById('rptMsg');
    const btn = document.getElementById('rptGo');
    const say = (t, k = '') => { msg.className = 'form-msg' + (k ? ' ' + k : ''); msg.textContent = t; };

    const fromV = document.getElementById('rptFrom').value;
    const toV = document.getElementById('rptTo').value;
    const sensors = [...document.querySelectorAll('#rptSensors input:checked')].map(i => i.value);
    const codes = [...document.querySelectorAll('#rptTypes input:checked')].map(i => i.value);
    const opts = {
      summary: document.getElementById('rptSummary').checked,
      status: document.getElementById('rptStatus').checked,
      daily: document.getElementById('rptDaily').checked,
      notes: document.getElementById('rptNotes').value.trim(),
    };
    if (!fromV || !toV) return say('Choose a start and end date.', 'error');
    const from = new Date(fromV + 'T00:00:00');
    const to = new Date(toV + 'T23:59:59');
    if (to < from) return say('The end date is before the start date.', 'error');
    if (!sensors.length) return say('Tick at least one sensor.', 'error');
    if (!codes.length) return say('Tick at least one reading type.', 'error');

    const inPeriod = r => { const t = parseTs(r.recorded_at); return t >= from && t <= to; };
    const data = Object.fromEntries(sensors.map(s => [s, rowsOf(s).filter(inPeriod)]));
    if (!sensors.some(s => data[s].length)) return say('No readings from those sensors in that period.', 'error');

    btn.disabled = true;
    try {
      say('Loading Word library…');
      const D = await loadDocx();
      say('Drawing charts…');
      const charts = [];
      for (const code of codes) charts.push({ code, png: await chartPng(code, sensors, data) });
      say('Building report…');
      const logo = await fetchBytes(LOGO_URL);
      const doc = buildDoc(D, { from, to, sensors, codes, data, charts, logo, opts });
      const blob = await stripFolderEntries(await D.Packer.toBlob(doc));
      const safe = (job.site_name || `Job ${job.id}`).replace(/[\\/:*?"<>|]+/g, '-').slice(0, 80).trim();
      download(blob, `Drying report - ${safe} - ${isoDay(new Date()).split('-').reverse().join('-')}.docx`);
      say('Report downloaded.', 'ok');
    } catch (e) {
      console.error(e);
      say('Couldn’t create the report: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  };

  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ── Chart image (always light, print-friendly) ──────────
  async function chartPng(code, sensors, data) {
    const t = TYPES[code];
    const series = sensors.map(s => ({ s, byTs: new Map(data[s].map(r => [r.recorded_at, r])) })).filter(x => x.byTs.size);
    const timestamps = [...new Set(series.flatMap(x => [...x.byTs.keys()]))].sort((a, b) => parseTs(a) - parseTs(b));
    const labels = timestamps.map(ts => { const d = parseTs(ts); return `${pad(d.getDate())}/${pad(d.getMonth() + 1)} ${pad(d.getHours())}:00`; });

    const holder = document.createElement('div');
    holder.style.cssText = `position:fixed;left:-10000px;top:0;width:${CHART_PX_W}px;height:${CHART_PX_H}px`;
    const canvas = document.createElement('canvas');
    canvas.width = CHART_PX_W; canvas.height = CHART_PX_H;
    holder.appendChild(canvas); document.body.appendChild(holder);

    const font = { family: 'Arial, sans-serif', size: 20 };
    const chart = new Chart(canvas, {
      type: 'line',
      data: {
        labels,
        datasets: series.map(x => ({
          label: sensorLabel(x.s),
          data: timestamps.map(ts => valueOf(x.byTs.get(ts), code, x.s)),
          borderColor: PALETTE[devices.indexOf(x.s) % PALETTE.length],
          backgroundColor: PALETTE[devices.indexOf(x.s) % PALETTE.length],
          borderWidth: 2.5, pointRadius: 0, tension: 0, spanGaps: true,
        })),
      },
      options: {
        responsive: false, animation: false, devicePixelRatio: 1,
        plugins: {
          legend: { position: 'bottom', labels: { color: '#1F2D3D', font, boxWidth: 28, boxHeight: 4, padding: 22 } },
          tooltip: { enabled: false },
        },
        scales: {
          x: { ticks: { color: '#1F2D3D', font, maxRotation: 0, autoSkip: true, maxTicksLimit: 8 }, grid: { color: 'rgba(0,0,0,0.08)' }, border: { color: '#1F2D3D' } },
          y: {
            ticks: { color: '#1F2D3D', font }, grid: { color: 'rgba(0,0,0,0.08)' }, border: { color: '#1F2D3D' },
            title: { display: true, text: `${code} (${t.unit})`, color: '#1F2D3D', font: { ...font, weight: 'bold' } },
          },
        },
      },
      plugins: [{ id: 'whiteBg', beforeDraw: c => { const x = c.ctx; x.save(); x.fillStyle = '#FFFFFF'; x.fillRect(0, 0, c.width, c.height); x.restore(); } }],
    });
    chart.update();
    const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
    chart.destroy(); holder.remove();
    return new Uint8Array(await blob.arrayBuffer());
  }

  // ── Calculations ────────────────────────────────────────
  function stats(rows, code, src) {
    const vals = rows.map(r => valueOf(r, code, src)).filter(v => v != null);
    if (!vals.length) return null;
    return {
      start: vals[0], end: vals[vals.length - 1],
      min: Math.min(...vals), max: Math.max(...vals),
      avg: vals.reduce((a, b) => a + b, 0) / vals.length,
    };
  }

  // Same rules as the dashboard badge, judged at the end of the report period
  function statusAt(src, rows) {
    const last = rows[rows.length - 1];
    if (!last) return { reading: '—', target: '—', status: 'No readings' };
    const rh = valueOf(last, 'HUM', src), gpk = valueOf(last, 'GPK', src);
    const when = ukDateTime(parseTs(last.recorded_at));
    if (sensorType(src) === 'ambient') {
      const good = gpk != null && rh != null && gpk <= DRY_GPK && rh <= DRY_RH;
      return {
        reading: `RH ${fmt(rh, 0)} %, GPK ${fmt(gpk, 2)} g/kg (${when})`,
        target: `GPK ≤ ${DRY_GPK} g/kg and RH ≤ ${DRY_RH} %`,
        status: good ? 'Dry air' : 'Damp air',
      };
    }
    const target = targetOf(src);
    const all = rowsOf(src);
    const firstMs = all.length ? parseTs(all[0].recorded_at).getTime() : 0;
    const settling = (parseTs(last.recorded_at).getTime() - firstMs) / 3600000 < SETTLE_HOURS;
    return {
      reading: `Pocket RH ${fmt(rh, 0)} % (${when})`,
      target: `RH ≤ ${target} %`,
      status: settling ? 'Settling' : rh != null && rh <= target ? 'Dry' : 'Wet',
    };
  }

  function dailyAverages(rows, code, src) {
    const days = new Map();
    for (const r of rows) {
      const v = valueOf(r, code, src);
      if (v == null) continue;
      const k = isoDay(parseTs(r.recorded_at));
      const d = days.get(k) || { sum: 0, n: 0 };
      d.sum += v; d.n++; days.set(k, d);
    }
    return new Map([...days].map(([k, d]) => [k, d.sum / d.n]));
  }

  // ── Word document ───────────────────────────────────────
  function buildDoc(D, ctx) {
    const { Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle,
      ShadingType, AlignmentType, Header, Footer, PageNumber, HeadingLevel, LevelFormat, TabStopType,
      VerticalAlign, PageBreak, TableLayoutType } = D;

    const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
    const rule = (color = BRAND.rule, size = 4) => ({ style: BorderStyle.SINGLE, size, color });

    const run = (text, o = {}) => new TextRun({ text: String(text), font: 'Arial', size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color || BRAND.text, allCaps: o.caps });
    const para = (children, o = {}) => new Paragraph({
      children: Array.isArray(children) ? children : [typeof children === 'string' ? run(children, o) : children],
      spacing: { before: o.before ?? 0, after: o.after ?? 80 }, alignment: o.align, keepNext: o.keepNext,
      numbering: o.bullet != null ? { reference: 'bullets', level: o.bullet } : undefined,
    });
    const heading = (text, level = 1) => new Paragraph({
      heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
      children: [new TextRun({ text })],
    });

    // Generic data table: header row in navy, banded rows
    function dataTable(headers, rows, widths, opts = {}) {
      const total = widths.reduce((a, b) => a + b, 0);
      const cell = (text, i, isHead, band) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        shading: isHead ? { fill: BRAND.navy, type: ShadingType.CLEAR, color: 'auto' } : band ? { fill: BRAND.band, type: ShadingType.CLEAR, color: 'auto' } : undefined,
        margins: { top: 60, bottom: 60, left: 90, right: 90 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: i >= (opts.numFrom ?? 99) ? AlignmentType.RIGHT : AlignmentType.LEFT,
          children: [run(text, { size: opts.size || 18, bold: isHead || (opts.boldFirst && i === 0), color: isHead ? 'FFFFFF' : BRAND.text })],
        })],
      });
      return new Table({
        width: { size: total, type: WidthType.DXA },
        columnWidths: widths,
        layout: TableLayoutType.FIXED,
        borders: { top: rule(), bottom: rule(), left: NONE, right: NONE, insideHorizontal: rule(), insideVertical: NONE },
        rows: [
          new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
          ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((v, i) => cell(v, i, false, ri % 2 === 1)) })),
        ],
      });
    }

    const { from, to, sensors, codes, data, charts, logo, opts } = ctx;
    const now = new Date();
    const days = Math.round((new Date(isoDay(to)) - new Date(isoDay(from))) / 86400000) + 1;

    // ── Header & footer ──
    const headerTable = new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [900, 5000, CONTENT_W - 5900],
      layout: TableLayoutType.FIXED,
      borders: { ...noBorders, bottom: rule(BRAND.navy, 8) },
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: 900, type: WidthType.DXA }, borders: { top: NONE, left: NONE, right: NONE, bottom: rule(BRAND.navy, 8) }, verticalAlign: VerticalAlign.CENTER,
          margins: { bottom: 80 },
          children: [new Paragraph({ children: logo ? [new ImageRun({ type: 'png', data: logo, transformation: { width: 39, height: 48 } })] : [run('')] })] }),
        new TableCell({ width: { size: 5000, type: WidthType.DXA }, borders: { top: NONE, left: NONE, right: NONE, bottom: rule(BRAND.navy, 8) }, verticalAlign: VerticalAlign.CENTER,
          margins: { bottom: 80 },
          children: [
            para([run('Disaster Care Capital', { size: 26, bold: true, color: BRAND.navy })], { after: 0 }),
            para([run('Fire & Flood Restoration', { size: 18, color: BRAND.orange })], { after: 0 }),
          ] }),
        new TableCell({ width: { size: CONTENT_W - 5900, type: WidthType.DXA }, borders: { top: NONE, left: NONE, right: NONE, bottom: rule(BRAND.navy, 8) }, verticalAlign: VerticalAlign.CENTER,
          margins: { bottom: 80 },
          children: [para([run('Remote drying monitoring', { size: 18, color: BRAND.mid })], { align: AlignmentType.RIGHT, after: 0 })] }),
      ] })],
    });

    const footer = new Footer({ children: [new Paragraph({
      border: { top: rule(BRAND.navy, 6) },
      spacing: { before: 120 },
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
      children: [
        run('Generated by DCC Environmental Monitor', { size: 16, color: BRAND.mid }),
        new TextRun({ children: ['\tPage ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], font: 'Arial', size: 16, color: BRAND.mid }),
      ],
    })] });

    // ── Details block ──
    const LABEL_W = 2300, VALUE_W = CONTENT_W - LABEL_W;
    const detailRow = (label, children, gapAfter) => new TableRow({ cantSplit: true, children: [
      new TableCell({ width: { size: LABEL_W, type: WidthType.DXA }, borders: { top: NONE, bottom: NONE, left: NONE, right: rule(BRAND.navy, 6) },
        margins: { top: 40, bottom: gapAfter ? 200 : 40, right: 160 },
        children: [para([run(label, { size: 18, bold: true, caps: true, color: BRAND.navy })], { after: 0 })] }),
      new TableCell({ width: { size: VALUE_W, type: WidthType.DXA }, borders: noBorders,
        margins: { top: 40, bottom: gapAfter ? 200 : 40, left: 200 },
        children }),
    ] });
    const v = text => [para(text, { after: 0 })];

    const startD = job.job_start_date ? new Date(String(job.job_start_date).slice(0, 10) + 'T00:00:00') : null;
    const endD = job.job_end_date ? new Date(String(job.job_end_date).slice(0, 10) + 'T00:00:00') : null;
    const equipment = [
      para(`${sensors.length} temperature and humidity sensor${sensors.length === 1 ? '' : 's'}:`, { bullet: 0, after: 20 }),
      ...sensors.map(s => para(`${sensorLabel(s)} (${TYPE_LABEL[sensorType(s)]})`, { bullet: 1, after: 20 })),
    ];
    const notes = opts.notes
      ? opts.notes.split(/\r?\n/).map(l => para(l || ' ', { after: 40 }))
      : v('—');

    const details = new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [LABEL_W, VALUE_W],
      layout: TableLayoutType.FIXED,
      borders: noBorders,
      rows: [
        detailRow('Company', v('Disaster Care Capital Ltd')),
        detailRow('Job', v(`${job.site_name || 'Unnamed'} (Job ${job.id})`)),
        detailRow('Postcode', v(job.postcode || '—')),
        detailRow('Started', v(startD ? ukDate(startD) : '—')),
        detailRow('Ended', v(endD ? ukDate(endD) : 'Ongoing'), true),
        detailRow('Report period', v(`${days} day${days === 1 ? '' : 's'} (${ukDate(from)} – ${ukDate(to)})`), true),
        detailRow('Equipment', equipment, true),
        detailRow('Notes', notes),
      ],
    });

    // ── Body ──
    const body = [
      new Paragraph({
        border: { bottom: rule(BRAND.navy, 8) },
        spacing: { before: 120, after: 240 },
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [
          new TextRun({ text: 'Drying report', font: 'Arial', size: 40, bold: true, color: BRAND.navy }),
          new TextRun({ text: `\tGenerated on ${ukDateTime(now)}`, font: 'Arial', size: 18, color: BRAND.sky }),
        ],
      }),
      details,
    ];

    // Charts
    body.push(new Paragraph({ children: [new PageBreak()] }), heading('Data charts'));
    for (const c of charts) {
      const t = TYPES[c.code];
      body.push(heading(`${t.name} (${c.code}, ${t.unit})`, 2));
      body.push(new Paragraph({
        spacing: { after: 240 },
        children: [new ImageRun({ type: 'png', data: c.png, transformation: { width: 640, height: Math.round(640 * CHART_PX_H / CHART_PX_W) } })],
      }));
    }

    // Summary
    if (opts.summary) {
      const rows = [];
      for (const s of sensors) for (const code of codes) {
        const st = stats(data[s], code, s);
        if (!st) continue;
        const dp = TYPES[code].dp;
        rows.push([sensorLabel(s), `${code} (${TYPES[code].unit})`, fmt(st.start, dp), fmt(st.end, dp), fmt(st.min, dp), fmt(st.avg, dp), fmt(st.max, dp)]);
      }
      if (rows.length) {
        body.push(heading('Summary for the period'));
        body.push(dataTable(['Sensor', 'Reading', 'Start', 'End', 'Min', 'Average', 'Max'], rows,
          [2338, 1700, 1120, 1120, 1120, 1120, 1120], { numFrom: 2, boldFirst: true }));
        body.push(para([run('Start and End are the first and last readings in the period.', { size: 16, italics: true, color: BRAND.mid })], { before: 80, after: 200 }));
      }
    }

    // Wet / dry status
    if (opts.status) {
      body.push(heading(`Wet / dry status at ${ukDate(to)}`));
      body.push(dataTable(['Sensor', 'Sensor type', 'Latest reading', 'Dry when', 'Status'],
        sensors.map(s => { const r = statusAt(s, data[s]); return [sensorLabel(s), TYPE_LABEL[sensorType(s)], r.reading, r.target, r.status]; }),
        [2000, 1700, 2738, 2000, 1200], { boldFirst: true }));
      body.push(para([run(
        'Enclosed sensors are sealed against a material, so the RH in the pocket reflects the material’s moisture. ' +
        'Room air sensors show whether the air is good for drying. Sealed sensors are marked Settling for their first ' +
        `${SETTLE_HOURS} hours. These readings describe conditions only: confirm materials are dry with moisture meter readings.`,
        { size: 16, italics: true, color: BRAND.mid })], { before: 80, after: 200 }));
    }

    // Daily averages appendix
    if (opts.daily) {
      body.push(new Paragraph({ children: [new PageBreak()] }), heading('Appendix: daily averages'));
      const dayKeys = [];
      for (let d = new Date(from); d <= to; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) dayKeys.push(isoDay(d));
      const withData = sensors.filter(s => data[s].length);
      const colW = Math.floor((CONTENT_W - 1500) / Math.max(1, withData.length));
      for (const code of codes) {
        const avgs = withData.map(s => dailyAverages(data[s], code, s));
        const rows = dayKeys
          .map(k => [k.split('-').reverse().join('/'), ...avgs.map(m => fmt(m.get(k), TYPES[code].dp))])
          .filter(r => r.slice(1).some(x => x !== '—'));
        if (!rows.length) continue;
        body.push(heading(`${TYPES[code].name} (${code}, ${TYPES[code].unit})`, 2));
        const widths = [1500, ...withData.map((_, i) => i === withData.length - 1 ? CONTENT_W - 1500 - colW * (withData.length - 1) : colW)];
        body.push(dataTable(['Date', ...withData.map(sensorLabel)], rows, widths, { numFrom: 1, size: withData.length > 5 ? 16 : 18 }));
        body.push(para('', { after: 200 }));
      }
    }

    return new Document({
      creator: 'DCC Environmental Monitor',
      title: `Drying report: ${job.site_name || 'Job ' + job.id}`,
      styles: {
        default: { document: { run: { font: 'Arial', size: 20, color: BRAND.text } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: 'Arial', size: 30, bold: true, color: BRAND.navy },
            paragraph: { spacing: { before: 240, after: 140 }, keepNext: true, keepLines: true, outlineLevel: 0 } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: 'Arial', size: 22, bold: true, color: BRAND.orange },
            paragraph: { spacing: { before: 160, after: 80 }, keepNext: true, keepLines: true, outlineLevel: 1 } },
        ],
      },
      numbering: { config: [{ reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 300, hanging: 240 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 660, hanging: 240 } } } },
      ] }] },
      sections: [{
        properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: 500, footer: 500 } } },
        headers: { default: new Header({ children: [headerTable, new Paragraph({ children: [], spacing: { after: 120 } })] }) },
        footers: { default: footer },
        children: body,
      }],
    });
  }
})();
