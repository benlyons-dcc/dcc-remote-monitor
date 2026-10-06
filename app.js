// DCC Environmental Monitor, shared helpers.
// Load in <head> (not deferred) so the saved theme applies before the page paints.

(function () {
  // ── Theme ────────────────────────────────────────────
  // No saved choice → follow the device's light/dark setting.
  try {
    const saved = localStorage.getItem('dcc-theme');
    if (saved === 'dark' || saved === 'light') document.documentElement.dataset.theme = saved;
  } catch (e) { /* storage blocked: just follow the device */ }

  const media = window.matchMedia('(prefers-color-scheme: dark)');

  window.currentTheme = function () {
    return document.documentElement.dataset.theme || (media.matches ? 'dark' : 'light');
  };

  window.toggleTheme = function () {
    const next = currentTheme() === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem('dcc-theme', next); } catch (e) {}
    window.dispatchEvent(new Event('themechange'));
  };

  media.addEventListener('change', () => {
    if (!document.documentElement.dataset.theme) window.dispatchEvent(new Event('themechange'));
  });

  // Read a CSS variable, for chart colours
  window.cssVar = name => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

  // ── Timestamps ───────────────────────────────────────
  // Postgres "timestamp without time zone" comes back with no offset
  // (e.g. 2026-09-30T12:06:00). Browsers read that as local time, which
  // makes every reading look an hour old during BST. Treat it as UTC.
  window.parseTs = function (ts) {
    if (ts instanceof Date) return ts;
    if (typeof ts === 'string' && !/[zZ]$|[+-]\d\d:?\d\d$/.test(ts)) ts = ts.replace(' ', 'T') + 'Z';
    return new Date(ts);
  };

  // ── Sensor status ────────────────────────────────────
  // ACTIVE:  last reading within 3 of the sensor's usual reporting intervals
  //          (never less than 1 hour), so an hourly sensor isn't flagged early.
  // DELAYED: no reading for longer than that, but less than 24 hours.
  // OFFLINE: nothing for 24 hours or more.
  window.sensorStatus = function (timestamps, now = Date.now()) {
    const ms = timestamps.map(t => parseTs(t).getTime()).filter(n => !isNaN(n)).sort((a, b) => a - b);
    if (!ms.length) return { status: 'offline', label: 'OFFLINE', age: Infinity, interval: null, window: null };

    const recent = ms.slice(-21);
    const gaps = [];
    for (let i = 1; i < recent.length; i++) if (recent[i] > recent[i - 1]) gaps.push(recent[i] - recent[i - 1]);
    gaps.sort((a, b) => a - b);
    const interval = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 10 * 60000;
    const window_ = Math.max(3 * interval, 60 * 60000);
    const age = now - ms[ms.length - 1];

    const status = age <= window_ ? 'live' : age < 24 * 3600000 ? 'delayed' : 'offline';
    const label = { live: 'ACTIVE', delayed: 'DELAYED', offline: 'OFFLINE' }[status];
    return { status, label, age, interval, window: window_ };
  };

  window.fmtDuration = function (ms) {
    if (ms == null || !isFinite(ms)) return '—';
    const m = Math.max(0, Math.round(ms / 60000));
    if (m < 60) return `${m} min`;
    const h = Math.floor(m / 60), rm = m % 60;
    if (h < 48) return rm ? `${h} h ${rm} min` : `${h} h`;
    return `${Math.floor(h / 24)} days`;
  };
})();