// S2 Recovery Tools for Microsoft Excel — PWA
// Browser-side recovery for corrupt .xlsx / .xls files. No upload, no server.

(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const fmtBytes = (n) => {
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / 1024 / 1024).toFixed(2)} MB`;
  };

  const state = {
    file: null,
    bytes: null,
    format: null,        // 'xlsx' | 'xls' | 'unknown'
    workbook: null,
    repairedBlob: null,
    activeSheet: null,
  };

  // ---------- Logging ----------
  const log = $('log');
  function logEntry(level, msg) {
    const empty = log.querySelector('.empty');
    if (empty) empty.remove();
    const ts = new Date().toLocaleTimeString();
    const div = document.createElement('div');
    div.className = 'entry';
    div.innerHTML = `<span class="ts">${ts}</span><span class="${level}">${msg}</span>`;
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
  }
  const info = (m) => logEntry('info', m);
  const ok = (m) => logEntry('ok', m);
  const warn = (m) => logEntry('warn', m);
  const err = (m) => logEntry('err', m);

  // ---------- File intake ----------
  const drop = $('drop');
  const fileInput = $('file');

  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', (e) => { if (e.target.files[0]) handleFile(e.target.files[0]); });
  document.addEventListener('paste', (e) => {
    const f = e.clipboardData?.files?.[0];
    if (f) handleFile(f);
  });

  async function handleFile(file) {
    state.file = file;
    state.bytes = new Uint8Array(await file.arrayBuffer());
    state.workbook = null;
    state.repairedBlob = null;
    state.activeSheet = null;

    $('m-name').textContent = file.name;
    $('m-size').textContent = fmtBytes(file.size);
    $('m-status').textContent = 'Analyzing…';
    $('meta').classList.remove('hidden');
    $('pills').classList.remove('hidden');

    info(`Loaded "${file.name}" (${fmtBytes(file.size)})`);

    // First analysis step: shared S2 File Identifier. Reads magic numbers,
    // separates embedded/concatenated foreign files (offered for download in
    // the panel), and hands back the bytes this program should repair.
    let idReport = null;
    if (window.S2FileID) {
      try {
        idReport = S2FileID.analyze(state.bytes, { programKey: 'corruptexcelrec', fileName: file.name });
        S2FileID.renderPanel($('fileid'), idReport);
        info(`Identified: ${idReport.primary.description}`);
        if (idReport.mismatch) warn('This tool repairs .xls/.xlsx — the file appears to be a different type. See the recommendation above.');
        if (idReport.foreign.length) warn(`${idReport.foreign.length} embedded/concatenated file(s) of a different type separated — download them above before repairing the rest.`);
        if (idReport.multiple) info(`Split into ${idReport.segments.length} segments; repairing segment ${idReport.proceedSegment.index + 1} (${idReport.proceedSegment.ext}, ${fmtBytes(idReport.proceedSegment.length)}).`);
        state.bytes = idReport.proceedBytes; // every recovery strategy reads state.bytes
      } catch (e) {
        warn(`File identification failed (${e.message || e}) — continuing with raw bytes.`);
      }
    }

    state.format = detectFormat(state.bytes, file.name, idReport);
    $('m-format').textContent = state.format.toUpperCase();
    info(`Detected format: ${state.format}`);

    const pills = $('pills');
    pills.innerHTML = '';
    const diag = await diagnose(state.bytes, state.format);
    diag.forEach(d => {
      const span = document.createElement('span');
      span.className = `pill ${d.level}`;
      span.textContent = d.text;
      pills.appendChild(span);
    });

    $('m-status').textContent = diag.some(d => d.level === 'bad') ? 'Corrupt' :
      diag.some(d => d.level === 'warn') ? 'Suspicious' : 'Looks OK';

    renderMethods();
    $('section-recover').classList.remove('hidden');
  }

  function detectFormat(bytes, name, idReport) {
    // Prefer the S2 File Identifier verdict when it names a type we handle.
    if (idReport) {
      const ext = (idReport.proceedSegment || idReport.primary).ext;
      if (ext === 'xlsx' || ext === 'xls') return ext;
    }
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)) return 'xlsx';
    if (bytes.length >= 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) return 'xls';
    const lower = name.toLowerCase();
    if (lower.endsWith('.xlsx') || lower.endsWith('.xlsm') || lower.endsWith('.xltx') || lower.endsWith('.xltm')) return 'xlsx';
    if (lower.endsWith('.xls') || lower.endsWith('.xlsb')) return 'xls';
    return 'unknown';
  }

  async function diagnose(bytes, format) {
    const out = [];
    if (format === 'unknown') {
      out.push({ level: 'bad', text: 'Unknown format' });
      return out;
    }
    if (format === 'xlsx') {
      // Use the fault-tolerant unzipper for diagnosis — JSZip rejects truncated
      // central directories outright, which is exactly the failure we recover from.
      const { files, stats } = unzipImmortal(bytes);
      const names = Object.keys(files);
      if (!names.length) {
        out.push({ level: 'bad', text: 'ZIP container damaged' });
        return out;
      }
      out.push({ level: 'good', text: `${names.length} entries recovered` });
      if (stats.partial) out.push({ level: 'warn', text: `${stats.partial} partial` });
      if (!names.includes('[Content_Types].xml')) out.push({ level: 'bad', text: 'Missing [Content_Types].xml' });
      if (!names.some(n => n.startsWith('xl/worksheets/'))) out.push({ level: 'bad', text: 'No worksheets/' });
      if (!names.some(n => n === 'xl/workbook.xml')) out.push({ level: 'warn', text: 'Missing xl/workbook.xml' });
    } else if (format === 'xls') {
      out.push({ level: 'good', text: 'OLE2 container present' });
    }
    return out;
  }

  // ---------- Recovery methods ----------
  const methods = [
    { id: 'auto', name: 'Auto Repair', desc: 'Try every method, return the best result.', tag: 'recommended' },
    { id: 'strict', name: 'Strict Read', desc: 'Open with strict parsing. Confirms corruption.' },
    { id: 'xml', name: 'Lenient XML Repair', desc: 'Re-parses XML inside .xlsx with lax rules.' },
    { id: 'zip', name: 'ZIP Recovery', desc: 'Rebuilds the ZIP container if damaged.' },
    { id: 'salvage', name: 'Salvage Cells', desc: 'Pulls every readable cell into a fresh workbook.' },
    { id: 'csv', name: 'Convert to CSV', desc: 'Last resort: extract any readable cells to CSV.' },
  ];

  function renderMethods() {
    const wrap = $('methods');
    wrap.innerHTML = '';
    methods.forEach(m => {
      const btn = document.createElement('button');
      btn.className = 'method';
      btn.dataset.id = m.id;
      btn.innerHTML = `<h3>${m.name}${m.tag ? `<span class="tag">${m.tag}</span>` : ''}</h3><p>${m.desc}</p>`;
      btn.addEventListener('click', () => runMethod(m.id));
      wrap.appendChild(btn);
    });
  }

  async function runMethod(id) {
    if (!state.file) return;
    info(`▶ Running: ${methods.find(m => m.id === id).name}`);
    document.querySelectorAll('.method').forEach(b => b.disabled = true);
    try {
      let wb;
      switch (id) {
        case 'auto':    wb = await runAuto(); break;
        case 'strict':  wb = await runStrict(); break;
        case 'xml':     wb = await runLenientXML(); break;
        case 'zip':     wb = await runZipRecovery(); break;
        case 'salvage': wb = await runSalvage(); break;
        case 'csv':     wb = await runCsvExport(); return;
      }
      if (wb) {
        state.workbook = wb;
        renderResult(wb);
        ok(`Recovery succeeded — ${wb.SheetNames.length} sheet(s)`);
      }
    } catch (e) {
      err(`Failed: ${e.message || e}`);
    } finally {
      document.querySelectorAll('.method').forEach(b => b.disabled = false);
    }
  }

  async function runStrict() {
    return XLSX.read(state.bytes, { type: 'array', WTF: true });
  }

  async function runLenientXML() {
    if (state.format !== 'xlsx') {
      warn('Lenient XML only applies to .xlsx — falling back to standard read.');
      return XLSX.read(state.bytes, { type: 'array' });
    }
    // Decode with the Immortal unzipper so we don't choke on a damaged central
    // directory. Then heal each XML entry, then rebuild via JSZip (output
    // ZIP is well-formed).
    const { files, stats } = unzipImmortal(state.bytes, { onLog: (lvl, m) => lvl === 'warn' ? warn(m) : null });
    info(`Immortal unzip recovered ${stats.recovered} entry/entries (${stats.partial} partial)`);
    if (!Object.keys(files).length) throw new Error('No files recoverable from ZIP container');

    const dec = new TextDecoder('utf-8', { fatal: false });
    const enc = new TextEncoder();
    const z = new JSZip();
    let healed = 0;
    for (const [name, data] of Object.entries(files)) {
      if (name.endsWith('.xml') || name.endsWith('.rels')) {
        let txt = dec.decode(data);
        const before = txt;
        txt = txt
          .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
          .replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g, '&amp;');
        if (!/^<\?xml/.test(txt)) txt = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + txt;
        if (txt !== before) healed++;
        z.file(name, enc.encode(txt));
      } else {
        z.file(name, data);
      }
    }
    if (healed) ok(`Cleaned XML in ${healed} file(s)`);
    const out = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    state.repairedBlob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return XLSX.read(out, { type: 'array' });
  }

  async function runZipRecovery() {
    if (state.format !== 'xlsx') throw new Error('ZIP recovery only applies to .xlsx');
    // Pure Immortal Inflater path — works even when JSZip refuses to load the
    // file (truncated central directory, bad CRCs, damaged DEFLATE blocks).
    const { files, stats } = unzipImmortal(state.bytes, {
      onLog: (lvl, m) => lvl === 'warn' ? warn(m) : info(m),
    });
    info(`Scanned ZIP — ${stats.scanned} local headers, ${stats.recovered} recovered (${stats.partial} partial), ${stats.skipped} skipped`);
    if (!stats.recovered) throw new Error('No files could be recovered from the ZIP');

    const z = new JSZip();
    for (const [name, data] of Object.entries(files)) z.file(name, data);
    ok(`Rebuilt ZIP with ${stats.recovered} entry/entries`);
    const out = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    state.repairedBlob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return XLSX.read(out, { type: 'array' });
  }

  async function runSalvage() {
    let wb;
    try {
      wb = XLSX.read(state.bytes, { type: 'array', cellFormula: false, cellHTML: false, cellNF: false, cellStyles: false });
    } catch (e) {
      warn('Standard read failed during salvage; attempting lenient XML first');
      wb = await runLenientXML();
    }
    const fresh = XLSX.utils.book_new();
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false, defval: '' });
      const filtered = data.filter(row => Array.isArray(row) && row.some(c => c !== '' && c != null));
      const newWs = XLSX.utils.aoa_to_sheet(filtered);
      XLSX.utils.book_append_sheet(fresh, newWs, sheetName.slice(0, 31));
    }
    const out = XLSX.write(fresh, { type: 'array', bookType: 'xlsx' });
    state.repairedBlob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return fresh;
  }

  async function runCsvExport() {
    let wb;
    try { wb = XLSX.read(state.bytes, { type: 'array' }); }
    catch (e) { wb = await runSalvage(); }
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    triggerDownload(blob, baseName() + '.recovered.csv');
    ok(`Exported CSV (${fmtBytes(blob.size)})`);
    state.workbook = wb;
    renderResult(wb);
  }

  async function runAuto() {
    const order = ['strict', 'xml', 'zip', 'salvage'];
    for (const id of order) {
      try {
        info(`Auto: trying "${id}"…`);
        let wb;
        if (id === 'strict')   wb = await runStrict();
        if (id === 'xml')      wb = await runLenientXML();
        if (id === 'zip')      wb = await runZipRecovery();
        if (id === 'salvage')  wb = await runSalvage();
        if (wb && wb.SheetNames.length) {
          ok(`Auto: succeeded with "${id}"`);
          return wb;
        }
      } catch (e) {
        warn(`Auto: "${id}" failed (${e.message || e})`);
      }
    }
    throw new Error('All automatic strategies failed');
  }

  // ---------- Result rendering ----------
  function renderResult(wb) {
    const tabs = $('sheet-tabs');
    tabs.innerHTML = '';
    wb.SheetNames.forEach((name, i) => {
      const b = document.createElement('button');
      b.textContent = name;
      if (i === 0) b.classList.add('active');
      b.addEventListener('click', () => {
        tabs.querySelectorAll('button').forEach(x => x.classList.remove('active'));
        b.classList.add('active');
        renderSheet(wb.Sheets[name]);
      });
      tabs.appendChild(b);
    });
    state.activeSheet = wb.SheetNames[0];
    renderSheet(wb.Sheets[state.activeSheet]);
    renderDownloads();
    $('section-result').classList.remove('hidden');
    $('section-result').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderSheet(sheet) {
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    const preview = $('preview');
    preview.innerHTML = '';
    if (!data.length) { preview.innerHTML = '<p style="padding:20px; color:var(--text-dim);">Empty sheet.</p>'; return; }
    const max = Math.min(data.length, 200);
    const cols = Math.max(...data.slice(0, max).map(r => r.length));
    let html = '<table><thead><tr>';
    for (let c = 0; c < cols; c++) html += `<th>${XLSX.utils.encode_col(c)}</th>`;
    html += '</tr></thead><tbody>';
    for (let r = 0; r < max; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        const v = data[r][c];
        html += `<td>${v == null ? '' : String(v).replace(/[<>&]/g, ch => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[ch]))}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody></table>';
    preview.innerHTML = html;
    if (data.length > max) preview.insertAdjacentHTML('beforeend', `<p style="padding:8px; color:var(--text-dim); font-size:0.8rem; text-align:center;">Showing first ${max} of ${data.length} rows.</p>`);
  }

  function baseName() {
    const n = state.file?.name || 'workbook';
    return n.replace(/\.[^.]+$/, '');
  }

  function renderDownloads() {
    const row = $('downloads');
    row.innerHTML = '';
    const wb = state.workbook;
    if (!wb) return;
    const mkBtn = (label, primary, fn) => {
      const b = document.createElement('button');
      b.textContent = label;
      if (primary) b.className = 'primary';
      b.addEventListener('click', fn);
      row.appendChild(b);
    };
    mkBtn('⬇ Save as .xlsx', true, () => {
      const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      triggerDownload(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '.recovered.xlsx');
    });
    mkBtn('⬇ Save as .xls', false, () => {
      const out = XLSX.write(wb, { type: 'array', bookType: 'xls' });
      triggerDownload(new Blob([out], { type: 'application/vnd.ms-excel' }), baseName() + '.recovered.xls');
    });
    mkBtn('⬇ Save active sheet as .csv', false, () => {
      const csv = XLSX.utils.sheet_to_csv(wb.Sheets[state.activeSheet], { blankrows: false });
      triggerDownload(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `${baseName()}.${state.activeSheet}.csv`);
    });
    mkBtn('⬇ Save as .ods', false, () => {
      const out = XLSX.write(wb, { type: 'array', bookType: 'ods' });
      triggerDownload(new Blob([out], { type: 'application/vnd.oasis.opendocument.spreadsheet' }), baseName() + '.recovered.ods');
    });
  }

  function triggerDownload(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
    info(`💾 Downloaded ${name}`);
  }

  // ---------- PWA install + offline ----------
  let deferredPrompt = null;
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    $('btn-install').classList.remove('hidden');
  });
  $('btn-install').addEventListener('click', async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const choice = await deferredPrompt.userChoice;
    if (choice.outcome === 'accepted') ok('Installed!');
    deferredPrompt = null;
    $('btn-install').classList.add('hidden');
  });
  window.addEventListener('appinstalled', () => { $('btn-install').classList.add('hidden'); ok('App installed.'); });

  function updateOnline() {
    $('offline').classList.toggle('hidden', navigator.onLine);
  }
  window.addEventListener('online', updateOnline);
  window.addEventListener('offline', updateOnline);
  updateOnline();

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch(() => { /* PWA disabled, app still works */ });
    });
  }
})();
