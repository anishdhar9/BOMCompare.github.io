/* app.js — UI wiring for BOM Compare. All parsing/comparison logic lives in
 * js/compare.js and js/parsers/; this file only handles files, state and DOM. */
(function () {
  'use strict';

  const BC = window.BOMCompare;
  const $ = function (id) { return document.getElementById(id); };

  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
  }

  // "Mark reviewed" on "Parts needing attention" — persisted so a part you've
  // already looked into stays out of the way across reloads, not just for the
  // current session. Keyed by part number only (not per-check): reviewing a
  // part is "I looked at this part", independent of which check flagged it.
  const REVIEWED_LS_KEY = 'bomcompare.reviewed.v1';
  function loadReviewed() {
    try {
      const arr = JSON.parse(localStorage.getItem(REVIEWED_LS_KEY) || '[]');
      return new Set(Array.isArray(arr) ? arr : []);
    } catch (e) { return new Set(); }
  }

  // Per-part priority tag — the user's own triage flag, independent of the
  // fixed per-check severity ranking in js/findings.js (which still decides
  // which check "owns" a part and the sort order below; priority never
  // touches that). Same load/save/get/set shape as "mark reviewed" above,
  // persisted under its own key so it survives a "Clear files & start over".
  const PRIORITY_LS_KEY = 'bomcompare.priority.v1';
  const PRIORITY_LEVELS = ['low', 'medium', 'high'];
  const PRIORITY_LABELS = { low: 'Low', medium: 'Medium', high: 'High' };
  function loadPriority() {
    try {
      const obj = JSON.parse(localStorage.getItem(PRIORITY_LS_KEY) || '{}');
      const map = new Map();
      if (obj && typeof obj === 'object') {
        for (const pn of Object.keys(obj)) { if (PRIORITY_LEVELS.indexOf(obj[pn]) !== -1) map.set(pn, obj[pn]); }
      }
      return map;
    } catch (e) { return new Map(); }
  }

  // Which of the 11 checks to run this session — a display/aggregation
  // preference, same category as column visibility and priority tags (not
  // reset by "Clear files & start over"). Ids and labels are the exact
  // sheet names buildCompareWorkbook() already uses, so the selector, every
  // section heading, and the export all agree on the same 11 names.
  // Nothing about computation is gated by this — every check still runs;
  // only what's aggregated into the dashboard/findings registry/panels/
  // tabs/export is affected. See isCheckEnabled() below.
  const CHECKS_LIST = [
    { key: 'missing', label: 'Missing from Item Master' },
    { key: 'reference', label: 'Reference items' },
    { key: 'qty', label: 'Quantity mismatches' },
    { key: 'qtyCascade', label: 'Quantity cascades' },
    { key: 'imOnly', label: 'In Item Master only' },
    { key: 'imQc', label: 'Item Master QC' },
    { key: 'material', label: 'Material vs CAD' },
    { key: 'revision', label: 'Revision vs CAD' },
    { key: 'titleDesc', label: 'Description vs CAD' },
    { key: 'virtual', label: 'Virtual parts' },
    { key: 'lldbo', label: 'LLDBO check' },
  ];
  const CHECKS_LS_KEY = 'bomcompare.enabledChecks.v1';
  function loadEnabledChecks() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(CHECKS_LS_KEY)) || {}; } catch (e) { /* ignore */ }
    const prefs = {};
    for (const c of CHECKS_LIST) prefs[c.key] = saved[c.key] !== undefined ? !!saved[c.key] : true;
    return prefs;
  }

  const state = {
    cadSources: [],     // parsed CAD results (max 2: full structure + intended BOM)
    im: null,           // parsed Item Master result (+ fileName)
    imQc: null,         // js/imqc.js runChecks() result for state.im
    result: null,
    activeTab: 'missing',
    filter: '',
    mappingCtx: null,   // { aoa, indents, source, analysis, fileName }
    folderHandle: null, // FileSystemDirectoryHandle from "Load from folder", if used
    lldbo: null,         // parsed LLDBO result (+ fileName), optional
    lldboResult: null,   // js/lldbo-compare.js compareLldbo() result — null until an LLDBO file is loaded
    lldboCandidatesResult: null, // js/lldbo-compare.js detectLldboCandidates() result — set as soon as the Item Master loads
    ignoreList: null,    // parsed js/parsers/ignorelist.js result (+ fileName), optional
    materialResult: null, // js/material-compare.js compareMaterial() result
    revisionResult: null, // js/revision-compare.js compareRevision() result
    titleDescResult: null, // js/titledesc-compare.js compareTitleDescription() result
    virtualResult: null,   // js/virtual-parts.js detectVirtualParts() result
    findings: null,       // js/findings.js buildRegistry() — cross-check registry
    showGrouped: false,   // "parts needing attention": reveal grouped children
    ignoreBoughtOutRevision: false, // "Revision" section: hide bought-out (X-999-*) parts
    sectionOpen: new Map(),  // qc-section id -> user override, survives re-renders
    traceOpen: new Map(),    // part number -> provenance-detail toggle open?, survives re-renders
    reviewedSet: loadReviewed(), // part numbers checked off "Parts needing attention", survives reloads (localStorage)
    priorityMap: loadPriority(), // part number -> user-assigned priority ('low'|'medium'|'high'), survives reloads (localStorage)
    enabledChecks: loadEnabledChecks(), // check key -> included in dashboard/registry/panels/tabs/export?, survives reloads (localStorage)
    attentionVisibleCount: 0,    // unreviewed, non-grouped part count — set by renderAttention(), read by the dashboard tile
  };

  function saveEnabledChecks() {
    try { localStorage.setItem(CHECKS_LS_KEY, JSON.stringify(state.enabledChecks)); } catch (e) { /* ignore */ }
  }
  function isCheckEnabled(key) { return state.enabledChecks[key] !== false; }

  function saveReviewed() {
    try { localStorage.setItem(REVIEWED_LS_KEY, JSON.stringify(Array.from(state.reviewedSet))); } catch (e) { /* ignore */ }
  }
  function isReviewed(pn) { return state.reviewedSet.has(BC.normNumber(pn)); }
  // val: true marks reviewed (hides from "Parts needing attention"), false
  // brings it back. Re-renders the dashboard so both the attention list and
  // the "Already reviewed" section pick up the change immediately.
  function setReviewed(pn, val) {
    const key = BC.normNumber(pn);
    if (!key) return;
    if (val) state.reviewedSet.add(key); else state.reviewedSet.delete(key);
    saveReviewed();
    renderDashboard();
  }

  function savePriority() {
    try {
      const obj = {};
      state.priorityMap.forEach(function (val, pn) { obj[pn] = val; });
      localStorage.setItem(PRIORITY_LS_KEY, JSON.stringify(obj));
    } catch (e) { /* ignore */ }
  }
  function getPriority(pn) { return state.priorityMap.get(BC.normNumber(pn)) || 'none'; }
  // val: 'none'/'low'/'medium'/'high'. Purely a user-facing triage tag — does
  // not affect check severity, sort order, dashboard tiles, or exports.
  function setPriority(pn, val) {
    const key = BC.normNumber(pn);
    if (!key) return;
    if (val === 'none') state.priorityMap.delete(key); else state.priorityMap.set(key, val);
    savePriority();
    renderDashboard();
  }

  // Two visual variants, both matching a (pn) -> element signature so
  // swapping which one renders (see priorityControlEl below) is a one-line
  // change. Variant A extends the app's existing severity-color badge
  // vocabulary (.issue-badge's muted/amber/red tiers); deliberately no green
  // for "low" — green means "pass" everywhere else in this app.
  function priorityControlElBadge(pn) {
    const val = getPriority(pn);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'priority-btn' + (val !== 'none' ? ' ' + val : '');
    btn.textContent = val === 'none' ? '— priority' : PRIORITY_LABELS[val];
    btn.title = 'Click to cycle priority (None → Low → Medium → High)';
    btn.addEventListener('click', function (ev) {
      ev.stopPropagation();
      const order = ['none'].concat(PRIORITY_LEVELS);
      setPriority(pn, order[(order.indexOf(val) + 1) % order.length]);
    });
    return btn;
  }

  // Variant B: a plain <select>, same declared style as the mapping panel's
  // dropdowns, sized for table-row density.
  function priorityControlElSelect(pn) {
    const sel = document.createElement('select');
    sel.className = 'priority-select';
    const opt0 = document.createElement('option');
    opt0.value = 'none'; opt0.textContent = '— priority';
    sel.appendChild(opt0);
    for (const lvl of PRIORITY_LEVELS) {
      const opt = document.createElement('option');
      opt.value = lvl; opt.textContent = PRIORITY_LABELS[lvl];
      sel.appendChild(opt);
    }
    sel.value = getPriority(pn);
    sel.addEventListener('click', function (ev) { ev.stopPropagation(); });
    sel.addEventListener('change', function () { setPriority(pn, sel.value); });
    return sel;
  }

  // Ships Variant A (cycling badge). Swap to priorityControlElSelect for the
  // dropdown variant — every call site uses this one seam.
  const priorityControlEl = priorityControlElBadge;

  // Base name suffix from the Item Master's SPN/PN project key, e.g.
  // '_SPN016823_PN22426'. Empty string when no key was found.
  function projectKeySuffix() {
    const key = state.im && state.im.projectKey;
    if (!key) return '';
    const parts = [key.spn, key.pn].filter(Boolean);
    return parts.length ? '_' + parts.join('_') : '';
  }

  // The two roles a CAD file can play; a new upload replaces the previous
  // file of the same role, so at most one of each is kept.
  function cadRole(parsed) {
    return parsed.source === 'leveled-sheet' ? 'bom' : 'structure';
  }

  function addCadSource(parsed) {
    const role = cadRole(parsed);
    const idx = state.cadSources.findIndex(function (s) { return cadRole(s) === role; });
    if (idx >= 0) state.cadSources.splice(idx, 1, parsed);
    else state.cadSources.push(parsed);
    runMaterialCheck();
    runRevisionCheck();
    runTitleDescCheck();
  }

  /* ---------------- column visibility ---------------- */

  const COLUMNS = [
    { key: 'title', label: 'Title', def: true },
    { key: 'description', label: 'Description', def: true },
    { key: 'type', label: 'Type (assembly/part)', def: true },
    { key: 'qty', label: 'Quantity', def: true },
    { key: 'file', label: 'File', def: false },
    { key: 'material', label: 'Material', def: false },
    { key: 'row', label: 'Source row', def: false },
  ];
  const LS_KEY = 'bomcompare.columns.v1';

  function loadColumnPrefs() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(LS_KEY)) || {}; } catch (e) { /* ignore */ }
    const prefs = {};
    for (const c of COLUMNS) prefs[c.key] = saved[c.key] !== undefined ? !!saved[c.key] : c.def;
    return prefs;
  }
  let colPrefs = loadColumnPrefs();

  function saveColumnPrefs() {
    try { localStorage.setItem(LS_KEY, JSON.stringify(colPrefs)); } catch (e) { /* ignore */ }
  }

  function renderColumnsMenu() {
    const menu = $('columns-menu');
    menu.innerHTML = '';
    for (const c of COLUMNS) {
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = colPrefs[c.key];
      cb.addEventListener('change', function () {
        colPrefs[c.key] = cb.checked;
        saveColumnPrefs();
        renderResults();
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(' ' + c.label));
      menu.appendChild(label);
    }
  }

  /* ---------------- checks to run ---------------- */

  function renderChecksMenu() {
    const menu = $('checks-menu');
    menu.innerHTML = '';
    for (const c of CHECKS_LIST) {
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = isCheckEnabled(c.key);
      cb.addEventListener('change', function () {
        state.enabledChecks[c.key] = cb.checked;
        saveEnabledChecks();
        rerenderChecks();
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(' ' + c.label));
      menu.appendChild(label);
    }
  }

  // Re-renders everything whose content depends on state.enabledChecks —
  // called live when the "Checks to run" dropdown changes. Nothing is
  // recomputed (every check still ran); each function below already
  // no-ops safely when its underlying state hasn't loaded yet, same as
  // calling them during a normal render pass.
  function rerenderChecks() {
    renderImQc();
    renderMaterialPanel();
    renderRevisionPanel();
    renderTitleDescPanel();
    renderLldboPanel();
    if (state.result) renderResults(); else renderDashboard();
  }

  /* ---------------- upload handling ---------------- */

  function chipEls(box, chips) {
    (chips || []).forEach(function (ch) {
      const el = document.createElement('span');
      el.className = 'chip' + (ch.kind ? ' ' + ch.kind : '');
      el.textContent = ch.text;
      if (ch.title) el.title = ch.title;
      box.appendChild(el);
    });
  }

  // Item Master zone: single file
  function setImStatus(fileName, chips, error) {
    $('status-file-im').textContent = fileName || '';
    const chipBox = $('status-chips-im');
    chipBox.innerHTML = '';
    chipEls(chipBox, chips);
    $('status-error-im').textContent = error || '';
    const zone = $('zone-im');
    zone.classList.toggle('parsed', !error && !!fileName && !!state.im);
    zone.classList.toggle('error', !!error);
  }

  // CAD zone: one row per loaded source, with a remove button
  function renderCadSources(progressText, error) {
    const box = $('cad-sources');
    box.innerHTML = '';
    state.cadSources.forEach(function (src, i) {
      const row = document.createElement('div');
      row.className = 'dz-source';
      const name = document.createElement('span');
      name.className = 'dz-source-name';
      name.textContent = src.fileName || '(file)';
      row.appendChild(name);
      const chips = document.createElement('span');
      chips.className = 'dz-chips';
      chipEls(chips, chipsFor(src));
      row.appendChild(chips);
      const rm = document.createElement('button');
      rm.type = 'button';
      rm.className = 'dz-remove';
      rm.title = 'Remove this file';
      rm.textContent = '✕';
      rm.addEventListener('click', function (ev) {
        ev.stopPropagation();
        state.cadSources.splice(i, 1);
        renderCadSources('', '');
        updateCompareButton();
        runMaterialCheck();
        runRevisionCheck();
        runTitleDescCheck();
      });
      row.appendChild(rm);
      box.appendChild(row);
    });
    if (progressText) {
      const p = document.createElement('div');
      p.className = 'dz-source';
      p.textContent = progressText;
      box.appendChild(p);
    }
    $('status-error-cad').textContent = error || '';
    const zone = $('zone-cad');
    zone.classList.toggle('parsed', !error && state.cadSources.length > 0);
    zone.classList.toggle('error', !!error);
  }

  function chipsFor(parsed) {
    const chips = [];
    if (parsed.kind === 'itemmaster') {
      chips.push({ text: 'Item Master', kind: 'good' });
      chips.push({ text: parsed.rows.length + ' rows' });
      chips.push({ text: parsed.hasPaths ? 'hierarchy ✓' : 'no hierarchy', kind: parsed.hasPaths ? '' : 'warn' });
      if (parsed.hasRevision) chips.push({ text: 'revision ✓', kind: 'good' });
    } else {
      let srcLabel = { 'flat-xlsx': 'Vault flat export', 'pdf': 'Vault PDF (full structure)', 'leveled-sheet': 'Leveled table' }[parsed.source] || parsed.source;
      if (parsed.source === 'leveled-sheet' && parsed.hasStructure) srcLabel = 'Inventor BOM export';
      chips.push({ text: srcLabel, kind: 'good' });
      chips.push({ text: parsed.items.length + ' components' });
      chips.push({ text: parsed.hasQty ? 'quantities ✓' : 'no quantities', kind: parsed.hasQty ? '' : 'warn' });
      chips.push({ text: parsed.hasLevels ? 'levels ✓' : 'no levels', kind: parsed.hasLevels ? '' : 'warn' });
      if (parsed.hasMaterial) chips.push({ text: 'material ✓', kind: 'good' });
      if (parsed.hasRevision) chips.push({ text: 'revision ✓', kind: 'good' });
    }
    if (parsed.warnings && parsed.warnings.length) {
      chips.push({ text: parsed.warnings.length + ' note(s)', kind: 'warn', title: parsed.warnings.join('\n') });
    }
    return chips;
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      const r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(new Error('Could not read the file.')); };
      r.readAsArrayBuffer(file);
    });
  }

  function firstSheetAoa(wb, XLSX) {
    if (!wb.SheetNames.length) return null;
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: null });
  }

  async function handleFiles(role, files) {
    hideMapping();
    for (const file of files) {
      try {
        const buf = await readFileAsArrayBuffer(file);
        if (/\.pdf$/i.test(file.name)) {
          if (role === 'im') throw new Error('PDF is only supported for the CAD BOM. Use the Excel export for the Item Master.');
          await handleCadPdf(file, buf);
        } else {
          const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
          if (role === 'im') handleImWorkbook(file, wb);
          else handleCadWorkbook(file, wb);
        }
      } catch (e) {
        if (role === 'cad') renderCadSources('', file.name + ': ' + (e.message || String(e)));
        else { state.im = null; state.imQc = null; hideImQc(); setImStatus(file.name, [], e.message || String(e)); runMaterialCheck(); runRevisionCheck(); runTitleDescCheck(); runLldboCheck(); }
      }
    }
    updateCompareButton();
  }

  function handleImWorkbook(file, wb) {
    const im = BC.detect.parseItemMasterFromWorkbook(wb, XLSX);
    if (!im) {
      const asCad = BC.detect.parseCadFromWorkbook(wb, XLSX);
      if (asCad && asCad.ok) {
        throw new Error('This looks like a CAD BOM export — drop it on the CAD BOM box on the left.');
      }
      const aoa = firstSheetAoa(wb, XLSX);
      if (aoa && BC.detect.looksLikeLldbo(aoa)) {
        throw new Error('This looks like an LLDBO (long-lead parts) file — drop it on the Long-Lead Parts box below instead.');
      }
      throw new Error('No "Number" header row found. Expected the Vault/ERP Item Master BOM export.');
    }
    im.fileName = file.name;
    state.im = im;
    setImStatus(file.name, chipsFor(im), '');
    state.imQc = BC.imQc.runChecks(im);
    renderImQc();
    runLldboCheck(); // runs even without an LLDBO file yet — shows candidate parts as a preview
    runMaterialCheck();
    runRevisionCheck();
    runTitleDescCheck();
  }

  function handleCadWorkbook(file, wb) {
    const res = BC.detect.parseCadFromWorkbook(wb, XLSX);
    if (res && res.ok) {
      const cad = res.ok;
      cad.fileName = file.name;
      addCadSource(cad);
      renderCadSources('', '');
      if (cad.imShaped) {
        notice('warn', 'The file in the CAD BOM box looks like an Item Master export (it has Vault item columns such as “Row Order”/“Row Type”). If that was a mistake, drop it on the right-hand box instead.');
      }
      return;
    }
    if (res && res.needsMapping) {
      const asIm = BC.detect.parseItemMasterFromWorkbook(wb, XLSX);
      if (asIm) {
        throw new Error('This looks like an Item Master export — drop it on the Item Master box on the right.');
      }
      if (BC.detect.looksLikeLldbo(res.needsMapping.aoa)) {
        throw new Error('This looks like an LLDBO (long-lead parts) file — drop it on the Long-Lead Parts box below instead.');
      }
      renderCadSources('', '');
      showMapping({
        aoa: res.needsMapping.aoa,
        indents: null,
        source: 'leveled-sheet',
        analysis: res.needsMapping.analysis,
        fileName: file.name,
      });
      return;
    }
    throw new Error('Could not find BOM data in this file.');
  }

  /* ---------------- LLDBO (long-lead parts) handling ---------------- */

  function setLldboStatus(fileName, chips, error) {
    $('status-file-lldbo').textContent = fileName || '';
    const chipBox = $('status-chips-lldbo');
    chipBox.innerHTML = '';
    chipEls(chipBox, chips);
    $('status-error-lldbo').textContent = error || '';
    const zone = $('zone-lldbo');
    zone.classList.toggle('parsed', !error && !!fileName && !!state.lldbo);
    zone.classList.toggle('error', !!error);
  }

  function chipsForLldbo(lldbo) {
    const chips = [{ text: 'LLDBO', kind: 'good' }, { text: lldbo.rows.length + ' rows' }];
    const withPn = lldbo.rows.filter(function (r) { return r.partNo; }).length;
    chips.push({ text: withPn + ' with Part No' });
    if (lldbo.projectKey) chips.push({ text: lldbo.projectKey.spn + ' / ' + lldbo.projectKey.pn });
    if (lldbo.warnings && lldbo.warnings.length) {
      chips.push({ text: lldbo.warnings.length + ' note(s)', kind: 'warn', title: lldbo.warnings.join('\n') });
    }
    return chips;
  }

  function handleLldboWorkbook(file, wb) {
    const lldbo = BC.detect.parseLldboFromWorkbook(wb, XLSX);
    if (!lldbo) {
      const asIm = BC.detect.parseItemMasterFromWorkbook(wb, XLSX);
      if (asIm) throw new Error('This looks like an Item Master export — drop it on the Item Master box above instead.');
      const asCad = BC.detect.parseCadFromWorkbook(wb, XLSX);
      if (asCad && asCad.ok) throw new Error('This looks like a CAD BOM export — drop it on the CAD BOM box above instead.');
      throw new Error('No "PART NO" header row found. Expected the LLDBO (Long Lead Direct Bought Out) list.');
    }
    lldbo.fileName = file.name;
    state.lldbo = lldbo;
    setLldboStatus(file.name, chipsForLldbo(lldbo), '');
    if (state.im) runLldboCheck();
  }

  async function handleLldboFile(file) {
    try {
      if (/\.pdf$/i.test(file.name)) throw new Error('LLDBO is an Excel file, not a PDF.');
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
      handleLldboWorkbook(file, wb);
    } catch (e) {
      state.lldbo = null;
      runLldboCheck(); // falls back to the Item-Master-only preview instead of hiding everything
      setLldboStatus(file.name, [], e.message || String(e));
    }
  }

  /* ---------------- Ignore List handling ---------------- */

  function setIgnoreListStatus(fileName, chips, error) {
    $('status-file-ignorelist').textContent = fileName || '';
    const chipBox = $('status-chips-ignorelist');
    chipBox.innerHTML = '';
    chipEls(chipBox, chips);
    $('status-error-ignorelist').textContent = error || '';
    const zone = $('zone-ignorelist');
    zone.classList.toggle('parsed', !error && !!fileName && !!state.ignoreList);
    zone.classList.toggle('error', !!error);
  }

  function chipsForIgnoreList(ignoreList) {
    const chips = [{ text: 'Ignore List', kind: 'good' }, { text: ignoreList.rows.length + ' parts' }];
    if (ignoreList.warnings && ignoreList.warnings.length) {
      chips.push({ text: ignoreList.warnings.length + ' note(s)', kind: 'warn', title: ignoreList.warnings.join('\n') });
    }
    return chips;
  }

  // Shows/hides the "unrecognized From value" notice right at upload time —
  // this only depends on the ignore list's own data plus the fixed category
  // table, not on a comparison having run, so it doesn't wait for renderDashboard.
  function showIgnoreListCategoryWarning(ignoreList) {
    const warnEl = $('ignorelist-category-warning');
    const idx = BC.ignoreListCompare.buildIgnoreIndex(ignoreList);
    if (idx.unrecognized.length) {
      const names = Array.from(new Set(idx.unrecognized.map(function (w) { return '"' + w.from + '"'; }))).join(', ');
      warnEl.textContent = '⚠ ' + idx.unrecognized.length + ' row(s) have an unrecognized "From" value (' + names +
        ') — not applied. Recognized values: "CAD vs Item compare", "Quantity Mismatch", "Revision", "LLDBO Candidate", "All".';
      warnEl.classList.remove('hidden');
    } else {
      warnEl.classList.add('hidden');
      warnEl.textContent = '';
    }
  }

  // The ignore list only takes effect through compareAll(), compareRevision()
  // and detectLldboCandidates() (see js/compare.js, js/revision-compare.js,
  // js/lldbo-compare.js and js/ignorelist-compare.js), so re-applying it
  // after a change means re-running all three, not just re-rendering stale
  // results — the same runCompare() the "Compare BOMs" button and folder
  // auto-load use, plus runRevisionCheck()/runLldboCheck() (harmless no-ops
  // if no Item Master is loaded yet).
  function applyIgnoreListChange() {
    runRevisionCheck();
    runLldboCheck();
    if (state.result) runCompare();
    else renderDashboard();
  }

  function handleIgnoreListWorkbook(file, wb) {
    const ignoreList = BC.detect.parseIgnoreListFromWorkbook(wb, XLSX);
    if (!ignoreList) {
      const asIm = BC.detect.parseItemMasterFromWorkbook(wb, XLSX);
      if (asIm) throw new Error('This looks like an Item Master export — drop it on the Item Master box above instead.');
      throw new Error('No "Part Number" header row found. Expected the Ignore List file ("IgnoreList*").');
    }
    ignoreList.fileName = file.name;
    state.ignoreList = ignoreList;
    setIgnoreListStatus(file.name, chipsForIgnoreList(ignoreList), '');
    showIgnoreListCategoryWarning(ignoreList);
    applyIgnoreListChange();
  }

  async function handleIgnoreListFile(file) {
    try {
      if (/\.pdf$/i.test(file.name)) throw new Error('Ignore List is an Excel file, not a PDF.');
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
      handleIgnoreListWorkbook(file, wb);
    } catch (e) {
      state.ignoreList = null;
      setIgnoreListStatus(file.name, [], e.message || String(e));
      $('ignorelist-category-warning').classList.add('hidden');
      applyIgnoreListChange();
    }
  }

  async function handleCadPdf(file, buf) {
    const grid = await BC.pdfExtract.extractGrid(buf, {
      pdfjsLib: window.pdfjsLib,
      onProgress: function (p, total) {
        renderCadSources('reading ' + file.name + ' — page ' + p + '/' + total + '…', '');
      },
    });
    if (grid.rows.length < 2) {
      throw new Error((grid.warnings[0] || 'No table found in the PDF.') +
        '\nExpected the Vault web client\'s multi-level BOM ("Uses") PDF report.');
    }
    const parsed = BC.cadLeveledParser.parse(grid.rows, { indents: grid.indents, pageOf: grid.pageOf, source: 'pdf' });
    if (!parsed) {
      renderCadSources('', '');
      showMapping({
        aoa: grid.rows,
        indents: grid.indents,
        pageOf: grid.pageOf,
        source: 'pdf',
        analysis: BC.cadLeveledParser.analyze(grid.rows),
        fileName: file.name,
      });
      return;
    }
    parsed.fileName = file.name;
    parsed.warnings = (grid.warnings || []).concat(parsed.warnings || []);
    addCadSource(parsed);
    renderCadSources('', '');
  }

  /* ---------------- column-mapping panel ---------------- */

  const MAPPING_FIELDS = [
    { key: 'number', label: 'Part Number *' },
    { key: 'qty', label: 'Quantity' },
    { key: 'level', label: 'Level' },
    { key: 'pos', label: 'Item / Position (dotted)' },
    { key: 'title', label: 'Title' },
    { key: 'description', label: 'Description' },
    { key: 'file', label: 'File name' },
  ];

  function showMapping(ctx) {
    state.mappingCtx = ctx;
    const panel = $('mapping-panel');
    const fieldsBox = $('mapping-fields');
    fieldsBox.innerHTML = '';
    for (const f of MAPPING_FIELDS) {
      const label = document.createElement('label');
      label.textContent = f.label;
      const sel = document.createElement('select');
      sel.dataset.field = f.key;
      const none = document.createElement('option');
      none.value = '-1';
      none.textContent = '(none)';
      sel.appendChild(none);
      for (const col of ctx.analysis.columns) {
        const opt = document.createElement('option');
        opt.value = String(col.index);
        opt.textContent = col.header + (col.samples.length ? ' — e.g. ' + col.samples[0].slice(0, 24) : '');
        sel.appendChild(opt);
      }
      const detected = ctx.analysis.detected[f.key];
      if (detected !== undefined) sel.value = String(detected);
      label.appendChild(sel);
      fieldsBox.appendChild(label);
    }
    // preview table
    const prev = $('mapping-preview');
    prev.innerHTML = '';
    const hdrTr = document.createElement('tr');
    for (const col of ctx.analysis.columns) {
      const th = document.createElement('th');
      th.textContent = col.header;
      hdrTr.appendChild(th);
    }
    prev.appendChild(hdrTr);
    const start = ctx.analysis.headerRow + 1;
    for (let r = start; r < Math.min(start + 6, ctx.aoa.length); r++) {
      const tr = document.createElement('tr');
      for (const col of ctx.analysis.columns) {
        const td = document.createElement('td');
        const v = (ctx.aoa[r] || [])[col.index];
        td.textContent = v === null || v === undefined ? '' : String(v).slice(0, 30);
        tr.appendChild(td);
      }
      prev.appendChild(tr);
    }
    panel.classList.remove('hidden');
  }

  function hideMapping() {
    $('mapping-panel').classList.add('hidden');
    state.mappingCtx = null;
  }

  $('mapping-apply').addEventListener('click', function () {
    const ctx = state.mappingCtx;
    if (!ctx) return;
    const mapping = {};
    $('mapping-fields').querySelectorAll('select').forEach(function (sel) {
      mapping[sel.dataset.field] = parseInt(sel.value, 10);
    });
    if (mapping.number < 0) { alert('Part Number column is required.'); return; }
    const parsed = BC.cadLeveledParser.parse(ctx.aoa, {
      mapping: mapping,
      headerRow: ctx.analysis.headerRow,
      indents: ctx.indents,
      pageOf: ctx.pageOf,
      source: ctx.source,
    });
    if (!parsed || !parsed.items.length) { alert('No rows with a part number found in that column.'); return; }
    parsed.fileName = ctx.fileName;
    addCadSource(parsed);
    renderCadSources('', '');
    hideMapping();
    updateCompareButton();
  });

  /* ---------------- notices ---------------- */

  function clearNotices() { $('notices').innerHTML = ''; }
  function notice(kind, text) {
    const el = document.createElement('div');
    el.className = 'notice ' + kind;
    el.textContent = text;
    $('notices').appendChild(el);
  }

  /* ---------------- Item Master QC panel ---------------- */

  function qcKindLabel(kind) {
    return {
      'both-missing': 'Title AND Description both blank',
      'title-missing': 'Title blank',
      'description-missing': 'Description blank',
    }[kind] || kind;
  }

  // cols entries are [dataKey, columnLabel, formatter?] — formatter(value, row)
  // defaults to the raw value.
  function qcCellValue(col, row) {
    const value = row[col[0]];
    return col[2] ? col[2](value, row) : value;
  }

  // Every check's columns include Row # (the row's position in the uploaded
  // file) and Parent Number/Parent Title (the immediate parent assembly, from
  // Row Order) so a flagged row can be found and placed in context directly
  // from the table/export, without cross-referencing the source file by hand.
  const LOCATION_COLS = [['sourceRow', 'Row #'], ['parentNumber', 'Parent Number'], ['parentTitle', 'Parent Title']];

  const QC_CHECKS = [
    { key: 'c1', title: 'Producer ↔ Description Match (top-level row only)',
      desc: 'The top-level row’s Producer and Producer Number should appear inside its Description.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat(
        [['producer', 'Producer'], ['producerNumber', 'Producer Number'], ['description', 'Description'], ['issue', 'Issue']]) },
    { key: 'c2', title: 'End of Line Integrity',
      desc: 'The "END OF LINE" entry should have Number ' + BC.imQc.END_OF_LINE_NUMBER + ' and a whole-number Row Order.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat([['issue', 'Issue']]) },
    { key: 'c3', title: 'Quantity vs Item Qty',
      desc: 'The numeric value inside Quantity should equal Item Qty on every row — a mismatch usually means one was edited manually without updating the other.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat(
        [['title', 'Title'], ['quantity', 'Quantity'], ['itemQty', 'Item Qty']]) },
    { key: 'c4', title: 'Entity Icon Status',
      desc: 'Every row’s Entity Icon should read "Normal".',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat([['icon', 'Entity Icon']]) },
    { key: 'c5', title: 'Title / Description Completeness',
      desc: 'Every row should have a Title and Description. Purchased/catalog parts (X-999-…) are only flagged when both are blank; other parts are flagged if either is missing.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat(
        [['title', 'Title'], ['description', 'Description'], ['kind', 'Issue', qcKindLabel]]) },
    { key: 'c6', title: 'Material Completeness',
      desc: 'Every non-assembly row should have a Material. Assemblies/weldments are excluded (they legitimately carry no material of their own).',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat([['title', 'Title']]) },
    { key: 'c7', title: 'Revision Consistency',
      desc: 'The same part number used at more than one BOM position should carry the same Revision everywhere it appears.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat(
        [['revision', 'Revision'], ['conflictsWith', 'Also Recorded As']]) },
    { key: 'c8', title: 'Sketch Parts in Item Master (7-333-…)',
      desc: '7-333-… numbers are rough-sketch models. They may exist in CAD, but only ever as Reference — one reaching the released Item Master is a serious release error. Every hit below shows its Row # and the full parent trail so you can trace how it got in.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat(
        [['title', 'Title'], ['state', 'State'], ['trail', 'Full Parent Trail']]) },
    { key: 'c9', title: 'Item State (obsolete / invalid revisions)',
      desc: 'A released BOM should contain Certified items only. Obsolete, Invalid and Phased Out are errors — the part was released against a dead revision. "New" is shown as a warning, not a failure.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order']].concat(LOCATION_COLS).concat(
        [['title', 'Title'], ['state', 'State'], ['severity', 'Severity']]) },
  ];

  function hideImQc() {
    $('im-qc').classList.add('hidden');
    $('qc-summary').innerHTML = '';
    $('qc-sections').innerHTML = '';
    $('qc-callout').classList.add('hidden');
    $('qc-callout').textContent = '';
    renderDashboard();
  }

  // A check can PASS on its own terms (e.g. every part has a Material
  // value) while a related cross-source comparison still finds a real
  // problem (e.g. that value doesn't match the CAD model) — surfacing that
  // as plain green would hide it. Returns a short message to show instead
  // of "clean" when that's the case, or null when there's nothing to add.
  function relatedWarningFor(check) {
    if (check.key === 'c6' && state.materialResult && state.materialResult.applicable && state.materialResult.mismatches.length) {
      return state.materialResult.mismatches.length + ' material mismatch(es) vs CAD — see the Material vs CAD section below';
    }
    if (check.key === 'c7' && state.revisionResult && state.revisionResult.applicable && state.revisionResult.mismatches.length) {
      return state.revisionResult.mismatches.length + ' revision mismatch(es) vs CAD — see the Revision vs CAD section below';
    }
    return null;
  }

  // A check that separates hard errors from warnings (Check 9: state) reports
  // errorCount/warnCount. When it flagged rows but none is an error, the card
  // is amber and reads "not yet certified" rather than counting as a failure.
  function warnOnly(result) {
    return result.applicable && result.errorCount === 0 && result.warnCount > 0;
  }

  function qcCardFor(check, result, warnMsg) {
    const soft = warnOnly(result);
    const cls = !result.applicable ? 'na' : (result.fail.length && !soft ? 'red' : ((soft || warnMsg) ? 'amber' : 'pass'));
    const num = !result.applicable ? 'N/A' : (result.fail.length ? String(result.fail.length) : (warnMsg ? '⚠' : 'PASS'));
    const sub = !result.applicable ? result.reason
      : (soft ? result.warnCount + ' not yet certified (no obsolete/invalid)'
        : (result.fail.length
          ? (result.errorCount !== undefined ? result.errorCount + ' error(s)' + (result.warnCount ? ' + ' + result.warnCount + ' warning(s)' : '') : 'flagged')
          : (warnMsg || 'clean')));
    const el = document.createElement('div');
    el.className = 'card ' + cls;
    el.innerHTML = '<div class="num"></div><div class="lbl"></div>';
    el.querySelector('.num').textContent = num;
    el.querySelector('.lbl').textContent = check.title.replace(/\s*\(.*\)$/, '') + ' — ' + sub;
    return el;
  }

  // Reveals a section's description text on demand instead of always showing
  // it — stopPropagation matters because the surrounding section head already
  // has its own click listener that toggles the results body.
  function wireDescToggle(btn, descEl) {
    btn.setAttribute('aria-expanded', 'false');
    btn.addEventListener('click', function (ev) {
      ev.stopPropagation();
      const open = descEl.classList.toggle('open');
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
  }

  // Section title + a collapsed-by-default description behind a small (i)
  // toggle. Shared by qcSectionFor/lldboSectionFor/boughtOutSection so the
  // "too much text" fix lives in one place instead of three.
  function buildQcTitleBox(titleText, descText) {
    const box = document.createElement('div');
    box.innerHTML = '<div class="qc-section-title-row"><span class="qc-section-title"></span>' +
      '<button type="button" class="qc-info-toggle" aria-label="Show description" title="Show description">i</button></div>' +
      '<p class="qc-section-desc"></p>';
    box.querySelector('.qc-section-title').textContent = titleText;
    box.querySelector('.qc-section-desc').textContent = descText;
    wireDescToggle(box.querySelector('.qc-info-toggle'), box.querySelector('.qc-section-desc'));
    return box;
  }

  // A check turned off in "Checks to run" — distinct from the existing "N/A"
  // treatment (that means the loaded files don't carry the data a check
  // needs; this means the user chose not to run it). The section stays
  // visible rather than disappearing, so it's easy to find and re-enable.
  function skippedSection(title) {
    const box = document.createElement('div');
    box.className = 'qc-section';
    const head = document.createElement('div');
    head.className = 'qc-section-head';
    head.innerHTML = '<div class="qc-section-title"></div><span class="qc-pill na">SKIPPED</span>';
    head.querySelector('.qc-section-title').textContent = title;
    box.appendChild(head);
    const body = document.createElement('div');
    body.className = 'qc-section-body open';
    body.innerHTML = '<div class="empty-state">Turned off in "Checks to run ▾" — enable it there to see this check.</div>';
    box.appendChild(body);
    return box;
  }

  function qcSectionFor(check, result, warnMsg) {
    const box = document.createElement('div');
    box.className = 'qc-section';
    const head = document.createElement('div');
    head.className = 'qc-section-head';
    head.appendChild(buildQcTitleBox(check.title, check.desc));
    const soft = warnOnly(result);
    const pill = document.createElement('span');
    pill.className = 'qc-pill ' + (!result.applicable ? 'na' : (result.fail.length && !soft ? 'fail' : ((soft || warnMsg) ? 'warn' : 'pass')));
    pill.textContent = !result.applicable ? 'N/A'
      : (soft ? result.warnCount + ' TO REVIEW'
        : (result.fail.length ? result.fail.length + ' FLAGGED' : (warnMsg ? 'OK, BUT SEE BELOW' : 'OK')));
    head.appendChild(pill);
    box.appendChild(head);

    const body = document.createElement('div');
    body.className = 'qc-section-body';
    let contentOpen = false;
    if (!result.applicable) {
      body.innerHTML = '<div class="empty-state">' + result.reason + '</div>';
    } else if (!result.fail.length) {
      body.innerHTML = '<div class="empty-state">' + (warnMsg ? '⚠ ' + warnMsg : '✓ No issues found' + (check.key === 'c1' && result.applicableCount ? ' across ' + result.applicableCount + ' applicable row(s)' : '') + '.') + '</div>';
    } else {
      const table = document.createElement('table');
      table.className = 'results-table';
      const htr = document.createElement('tr');
      for (const [, label] of check.cols) addTh(htr, label);
      table.appendChild(htr);
      for (const row of result.fail) {
        const tr = document.createElement('tr');
        for (const col of check.cols) addTd(tr, qcCellValue(col, row));
        // Check 9's warning rows are ranked under their own key in the
        // findings registry (see findings.js), so match that here or every
        // warning row would look like a demoted duplicate of itself.
        const key = (check.key === 'c9' && !/^ERROR/.test(row.severity || '')) ? 'c9warn' : check.key;
        decorateSecondary(tr, key, row.number);
        table.appendChild(tr);
      }
      body.appendChild(table);
      contentOpen = true;
    }
    box.appendChild(body);

    // A manual toggle overrides the content-driven default above, and must
    // survive this box being torn down and rebuilt on every re-render (this
    // function runs fresh every time — see state.sectionOpen).
    const id = check.key;
    const isOpen = state.sectionOpen.has(id) ? state.sectionOpen.get(id) : contentOpen;
    if (isOpen) body.classList.add('open');
    head.addEventListener('click', function () { state.sectionOpen.set(id, body.classList.toggle('open')); });
    return box;
  }

  function renderImQc() {
    const qc = state.imQc;
    if (!qc) { hideImQc(); return; }
    rebuildFindings(); // rows below are decorated against it
    $('im-qc').classList.remove('hidden');
    setSourceLine($('qc-source'), { imOnly: true });
    const summary = $('qc-summary');
    summary.innerHTML = '';
    const sections = $('qc-sections');
    sections.innerHTML = '';
    $('qc-callout').classList.add('hidden');
    if (!isCheckEnabled('imQc')) {
      sections.appendChild(skippedSection('Item Master QC'));
      renderDashboard();
      return;
    }
    for (const check of QC_CHECKS) {
      const result = qc[check.key];
      const warnMsg = result.applicable && !result.fail.length ? relatedWarningFor(check) : null;
      summary.appendChild(qcCardFor(check, result, warnMsg));
      sections.appendChild(qcSectionFor(check, result, warnMsg));
    }

    const callout = $('qc-callout');
    const bits = [];
    if (qc.c8 && qc.c8.applicable && qc.c8.fail.length) {
      bits.push('🚨 ' + qc.c8.fail.length + ' sketch part(s) (7-333-…) made it into the Item Master — these must never be released');
    }
    if (qc.c9 && qc.c9.applicable && qc.c9.errorCount) {
      bits.push('🚨 ' + qc.c9.errorCount + ' part(s) released against an obsolete/invalid state');
    }
    if (qc.c5.applicable && qc.c5.fail.length) {
      bits.push(qc.c5.fail.length + ' row(s) need Title and/or Description populated');
    }
    if (qc.c6.applicable && qc.c6.fail.length) {
      bits.push(qc.c6.fail.length + ' part(s) need Material populated');
    }
    if (qc.c7.applicable && qc.c7.fail.length) {
      bits.push(qc.c7.fail.length + ' row(s) have a Revision that conflicts with another occurrence of the same part');
    }
    if (bits.length) {
      callout.textContent = '⚠ ' + bits.join(' · ') + ' — see below, or download the QC report.';
      callout.classList.remove('hidden');
    } else {
      callout.textContent = '';
      callout.classList.add('hidden');
    }
    renderDashboard();
  }

  /* ---------------- overview dashboard + source provenance ---------------- */

  // Human label for a CAD source, including its file name, so a result can
  // always be traced back to the exact export that produced it — different
  // people export with different Vault columns enabled, which changes which
  // checks are even available.
  function cadSourceLabel(s) {
    let lbl = { 'flat-xlsx': 'Vault flat export', 'pdf': 'Vault PDF', 'leveled-sheet': 'Leveled table' }[s.source] || s.source;
    if (s.source === 'leveled-sheet' && s.hasStructure) lbl = 'Inventor BOM export';
    return lbl + (s.fileName ? ' — ' + s.fileName : '');
  }

  function imSourceLabel() {
    if (!state.im) return '';
    return (state.im.fileName || 'Item Master') + ' · ' + state.im.rows.length + ' rows';
  }

  // Fills a .source-line element with "CAD BOM: … ↔ Item Master: …" (built
  // with DOM nodes, not innerHTML, since file names are untrusted text).
  // opts.imOnly shows just the Item Master side (for the IM-only QC section).
  function setSourceLine(el, opts) {
    if (!el) return;
    opts = opts || {};
    el.textContent = '';
    const add = function (label, value) {
      if (el.childNodes.length) el.appendChild(document.createTextNode('  ↔  '));
      const b = document.createElement('strong');
      b.textContent = label + ' ';
      el.appendChild(b);
      el.appendChild(document.createTextNode(value));
    };
    if (!opts.imOnly && state.cadSources.length) {
      add('CAD BOM:', state.cadSources.map(cadSourceLabel).join('  +  '));
    }
    if (state.im) add('Item Master:', imSourceLabel());
  }

  function jumpTo(sectionId, tab) {
    const sec = $(sectionId);
    if (!sec) return;
    if (sectionId === 'results') {
      if (!state.result) return; // nothing to jump to until a compare has run
      if (tab) switchTab(tab);
    }
    sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    sec.classList.remove('section-flash');
    void sec.offsetWidth; // reflow so the flash animation retriggers on repeat clicks
    sec.classList.add('section-flash');
  }

  function dashTileEl(t) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'flag-tile' + (t.cls ? ' ' + t.cls : '');
    const num = document.createElement('div'); num.className = 'num'; num.textContent = t.num;
    const lbl = document.createElement('div'); lbl.className = 'lbl'; lbl.textContent = t.label;
    b.appendChild(num); b.appendChild(lbl);
    if (t.sub) { const s = document.createElement('div'); s.className = 'sub'; s.textContent = t.sub; b.appendChild(s); }
    b.addEventListener('click', t.onClick);
    return b;
  }

  // The top-of-page flag summary: one clickable tile per flag value, each
  // jumping to (and flashing) its detail section below. Rebuilt from full
  // state on every change, so it always reflects whatever has been loaded and
  // compared so far — QC/material/revision appear as soon as the Item Master
  // (and a CAD source) load; the compare tiles appear once "Compare" is run.
  // Rebuilds the cross-check registry from current state. Must run BEFORE any
  // renderer that calls decorateSecondary(), otherwise those rows are decorated
  // against a stale registry (e.g. a quantity mismatch marked secondary to a
  // check that only looked primary because the comparison hadn't run yet).
  function rebuildFindings() {
    state.findings = BC.findings.buildRegistry({
      result: state.result,
      imQc: state.imQc,
      materialResult: state.materialResult,
      revisionResult: state.revisionResult,
      titleDescResult: state.titleDescResult,
      virtualResult: state.virtualResult,
      lldboResult: state.lldboResult,
      lldboCandidatesResult: state.lldboCandidatesResult,
    }, isCheckEnabled);
  }

  function renderDashboard() {
    const dash = $('dashboard');
    rebuildFindings();
    renderAttention(); // sets state.attentionVisibleCount, used by the tile below
    renderIgnoredFindings();
    renderReviewedParts();
    if (!state.im && !state.cadSources.length) { dash.classList.add('hidden'); return; }
    dash.classList.remove('hidden');
    setSourceLine($('dash-source'));

    const wrap = $('dash-tiles');
    wrap.innerHTML = '';
    const tiles = [];

    const reg = state.findings;
    if (reg && reg.counts.parts) {
      tiles.push({
        num: state.attentionVisibleCount,
        label: 'Parts needing attention',
        sub: reg.counts.grouped ? '+' + reg.counts.grouped + ' grouped under a parent' : 'unique parts, worst issue first',
        cls: state.attentionVisibleCount ? 'red' : 'pass',
        onClick: function () { jumpTo('attention'); },
      });
    }

    const res = state.result;
    if (res) {
      if (isCheckEnabled('missing')) {
        tiles.push({ num: res.actionableCount, label: 'Findings needing action',
          cls: res.actionableCount ? 'red' : 'pass', onClick: function () { jumpTo('results', 'missing'); } });
      }
      if (isCheckEnabled('qty')) {
        tiles.push({ num: res.qtyMismatches === null ? '—' : res.qtyMismatches.length, label: 'Quantity mismatches',
          cls: res.qtyMismatches === null ? 'na' : (res.qtyMismatches.length ? 'amber' : 'pass'),
          onClick: function () { jumpTo('results', 'qty'); } });
      }
      const vres = state.virtualResult;
      if (isCheckEnabled('virtual') && vres && vres.applicable) {
        const vn = vres.confirmed.length + vres.suspected.length;
        tiles.push({ num: vn, label: 'Virtual parts',
          sub: vn ? 'no CAD file behind them — child BOM never compared'
                  : (vres.hasThumbnailColumn ? 'none found ✓' : 'none inferred (no Thumbnail column)'),
          cls: vn ? 'red' : 'pass', onClick: function () { jumpTo('results', 'virtual'); } });
      }
      if (isCheckEnabled('qtyCascade') && res.qtyCascades.applicable) {
        tiles.push({ num: res.qtyCascades.roots.length, label: 'Quantity cascades',
          sub: res.qtyCascades.roots.length ? 'one root cause explains many mismatches — fix it first' : 'no uniform-ratio subtree found ✓',
          cls: res.qtyCascades.roots.length ? 'red' : 'pass', onClick: function () { jumpTo('results', 'cascade'); } });
      }
    }

    const mat = state.materialResult;
    if (isCheckEnabled('material') && mat) tiles.push({ num: mat.applicable ? mat.mismatches.length : '—', label: 'Material mismatches vs CAD',
      cls: mat.applicable ? (mat.mismatches.length ? 'amber' : 'pass') : 'na',
      onClick: function () { jumpTo('material-sections'); } });

    const rev = state.revisionResult;
    if (isCheckEnabled('revision') && rev) tiles.push({ num: rev.applicable ? rev.mismatches.length : '—', label: 'Revision mismatches vs CAD',
      cls: rev.applicable ? (rev.mismatches.length ? 'amber' : 'pass') : 'na',
      onClick: function () { jumpTo('revision-sections'); } });

    const qc = state.imQc;
    if (isCheckEnabled('imQc') && qc) {
      // The two release-blocking checks get their own tiles — a sketch part in
      // the Item Master, or a part released against a dead revision, is a
      // different class of problem from general data-quality tidiness.
      if (qc.c8 && qc.c8.applicable) {
        tiles.push({ num: qc.c8.fail.length, label: 'Sketch parts (7-333-) in Item Master',
          sub: qc.c8.fail.length ? 'must never be released — trace below' : 'none present ✓',
          cls: qc.c8.fail.length ? 'red' : 'pass', onClick: function () { jumpTo('im-qc'); } });
      }
      if (qc.c9 && qc.c9.applicable) {
        tiles.push({ num: qc.c9.errorCount, label: 'Obsolete / invalid states',
          sub: qc.c9.warnCount ? qc.c9.warnCount + ' more not yet certified' : (qc.c9.errorCount ? 'released against a dead revision' : 'all states valid ✓'),
          cls: qc.c9.errorCount ? 'red' : (qc.c9.warnCount ? 'amber' : 'pass'),
          onClick: function () { jumpTo('im-qc'); } });
      }
      let flagged = 0, applicable = 0;
      for (const check of QC_CHECKS) { const r = qc[check.key]; if (r.applicable) { applicable++; if (r.fail.length) flagged++; } }
      tiles.push({ num: flagged, label: 'Item Master quality issues',
        sub: flagged ? 'of ' + applicable + ' applicable checks' : 'all ' + applicable + ' checks clean',
        cls: flagged ? 'red' : 'pass', onClick: function () { jumpTo('im-qc'); } });
    }

    const lld = state.lldboResult;
    if (isCheckEnabled('lldbo') && lld) {
      // Confident-tier candidates only count once cross-checked against a
      // loaded LLDBO file (cand.crossChecked is always true here, since lld
      // is only non-null once state.lldbo exists too) — the Item-Master-only
      // preview never contributes to this tile.
      const cand = state.lldboCandidatesResult;
      const candMissing = (cand && cand.crossChecked) ? cand.confident.length : 0;
      const n = lld.missingFromIm.length + lld.qtyMismatches.length + candMissing;
      tiles.push({ num: n, label: 'Long-lead (LLDBO) issues',
        cls: lld.projectKeyMismatch ? 'red' : (n ? 'amber' : 'pass'),
        onClick: function () { jumpTo('lldbo-panel'); } });
    }

    for (const t of tiles) wrap.appendChild(dashTileEl(t));
  }

  /* ----- consolidated "parts needing attention" + cross-check demotion ----- */

  function issueBadgeEl(issue) {
    const b = document.createElement('span');
    b.className = 'issue-badge' + (issue.primary ? ' primary' : (issue.severity >= 60 ? ' sev-mid' : ''));
    b.textContent = issue.label + (issue.occurrences > 1 ? ' ×' + issue.occurrences : '');
    if (issue.detail) b.title = issue.detail;
    return b;
  }

  // One row per part, worst issue first. A part flagged by several checks is
  // listed once here with every issue on it, instead of once per check.
  function renderAttention() {
    const sec = $('attention');
    const body = $('attention-body');
    const reg = state.findings;
    body.innerHTML = '';
    if (!reg || !reg.counts.parts) { sec.classList.add('hidden'); state.attentionVisibleCount = 0; return; }
    sec.classList.remove('hidden');

    const cols = visibleCols();
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, 'Reviewed');
    addTh(htr, 'Priority');
    addTh(htr, 'Part Number');
    if (cols.title) addTh(htr, 'Title');
    addTh(htr, 'Issues found');
    addTh(htr, 'Detail');
    if (cols.row) addTh(htr, 'Row #');
    addTh(htr, 'Grouped under');
    table.appendChild(htr);

    // Reviewed parts are pulled out before the grouped/filter split below, so
    // they never count toward "needing attention" — see "Already reviewed"
    // for how to bring one back.
    const unreviewed = reg.parts.filter(function (p) { return !isReviewed(p.number); });
    const reviewedCount = reg.parts.length - unreviewed.length;
    state.attentionVisibleCount = unreviewed.filter(function (p) { return !p.grouped; }).length;
    $('attention-count').textContent = state.attentionVisibleCount + ' PART(S)' +
      (reg.counts.grouped ? ' · ' + reg.counts.grouped + ' GROUPED' : '') +
      (reviewedCount ? ' · ' + reviewedCount + ' REVIEWED' : '');

    const shown = unreviewed.filter(function (p) {
      if (!state.showGrouped && p.grouped) return false;
      return matchesFilter([p.number, p.title, p.description]);
    });
    for (const part of shown) {
      const tr = document.createElement('tr');
      tr.className = 'attention-row' + (part.grouped ? ' row-child' : '');
      const tdReviewed = document.createElement('td');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.title = 'Mark reviewed — hides it here; bring it back from "Already reviewed" below';
      cb.addEventListener('click', function (ev) {
        ev.stopPropagation();
        setReviewed(part.number, true);
      });
      tdReviewed.appendChild(cb);
      tr.appendChild(tdReviewed);
      const tdPriority = document.createElement('td');
      tdPriority.appendChild(priorityControlEl(part.number));
      tr.appendChild(tdPriority);
      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      tdNum.textContent = part.number;
      tr.appendChild(tdNum);
      if (cols.title) addTd(tr, part.title);
      const tdIssues = document.createElement('td');
      for (const i of part.issues) tdIssues.appendChild(issueBadgeEl(i));
      tr.appendChild(tdIssues);
      const tdDetail = document.createElement('td');
      tdDetail.className = 'issue-detail';
      tdDetail.textContent = part.primary ? part.primary.detail : '';
      tr.appendChild(tdDetail);
      if (cols.row) addTd(tr, part.primary ? part.primary.sourceRow : '');
      addTd(tr, part.groupedUnder || '');
      // clicking jumps to wherever the part's worst issue actually lives
      tr.addEventListener('click', function () {
        if (part.primary) jumpTo(part.primary.section, part.primary.tab);
      });
      table.appendChild(tr);
    }
    if (!shown.length) {
      const msg = state.filter ? 'Nothing matches the filter.'
        : (unreviewed.length ? 'Nothing shown — everything left is grouped under a flagged parent.'
          : (reviewedCount ? '✓ All flagged parts are marked reviewed — see "Already reviewed" below.' : '✓ No parts flagged by any check.'));
      body.innerHTML = '<div class="empty-state">' + msg + '</div>';
    } else {
      body.appendChild(table);
    }
    if (reg.counts.grouped) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'attention-toggle';
      btn.textContent = state.showGrouped
        ? 'Hide ' + reg.counts.grouped + ' parts grouped under a flagged parent'
        : 'Show ' + reg.counts.grouped + ' parts grouped under a flagged parent';
      btn.addEventListener('click', function () { state.showGrouped = !state.showGrouped; renderAttention(); });
      body.appendChild(btn);
    }
  }

  // Collapsible toggle for a whole section, wired once (unlike the
  // qc-section pattern, these bodies are never replaced by re-render — only
  // their innerHTML — so the 'open' class seeded in index.html survives every
  // re-render with no state bookkeeping needed). Also flips the .expander
  // chevron glyph so the section visibly reads as collapsible.
  function wireHeadToggle(headId, bodyId, chevronId) {
    $(headId).addEventListener('click', function () {
      const open = $(bodyId).classList.toggle('open');
      $(chevronId).textContent = open ? '▾' : '▸';
    });
  }
  wireHeadToggle('attention-head', 'attention-body', 'attention-chevron');
  wireHeadToggle('reviewed-parts-head', 'reviewed-parts-body', 'reviewed-parts-chevron');
  wireHeadToggle('ignored-findings-head', 'ignored-findings-body', 'ignored-findings-chevron');

  // Static panel intros (Item Master QC / LLDBO / Ignore List) — the
  // biggest always-visible prose blocks — get the same (i) reveal-on-demand
  // treatment as the per-check section descriptions above.
  [['im-qc-info-btn', 'im-qc-desc'], ['lldbo-info-btn', 'lldbo-desc'], ['ignorelist-info-btn', 'ignorelist-desc']]
    .forEach(function (p) { wireDescToggle($(p[0]), $(p[1])); });

  const IGNORED_FINDINGS_COLS = [
    ['number', 'Part Number'], ['title', 'Title'], ['description', 'Description'],
    ['label', 'Would have been flagged as'], ['sourceRow', 'Row #'], ['parentNumber', 'Parent Number'], ['parentTitle', 'Parent Title'],
  ];

  // Parts suppressed from the checks that support it — the file-based Ignore
  // List (js/compare.js's compareAll() and js/revision-compare.js's
  // compareRevision(), via js/ignorelist-compare.js) plus the in-session
  // "ignore bought-out parts" toggle on the Revision section — listed here
  // for transparency instead of just vanishing, but collapsed by default: a
  // review/audit list, not something needing action. Only visible once
  // there's something that could be suppressing a finding — before that
  // there's nothing to show.
  function renderIgnoredFindings() {
    const sec = $('ignored-findings');
    const body = $('ignored-findings-body');
    body.innerHTML = '';
    if (!state.ignoreList && !state.ignoreBoughtOutRevision) { sec.classList.add('hidden'); return; }
    sec.classList.remove('hidden');
    const list = ((state.result && state.result.ignoredFindings) || [])
      .concat((state.revisionResult && state.revisionResult.ignoredFindings) || [])
      .concat((state.lldboCandidatesResult && state.lldboCandidatesResult.ignoredFindings) || []);
    $('ignored-findings-count').textContent = list.length + ' SUPPRESSED';
    if (!list.length) {
      body.innerHTML = '<div class="empty-state">' + (state.result || state.revisionResult
        ? 'No currently-flaggable part matches the Ignore List — nothing is being suppressed right now.'
        : 'Run "Compare BOMs" to apply the Ignore List.') + '</div>';
      return;
    }
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    for (const [, label] of IGNORED_FINDINGS_COLS) addTh(htr, label);
    table.appendChild(htr);
    for (const row of list) {
      const tr = document.createElement('tr');
      for (const [key] of IGNORED_FINDINGS_COLS) {
        addTd(tr, key === 'label' ? (BC.findings.CHECK_BY_KEY[row.checkKey] || {}).label || row.checkKey : row[key]);
      }
      table.appendChild(tr);
    }
    body.appendChild(table);
  }

  const REVIEWED_PARTS_COLS = [['number', 'Part Number'], ['title', 'Title']];

  // Parts checked off "Parts needing attention" — the undo bin for that
  // checkbox. Driven by state.reviewedSet (persisted, see loadReviewed()
  // above), cross-referenced against the current findings registry for
  // display; a part no longer flagged by anything (e.g. the underlying issue
  // was fixed) still shows here so it can be un-reviewed and cleared out.
  function renderReviewedParts() {
    const sec = $('reviewed-parts');
    const body = $('reviewed-parts-body');
    body.innerHTML = '';
    if (!state.reviewedSet.size) { sec.classList.add('hidden'); return; }
    sec.classList.remove('hidden');
    const reg = state.findings;
    const rows = Array.from(state.reviewedSet).sort().map(function (pn) {
      return { pn: pn, part: reg ? reg.byPn.get(pn) : null };
    });
    $('reviewed-parts-count').textContent = rows.length + ' REVIEWED';
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, '');
    addTh(htr, 'Priority');
    for (const [, label] of REVIEWED_PARTS_COLS) addTh(htr, label);
    addTh(htr, 'Issue(s) when reviewed');
    table.appendChild(htr);
    for (const row of rows) {
      const tr = document.createElement('tr');
      const tdBtn = document.createElement('td');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'attention-toggle';
      btn.textContent = 'Un-review';
      btn.title = 'Bring this part back into "Parts needing attention"';
      btn.addEventListener('click', function () { setReviewed(row.pn, false); });
      tdBtn.appendChild(btn);
      tr.appendChild(tdBtn);
      const tdPriority = document.createElement('td');
      tdPriority.appendChild(priorityControlEl(row.pn));
      tr.appendChild(tdPriority);
      for (const [key] of REVIEWED_PARTS_COLS) addTd(tr, row.part ? row.part[key] : (key === 'number' ? row.pn : ''));
      const tdIssues = document.createElement('td');
      if (row.part) { for (const i of row.part.issues) tdIssues.appendChild(issueBadgeEl(i)); }
      else tdIssues.textContent = 'no longer flagged by any check';
      tr.appendChild(tdIssues);
      table.appendChild(tr);
    }
    body.appendChild(table);
  }

  // AOA rows for the "Parts needing attention" export sheet — one row per part
  // with every issue found on it, so the same part isn't chased across sheets.
  function attentionSheetRows(reg) {
    const rows = [['Action needed', 'Part Number', 'Title', 'Primary issue', 'Detail',
      'All issues found', 'Row #', 'Parent Number', 'Grouped under']];
    for (const p of reg.parts) {
      rows.push([
        p.grouped ? 'grouped (parent also flagged)' : 'YES',
        p.number, p.title,
        p.primary ? p.primary.label : '',
        p.primary ? p.primary.detail : '',
        p.issues.map(function (i) { return i.label + (i.occurrences > 1 ? ' x' + i.occurrences : ''); }).join('; '),
        p.primary ? p.primary.sourceRow : '',
        p.primary ? p.primary.parentNumber : '',
        p.groupedUnder || '',
      ]);
    }
    return rows;
  }

  // Mutes a detail row whose part is already reported, more seriously,
  // somewhere else — so the same part read across sections doesn't look like
  // several unrelated problems. The row stays (nothing is hidden), it just
  // points at the finding that owns it.
  function decorateSecondary(tr, checkKey, number, cell) {
    const reg = state.findings;
    if (!reg || !number || !reg.isSecondary(checkKey, number)) return;
    tr.classList.add('row-demoted');
    const primary = reg.primaryFor(number);
    if (!primary) return;
    tr.title = 'Also reported under: ' + primary.label;
    const target = cell || tr.firstChild;
    if (!target) return;
    const badge = document.createElement('span');
    badge.className = 'dup-badge';
    badge.textContent = 'also in: ' + primary.label;
    target.appendChild(badge);
  }

  // AOA rows for the "Item Master QC" export sheet, shared by the main
  // Compare export and the standalone QC-only export.
  function qcSheetRows(qc) {
    const rows = [['Item Master data quality checks', '']];
    for (const check of QC_CHECKS) {
      const result = qc[check.key];
      rows.push([]);
      rows.push([check.title, !result.applicable ? 'N/A — ' + result.reason : (result.fail.length ? result.fail.length + ' flagged' : 'OK')]);
      if (result.applicable && result.fail.length) {
        rows.push(check.cols.map(function (c) { return c[1]; }));
        for (const row of result.fail) rows.push(check.cols.map(function (c) { return qcCellValue(c, row); }));
      }
    }
    return rows;
  }

  $('btn-qc-export').addEventListener('click', function () {
    if (!state.imQc) return;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qcSheetRows(state.imQc)), 'Item Master QC');
    if (state.materialResult) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(materialSheetRows(state.materialResult)), 'Material vs CAD');
    }
    if (state.revisionResult) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(revisionSheetRows(state.revisionResult)), 'Revision vs CAD');
    }
    XLSX.writeFile(wb, 'ItemMaster-QC-report' + projectKeySuffix() + '.xlsx');
  });

  /* ---------------- LLDBO (long-lead parts) check ---------------- */

  function hideLldboResults() {
    $('lldbo-project-warning').classList.add('hidden');
    $('lldbo-project-warning').textContent = '';
    $('lldbo-summary').innerHTML = '';
    $('lldbo-sections').innerHTML = '';
    renderDashboard();
  }

  // Unlike compareLldbo() (needs both files), the candidate check below only
  // needs the Item Master — so this gate is intentionally looser than the
  // old "both files" one, and state.lldboResult itself may be null while
  // state.lldboCandidatesResult is populated (Item-Master-only preview).
  function runLldboCheck() {
    if (!state.im) {
      state.lldboResult = null;
      state.lldboCandidatesResult = null;
      hideLldboResults();
      return;
    }
    state.lldboResult = state.lldbo ? BC.lldboCompare.compareLldbo(state.lldbo, state.im, BC.indexItemMaster) : null;
    const ignoreIdx = BC.ignoreListCompare.buildIgnoreIndex(state.ignoreList);
    state.lldboCandidatesResult = BC.lldboCompare.detectLldboCandidates(state.im, state.lldbo, { isIgnored: ignoreIdx.isIgnored });
    renderLldboPanel();
  }

  function lldboCardFor(label, count, cls) {
    const el = document.createElement('div');
    el.className = 'card' + (cls ? ' ' + cls : '');
    el.innerHTML = '<div class="num"></div><div class="lbl"></div>';
    el.querySelector('.num').textContent = count;
    el.querySelector('.lbl').textContent = label;
    return el;
  }

  // `checkKey` (optional) ties the rows back to the findings registry so a part
  // already reported more seriously elsewhere renders muted with a cross-ref.
  // `traceable` (optional): adds a trace toggle to the Part Number cell (cols
  // must include a 'number' entry) revealing full cross-source provenance —
  // opt-in per call site, not every list built from this function needs it.
  // `rowClassFor` (optional): (row) -> className|falsy, applied to each <tr>
  // — e.g. the Revision section's amber "Item Master behind CAD" tint.
  function lldboSectionFor(title, desc, cols, fail, emptyText, checkKey, traceable, rowClassFor) {
    const box = document.createElement('div');
    box.className = 'qc-section';
    const head = document.createElement('div');
    head.className = 'qc-section-head';
    head.appendChild(buildQcTitleBox(title, desc));
    const pill = document.createElement('span');
    pill.className = 'qc-pill ' + (fail.length ? 'fail' : 'pass');
    pill.textContent = fail.length ? fail.length + ' FLAGGED' : 'OK';
    head.appendChild(pill);
    box.appendChild(head);

    const body = document.createElement('div');
    body.className = 'qc-section-body';
    let contentOpen = false;
    if (!fail.length) {
      body.innerHTML = '<div class="empty-state">' + emptyText + '</div>';
    } else {
      const table = document.createElement('table');
      table.className = 'results-table';
      const htr = document.createElement('tr');
      for (const [, label] of cols) addTh(htr, label);
      table.appendChild(htr);
      for (const row of fail) {
        const tr = document.createElement('tr');
        if (rowClassFor) {
          const cls = rowClassFor(row);
          if (cls) tr.className = cls;
        }
        for (const col of cols) {
          if (traceable && col[0] === 'number') {
            const td = document.createElement('td');
            td.appendChild(document.createTextNode(row.number + ' '));
            td.appendChild(traceToggleEl(row.number));
            tr.appendChild(td);
          } else {
            addTd(tr, row[col[0]]);
          }
        }
        if (checkKey) decorateSecondary(tr, checkKey, row.number);
        table.appendChild(tr);
        if (traceable) table.appendChild(traceDetailRow(row.number, cols.length));
      }
      if (traceable) wireTraceToggles(table);
      body.appendChild(table);
      contentOpen = true;
    }
    box.appendChild(body);

    // See qcSectionFor's comment — same override-survives-rebuild mechanism.
    // Falls back to the title text when no checkKey is given so unrelated
    // sections never accidentally share one state.sectionOpen slot.
    const id = checkKey || title;
    const isOpen = state.sectionOpen.has(id) ? state.sectionOpen.get(id) : contentOpen;
    if (isOpen) body.classList.add('open');
    head.addEventListener('click', function () { state.sectionOpen.set(id, body.classList.toggle('open')); });
    return box;
  }

  const LLDBO_MISSING_COLS = [['number', 'Part Number'], ['description', 'Description'], ['qtyText', 'LLDBO Qty'], ['sourceRow', 'LLDBO Row #']];
  const LLDBO_QTY_COLS = [['number', 'Part Number'], ['description', 'Description'], ['lldboQty', 'LLDBO Qty'], ['imQty', 'Item Master Qty'],
    ['sourceRow', 'LLDBO Row #'], ['foundUnder', 'Found Under (Item Master)']];
  const LLDBO_CANDIDATE_COLS = [['number', 'Part Number'], ['title', 'Title'], ['description', 'Description'],
    ['quantity', 'Quantity'], ['keyword', 'Matched Keyword']].concat(LOCATION_COLS);

  function renderLldboPanel() {
    const res = state.lldboResult;             // null until an LLDBO file is also loaded
    const cand = state.lldboCandidatesResult;   // null only when no Item Master is loaded
    if (!cand) { hideLldboResults(); return; }
    rebuildFindings(); // rows below are decorated against it

    if (!isCheckEnabled('lldbo')) {
      $('lldbo-project-warning').classList.add('hidden');
      $('lldbo-summary').innerHTML = '';
      const skippedSections = $('lldbo-sections');
      skippedSections.innerHTML = '';
      skippedSections.appendChild(skippedSection('Long-Lead Parts (LLDBO) check'));
      renderDashboard();
      return;
    }

    const warnEl = $('lldbo-project-warning');
    if (res && res.projectKeyMismatch) {
      warnEl.textContent = '⚠ This LLDBO document is for ' + res.projectKeyMismatch.lldbo.pn +
        ' (' + res.projectKeyMismatch.lldbo.spn + ') but the loaded Item Master is for ' +
        res.projectKeyMismatch.im.pn + ' (' + res.projectKeyMismatch.im.spn + ') — are these the right files for the same project?';
      warnEl.classList.remove('hidden');
    } else {
      warnEl.classList.add('hidden');
      warnEl.textContent = '';
    }

    const summary = $('lldbo-summary');
    summary.innerHTML = '';
    if (res) {
      summary.appendChild(lldboCardFor('long-lead parts (with Part No)', res.totalLldboItems));
      summary.appendChild(lldboCardFor('missing from Item Master', res.missingFromIm.length, res.missingFromIm.length ? 'red' : 'pass'));
      summary.appendChild(lldboCardFor('quantity mismatches', res.qtyMismatches.length, res.qtyMismatches.length ? 'amber' : 'pass'));
      if (res.noPartNumberCount) {
        summary.appendChild(lldboCardFor('not yet assigned a Part No', res.noPartNumberCount, 'na'));
      }
    }
    summary.appendChild(lldboCardFor(
      cand.crossChecked ? 'candidate long-lead parts missing from LLDBO' : 'candidate long-lead parts found (load LLDBO to verify)',
      cand.confident.length,
      cand.crossChecked ? (cand.confident.length ? 'red' : 'pass') : 'na'
    ));
    summary.appendChild(lldboCardFor('candidates needing manual review', cand.review.length, cand.review.length ? 'amber' : 'pass'));
    if (cand.crossChecked && cand.trackedCount) {
      summary.appendChild(lldboCardFor('candidates already tracked in LLDBO', cand.trackedCount, 'na'));
    }

    const sections = $('lldbo-sections');
    sections.innerHTML = '';
    if (res) {
      sections.appendChild(lldboSectionFor(
        'Missing from Item Master',
        'Long-lead parts released early to procurement that never made it into the Item Master — the part may quietly never get ordered through the normal channel.',
        LLDBO_MISSING_COLS, res.missingFromIm,
        '✓ Every long-lead part with a Part No is present in the Item Master.',
        'lldboMissing'
      ));
      sections.appendChild(lldboSectionFor(
        'Quantity mismatches',
        'The long-lead quantity (summed across all LLDBO rows for that part) should equal the Item Master\'s rolled-up total.',
        LLDBO_QTY_COLS, res.qtyMismatches,
        '✓ Quantities agree for every part found in both.',
        'lldboQty'
      ));
    }
    sections.appendChild(lldboSectionFor(
      'Candidate long-lead parts (motors / actuators / seals / nozzles)',
      cand.crossChecked
        ? 'Purchased-part-numbered (X-999-*) rows whose Title/Description match a long-lead keyword and are NOT in the loaded LLDBO file — a forgotten long-lead part that may never get ordered.'
        : 'Purchased-part-numbered (X-999-*) rows whose Title/Description match a long-lead keyword (motor, actuator, seal, gasket, nozzle...). Load an LLDBO file to check which of these are actually missing from it.',
      LLDBO_CANDIDATE_COLS, cand.confident,
      cand.crossChecked ? '✓ Every candidate long-lead part is already tracked in the LLDBO file.' : 'No candidate long-lead parts found by keyword in this Item Master.',
      'lldboCandidate'
    ));
    sections.appendChild(lldboSectionFor(
      'Candidate long-lead parts needing manual review',
      'Keyword matched, but the part number is not in the X-999-* purchased-part convention — could be a genuinely bought part numbered outside convention, or a manufactured part/assembly that merely mentions the keyword. Review individually; not counted as a finding on its own.',
      LLDBO_CANDIDATE_COLS, cand.review,
      'No parts need manual review.',
      null // deliberately no checkKey — a section-open-state id distinct from the confident section, and review-tier rows never claim findings-registry ownership
    ));
    renderDashboard();
  }

  // AOA rows for the LLDBO check export sheet. `cand` is present once the
  // Item Master is loaded; `res` stays null until an LLDBO file also loads.
  function lldboSheetRows(res, cand) {
    const rows = [['Long-Lead Parts (LLDBO) check', '']];
    if (res) {
      if (res.projectKeyMismatch) {
        rows.push(['⚠ Project key mismatch', 'LLDBO: ' + res.projectKeyMismatch.lldbo.pn + ' (' + res.projectKeyMismatch.lldbo.spn +
          ') vs Item Master: ' + res.projectKeyMismatch.im.pn + ' (' + res.projectKeyMismatch.im.spn + ')']);
      }
      rows.push([]);
      rows.push(['Long-lead parts (with Part No)', res.totalLldboItems]);
      rows.push(['Not yet assigned a Part No', res.noPartNumberCount]);
      rows.push(['Missing from Item Master', res.missingFromIm.length]);
      rows.push(['Quantity mismatches', res.qtyMismatches.length]);
      if (res.missingFromIm.length) {
        rows.push([]);
        rows.push(['Missing from Item Master', '', '']);
        rows.push(LLDBO_MISSING_COLS.map(function (c) { return c[1]; }));
        for (const r of res.missingFromIm) rows.push(LLDBO_MISSING_COLS.map(function (c) { return r[c[0]]; }));
      }
      if (res.qtyMismatches.length) {
        rows.push([]);
        rows.push(['Quantity mismatches', '', '', '']);
        rows.push(LLDBO_QTY_COLS.map(function (c) { return c[1]; }));
        for (const r of res.qtyMismatches) rows.push(LLDBO_QTY_COLS.map(function (c) { return r[c[0]]; }));
      }
    }
    if (cand) {
      rows.push([]);
      rows.push(['Candidate long-lead parts (motors/actuators/seals/nozzles)',
        cand.crossChecked ? cand.confident.length + ' missing from LLDBO' : cand.confident.length + ' found — not yet cross-checked against an LLDBO file']);
      if (cand.confident.length) {
        rows.push([]);
        rows.push(LLDBO_CANDIDATE_COLS.map(function (c) { return c[1]; }));
        for (const r of cand.confident) rows.push(LLDBO_CANDIDATE_COLS.map(function (c) { return r[c[0]]; }));
      }
      rows.push([]);
      rows.push(['Candidates needing manual review (keyword match, non-standard numbering)', cand.review.length]);
      if (cand.review.length) {
        rows.push([]);
        rows.push(LLDBO_CANDIDATE_COLS.map(function (c) { return c[1]; }));
        for (const r of cand.review) rows.push(LLDBO_CANDIDATE_COLS.map(function (c) { return r[c[0]]; }));
      }
    }
    return rows;
  }

  $('btn-lldbo-export').addEventListener('click', function () {
    if (!state.lldboCandidatesResult) return;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(lldboSheetRows(state.lldboResult, state.lldboCandidatesResult)), 'LLDBO check');
    XLSX.writeFile(wb, 'LLDBO-check' + projectKeySuffix() + '.xlsx');
  });

  /* ---------------- material: CAD vs Item Master ---------------- */

  function hideMaterialResults() {
    $('material-sections').innerHTML = '';
    renderDashboard();
  }

  function runMaterialCheck() {
    if (!state.im) { state.materialResult = null; hideMaterialResults(); return; }
    state.materialResult = BC.materialCompare.compareMaterial(state.cadSources, state.im);
    renderMaterialPanel();
    if (state.imQc) renderImQc(); // Check 6's card reflects state.materialResult too — see relatedWarningFor()
  }

  const MATERIAL_MISMATCH_COLS = [['number', 'Part Number'], ['title', 'Title']].concat(LOCATION_COLS).concat(
    [['imMaterial', 'Item Master Material'], ['cadMaterial', 'CAD Material']]);

  function boughtOutSection(boughtOut) {
    const box = document.createElement('div');
    box.className = 'qc-section';
    const head = document.createElement('div');
    head.className = 'qc-section-head';
    const mismatchCount = boughtOut.filter(function (b) { return b.mismatch; }).length;
    const missingCount = boughtOut.filter(function (b) { return b.missingMaterial; }).length;
    const boughtOutDesc = 'Reference listing of every purchased/catalog part — not counted toward the checks above. ' +
      (mismatchCount || missingCount ? mismatchCount + ' material mismatch(es), ' + missingCount + ' missing material — marked below.' : 'No material issues noted.');
    head.appendChild(buildQcTitleBox('Bought-Out Parts (7-999-*)', boughtOutDesc));
    const pill = document.createElement('span');
    pill.className = 'qc-pill na';
    pill.textContent = boughtOut.length + ' PARTS';
    head.appendChild(pill);
    box.appendChild(head);

    const body = document.createElement('div');
    body.className = 'qc-section-body'; // always starts collapsed — reference list, not an action item
    if (!boughtOut.length) {
      body.innerHTML = '<div class="empty-state">No purchased/catalog (7-999-*) parts found in this Item Master.</div>';
    } else {
      const table = document.createElement('table');
      table.className = 'results-table';
      const htr = document.createElement('tr');
      addTh(htr, 'Part Number'); addTh(htr, 'Title'); addTh(htr, 'Row #'); addTh(htr, 'Parent Number'); addTh(htr, 'Parent Title');
      addTh(htr, 'Item Master Material'); addTh(htr, 'CAD Material');
      table.appendChild(htr);
      for (const part of boughtOut) {
        const tr = document.createElement('tr');
        addTd(tr, part.number);
        addTd(tr, part.title);
        addTd(tr, part.sourceRow);
        addTd(tr, part.parentNumber);
        addTd(tr, part.parentTitle);
        const tdIm = document.createElement('td');
        tdIm.textContent = part.missingMaterial ? '(missing)' : part.imMaterial;
        if (part.missingMaterial) tdIm.className = 'mat-missing';
        tr.appendChild(tdIm);
        const tdCad = document.createElement('td');
        tdCad.textContent = part.cadMaterial || '';
        if (part.mismatch) tdCad.className = 'mat-mismatch';
        tr.appendChild(tdCad);
        table.appendChild(tr);
      }
      body.appendChild(table);
    }
    box.appendChild(body);

    // Singleton section (called once) — always starts collapsed regardless
    // of content, per the comment above, unless the user has toggled it.
    const id = 'boughtOut';
    if (state.sectionOpen.get(id)) body.classList.add('open');
    head.addEventListener('click', function () { state.sectionOpen.set(id, body.classList.toggle('open')); });
    return box;
  }

  function renderMaterialPanel() {
    const res = state.materialResult;
    const sections = $('material-sections');
    sections.innerHTML = '';
    if (!res) return;
    rebuildFindings(); // rows below are decorated against it

    if (!isCheckEnabled('material')) {
      sections.appendChild(skippedSection('Material: CAD vs Item Master'));
    } else if (res.applicable) {
      sections.appendChild(lldboSectionFor(
        'Material: CAD vs Item Master',
        'Compares CAD material (' + (res.cadSourceFileName ? '"' + res.cadSourceFileName + '"' : 'the flat Vault export') +
          ') against the Item Master, for manufactured (non-purchased) parts. Naming-convention differences (e.g. "1.4301" vs "AISI 304") are normalized and not flagged; a genuine grade difference (e.g. 304 vs 304L) is.',
        MATERIAL_MISMATCH_COLS, res.mismatches,
        '✓ CAD and Item Master materials agree for every shared manufactured part.',
        'material'
      ));
    } else {
      const note = document.createElement('div');
      note.className = 'qc-section';
      const head = document.createElement('div');
      head.className = 'qc-section-head';
      head.innerHTML = '<div class="qc-section-title">Material: CAD vs Item Master</div><span class="qc-pill na">N/A</span>';
      note.appendChild(head);
      const body = document.createElement('div');
      body.className = 'qc-section-body open';
      body.innerHTML = '<div class="empty-state">' + res.reason + '</div>';
      note.appendChild(body);
      sections.appendChild(note);
    }

    sections.appendChild(boughtOutSection(res.boughtOut));
    renderDashboard();
  }

  function materialSheetRows(res) {
    const rows = [['Material: CAD vs Item Master', '']];
    if (!res.applicable) {
      rows.push(['N/A', res.reason]);
    } else {
      rows.push(['Mismatches (manufactured parts)', res.mismatches.length]);
      if (res.mismatches.length) {
        rows.push([]);
        rows.push(MATERIAL_MISMATCH_COLS.map(function (c) { return c[1]; }));
        for (const m of res.mismatches) rows.push(MATERIAL_MISMATCH_COLS.map(function (c) { return m[c[0]]; }));
      }
    }
    rows.push([]);
    rows.push(['Bought-Out Parts (7-999-*)', res.boughtOut.length + ' parts — reference listing, not counted as findings']);
    rows.push(['Part Number', 'Title', 'Row #', 'Parent Number', 'Parent Title', 'Item Master Material', 'CAD Material', 'Note']);
    for (const b of res.boughtOut) {
      rows.push([b.number, b.title, b.sourceRow || '', b.parentNumber || '', b.parentTitle || '',
        b.missingMaterial ? '(missing)' : b.imMaterial, b.cadMaterial || '',
        b.mismatch ? 'mismatch' : (b.missingMaterial ? 'missing material' : '')]);
    }
    return rows;
  }

  /* ---------------- revision: CAD vs Item Master ---------------- */

  function hideRevisionResults() {
    $('revision-sections').innerHTML = '';
    renderDashboard();
  }

  // Combines the file-based Ignore List with the in-session "ignore
  // bought-out parts" toggle below into the single isIgnored(pn, checkKey)
  // predicate compareRevision() expects — same shape runCompare() builds
  // for compareAll(), just with one extra source.
  function runRevisionCheck() {
    if (!state.im) { state.revisionResult = null; hideRevisionResults(); return; }
    const ignoreIdx = BC.ignoreListCompare.buildIgnoreIndex(state.ignoreList);
    const isIgnored = function (pn, checkKey) {
      if (checkKey === 'revision' && state.ignoreBoughtOutRevision && BC.imQc.PURCHASED_PART_RE.test(pn)) return true;
      return ignoreIdx.isIgnored(pn, checkKey);
    };
    state.revisionResult = BC.revisionCompare.compareRevision(state.cadSources, state.im, { isIgnored: isIgnored });
    renderRevisionPanel();
    renderResults(); // the main summary row's Revision mismatches card depends on this
    if (state.imQc) renderImQc(); // Check 7's card reflects state.revisionResult too — see relatedWarningFor()
  }

  const REVISION_MISMATCH_COLS = [['number', 'Part Number'], ['title', 'Title']].concat(LOCATION_COLS).concat(
    [['imRevision', 'Item Master Revision'], ['cadRevision', 'CAD Revision'], ['staleDesign', 'Stale Design?']]);

  // Quick in-session toggle for bought-out/purchased parts (X-999-*, see
  // js/imqc.js's PURCHASED_PART_RE) — these often carry a vendor's own
  // revision scheme rather than this org's, so a mismatch here is frequently
  // expected noise. Separate from the file-based Ignore List: no upload
  // needed, just a click, and it re-runs through the same isIgnored
  // predicate so it shows up in "Ignored findings" like any other suppression.
  function revisionBoughtOutToggleEl() {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'attention-toggle';
    btn.textContent = state.ignoreBoughtOutRevision
      ? 'Bought-out parts (X-999-*) hidden from this list — click to show them'
      : 'Ignore bought-out parts (X-999-*) from this list';
    btn.addEventListener('click', function () {
      state.ignoreBoughtOutRevision = !state.ignoreBoughtOutRevision;
      runRevisionCheck();
    });
    return btn;
  }

  function renderRevisionPanel() {
    const res = state.revisionResult;
    const sections = $('revision-sections');
    sections.innerHTML = '';
    if (!res) return;
    rebuildFindings(); // rows below are decorated against it

    if (!isCheckEnabled('revision')) {
      sections.appendChild(skippedSection('Revision: CAD vs Item Master'));
    } else if (res.applicable) {
      sections.appendChild(revisionBoughtOutToggleEl());
      sections.appendChild(lldboSectionFor(
        'Revision: CAD vs Item Master',
        'Compares CAD revision (' + (res.cadSourceFileName ? '"' + res.cadSourceFileName + '"' : 'the loaded CAD source') +
          ') against the Item Master, for every shared part. Revision values are compared directly (not normalized) — any difference is a genuine finding. ' +
          'Rows shaded amber have an Item Master revision older than CAD\'s — the design may have moved on since this part was released. Only asserted when both sides are plain numbers; purchased-part placeholder text (e.g. "ANY"/"NONE") is never guessed at.',
        REVISION_MISMATCH_COLS, res.mismatches,
        '✓ CAD and Item Master revisions agree for every shared part.',
        'revision', false,
        function (row) { return row.imBehindCad ? 'row-revision-behind' : null; }
      ));
    } else {
      const note = document.createElement('div');
      note.className = 'qc-section';
      const head = document.createElement('div');
      head.className = 'qc-section-head';
      head.innerHTML = '<div class="qc-section-title">Revision: CAD vs Item Master</div><span class="qc-pill na">N/A</span>';
      note.appendChild(head);
      const body = document.createElement('div');
      body.className = 'qc-section-body open';
      body.innerHTML = '<div class="empty-state">' + res.reason + '</div>';
      note.appendChild(body);
      sections.appendChild(note);
    }
    renderDashboard();
  }

  function revisionSheetRows(res) {
    const rows = [['Revision: CAD vs Item Master', '']];
    if (!res.applicable) {
      rows.push(['N/A', res.reason]);
    } else {
      rows.push(['Mismatches', res.mismatches.length]);
      if (res.mismatches.length) {
        rows.push([]);
        rows.push(REVISION_MISMATCH_COLS.map(function (c) { return c[1]; }));
        for (const m of res.mismatches) rows.push(REVISION_MISMATCH_COLS.map(function (c) { return m[c[0]]; }));
      }
    }
    return rows;
  }

  /* ---------------- description: CAD vs Item Master ---------------- */

  function hideTitleDescResults() {
    $('titledesc-sections').innerHTML = '';
    renderDashboard();
  }

  function runTitleDescCheck() {
    if (!state.im) { state.titleDescResult = null; hideTitleDescResults(); return; }
    state.titleDescResult = BC.titleDescCompare.compareTitleDescription(state.cadSources, state.im);
    renderTitleDescPanel();
  }

  const TITLEDESC_MISMATCH_COLS = [['number', 'Part Number'], ['title', 'Title']].concat(LOCATION_COLS).concat(
    [['imDescription', 'Item Master Description'], ['cadDescription', 'CAD Description']]);

  // Parts where the CAD description is the Item Master description with
  // extra text appended at the end (most often a material grade) — not a
  // mismatch, since nothing in the shared text changed (see
  // isAutoMatchedAppend in js/titledesc-compare.js). Modeled on
  // boughtOutSection() above, not lldboSectionFor(): that helper colors its
  // pill red/green off fail.length, which would wrongly imply these rows
  // are still a problem — this is a reference list, same as Bought-Out Parts.
  function titleDescAutoMatchSection(autoMatched) {
    const box = document.createElement('div');
    box.className = 'qc-section';
    const head = document.createElement('div');
    head.className = 'qc-section-head';
    head.appendChild(buildQcTitleBox('Description: CAD adds extra text (not flagged)',
      'Parts where the CAD description is the Item Master description with extra text appended at the ' +
      'end — most often a material grade, e.g. "GSF PRO 180" → "GSF PRO 180 AISI 304" — and nothing in the ' +
      'shared text changed. Listed here for transparency, not counted as a finding.'));
    const pill = document.createElement('span');
    pill.className = 'qc-pill na';
    pill.textContent = autoMatched.length + ' PARTS';
    head.appendChild(pill);
    box.appendChild(head);

    const body = document.createElement('div');
    body.className = 'qc-section-body'; // always starts collapsed — reference list, not an action item
    if (!autoMatched.length) {
      body.innerHTML = '<div class="empty-state">No parts matched this pattern.</div>';
    } else {
      const table = document.createElement('table');
      table.className = 'results-table';
      const htr = document.createElement('tr');
      for (const [, label] of TITLEDESC_MISMATCH_COLS) addTh(htr, label);
      table.appendChild(htr);
      for (const row of autoMatched) {
        const tr = document.createElement('tr');
        for (const col of TITLEDESC_MISMATCH_COLS) addTd(tr, row[col[0]]);
        table.appendChild(tr);
      }
      body.appendChild(table);
    }
    box.appendChild(body);

    // Singleton section (called once) — always starts collapsed regardless
    // of content, per the comment above, unless the user has toggled it.
    const id = 'titleDescAutoMatch';
    if (state.sectionOpen.get(id)) body.classList.add('open');
    head.addEventListener('click', function () { state.sectionOpen.set(id, body.classList.toggle('open')); });
    return box;
  }

  function renderTitleDescPanel() {
    const res = state.titleDescResult;
    const sections = $('titledesc-sections');
    sections.innerHTML = '';
    if (!res) return;
    rebuildFindings(); // rows below are decorated against it

    if (!isCheckEnabled('titleDesc')) {
      sections.appendChild(skippedSection('Description: CAD vs Item Master'));
    } else if (res.applicable) {
      sections.appendChild(lldboSectionFor(
        'Description: CAD vs Item Master',
        'Compares the CAD Description (' + (res.cadSourceFileName ? '"' + res.cadSourceFileName + '"' : 'the loaded CAD source') +
          ') against the Item Master Description, for every shared part. Case and spacing are ignored; ' +
          'punctuation and digits are not, so a changed dimension or thread size is still reported. ' +
          'Purchased "X-999-" parts and parts procured at another site (numbers not starting "7-") are excluded.',
        TITLEDESC_MISMATCH_COLS, res.mismatches,
        '✓ CAD and Item Master descriptions agree for every shared part.',
        'titleDesc', true
      ));
      sections.appendChild(titleDescAutoMatchSection(res.autoMatched));
    } else {
      const note = document.createElement('div');
      note.className = 'qc-section';
      const head = document.createElement('div');
      head.className = 'qc-section-head';
      head.innerHTML = '<div class="qc-section-title">Description: CAD vs Item Master</div><span class="qc-pill na">N/A</span>';
      note.appendChild(head);
      const body = document.createElement('div');
      body.className = 'qc-section-body open';
      body.innerHTML = '<div class="empty-state">' + res.reason + '</div>';
      note.appendChild(body);
      sections.appendChild(note);
    }
    renderDashboard();
  }

  function virtualSheetRows(res) {
    const rows = [['Virtual parts (no CAD file behind them)', '']];
    if (!res.applicable) {
      rows.push(['N/A', res.reason]);
      return rows;
    }
    rows.push(['Confirmed (Thumbnail column says "(NULL)")', res.confirmed.length]);
    rows.push(['Suspected (inferred, no Thumbnail column)', res.suspected.length]);
    if (res.reason) rows.push(['Note', res.reason]);
    const all = res.confirmed.concat(res.suspected);
    if (all.length) {
      rows.push([]);
      rows.push(['Confidence', 'Part Number', 'Title', 'Child count', 'Row #', 'Child Number', 'Child Title', 'Child Row #']);
      for (const v of all) {
        rows.push([v.confidence, v.number, v.title, v.childCount, v.sourceRow, '', '', '']);
        for (const c of v.children || []) {
          rows.push(['', '', '', '', '', c.number, c.title, c.sourceRow]);
        }
      }
    }
    return rows;
  }

  function titleDescSheetRows(res) {
    const rows = [['Description: CAD vs Item Master', '']];
    if (!res.applicable) {
      rows.push(['N/A', res.reason]);
    } else {
      rows.push(['Mismatches', res.mismatches.length]);
      rows.push(['Parts compared', res.eligibleCount]);
      if (res.mismatches.length) {
        rows.push([]);
        rows.push(TITLEDESC_MISMATCH_COLS.map(function (c) { return c[1]; }));
        for (const m of res.mismatches) rows.push(TITLEDESC_MISMATCH_COLS.map(function (c) { return m[c[0]]; }));
      }
      rows.push([]);
      rows.push(['Description: CAD adds extra text (not flagged)', (res.autoMatched || []).length + ' parts — reference listing, not counted as a finding']);
      if (res.autoMatched && res.autoMatched.length) {
        rows.push(TITLEDESC_MISMATCH_COLS.map(function (c) { return c[1]; }));
        for (const m of res.autoMatched) rows.push(TITLEDESC_MISMATCH_COLS.map(function (c) { return m[c[0]]; }));
      }
    }
    return rows;
  }

  /* ---------------- compare & render ---------------- */

  function updateCompareButton() {
    $('btn-compare').disabled = !(state.cadSources.length && state.im);
  }

  // Runs the CAD-vs-Item-Master comparison from current state and renders
  // it. Shared by the manual "Compare BOMs" button and the folder
  // auto-compare path. Returns true if a comparison was actually run.
  function runCompare() {
    if (!state.cadSources.length || !state.im) return false;
    clearNotices();
    // Virtual parts are detected first: their anchors let the "In Item Master
    // only" rollup group a virtual subassembly's orphaned children under it,
    // instead of scattering them as one unexplained finding each.
    state.virtualResult = BC.virtualParts.detectVirtualParts(state.cadSources, state.im, BC.indexItemMaster);
    const ignoreIdx = BC.ignoreListCompare.buildIgnoreIndex(state.ignoreList);
    state.result = BC.compareAll(state.cadSources, state.im,
      { virtualAnchorRows: state.virtualResult.anchorRows, isIgnored: ignoreIdx.isIgnored });
    const res = state.result;
    if (!res.hasQty) {
      notice('warn', 'None of the CAD BOM sources has a quantity column, so quantity comparison is disabled. ' +
        'Add the Inventor BOM export (.xlsx with a QTY column) to enable it.');
    }
    if (!res.hasLevels) {
      notice('info', 'The CAD BOM has no level information — grouping of reference-assembly children is inferred ' +
        'from the export’s depth-first order and the Item Master hierarchy.');
    }
    if (res.referenceRoots === null) {
      notice('info', 'Upload the Vault multi-level BOM PDF and the Inventor BOM export together to also get the ' +
        '“Reference items” review list (everything currently flagged BOM Reference in the model).');
    }
    for (const w of res.warnings || []) notice('info', w);
    $('results').classList.remove('hidden');
    renderResults();
    return true;
  }

  $('btn-compare').addEventListener('click', function () {
    if (runCompare()) $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  // "Clear files & start over" — resets everything about the CURRENT set of
  // loaded files and their comparison results, but deliberately leaves the
  // Ignore List (state.ignoreList + its dropzone) untouched: runCompare()/
  // runRevisionCheck() already rebuild the ignore index fresh on every run,
  // so a previously-loaded Ignore List keeps applying to the next comparison
  // with no extra code. Also untouched: state.priorityMap, state.reviewedSet,
  // and state.sectionOpen/traceOpen (the user's own triage tags and collapse
  // preferences, none of which are tied to any particular set of files).
  function clearAll() {
    state.cadSources = [];
    state.im = null; state.imQc = null; state.result = null;
    state.lldbo = null; state.lldboResult = null; state.lldboCandidatesResult = null;
    state.materialResult = null; state.revisionResult = null; state.titleDescResult = null;
    state.virtualResult = null; state.findings = null;
    state.filter = ''; state.showGrouped = false; state.ignoreBoughtOutRevision = false;
    state.folderHandle = null;

    $('file-cad').value = ''; $('file-im').value = ''; $('file-lldbo').value = '';
    renderCadSources('', '');
    setImStatus('', [], '');
    setLldboStatus('', [], '');
    $('filter').value = '';
    $('folder-status').textContent = '';
    $('folder-status').className = 'folder-status';
    switchTab('missing');

    hideImQc();            // also clears qc-summary/qc-sections/qc-callout
    hideLldboResults();    // clears LLDBO results only — #lldbo-panel's dropzone stays usable
    hideMaterialResults();
    hideRevisionResults();
    hideTitleDescResults();
    $('results').classList.add('hidden');
    clearNotices();
    updateCompareButton();
  }

  $('btn-clear-all').addEventListener('click', clearAll);

  function fmtQty(v) {
    if (v === null || v === undefined) return '';
    return Math.abs(v - Math.round(v)) < 1e-9 ? String(Math.round(v)) : String(Math.round(v * 1000) / 1000);
  }

  function matchesFilter(strings) {
    if (!state.filter) return true;
    const f = state.filter.toLowerCase();
    return strings.some(function (s) { return s && String(s).toLowerCase().indexOf(f) !== -1; });
  }

  function renderResults() {
    const res = state.result;
    if (!res) { renderDashboard(); return; }
    rebuildFindings(); // the qty / in-Item-Master-only rows are decorated against it
    setSourceLine($('results-source'));

    // summary cards
    const summary = $('summary');
    summary.innerHTML = '';
    const cards = [
      { num: res.cadUniqueCount, lbl: 'unique parts in CAD BOM' },
      { num: res.imUniqueCount, lbl: 'unique parts in Item Master' },
      { num: res.actionableCount, lbl: 'findings needing action', cls: res.actionableCount ? 'red' : '' },
      { num: res.missingTotal, lbl: 'missing incl. grouped children' },
      { num: res.qtyMismatches === null ? '—' : res.qtyMismatches.length, lbl: 'quantity mismatches', cls: res.qtyMismatches && res.qtyMismatches.length ? 'amber' : '' },
      { num: res.imOnly.length, lbl: 'in Item Master only' },
    ];
    const rev = state.revisionResult;
    cards.splice(5, 0, {
      num: rev && rev.applicable ? rev.mismatches.length : '—',
      lbl: 'revision mismatches',
      cls: rev && rev.applicable && rev.mismatches.length ? 'amber' : '',
    });
    if (res.referenceRoots !== null) {
      cards.splice(4, 0, { num: res.referenceTotal, lbl: 'reference components in CAD' });
    }
    for (const c of cards) {
      const el = document.createElement('div');
      el.className = 'card' + (c.cls ? ' ' + c.cls : '');
      el.innerHTML = '<div class="num"></div><div class="lbl"></div>';
      el.querySelector('.num').textContent = c.num;
      el.querySelector('.lbl').textContent = c.lbl;
      summary.appendChild(el);
    }

    $('count-missing').textContent = res.actionableCount;
    $('count-ref').textContent = res.referenceRoots === null ? '' : res.referenceRoots.length;
    $('count-qty').textContent = res.qtyMismatches === null ? '—' : res.qtyMismatches.length;
    $('count-cascade').textContent = res.qtyCascades.applicable ? res.qtyCascades.roots.length : '—';
    const virt = state.virtualResult;
    $('count-virtual').textContent = virt && virt.applicable
      ? (virt.confirmed.length + virt.suspected.length) : '—';
    $('count-imonly').textContent = res.imOnlyRoots ? res.imOnlyRoots.length : res.imOnly.length;

    // Tab visibility: a tab is hidden either for the existing data reasons
    // (only "Reference items" has one today — needs both CAD sources) or
    // because its check is off in "Checks to run ▾". Re-evaluated on every
    // render (not just after a fresh Compare click) so toggling a check
    // updates the tab bar live. Falls back to the first still-visible tab
    // if the currently active one just became hidden.
    const tabAvailable = {
      missing: isCheckEnabled('missing'),
      ref: isCheckEnabled('reference') && res.referenceRoots !== null,
      qty: isCheckEnabled('qty'),
      cascade: isCheckEnabled('qtyCascade'),
      virtual: isCheckEnabled('virtual'),
      imonly: isCheckEnabled('imOnly'),
    };
    for (const name of Object.keys(tabAvailable)) $('tab-' + name).classList.toggle('hidden', !tabAvailable[name]);
    if (!tabAvailable[state.activeTab]) {
      const fallback = ['missing', 'qty', 'cascade', 'virtual', 'imonly', 'ref'].find(function (n) { return tabAvailable[n]; });
      if (fallback) switchTab(fallback);
    }

    renderMissingTab();
    renderRefTab();
    renderQtyTab();
    renderCascadeTab();
    renderVirtualTab();
    renderImOnlyTab();
    // The Item Master / material / revision sections were rendered when their
    // files loaded — before any comparison existed — so their rows were
    // decorated against a registry that didn't yet know about missing/qty
    // findings. Re-render them now that it does. Guarded because these
    // renderers each call renderDashboard() on the way out.
    if (!reRendering) {
      reRendering = true;
      try {
        if (state.imQc) renderImQc();
        if (state.materialResult) renderMaterialPanel();
        if (state.revisionResult) renderRevisionPanel();
        if (state.titleDescResult) renderTitleDescPanel();
        if (state.lldboResult) renderLldboPanel();
      } finally {
        reRendering = false;
      }
    }
    renderDashboard();
  }
  let reRendering = false;

  /* ----- tree tables (missing + reference tabs) ----- */

  // opts: { rowClass, groupedBadge, extraBadge(node, depth) -> {text, cls}|null }
  function renderTree(pane, roots, opts) {
    const res = state.result;
    pane.innerHTML = '';
    const table = document.createElement('table');
    table.className = 'results-table';
    const cols = visibleCols();
    const thead = document.createElement('thead');
    const htr = document.createElement('tr');
    addTh(htr, '');
    addTh(htr, 'Part Number');
    if (cols.title) addTh(htr, 'Title');
    if (cols.description) addTh(htr, 'Description');
    if (cols.type) addTh(htr, 'Type');
    if (cols.qty && res.hasQty) addTh(htr, 'Qty (CAD)');
    if (cols.file) addTh(htr, 'File');
    if (cols.material) addTh(htr, 'Material');
    if (cols.row) addTh(htr, 'CAD row');
    thead.appendChild(htr);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    const colSpan = htr.children.length;

    let uid = 0;
    const renderNode = function (node, depth, parentId) {
      const id = opts.idPrefix + (uid++);
      const it = node.item;
      if (state.filter && !nodeTreeMatches(node)) return;

      const tr = document.createElement('tr');
      tr.className = depth === 0 ? opts.rowClass : 'row-child';
      tr.dataset.id = id;
      if (parentId) {
        tr.dataset.parent = parentId;
        // children start hidden unless a filter is active (then show match paths)
        if (!state.filter) tr.classList.add('hidden-row');
      }

      const tdExp = document.createElement('td');
      if (node.children.length) {
        const exp = document.createElement('span');
        exp.className = 'expander';
        exp.textContent = state.filter ? '▾' : '▸';
        exp.dataset.for = id;
        tdExp.appendChild(exp);
      }
      tr.appendChild(tdExp);

      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      const indent = document.createElement('span');
      indent.className = 'indent';
      indent.style.width = (depth * 22) + 'px';
      tdNum.appendChild(indent);
      tdNum.appendChild(document.createTextNode(it.number + ' '));
      if (depth === 0 && it.isAssembly && node.children.length) {
        const b = document.createElement('span');
        b.className = 'badge asm';
        b.textContent = 'assembly';
        tdNum.appendChild(b);
        tdNum.appendChild(document.createTextNode(' '));
      }
      if (depth === 0 && node.children.length) {
        const g = document.createElement('span');
        g.className = 'badge grouped';
        g.textContent = '+' + BC.countDescendants(node) + ' ' + opts.groupedBadge;
        tdNum.appendChild(g);
        tdNum.appendChild(document.createTextNode(' '));
      }
      const extra = opts.extraBadge ? opts.extraBadge(node, depth) : null;
      if (extra) {
        const x = document.createElement('span');
        x.className = 'badge ' + extra.cls;
        x.textContent = extra.text;
        tdNum.appendChild(x);
      }
      tdNum.appendChild(document.createTextNode(' '));
      tdNum.appendChild(traceToggleEl(it.number));
      tr.appendChild(tdNum);

      if (cols.title) addTd(tr, it.title);
      if (cols.description) addTd(tr, it.description);
      if (cols.type) addTd(tr, it.isAssembly === null ? '' : (it.isAssembly ? 'Assembly' : 'Part'));
      if (cols.qty && res.hasQty) addTd(tr, fmtQty(it.qty));
      if (cols.file) addTd(tr, it.file);
      if (cols.material) addTd(tr, it.material || '');
      if (cols.row) addTd(tr, String(it.sourceRow || ''));

      tbody.appendChild(tr);
      tbody.appendChild(traceDetailRow(it.number, colSpan));
      for (const c of node.children) renderNode(c, depth + 1, id);
    };

    for (const rootNode of roots) renderNode(rootNode, 0, null);
    table.appendChild(tbody);
    pane.appendChild(table);

    wireExpanders(tbody);
    wireTraceToggles(tbody);
  }

  function renderMissingTab() {
    const pane = $('pane-missing');
    const res = state.result;
    if (!res.missingRoots.length) {
      pane.innerHTML = '<div class="empty-state">🎉 Every CAD part number exists in the Item Master.</div>';
      return;
    }
    renderTree(pane, res.missingRoots, {
      idPrefix: 'mn',
      rowClass: 'row-missing',
      groupedBadge: 'children grouped (reference subtree)',
    });
  }

  function renderRefTab() {
    const pane = $('pane-ref');
    const res = state.result;
    pane.innerHTML = '';
    if (res.referenceRoots === null) return;
    const intro = document.createElement('p');
    intro.className = 'pane-intro';
    intro.textContent = 'Components that are in the CAD structure but not in the Inventor BOM export — i.e. everything ' +
      'currently flagged “BOM Reference” in the model. Review each one: should it really be reference?';
    pane.appendChild(intro);
    if (!res.referenceRoots.length) {
      const d = document.createElement('div');
      d.className = 'empty-state';
      d.textContent = 'No reference components — every CAD component is part of the intended BOM.';
      pane.appendChild(d);
      return;
    }
    const holder = document.createElement('div');
    pane.appendChild(holder);
    renderTree(holder, res.referenceRoots, {
      idPrefix: 'rn',
      rowClass: 'row-ref',
      groupedBadge: 'children (reference subtree)',
      extraBadge: function (node, depth) {
        if (depth !== 0) return null;
        return node.inItemMaster
          ? { text: 'in Item Master ✓', cls: 'ok' }
          : { text: 'not in Item Master', cls: 'missing' };
      },
    });
  }

  function hideDescendants(tbody, id) {
    tbody.querySelectorAll('tr[data-parent="' + id + '"]').forEach(function (tr) {
      tr.classList.add('hidden-row');
      const exp = tr.querySelector('.expander');
      if (exp) exp.textContent = '▸';
      hideDescendants(tbody, tr.dataset.id);
    });
  }

  // Expand/collapse for any container holding rows tagged with
  // data-id / data-parent and an .expander cell (renderTree's tbody, and the
  // "In Item Master only" tree).
  function wireExpanders(container) {
    container.addEventListener('click', function (ev) {
      const exp = ev.target.closest('.expander');
      if (!exp) return;
      const id = exp.dataset.for;
      if (exp.textContent === '▾') {
        exp.textContent = '▸';
        hideDescendants(container, id);
      } else {
        exp.textContent = '▾';
        container.querySelectorAll('tr[data-parent="' + id + '"]').forEach(function (tr) {
          tr.classList.remove('hidden-row');
        });
      }
    });
  }

  function nodeTreeMatches(node) {
    if (matchesFilter([node.item.number, node.item.title, node.item.description])) return true;
    return node.children.some(nodeTreeMatches);
  }

  /* ----- cross-source traceability -----
   * Lets a flagged row in any results list answer "which uploaded file(s)
   * was this part found in or absent from, and what row/page do I check by
   * hand" without the user reopening the CAD source(s) or Item Master
   * themselves. Built from compareAll()'s cadOccurrences/imOccurrences
   * (js/compare.js) — an absent map entry for a PN already means "not found
   * in that source", the same signal every other check is built from.
   */

  // Small link-style control that reveals a hidden detail row with full
  // cross-source provenance for one part number. Kept structurally separate
  // from wireExpanders()'s tree-row disclosure (different dataset
  // attributes), since a single row can carry both a children-expander and
  // a trace toggle at once.
  function traceToggleEl(pn) {
    const span = document.createElement('span');
    span.className = 'trace-toggle';
    span.textContent = 'trace';
    span.dataset.traceFor = BC.normNumber(pn);
    return span;
  }

  function occurrenceBreakdownEl(pn) {
    const res = state.result;
    const box = document.createElement('div');
    // The Description-vs-CAD panel can render before "Compare BOMs" is ever
    // clicked (it reruns as soon as a CAD file loads) — cadOccurrences/
    // imOccurrences only exist once compareAll() has actually run.
    if (!res) {
      const note = document.createElement('div');
      note.className = 'trace-none';
      note.textContent = 'Run "Compare BOMs" to see where this part is found.';
      box.appendChild(note);
      return box;
    }
    const key = BC.normNumber(pn);
    box.appendChild(breakdownTable('Found in CAD BOM', res.cadOccurrences.get(key), 'Not found in the CAD BOM.'));
    box.appendChild(breakdownTable('Found in Item Master', res.imOccurrences.get(key), 'Not found in the Item Master.'));
    return box;
  }

  function traceDetailRow(pn, colSpan) {
    const tr = document.createElement('tr');
    tr.className = 'hidden-row';
    tr.dataset.traceId = BC.normNumber(pn);
    const td = document.createElement('td');
    td.colSpan = colSpan;
    td.appendChild(occurrenceBreakdownEl(pn));
    tr.appendChild(td);
    return tr;
  }

  // Wires every '.trace-toggle' inside container to its sibling detail row
  // (matched via data-trace-id, compared here in JS rather than built into a
  // querySelector string, since a part number could contain characters that
  // would need escaping there). Persists open/closed per part number in
  // state.traceOpen so it survives this container being rebuilt on the next
  // render — unlike state.sectionOpen, this Map is keyed by part number, not
  // a fixed section id, and defaults to closed: a drill-down aid, not
  // primary content.
  function wireTraceToggles(container) {
    const rowByKey = new Map();
    container.querySelectorAll('tr[data-trace-id]').forEach(function (tr) { rowByKey.set(tr.dataset.traceId, tr); });
    container.querySelectorAll('.trace-toggle').forEach(function (el) {
      const key = el.dataset.traceFor;
      const row = rowByKey.get(key);
      if (!row) return;
      if (state.traceOpen.get(key)) row.classList.remove('hidden-row');
      el.addEventListener('click', function () {
        const nowHidden = row.classList.toggle('hidden-row');
        state.traceOpen.set(key, !nowHidden);
      });
    });
  }

  /* ----- qty tab ----- */

  function renderQtyTab() {
    const pane = $('pane-qty');
    const res = state.result;
    pane.innerHTML = '';
    if (res.qtyMismatches === null) {
      pane.innerHTML = '<div class="empty-state">Quantity comparison is unavailable: the CAD BOM source has no quantity column.<br>' +
        'Upload the multi-level BOM PDF from Vault (with QTY visible) or a leveled export instead.</div>';
      return;
    }
    const rows = res.qtyMismatches.filter(function (m) {
      return matchesFilter([m.number, m.title, m.description]);
    });
    if (!rows.length) {
      pane.innerHTML = '<div class="empty-state">' + (state.filter ? 'No quantity mismatches match the filter.' : '✔ All shared parts have matching rolled-up quantities.') + '</div>';
      return;
    }
    const cols = visibleCols();
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, '');
    addTh(htr, 'Part Number');
    if (cols.title) addTh(htr, 'Title');
    if (cols.description) addTh(htr, 'Description');
    addTh(htr, 'Qty in CAD');
    addTh(htr, 'Qty in Item Master');
    addTh(htr, 'Difference');
    table.appendChild(htr);

    let uid = 0;
    for (const m of rows) {
      const id = 'qm' + (uid++);
      const tr = document.createElement('tr');
      tr.className = 'row-qty';
      const tdExp = document.createElement('td');
      const exp = document.createElement('span');
      exp.className = 'expander';
      exp.textContent = '▸';
      exp.dataset.for = id;
      tdExp.appendChild(exp);
      tr.appendChild(tdExp);
      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      tdNum.textContent = m.number;
      tr.appendChild(tdNum);
      decorateSecondary(tr, 'qty', m.number, tdNum);
      if (cols.title) addTd(tr, m.title);
      if (cols.description) addTd(tr, m.description);
      addTd(tr, fmtQty(m.cadQty));
      addTd(tr, fmtQty(m.imQty));
      const diff = m.cadQty - m.imQty;
      const tdDiff = document.createElement('td');
      tdDiff.className = 'qty-diff';
      tdDiff.textContent = (diff > 0 ? '+' : '') + fmtQty(diff);
      tr.appendChild(tdDiff);
      table.appendChild(tr);

      // breakdown row
      const br = document.createElement('tr');
      br.className = 'hidden-row';
      br.dataset.id = id;
      const td = document.createElement('td');
      td.colSpan = htr.children.length;
      td.appendChild(breakdownTable('Where it appears in the CAD BOM', m.cadBreakdown));
      td.appendChild(breakdownTable('Where it appears in the Item Master', m.imBreakdown));
      br.appendChild(td);
      table.appendChild(br);
    }
    pane.appendChild(table);
    table.addEventListener('click', function (ev) {
      const exp = ev.target.closest('.expander');
      if (!exp) return;
      const row = table.querySelector('tr[data-id="' + exp.dataset.for + '"]');
      const open = exp.textContent === '▾';
      exp.textContent = open ? '▸' : '▾';
      if (row) row.classList.toggle('hidden-row', open);
    });
  }

  // notFoundText (optional): shown instead of an empty table when entries is
  // empty/absent — the qty tab's two callers never hit this (a quantity
  // mismatch requires the part to be in both sources), but it's what makes
  // this function reusable for occurrenceBreakdownEl's "not found in X" case.
  function breakdownTable(caption, entries, notFoundText) {
    const box = document.createElement('div');
    box.className = 'qty-breakdown';
    const cap = document.createElement('div');
    cap.textContent = caption + ':';
    box.appendChild(cap);
    if (!entries || !entries.length) {
      const none = document.createElement('div');
      none.className = 'trace-none';
      none.textContent = notFoundText || 'None found.';
      box.appendChild(none);
      return box;
    }
    const t = document.createElement('table');
    const h = document.createElement('tr');
    ['Where', 'Parent assembly', 'Parent title', 'Qty per parent', 'Effective qty', 'File'].forEach(function (x) { addTh(h, x); });
    t.appendChild(h);
    for (const e of entries) {
      const tr = document.createElement('tr');
      addTd(tr, e.page ? 'PDF page ' + e.page + ' (row ' + (e.sourceRow || '') + ')' : (e.sourceRow ? 'row ' + e.sourceRow : ''));
      addTd(tr, e.parentNumber || '(top level)');
      addTd(tr, e.parentTitle || '');
      addTd(tr, fmtQty(e.qty));
      addTd(tr, fmtQty(e.effQty));
      addTd(tr, e.file || '');
      t.appendChild(tr);
    }
    box.appendChild(t);
    return box;
  }

  /* ----- quantity-cascade tab ----- */

  // A quantity cascade is a whole Item Master subtree released at one clean,
  // uniform ratio of its CAD-required quantity — one root-cause error that
  // would otherwise flood the plain "Quantity mismatches" tab with every
  // descendant it explains. Rendered as a tree over the same Item Master
  // hierarchy renderImOnlyTab uses, since qtyCascades.roots share that shape.
  function renderCascadeTab() {
    const pane = $('pane-cascade');
    const res = state.result;
    pane.innerHTML = '';
    if (!res.qtyCascades.applicable) {
      pane.innerHTML = '<div class="empty-state">Quantity-cascade detection is unavailable: it needs a leveled CAD ' +
        'source with quantities (the Inventor BOM export) — the same requirement as "Quantity mismatches".</div>';
      return;
    }
    const intro = document.createElement('p');
    intro.className = 'pane-intro';
    intro.textContent = 'A whole assembly whose direct children are ALL released at the same clean multiple of their ' +
      'CAD quantity (e.g. every child at exactly 2x) usually traces to one data-entry error, not many. Fix the one ' +
      'root cause below and most of the "Quantity mismatches" tab should clear with it.';
    pane.appendChild(intro);
    const roots = res.qtyCascades.roots;
    if (!roots.length) {
      const d = document.createElement('div');
      d.className = 'empty-state';
      d.textContent = 'No quantity cascades found — no assembly has its whole set of children released at one uniform ratio.';
      pane.appendChild(d);
      return;
    }
    const cols = visibleCols();
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, '');
    addTh(htr, 'Part Number');
    if (cols.title) addTh(htr, 'Title');
    if (cols.row) addTh(htr, 'Row #');
    table.appendChild(htr);

    let uid = 0;
    const renderNode = function (node, depth, parentId) {
      const r = node.item;
      if (state.filter && !imOnlyTreeMatches(node)) return;
      const id = 'ca' + (uid++);
      const tr = document.createElement('tr');
      tr.dataset.id = id;
      if (depth > 0) {
        tr.className = 'row-child';
        tr.dataset.parent = parentId;
        if (!state.filter) tr.classList.add('hidden-row');
      } else {
        tr.className = 'row-cascade';
      }
      const tdExp = document.createElement('td');
      if (node.children && node.children.length) {
        const exp = document.createElement('span');
        exp.className = 'expander';
        exp.textContent = state.filter ? '▾' : '▸';
        exp.dataset.for = id;
        tdExp.appendChild(exp);
      }
      tr.appendChild(tdExp);
      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      const indent = document.createElement('span');
      indent.className = 'indent';
      indent.style.width = (depth * 14) + 'px';
      tdNum.appendChild(indent);
      tdNum.appendChild(document.createTextNode(r.number));
      if (depth === 0) {
        const ratio = document.createElement('span');
        ratio.className = 'badge red';
        ratio.textContent = r.cascadeMismatchedChildCount + ' of ' + r.cascadeChildCount + ' children at ' + r.cascadeRatio + '×';
        tdNum.appendChild(document.createTextNode(' '));
        tdNum.appendChild(ratio);
      }
      if (node.children && node.children.length) {
        const badge = document.createElement('span');
        badge.className = 'badge grouped';
        badge.textContent = '+' + BC.countDescendants(node) + ' descendants grouped';
        tdNum.appendChild(document.createTextNode(' '));
        tdNum.appendChild(badge);
      }
      tr.appendChild(tdNum);
      if (depth === 0) decorateSecondary(tr, 'qtyCascade', r.number, tdNum);
      if (cols.title) addTd(tr, r.title);
      if (cols.row) addTd(tr, r.sourceRow || '');
      table.appendChild(tr);
      for (const c of node.children || []) renderNode(c, depth + 1, id);
    };
    for (const n of roots) renderNode(n, 0, null);
    pane.appendChild(table);
    wireExpanders(table);
  }

  /* ----- virtual parts tab ----- */

  // A virtual part is in the CAD BOM but has no CAD file behind it, and its
  // whole child BOM lives only in the Item Master. Every other check passes
  // it, so without this tab it is invisible and its children read as
  // unexplained "In Item Master only" rows.
  function renderVirtualTab() {
    const pane = $('pane-virtual');
    const res = state.virtualResult;
    pane.innerHTML = '';
    if (!res || !res.applicable) {
      pane.innerHTML = '<div class="empty-state">' +
        ((res && res.reason) || 'Load a CAD BOM and an Item Master to check for virtual parts.') + '</div>';
      return;
    }
    const intro = document.createElement('p');
    intro.className = 'pane-intro';
    intro.textContent = 'A virtual part is a BOM entry with no CAD file behind it. It passes every other check ' +
      '(it is in both BOMs), but the parts underneath it exist only in the Item Master, so they never get compared. ' +
      (res.hasThumbnailColumn
        ? 'Detected exactly, from the Inventor export’s Thumbnail column.'
        : res.reason);
    pane.appendChild(intro);

    const all = (res.confirmed || []).concat(res.suspected || []);
    if (!all.length) {
      const d = document.createElement('div');
      d.className = 'empty-state';
      d.textContent = 'No virtual parts found — every CAD assembly has its child parts in the CAD BOM too.';
      pane.appendChild(d);
      return;
    }
    const rows = all.filter(function (v) { return matchesFilter([v.number, v.title, v.description]); });
    if (!rows.length) {
      pane.innerHTML = '<div class="empty-state">Nothing matches the filter.</div>';
      return;
    }
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, '');
    addTh(htr, 'Part Number');
    addTh(htr, 'Title');
    addTh(htr, 'Children only in Item Master');
    addTh(htr, 'Row #');
    table.appendChild(htr);

    let uid = 0;
    for (const v of rows) {
      const id = 'vp' + (uid++);
      const tr = document.createElement('tr');
      tr.className = 'row-missing';
      const tdExp = document.createElement('td');
      const exp = document.createElement('span');
      exp.className = 'expander';
      exp.textContent = '▸';
      exp.dataset.for = id;
      tdExp.appendChild(exp);
      tr.appendChild(tdExp);
      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      tdNum.textContent = v.number + ' ';
      if (v.confidence === 'suspected') {
        const b = document.createElement('span');
        b.className = 'badge grouped';
        b.textContent = 'suspected';
        b.title = 'Inferred from BOM structure — include the Thumbnail column in the Inventor export to confirm.';
        tdNum.appendChild(b);
      }
      tr.appendChild(tdNum);
      decorateSecondary(tr, 'virtualPart', v.number, tdNum);
      addTd(tr, v.title);
      addTd(tr, String(v.childCount));
      addTd(tr, v.sourceRow || '');
      table.appendChild(tr);

      const br = document.createElement('tr');
      br.className = 'hidden-row';
      br.dataset.id = id;
      const td = document.createElement('td');
      td.colSpan = htr.children.length;
      const box = document.createElement('div');
      box.className = 'qty-breakdown';
      const cap = document.createElement('div');
      cap.textContent = 'These parts sit under it in the Item Master but are absent from the CAD BOM:';
      box.appendChild(cap);
      const t = document.createElement('table');
      const h = document.createElement('tr');
      ['Part Number', 'Title', 'Item Master Row #'].forEach(function (x) { addTh(h, x); });
      t.appendChild(h);
      for (const c of v.children || []) {
        const ctr = document.createElement('tr');
        addTd(ctr, c.number);
        addTd(ctr, c.title);
        addTd(ctr, c.sourceRow || '');
        t.appendChild(ctr);
      }
      box.appendChild(t);
      td.appendChild(box);
      // The table above is entirely Item-Master-side (this part's orphaned
      // children); it says nothing about where the virtual part ITSELF sits
      // in the CAD BOM — the "Row #" column on the visible row is also
      // Item-Master-side (js/virtual-parts.js reads it from imIndex, not the
      // CAD source). Append the CAD-side occurrence(s) so that's covered too.
      td.appendChild(occurrenceBreakdownEl(v.number));
      br.appendChild(td);
      table.appendChild(br);
    }
    pane.appendChild(table);
    table.addEventListener('click', function (ev) {
      const e = ev.target.closest('.expander');
      if (!e) return;
      const row = table.querySelector('tr[data-id="' + e.dataset.for + '"]');
      const open = e.textContent === '▾';
      e.textContent = open ? '▸' : '▾';
      if (row) row.classList.toggle('hidden-row', open);
    });
  }

  /* ----- IM-only tab ----- */

  // Rendered as a tree over the Item Master's own hierarchy: when a whole
  // subassembly is absent from the CAD BOM, every part under it is flagged, so
  // the children are grouped under the parent (collapsed) rather than listed as
  // hundreds of separate findings. Falls back to a flat list for results built
  // before imOnlyRoots existed, or when the export has no Row Order column.
  function renderImOnlyTab() {
    const pane = $('pane-imonly');
    const res = state.result;
    pane.innerHTML = '';
    const roots = res.imOnlyRoots || (res.imOnly || []).map(function (r) { return { item: r, children: [] }; });
    if (!roots.length) {
      pane.innerHTML = '<div class="empty-state">Every Item Master entry also appears in the CAD BOM.</div>';
      return;
    }
    const cols = visibleCols();
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, '');
    addTh(htr, 'Part Number');
    if (cols.title) addTh(htr, 'Title');
    if (cols.description) addTh(htr, 'Description');
    if (cols.qty) addTh(htr, 'Qty');
    addTh(htr, 'Row type');
    if (cols.row) addTh(htr, 'Row #');
    addTh(htr, 'Parent Number');
    addTh(htr, 'Parent Title');
    table.appendChild(htr);
    const colSpan = htr.children.length;

    let uid = 0;
    let anyShown = false;
    const renderNode = function (node, depth, parentId) {
      const r = node.item;
      if (state.filter && !imOnlyTreeMatches(node)) return;
      const id = 'io' + (uid++);
      anyShown = true;
      const tr = document.createElement('tr');
      tr.dataset.id = id;
      if (depth > 0) {
        tr.className = 'row-child';
        tr.dataset.parent = parentId;
        if (!state.filter) tr.classList.add('hidden-row');
      }
      const tdExp = document.createElement('td');
      if (node.children && node.children.length) {
        const exp = document.createElement('span');
        exp.className = 'expander';
        exp.textContent = state.filter ? '▾' : '▸';
        exp.dataset.for = id;
        tdExp.appendChild(exp);
      }
      tr.appendChild(tdExp);
      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      const indent = document.createElement('span');
      indent.className = 'indent';
      indent.style.width = (depth * 14) + 'px';
      tdNum.appendChild(indent);
      tdNum.appendChild(document.createTextNode(r.number));
      if (node.children && node.children.length) {
        const badge = document.createElement('span');
        badge.className = 'badge grouped';
        badge.textContent = '+' + BC.countDescendants(node) + ' children grouped';
        tdNum.appendChild(document.createTextNode(' '));
        tdNum.appendChild(badge);
      }
      tdNum.appendChild(document.createTextNode(' '));
      tdNum.appendChild(traceToggleEl(r.number));
      tr.appendChild(tdNum);
      if (depth === 0) decorateSecondary(tr, 'imOnly', r.number, tdNum);
      if (cols.title) addTd(tr, r.title);
      if (cols.description) addTd(tr, r.description);
      if (cols.qty) addTd(tr, fmtQty(r.qty));
      addTd(tr, r.rowType || '');
      if (cols.row) addTd(tr, r.sourceRow || '');
      addTd(tr, r.parentNumber || '');
      addTd(tr, r.parentTitle || '');
      table.appendChild(tr);
      table.appendChild(traceDetailRow(r.number, colSpan));
      for (const c of node.children || []) renderNode(c, depth + 1, id);
    };
    for (const n of roots) renderNode(n, 0, null);

    if (!anyShown) {
      pane.innerHTML = '<div class="empty-state">Nothing matches the filter.</div>';
      return;
    }
    pane.appendChild(table);
    wireExpanders(table);
    wireTraceToggles(table);
  }

  function imOnlyTreeMatches(node) {
    const r = node.item;
    if (matchesFilter([r.number, r.title, r.description])) return true;
    return (node.children || []).some(imOnlyTreeMatches);
  }

  /* ----- shared helpers ----- */

  function visibleCols() {
    const v = {};
    for (const c of COLUMNS) v[c.key] = colPrefs[c.key];
    return v;
  }

  function addTh(tr, text) {
    const th = document.createElement('th');
    th.textContent = text;
    tr.appendChild(th);
  }
  function addTd(tr, text) {
    const td = document.createElement('td');
    td.textContent = text === null || text === undefined ? '' : text;
    tr.appendChild(td);
  }

  /* ---------------- tabs, filter, columns dropdown ---------------- */

  function switchTab(name) {
    state.activeTab = name;
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === name); });
    ['missing', 'ref', 'qty', 'cascade', 'virtual', 'imonly'].forEach(function (n) {
      $('pane-' + n).classList.toggle('hidden', n !== name);
    });
  }

  $('tabs').addEventListener('click', function (ev) {
    const tab = ev.target.closest('.tab');
    if (!tab) return;
    switchTab(tab.dataset.tab);
  });

  $('filter').addEventListener('input', function () {
    state.filter = this.value.trim();
    renderResults();
  });

  $('columns-btn').addEventListener('click', function (ev) {
    ev.stopPropagation();
    $('columns-menu').classList.toggle('hidden');
  });
  $('checks-btn').addEventListener('click', function (ev) {
    ev.stopPropagation();
    $('checks-menu').classList.toggle('hidden');
  });
  document.addEventListener('click', function (ev) {
    if (!ev.target.closest('#columns-dropdown')) $('columns-menu').classList.add('hidden');
    if (!ev.target.closest('#checks-dropdown')) $('checks-menu').classList.add('hidden');
  });

  /* ---------------- export ---------------- */

  // Builds the full Compare workbook from current state. Shared by the
  // manual "Download .xlsx" button and the folder auto-save path so both
  // produce byte-identical output.
  function buildCompareWorkbook() {
    const res = state.result;
    const wb = XLSX.utils.book_new();

    const summary = [
      ['BOM Compare results', ''],
      ['CAD BOM file(s)', state.cadSources.map(function (s) { return s.fileName || ''; }).join(' + ')],
      ['Item Master file', state.im.fileName || ''],
      ['Compared on', new Date().toISOString().slice(0, 16).replace('T', ' ')],
      [],
      ['Unique parts in CAD BOM', res.cadUniqueCount],
      ['Unique parts in Item Master', res.imUniqueCount],
      ['Findings needing action', res.actionableCount],
      ['Missing incl. grouped children', res.missingTotal],
      ['Reference components', res.referenceRoots === null ? 'n/a (needs PDF + Inventor export together)' : res.referenceTotal],
      ['Quantity mismatches', res.qtyMismatches === null ? 'n/a (no CAD source has quantities)' : res.qtyMismatches.length],
      ['Quantity cascades', res.qtyCascades.applicable ? res.qtyCascades.roots.length : 'n/a (no CAD source has quantities)'],
      ['In Item Master only', res.imOnly.length + (res.imOnlyRoots ? ' (' + res.imOnlyRoots.length + ' incl. grouping)' : '')],
    ];
    if (state.findings) {
      summary.push(['Unique parts needing attention', state.findings.counts.actionable]);
      summary.push(['…plus grouped under a flagged parent', state.findings.counts.grouped]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Summary');

    if (state.findings) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(attentionSheetRows(state.findings)), 'Parts needing attention');
    }

    const missing = [['Action needed', 'Parent Number', 'Parent Title', 'Level', 'Part Number', 'Title', 'Description', 'Type', 'File', 'Material', 'CAD Row #']];
    const walk = function (node, depth, rootNumber, rootTitle) {
      const it = node.item;
      missing.push([
        depth === 0 ? 'YES' : 'grouped (reference subtree)',
        depth === 0 ? '' : rootNumber,
        depth === 0 ? '' : rootTitle,
        depth + 1,
        it.number, it.title, it.description,
        it.isAssembly === null ? '' : (it.isAssembly ? 'Assembly' : 'Part'),
        it.file, it.material || '', it.sourceRow || '',
      ]);
      for (const c of node.children) walk(c, depth + 1, rootNumber, rootTitle);
    };
    for (const n of res.missingRoots) walk(n, 0, n.item.number, n.item.title);
    if (isCheckEnabled('missing')) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(missing), 'Missing from Item Master');

    if (isCheckEnabled('reference') && res.referenceRoots !== null) {
      const ref = [['Reference component', 'Parent Number', 'Parent Title', 'In Item Master?', 'Part Number', 'Title', 'Description', 'Type', 'File', 'CAD Row #']];
      const walkRef = function (node, depth, rootNumber, rootTitle) {
        const it = node.item;
        ref.push([
          depth === 0 ? 'YES' : 'child of reference subtree',
          depth === 0 ? '' : rootNumber,
          depth === 0 ? '' : rootTitle,
          node.inItemMaster ? 'yes' : 'NO',
          it.number, it.title, it.description,
          it.isAssembly === null ? '' : (it.isAssembly ? 'Assembly' : 'Part'),
          it.file, it.sourceRow || '',
        ]);
        for (const c of node.children) walkRef(c, depth + 1, rootNumber, rootTitle);
      };
      for (const n of res.referenceRoots) walkRef(n, 0, n.item.number, n.item.title);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ref), 'Reference items');
    }

    const breakdownText = function (list) {
      return (list || [])
        .map(function (b) {
          const where = b.parentNumber ? (b.parentNumber + (b.parentTitle ? ' (' + b.parentTitle + ')' : '')) : '(top-level)';
          const row = b.sourceRow ? ' [row ' + b.sourceRow + ']' : '';
          return where + row + ': ' + fmtQty(b.qty);
        })
        .join(' + ');
    };
    const rowNumsText = function (list) {
      return (list || []).map(function (b) { return b.sourceRow; }).filter(Boolean).join(', ');
    };
    const qty = [['Part Number', 'Title', 'Qty in CAD', 'Qty in Item Master', 'Difference',
      'Item Master Row #(s)', 'Found Under (CAD)', 'Found Under (Item Master)']];
    if (res.qtyMismatches) {
      for (const m of res.qtyMismatches) {
        qty.push([m.number, m.title, m.cadQty, m.imQty, m.cadQty - m.imQty,
          rowNumsText(m.imBreakdown), breakdownText(m.cadBreakdown), breakdownText(m.imBreakdown)]);
      }
    } else {
      qty.push(['n/a — the CAD BOM source has no quantity column']);
    }
    if (isCheckEnabled('qty')) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qty), 'Qty mismatches');

    const cascades = [['Action needed', 'Cascade ratio', 'Children at ratio', 'Grouped under', 'Part Number', 'Title', 'Row #']];
    const walkCascade = function (node, depth, rootNumber) {
      const it = node.item;
      cascades.push([
        depth === 0 ? 'YES — fix this root cause' : 'grouped (explained by cascade above)',
        depth === 0 ? it.cascadeRatio + '×' : '',
        depth === 0 ? it.cascadeMismatchedChildCount + ' of ' + it.cascadeChildCount : '',
        depth === 0 ? '' : rootNumber,
        it.number, it.title, it.sourceRow || '',
      ]);
      for (const c of node.children || []) walkCascade(c, depth + 1, rootNumber);
    };
    if (res.qtyCascades.applicable) {
      for (const n of res.qtyCascades.roots) walkCascade(n, 0, n.item.number);
    } else {
      cascades.push(['n/a — the CAD BOM source has no quantity column']);
    }
    if (isCheckEnabled('qtyCascade')) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cascades), 'Quantity cascades');

    // Grouped-under column mirrors the Missing sheet's "grouped" marker so a
    // whole subassembly absent from the CAD BOM reads as one finding.
    const groupedUnder = new Map();
    if (res.imOnlyRoots) {
      const walkIo = function (node, rootNumber) {
        for (const c of node.children || []) {
          groupedUnder.set(BC.normNumber(c.item.number), rootNumber);
          walkIo(c, rootNumber);
        }
      };
      for (const n of res.imOnlyRoots) walkIo(n, n.item.number);
    }
    const imonly = [['Action needed', 'Grouped under', 'Part Number', 'Title', 'Description', 'Qty', 'Row type', 'Row #', 'Parent Number', 'Parent Title']];
    for (const r of res.imOnly) {
      const under = groupedUnder.get(BC.normNumber(r.number)) || '';
      imonly.push([under ? 'grouped (parent also absent)' : 'YES', under,
        r.number, r.title, r.description, r.qty, r.rowType || '', r.sourceRow || '', r.parentNumber || '', r.parentTitle || '']);
    }
    if (isCheckEnabled('imOnly')) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(imonly), 'In Item Master only');

    if (isCheckEnabled('imQc') && state.imQc) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qcSheetRows(state.imQc)), 'Item Master QC');
    }

    if (isCheckEnabled('material') && state.materialResult) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(materialSheetRows(state.materialResult)), 'Material vs CAD');
    }

    if (isCheckEnabled('revision') && state.revisionResult) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(revisionSheetRows(state.revisionResult)), 'Revision vs CAD');
    }

    if (isCheckEnabled('titleDesc') && state.titleDescResult) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(titleDescSheetRows(state.titleDescResult)), 'Description vs CAD');
    }

    if (isCheckEnabled('virtual') && state.virtualResult) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(virtualSheetRows(state.virtualResult)), 'Virtual parts');
    }

    if (isCheckEnabled('lldbo') && (state.lldboResult || state.lldboCandidatesResult)) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(lldboSheetRows(state.lldboResult, state.lldboCandidatesResult)), 'LLDBO check');
    }

    const allIgnored = (res.ignoredFindings || [])
      .concat((state.revisionResult && state.revisionResult.ignoredFindings) || [])
      .concat((state.lldboCandidatesResult && state.lldboCandidatesResult.ignoredFindings) || []);
    if (allIgnored.length) {
      const ignored = [['Part Number', 'Title', 'Description', 'Would have been flagged as', 'Row #', 'Parent Number', 'Parent Title']];
      for (const r of allIgnored) {
        ignored.push([r.number, r.title, r.description, (BC.findings.CHECK_BY_KEY[r.checkKey] || {}).label || r.checkKey,
          r.sourceRow || '', r.parentNumber || '', r.parentTitle || '']);
      }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ignored), 'Ignored findings');
    }

    return wb;
  }

  function compareResultFilename() {
    return 'BOM-compare-results' + projectKeySuffix() + '.xlsx';
  }

  $('btn-export').addEventListener('click', function () {
    if (!state.result) return;
    XLSX.writeFile(buildCompareWorkbook(), compareResultFilename());
  });

  /* ---------------- dropzone wiring ---------------- */

  function wireZone(role) {
    const zone = $('zone-' + role);
    const input = $('file-' + role);
    const pick = $('pick-' + role);
    const open = function () { input.click(); };
    zone.addEventListener('click', function (ev) {
      if (ev.target.closest('button')) return;
      open();
    });
    pick.addEventListener('click', function (ev) { ev.stopPropagation(); open(); });
    zone.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); open(); }
    });
    input.addEventListener('change', function () {
      if (input.files && input.files.length) handleFiles(role, Array.from(input.files));
      input.value = '';
    });
    zone.addEventListener('dragover', function (ev) { ev.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('dragover'); });
    zone.addEventListener('drop', function (ev) {
      ev.preventDefault();
      zone.classList.remove('dragover');
      if (ev.dataTransfer.files && ev.dataTransfer.files.length) handleFiles(role, Array.from(ev.dataTransfer.files));
    });
  }

  // LLDBO is a single-file, optional zone with its own handler (not routed
  // through the CAD/IM-specific handleFiles dispatcher).
  function wireLldboZone() {
    const zone = $('zone-lldbo');
    const input = $('file-lldbo');
    const pick = $('pick-lldbo');
    const open = function () { input.click(); };
    zone.addEventListener('click', function (ev) { if (!ev.target.closest('button')) open(); });
    pick.addEventListener('click', function (ev) { ev.stopPropagation(); open(); });
    zone.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); open(); }
    });
    input.addEventListener('change', function () {
      if (input.files && input.files[0]) handleLldboFile(input.files[0]);
      input.value = '';
    });
    zone.addEventListener('dragover', function (ev) { ev.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('dragover'); });
    zone.addEventListener('drop', function (ev) {
      ev.preventDefault();
      zone.classList.remove('dragover');
      if (ev.dataTransfer.files && ev.dataTransfer.files[0]) handleLldboFile(ev.dataTransfer.files[0]);
    });
  }

  // Ignore List: same single-file, optional-zone shape as LLDBO above.
  function wireIgnoreListZone() {
    const zone = $('zone-ignorelist');
    const input = $('file-ignorelist');
    const pick = $('pick-ignorelist');
    const open = function () { input.click(); };
    zone.addEventListener('click', function (ev) { if (!ev.target.closest('button')) open(); });
    pick.addEventListener('click', function (ev) { ev.stopPropagation(); open(); });
    zone.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); open(); }
    });
    input.addEventListener('change', function () {
      if (input.files && input.files[0]) handleIgnoreListFile(input.files[0]);
      input.value = '';
    });
    zone.addEventListener('dragover', function (ev) { ev.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('dragover'); });
    zone.addEventListener('drop', function (ev) {
      ev.preventDefault();
      zone.classList.remove('dragover');
      if (ev.dataTransfer.files && ev.dataTransfer.files[0]) handleIgnoreListFile(ev.dataTransfer.files[0]);
    });
  }

  /* ---------------- folder auto-load / auto-save ---------------- */
  // File System Access API: a browser cannot read a typed filesystem path
  // (no such API exists, in any browser, for security reasons). One click
  // opens the OS's own folder picker; the user browses to the PNxxxx
  // project folder (a mapped NAS drive or UNC share both work, since it's
  // the OS picker doing the browsing); everything after that — scanning,
  // loading, comparing, saving the result back — is automatic. Chrome/Edge
  // only; Firefox/Safari fall back to the manual dropzones below.

  function folderStatus(text, kind) {
    const el = $('folder-status');
    el.textContent = text || '';
    el.className = 'folder-status' + (kind ? ' ' + kind : '');
  }

  // loader: async (file) => void — either handleFiles(role, [file]) for CAD/IM
  // or handleLldboFile(file) for the LLDBO's own single-file path.
  async function loadFolderMatch(entries, label, loader, optional) {
    if (entries.length === 1) {
      const file = await entries[0].getFile();
      await loader(file);
      return file.name;
    }
    if (entries.length === 0) return optional ? 'none found (optional)' : 'no ' + label + ' found — drop it manually';
    return entries.length + ' possible ' + label + ' files found — ambiguous, drop the right one manually';
  }

  async function saveResultToFolder(matchSummary) {
    if (!state.folderHandle || !state.result) return;
    const filename = compareResultFilename();
    try {
      const buf = XLSX.write(buildCompareWorkbook(), { bookType: 'xlsx', type: 'array' });
      const fileHandle = await state.folderHandle.getFileHandle(filename, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(buf);
      await writable.close();
      folderStatus(matchSummary + ' — saved "' + filename + '" to "' + state.folderHandle.name + '" ✓', 'ok');
    } catch (e) {
      folderStatus(matchSummary + ' — compared, but could not save the result into the folder (' +
        (e.message || e) + ') — use Download .xlsx instead.', 'error');
    }
  }

  async function handlePickFolder() {
    let handle;
    try {
      handle = await window.showDirectoryPicker();
    } catch (e) {
      if (e && e.name === 'AbortError') return; // user cancelled the picker
      folderStatus('Could not open that folder (' + (e.message || e) + ').', 'error');
      return;
    }
    state.folderHandle = handle;
    folderStatus('Scanning "' + handle.name + '"…');

    let found;
    try {
      found = await BC.folder.scanFolder(handle);
    } catch (e) {
      folderStatus('Could not read "' + handle.name + '" (' + (e.message || e) + ').', 'error');
      return;
    }

    const cadMsg = await loadFolderMatch(found['cad-pdf'], 'CAD BOM PDF', function (f) { return handleFiles('cad', [f]); }, false);
    const imMsg = await loadFolderMatch(found['item-master'], 'Item Master (EBOM_*)', function (f) { return handleFiles('im', [f]); }, false);
    const invMsg = await loadFolderMatch(found['inventor-bom'], 'Inventor BOM export (INVENTOR_BOM_*)', function (f) { return handleFiles('cad', [f]); }, true);
    const lldboMsg = await loadFolderMatch(found['lldbo'], 'LLDBO file', handleLldboFile, true);
    const ignoreMsg = await loadFolderMatch(found['ignore-list'], 'Ignore List file', handleIgnoreListFile, true);
    const matchSummary = '"' + handle.name + '" — CAD: ' + cadMsg + ' · Item Master: ' + imMsg + ' · Inventor BOM: ' + invMsg +
      ' · LLDBO: ' + lldboMsg + ' · Ignore List: ' + ignoreMsg;
    folderStatus(matchSummary);

    if (state.cadSources.length && state.im) {
      if (runCompare()) await saveResultToFolder(matchSummary);
    }
  }

  function setupFolderFeature() {
    const supported = typeof window.showDirectoryPicker === 'function';
    $('btn-pick-folder').hidden = !supported;
    $('folder-unsupported').classList.toggle('hidden', supported);
    if (supported) $('btn-pick-folder').addEventListener('click', handlePickFolder);
  }

  wireZone('cad');
  wireZone('im');
  wireLldboZone();
  wireIgnoreListZone();
  renderColumnsMenu();
  renderChecksMenu();
  setupFolderFeature();
})();
