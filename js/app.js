/* app.js — UI wiring for BOM Compare. All parsing/comparison logic lives in
 * js/compare.js and js/parsers/; this file only handles files, state and DOM. */
(function () {
  'use strict';

  const BC = window.BOMCompare;
  const $ = function (id) { return document.getElementById(id); };

  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
  }

  const state = {
    cadSources: [],     // parsed CAD results (max 2: full structure + intended BOM)
    im: null,           // parsed Item Master result (+ fileName)
    imQc: null,         // js/imqc.js runChecks() result for state.im
    result: null,
    activeTab: 'missing',
    filter: '',
    mappingCtx: null,   // { aoa, indents, source, analysis, fileName }
  };

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
    } else {
      let srcLabel = { 'flat-xlsx': 'Vault flat export', 'pdf': 'Vault PDF (full structure)', 'leveled-sheet': 'Leveled table' }[parsed.source] || parsed.source;
      if (parsed.source === 'leveled-sheet' && parsed.hasStructure) srcLabel = 'Inventor BOM export';
      chips.push({ text: srcLabel, kind: 'good' });
      chips.push({ text: parsed.items.length + ' components' });
      chips.push({ text: parsed.hasQty ? 'quantities ✓' : 'no quantities', kind: parsed.hasQty ? '' : 'warn' });
      chips.push({ text: parsed.hasLevels ? 'levels ✓' : 'no levels', kind: parsed.hasLevels ? '' : 'warn' });
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
        else { state.im = null; state.imQc = null; hideImQc(); setImStatus(file.name, [], e.message || String(e)); }
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
      throw new Error('No "Number" header row found. Expected the Vault/ERP Item Master BOM export.');
    }
    im.fileName = file.name;
    state.im = im;
    setImStatus(file.name, chipsFor(im), '');
    state.imQc = BC.imQc.runChecks(im);
    renderImQc();
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
    const parsed = BC.cadLeveledParser.parse(grid.rows, { indents: grid.indents, source: 'pdf' });
    if (!parsed) {
      renderCadSources('', '');
      showMapping({
        aoa: grid.rows,
        indents: grid.indents,
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

  const QC_CHECKS = [
    { key: 'c1', title: 'Producer ↔ Description Match (top-level row only)',
      desc: 'The top-level row’s Producer and Producer Number should appear inside its Description.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order'], ['producer', 'Producer'], ['producerNumber', 'Producer Number'], ['description', 'Description']] },
    { key: 'c2', title: 'End of Line Integrity',
      desc: 'The "END OF LINE" entry should have Number ' + BC.imQc.END_OF_LINE_NUMBER + ' and a whole-number Row Order.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order'], ['issue', 'Issue']] },
    { key: 'c3', title: 'Quantity vs Item Qty',
      desc: 'The numeric value inside Quantity should equal Item Qty on every row — a mismatch usually means one was edited manually without updating the other.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order'], ['title', 'Title'], ['quantity', 'Quantity'], ['itemQty', 'Item Qty']] },
    { key: 'c4', title: 'Entity Icon Status',
      desc: 'Every row’s Entity Icon should read "Normal".',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order'], ['icon', 'Entity Icon']] },
    { key: 'c5', title: 'Title / Description Completeness',
      desc: 'Every row should have a Title and Description. Purchased/catalog parts (X-999-…) are only flagged when both are blank; other parts are flagged if either is missing.',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order'], ['title', 'Title'], ['description', 'Description'], ['kind', 'Issue', qcKindLabel]] },
    { key: 'c6', title: 'Material Completeness',
      desc: 'Every non-assembly row should have a Material. Assemblies/weldments are excluded (they legitimately carry no material of their own).',
      cols: [['number', 'Number'], ['rowOrder', 'Row Order'], ['title', 'Title']] },
  ];

  function hideImQc() {
    $('im-qc').classList.add('hidden');
    $('qc-summary').innerHTML = '';
    $('qc-sections').innerHTML = '';
    $('qc-callout').classList.add('hidden');
    $('qc-callout').textContent = '';
  }

  function qcCardFor(check, result) {
    const cls = !result.applicable ? 'na' : (result.fail.length ? 'fail' : 'pass');
    const num = !result.applicable ? 'N/A' : (result.fail.length ? String(result.fail.length) : 'PASS');
    const sub = !result.applicable ? result.reason : (result.fail.length ? 'flagged' : 'clean');
    const el = document.createElement('div');
    el.className = 'card ' + cls;
    el.innerHTML = '<div class="num"></div><div class="lbl"></div>';
    el.querySelector('.num').textContent = num;
    el.querySelector('.lbl').textContent = check.title.replace(/\s*\(.*\)$/, '') + ' — ' + sub;
    return el;
  }

  function qcSectionFor(check, result) {
    const box = document.createElement('div');
    box.className = 'qc-section';
    const head = document.createElement('div');
    head.className = 'qc-section-head';
    const titleBox = document.createElement('div');
    titleBox.innerHTML = '<div class="qc-section-title"></div><p class="qc-section-desc"></p>';
    titleBox.querySelector('.qc-section-title').textContent = check.title;
    titleBox.querySelector('.qc-section-desc').textContent = check.desc;
    head.appendChild(titleBox);
    const pill = document.createElement('span');
    pill.className = 'qc-pill ' + (!result.applicable ? 'na' : (result.fail.length ? 'fail' : 'pass'));
    pill.textContent = !result.applicable ? 'N/A' : (result.fail.length ? result.fail.length + ' FLAGGED' : 'OK');
    head.appendChild(pill);
    box.appendChild(head);

    const body = document.createElement('div');
    body.className = 'qc-section-body';
    if (!result.applicable) {
      body.innerHTML = '<div class="empty-state">' + result.reason + '</div>';
    } else if (!result.fail.length) {
      body.innerHTML = '<div class="empty-state">✓ No issues found' + (check.key === 'c1' && result.applicableCount ? ' across ' + result.applicableCount + ' applicable row(s)' : '') + '.</div>';
    } else {
      const table = document.createElement('table');
      table.className = 'results-table';
      const htr = document.createElement('tr');
      for (const [, label] of check.cols) addTh(htr, label);
      table.appendChild(htr);
      for (const row of result.fail) {
        const tr = document.createElement('tr');
        for (const col of check.cols) addTd(tr, qcCellValue(col, row));
        table.appendChild(tr);
      }
      body.appendChild(table);
      body.classList.add('open');
    }
    box.appendChild(body);

    head.addEventListener('click', function () { body.classList.toggle('open'); });
    return box;
  }

  function renderImQc() {
    const qc = state.imQc;
    if (!qc) { hideImQc(); return; }
    $('im-qc').classList.remove('hidden');
    const summary = $('qc-summary');
    summary.innerHTML = '';
    const sections = $('qc-sections');
    sections.innerHTML = '';
    for (const check of QC_CHECKS) {
      const result = qc[check.key];
      summary.appendChild(qcCardFor(check, result));
      sections.appendChild(qcSectionFor(check, result));
    }

    const callout = $('qc-callout');
    const bits = [];
    if (qc.c5.applicable && qc.c5.fail.length) {
      bits.push(qc.c5.fail.length + ' row(s) need Title and/or Description populated');
    }
    if (qc.c6.applicable && qc.c6.fail.length) {
      bits.push(qc.c6.fail.length + ' part(s) need Material populated');
    }
    if (bits.length) {
      callout.textContent = '⚠ ' + bits.join(' · ') + ' — see below, or download the QC report for a highlighted copy of the Item Master.';
      callout.classList.remove('hidden');
    } else {
      callout.textContent = '';
      callout.classList.add('hidden');
    }
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

  // Delegates to js/imqc-export.js (shared with the Node test suite, which
  // asserts the fills actually round-trip through a real .xlsx).
  function buildStyledImSheet(im, qc) {
    return BC.imQcExport.buildStyledImSheet(XLSX, im, qc);
  }

  $('btn-qc-export').addEventListener('click', function () {
    if (!state.imQc) return;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qcSheetRows(state.imQc)), 'Item Master QC');
    XLSX.utils.book_append_sheet(wb, buildStyledImSheet(state.im, state.imQc), 'Item Master — data quality');
    XLSX.writeFile(wb, 'ItemMaster-QC-report' + projectKeySuffix() + '.xlsx');
  });

  /* ---------------- compare & render ---------------- */

  function updateCompareButton() {
    $('btn-compare').disabled = !(state.cadSources.length && state.im);
  }

  $('btn-compare').addEventListener('click', function () {
    if (!state.cadSources.length || !state.im) return;
    clearNotices();
    state.result = BC.compareAll(state.cadSources, state.im);
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
    // reference tab visibility
    $('tab-ref').classList.toggle('hidden', res.referenceRoots === null);
    if (res.referenceRoots === null && state.activeTab === 'ref') switchTab('missing');
    $('results').classList.remove('hidden');
    renderResults();
    $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

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
    if (!res) return;

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
    $('count-imonly').textContent = res.imOnly.length;

    renderMissingTab();
    renderRefTab();
    renderQtyTab();
    renderImOnlyTab();
  }

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
      tr.appendChild(tdNum);

      if (cols.title) addTd(tr, it.title);
      if (cols.description) addTd(tr, it.description);
      if (cols.type) addTd(tr, it.isAssembly === null ? '' : (it.isAssembly ? 'Assembly' : 'Part'));
      if (cols.qty && res.hasQty) addTd(tr, fmtQty(it.qty));
      if (cols.file) addTd(tr, it.file);
      if (cols.material) addTd(tr, it.material || '');
      if (cols.row) addTd(tr, String(it.sourceRow || ''));

      tbody.appendChild(tr);
      for (const c of node.children) renderNode(c, depth + 1, id);
    };

    for (const rootNode of roots) renderNode(rootNode, 0, null);
    table.appendChild(tbody);
    pane.appendChild(table);

    // expand/collapse behaviour
    tbody.addEventListener('click', function (ev) {
      const exp = ev.target.closest('.expander');
      if (!exp) return;
      const id = exp.dataset.for;
      const expanded = exp.textContent === '▾';
      if (expanded) {
        exp.textContent = '▸';
        hideDescendants(tbody, id);
      } else {
        exp.textContent = '▾';
        tbody.querySelectorAll('tr[data-parent="' + id + '"]').forEach(function (tr) {
          tr.classList.remove('hidden-row');
        });
      }
    });
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

  function nodeTreeMatches(node) {
    if (matchesFilter([node.item.number, node.item.title, node.item.description])) return true;
    return node.children.some(nodeTreeMatches);
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

  function breakdownTable(caption, entries) {
    const box = document.createElement('div');
    box.className = 'qty-breakdown';
    const cap = document.createElement('div');
    cap.textContent = caption + ':';
    box.appendChild(cap);
    const t = document.createElement('table');
    const h = document.createElement('tr');
    ['Parent assembly', 'Parent title', 'Qty per parent', 'Effective qty'].forEach(function (x) { addTh(h, x); });
    t.appendChild(h);
    for (const e of entries) {
      const tr = document.createElement('tr');
      addTd(tr, e.parentNumber || '(top level)');
      addTd(tr, e.parentTitle || '');
      addTd(tr, fmtQty(e.qty));
      addTd(tr, fmtQty(e.effQty));
      t.appendChild(tr);
    }
    box.appendChild(t);
    return box;
  }

  /* ----- IM-only tab ----- */

  function renderImOnlyTab() {
    const pane = $('pane-imonly');
    const res = state.result;
    pane.innerHTML = '';
    const rows = res.imOnly.filter(function (r) {
      return matchesFilter([r.number, r.title, r.description]);
    });
    if (!rows.length) {
      pane.innerHTML = '<div class="empty-state">' + (state.filter ? 'Nothing matches the filter.' : 'Every Item Master entry also appears in the CAD BOM.') + '</div>';
      return;
    }
    const cols = visibleCols();
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, 'Part Number');
    if (cols.title) addTh(htr, 'Title');
    if (cols.description) addTh(htr, 'Description');
    if (cols.qty) addTh(htr, 'Qty');
    if (cols.row) addTh(htr, 'Row type');
    table.appendChild(htr);
    for (const r of rows) {
      const tr = document.createElement('tr');
      const tdNum = document.createElement('td');
      tdNum.className = 'num-cell';
      tdNum.textContent = r.number;
      tr.appendChild(tdNum);
      if (cols.title) addTd(tr, r.title);
      if (cols.description) addTd(tr, r.description);
      if (cols.qty) addTd(tr, fmtQty(r.qty));
      if (cols.row) addTd(tr, r.rowType || '');
      table.appendChild(tr);
    }
    pane.appendChild(table);
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
    ['missing', 'ref', 'qty', 'imonly'].forEach(function (n) {
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
  document.addEventListener('click', function (ev) {
    if (!ev.target.closest('#columns-dropdown')) $('columns-menu').classList.add('hidden');
  });

  /* ---------------- export ---------------- */

  $('btn-export').addEventListener('click', function () {
    const res = state.result;
    if (!res) return;
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
      ['In Item Master only', res.imOnly.length],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Summary');

    const missing = [['Action needed', 'Grouped under', 'Level', 'Part Number', 'Title', 'Description', 'Type', 'File', 'Material', 'CAD source row']];
    const walk = function (node, depth, rootNumber) {
      const it = node.item;
      missing.push([
        depth === 0 ? 'YES' : 'grouped (reference subtree)',
        depth === 0 ? '' : rootNumber,
        depth + 1,
        it.number, it.title, it.description,
        it.isAssembly === null ? '' : (it.isAssembly ? 'Assembly' : 'Part'),
        it.file, it.material || '', it.sourceRow || '',
      ]);
      for (const c of node.children) walk(c, depth + 1, rootNumber);
    };
    for (const n of res.missingRoots) walk(n, 0, n.item.number);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(missing), 'Missing from Item Master');

    if (res.referenceRoots !== null) {
      const ref = [['Reference component', 'Grouped under', 'In Item Master?', 'Part Number', 'Title', 'Description', 'Type', 'File']];
      const walkRef = function (node, depth, rootNumber) {
        const it = node.item;
        ref.push([
          depth === 0 ? 'YES' : 'child of reference subtree',
          depth === 0 ? '' : rootNumber,
          node.inItemMaster ? 'yes' : 'NO',
          it.number, it.title, it.description,
          it.isAssembly === null ? '' : (it.isAssembly ? 'Assembly' : 'Part'),
          it.file,
        ]);
        for (const c of node.children) walkRef(c, depth + 1, rootNumber);
      };
      for (const n of res.referenceRoots) walkRef(n, 0, n.item.number);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ref), 'Reference items');
    }

    const qty = [['Part Number', 'Title', 'Qty in CAD', 'Qty in Item Master', 'Difference']];
    if (res.qtyMismatches) {
      for (const m of res.qtyMismatches) qty.push([m.number, m.title, m.cadQty, m.imQty, m.cadQty - m.imQty]);
    } else {
      qty.push(['n/a — the CAD BOM source has no quantity column']);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qty), 'Qty mismatches');

    const imonly = [['Part Number', 'Title', 'Description', 'Qty', 'Row type']];
    for (const r of res.imOnly) imonly.push([r.number, r.title, r.description, r.qty, r.rowType || '']);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(imonly), 'In Item Master only');

    if (state.imQc) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(qcSheetRows(state.imQc)), 'Item Master QC');
      XLSX.utils.book_append_sheet(wb, buildStyledImSheet(state.im, state.imQc), 'Item Master — data quality');
    }

    XLSX.writeFile(wb, 'BOM-compare-results' + projectKeySuffix() + '.xlsx');
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

  wireZone('cad');
  wireZone('im');
  renderColumnsMenu();
})();
