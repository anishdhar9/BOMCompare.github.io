/* im-diff-app.js — UI wiring for the standalone Item Master diff page
 * (im-diff.html). All diff logic lives in js/im-diff-compare.js; this file
 * only handles files, state and DOM, mirroring js/app.js's own split. */
(function () {
  'use strict';

  const BC = window.BOMCompare;
  const $ = function (id) { return document.getElementById(id); };

  const state = {
    old: null,   // parsed Item Master result (+ fileName)
    new: null,   // parsed Item Master result (+ fileName)
    diff: null,  // js/im-diff-compare.js diffItemMasters() result
    filter: '',
    activeTab: 'added',
  };

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      const r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(new Error('Could not read the file.')); };
      r.readAsArrayBuffer(file);
    });
  }

  function chipEls(box, chips) {
    (chips || []).forEach(function (ch) {
      const el = document.createElement('span');
      el.className = 'chip' + (ch.kind ? ' ' + ch.kind : '');
      el.textContent = ch.text;
      if (ch.title) el.title = ch.title;
      box.appendChild(el);
    });
  }

  function chipsForIm(im) {
    const chips = [{ text: 'Item Master', kind: 'good' }, { text: im.rows.length + ' rows' }];
    if (im.projectKey) chips.push({ text: im.projectKey.spn + ' / ' + im.projectKey.pn });
    chips.push({ text: im.hasPaths ? 'hierarchy ✓' : 'no hierarchy', kind: im.hasPaths ? '' : 'warn' });
    return chips;
  }

  function setStatus(role, fileName, chips, error) {
    $('status-file-' + role).textContent = fileName || '';
    const chipBox = $('status-chips-' + role);
    chipBox.innerHTML = '';
    chipEls(chipBox, chips);
    $('status-error-' + role).textContent = error || '';
    const zone = $('zone-' + role);
    zone.classList.toggle('parsed', !error && !!fileName);
    zone.classList.toggle('error', !!error);
  }

  function updateCompareButton() {
    $('btn-diff').disabled = !(state.old && state.new);
  }

  function notice(kind, text) {
    const el = document.createElement('p');
    el.className = 'notice ' + kind;
    el.textContent = text;
    $('notices').appendChild(el);
  }

  async function handleFile(role, file) {
    try {
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
      const im = BC.itemMasterParser.parse(wb, XLSX);
      if (!im) throw new Error('No "Number" header row found. Expected a Vault/ERP Item Master BOM export.');
      im.fileName = file.name;
      state[role] = im;
      setStatus(role, file.name, chipsForIm(im), '');
    } catch (e) {
      state[role] = null;
      setStatus(role, file.name, [], e.message || String(e));
    }
    updateCompareButton();
  }

  function wireZone(role) {
    const zone = $('zone-' + role);
    const input = $('file-' + role);
    const pick = $('pick-' + role);
    const open = function () { input.click(); };
    zone.addEventListener('click', function (ev) { if (!ev.target.closest('button')) open(); });
    pick.addEventListener('click', function (ev) { ev.stopPropagation(); open(); });
    zone.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); open(); }
    });
    input.addEventListener('change', function () {
      if (input.files && input.files[0]) handleFile(role, input.files[0]);
      input.value = '';
    });
    zone.addEventListener('dragover', function (ev) { ev.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('dragover'); });
    zone.addEventListener('drop', function (ev) {
      ev.preventDefault();
      zone.classList.remove('dragover');
      if (ev.dataTransfer.files && ev.dataTransfer.files[0]) handleFile(role, ev.dataTransfer.files[0]);
    });
  }
  wireZone('old');
  wireZone('new');

  function matchesFilter(strings) {
    if (!state.filter) return true;
    const f = state.filter.toLowerCase();
    return strings.some(function (s) { return s && String(s).toLowerCase().indexOf(f) !== -1; });
  }

  function addTh(tr, text) { const th = document.createElement('th'); th.textContent = text; tr.appendChild(th); }
  function addTd(tr, text) { const td = document.createElement('td'); td.textContent = text === null || text === undefined ? '' : text; tr.appendChild(td); }

  function renderAddedRemovedTab(paneId, list, rowClass, emptyMsg) {
    const pane = $(paneId);
    pane.innerHTML = '';
    const rows = list.filter(function (r) { return matchesFilter([r.number, r.title, r.description]); });
    if (!rows.length) {
      pane.innerHTML = '<div class="empty-state">' + (state.filter ? 'Nothing matches the filter.' : emptyMsg) + '</div>';
      return;
    }
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, 'Part Number'); addTh(htr, 'Title'); addTh(htr, 'Description'); addTh(htr, 'Row #');
    table.appendChild(htr);
    rows.forEach(function (r) {
      const tr = document.createElement('tr');
      tr.className = rowClass;
      addTd(tr, r.number); addTd(tr, r.title); addTd(tr, r.description); addTd(tr, r.sourceRow);
      table.appendChild(tr);
    });
    pane.appendChild(table);
  }

  function renderChangedTab() {
    const pane = $('pane-changed');
    pane.innerHTML = '';
    const rows = state.diff.changed.filter(function (r) { return matchesFilter([r.number, r.title, r.description]); });
    if (!rows.length) {
      pane.innerHTML = '<div class="empty-state">' +
        (state.filter ? 'Nothing matches the filter.' : 'No field changes found on any shared part.') + '</div>';
      return;
    }
    const table = document.createElement('table');
    table.className = 'results-table';
    const htr = document.createElement('tr');
    addTh(htr, 'Part Number'); addTh(htr, 'Title'); addTh(htr, 'Field'); addTh(htr, 'Old'); addTh(htr, 'New'); addTh(htr, 'Row # (old / new)');
    table.appendChild(htr);
    rows.forEach(function (r) {
      r.fields.forEach(function (f, i) {
        const tr = document.createElement('tr');
        tr.className = 'row-qty';
        addTd(tr, i === 0 ? r.number : '');
        addTd(tr, i === 0 ? r.title : '');
        addTd(tr, f.field);
        addTd(tr, f.old);
        addTd(tr, f.new);
        addTd(tr, i === 0 ? (r.oldSourceRow + ' / ' + r.newSourceRow) : '');
        table.appendChild(tr);
      });
    });
    pane.appendChild(table);
  }

  function renderResults() {
    const diff = state.diff;
    if (!diff) return;
    $('results-source').textContent = 'Older: ' + (state.old.fileName || '') + '  →  Newer: ' + (state.new.fileName || '');

    const summary = $('summary');
    summary.innerHTML = '';
    const cards = [
      { num: diff.oldUniqueCount, lbl: 'unique parts in older file' },
      { num: diff.newUniqueCount, lbl: 'unique parts in newer file' },
      { num: diff.added.length, lbl: 'added', cls: diff.added.length ? 'amber' : '' },
      { num: diff.removed.length, lbl: 'removed', cls: diff.removed.length ? 'red' : '' },
      { num: diff.changed.length, lbl: 'changed', cls: diff.changed.length ? 'amber' : '' },
      { num: diff.unchangedCount, lbl: 'unchanged' },
    ];
    cards.forEach(function (c) {
      const el = document.createElement('div');
      el.className = 'card' + (c.cls ? ' ' + c.cls : '');
      el.innerHTML = '<div class="num"></div><div class="lbl"></div>';
      el.querySelector('.num').textContent = c.num;
      el.querySelector('.lbl').textContent = c.lbl;
      summary.appendChild(el);
    });

    $('count-added').textContent = diff.added.length;
    $('count-removed').textContent = diff.removed.length;
    $('count-changed').textContent = diff.changed.length;

    renderAddedRemovedTab('pane-added', diff.added, 'row-ref', 'No parts were added.');
    renderAddedRemovedTab('pane-removed', diff.removed, 'row-missing', 'No parts were removed.');
    renderChangedTab();
  }

  function runDiff() {
    $('notices').innerHTML = '';
    if (!state.old || !state.new) return;
    if (state.old.fileName === state.new.fileName && state.old.projectKey && state.new.projectKey &&
      state.old.projectKey.pn !== state.new.projectKey.pn) {
      notice('warn', 'The two files carry different SPN/PN project keys — check this is really the same BOM at two points in time.');
    } else if (state.old.projectKey && state.new.projectKey && state.old.projectKey.pn !== state.new.projectKey.pn) {
      notice('warn', 'The two files carry different SPN/PN project keys (' + state.old.projectKey.pn + ' vs ' +
        state.new.projectKey.pn + ') — check this is really the same BOM at two points in time.');
    }
    state.diff = BC.imDiffCompare.diffItemMasters(state.old, state.new, BC.indexItemMaster, BC.materialCompare.materialsMatch);
    $('results').classList.remove('hidden');
    renderResults();
  }

  $('btn-diff').addEventListener('click', function () {
    runDiff();
    $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  $('filter').addEventListener('input', function () {
    state.filter = this.value.trim();
    renderResults();
  });

  function switchTab(name) {
    state.activeTab = name;
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === name); });
    ['added', 'removed', 'changed'].forEach(function (n) {
      $('pane-' + n).classList.toggle('hidden', n !== name);
    });
  }
  $('tabs').addEventListener('click', function (ev) {
    const tab = ev.target.closest('.tab');
    if (!tab) return;
    switchTab(tab.dataset.tab);
  });

  function buildDiffWorkbook() {
    const diff = state.diff;
    const wb = XLSX.utils.book_new();
    const summary = [
      ['Item Master diff', ''],
      ['Older file', state.old.fileName || ''],
      ['Newer file', state.new.fileName || ''],
      ['Compared on', new Date().toISOString().slice(0, 16).replace('T', ' ')],
      [],
      ['Unique parts in older file', diff.oldUniqueCount],
      ['Unique parts in newer file', diff.newUniqueCount],
      ['Added', diff.added.length],
      ['Removed', diff.removed.length],
      ['Changed', diff.changed.length],
      ['Unchanged', diff.unchangedCount],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Summary');

    const added = [['Part Number', 'Title', 'Description', 'Row #']];
    diff.added.forEach(function (r) { added.push([r.number, r.title, r.description, r.sourceRow]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(added), 'Added');

    const removed = [['Part Number', 'Title', 'Description', 'Row #']];
    diff.removed.forEach(function (r) { removed.push([r.number, r.title, r.description, r.sourceRow]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(removed), 'Removed');

    const changed = [['Part Number', 'Title', 'Field', 'Old', 'New', 'Row # (old)', 'Row # (new)']];
    diff.changed.forEach(function (r) {
      r.fields.forEach(function (f) {
        changed.push([r.number, r.title, f.field, f.old, f.new, r.oldSourceRow, r.newSourceRow]);
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(changed), 'Changed');

    return wb;
  }

  $('btn-export').addEventListener('click', function () {
    if (!state.diff) return;
    const wb = buildDiffWorkbook();
    const oldName = (state.old.fileName || 'old').replace(/\.[^.]+$/, '');
    const newName = (state.new.fileName || 'new').replace(/\.[^.]+$/, '');
    XLSX.writeFile(wb, 'ItemMaster-diff_' + oldName + '_vs_' + newName + '.xlsx');
  });

  updateCompareButton();
})();
