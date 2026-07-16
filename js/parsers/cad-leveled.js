/*
 * cad-leveled.js — parser for multi-level CAD BOM tables with a header row:
 * a leveled Excel/CSV export, or the row grid reconstructed from the Vault
 * multi-level BOM PDF by pdf-extract.js.
 *
 * Column sets vary (Vault lets users pick visible fields), so columns are
 * auto-detected by header keywords and can be overridden by an explicit
 * `mapping` (from the app's column-mapping UI).
 *
 * Hierarchy level per row, in order of preference:
 *   1. an explicit Level column (numeric)
 *   2. a dotted item/position column ('1.2.3' -> level 3)
 *   3. row indentation (leading spaces, or x-offsets supplied by pdf-extract)
 *
 * Produces: { kind:'cad', source, hasQty, hasLevels, items:[...], columns,
 *             headerRow, warnings }
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function cellText(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function parseQty(v) {
    const s = cellText(v);
    if (!s || s === '-') return null;
    const m = s.replace(',', '.').match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  }

  const FIELD_KEYWORDS = {
    number: ['number', 'part number', 'part no', 'part no.', 'item number', 'document number', 'artikelnummer', 'teilenummer', 'sachnummer'],
    qty: ['qty', 'qty.', 'quantity', 'item qty', 'unit qty', 'menge', 'anzahl', 'stück', 'stck'],
    level: ['level', 'ebene', 'stufe', 'depth'],
    pos: ['item', 'pos', 'pos.', 'position', 'row order', 'bom structure position'],
    title: ['title', 'name', 'bezeichnung', 'benennung'],
    description: ['description', 'beschreibung'],
    file: ['file', 'file name', 'filename', 'dateiname', 'document'],
  };

  function matchField(headerText) {
    const h = headerText.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!h) return null;
    for (const field of Object.keys(FIELD_KEYWORDS)) {
      if (FIELD_KEYWORDS[field].indexOf(h) !== -1) return field;
    }
    // prefix matches for compound headers like 'Title (Item,CO)' / 'Qty per'
    for (const field of Object.keys(FIELD_KEYWORDS)) {
      for (const kw of FIELD_KEYWORDS[field]) {
        if (kw.length >= 3 && h.indexOf(kw) === 0) return field;
      }
    }
    return null;
  }

  function detectHeader(aoa) {
    let best = null;
    const limit = Math.min(aoa.length, 12);
    for (let r = 0; r < limit; r++) {
      const row = aoa[r] || [];
      const cols = {};
      let score = 0;
      for (let c = 0; c < row.length; c++) {
        const f = matchField(cellText(row[c]));
        if (f && cols[f] === undefined) { cols[f] = c; score++; }
      }
      if (cols.number === undefined) continue;
      if (score >= 2 && (!best || score > best.score)) best = { headerRow: r, cols: cols, score: score };
    }
    return best;
  }

  // Quantize indentation offsets (leading-space counts or PDF x-offsets)
  // into 1-based levels.
  function levelsFromIndents(indents) {
    const uniq = Array.from(new Set(indents.filter(function (v) { return v !== null; }))).sort(function (a, b) { return a - b; });
    if (uniq.length < 2) return null;
    // merge offsets closer than half the median step (PDF x jitter)
    const steps = [];
    for (let i = 1; i < uniq.length; i++) steps.push(uniq[i] - uniq[i - 1]);
    steps.sort(function (a, b) { return a - b; });
    const tol = steps[Math.floor(steps.length / 2)] / 2;
    const buckets = [uniq[0]];
    for (const v of uniq.slice(1)) {
      if (v - buckets[buckets.length - 1] > tol) buckets.push(v);
    }
    return indents.map(function (v) {
      if (v === null) return null;
      let lvl = 1;
      for (let i = 0; i < buckets.length; i++) if (v >= buckets[i] - tol) lvl = i + 1;
      return lvl;
    });
  }

  // For the mapping UI: header + sample values per column.
  function analyze(aoa) {
    const hdr = detectHeader(aoa);
    const headerRow = hdr ? hdr.headerRow : 0;
    const header = (aoa[headerRow] || []).map(cellText);
    const width = Math.max.apply(null, aoa.slice(0, 30).map(function (r) { return (r || []).length; }).concat([header.length]));
    const columns = [];
    for (let c = 0; c < width; c++) {
      const samples = [];
      for (let r = headerRow + 1; r < aoa.length && samples.length < 5; r++) {
        const t = cellText((aoa[r] || [])[c]);
        if (t) samples.push(t);
      }
      columns.push({ index: c, header: header[c] || '(column ' + (c + 1) + ')', samples: samples });
    }
    return { headerRow: headerRow, columns: columns, detected: hdr ? hdr.cols : { }, score: hdr ? hdr.score : 0 };
  }

  /**
   * @param aoa      array-of-arrays including the header row
   * @param opts     { mapping?: {number,qty,level,pos,title,description,file},
   *                   headerRow?: number,
   *                   indents?: (number|null)[] aligned to aoa rows (from pdf-extract),
   *                   source?: string }
   */
  function parse(aoa, opts) {
    opts = opts || {};
    const warnings = [];
    let cols, headerRow;
    if (opts.mapping) {
      cols = opts.mapping;
      headerRow = opts.headerRow !== undefined ? opts.headerRow : 0;
    } else {
      const hdr = detectHeader(aoa);
      if (!hdr) return null;
      cols = hdr.cols;
      headerRow = hdr.headerRow;
    }
    if (cols.number === undefined || cols.number === null || cols.number < 0) return null;

    const col = function (name) {
      return cols[name] !== undefined && cols[name] !== null && cols[name] >= 0 ? cols[name] : -1;
    };
    const cNumber = col('number'), cQty = col('qty'), cLevel = col('level'),
          cPos = col('pos'), cTitle = col('title'), cDesc = col('description'), cFile = col('file');

    const items = [];
    const rawIndents = [];
    for (let r = headerRow + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const rawNumberCell = row[cNumber];
      const number = cellText(rawNumberCell);
      if (!number) continue;
      // skip repeated header rows (multi-page PDFs)
      if (matchField(number) === 'number') continue;

      let level = null;
      if (cLevel >= 0) {
        const lv = parseQty(row[cLevel]);
        if (lv !== null) level = Math.round(lv);
      }
      if (level === null && cPos >= 0) {
        const pos = cellText(row[cPos]);
        if (/^\d+(\.\d+)*$/.test(pos)) level = pos.split('.').length;
      }
      let indent = null;
      if (opts.indents && opts.indents[r] !== undefined && opts.indents[r] !== null) {
        indent = opts.indents[r];
      } else {
        const rawStr = rawNumberCell === null || rawNumberCell === undefined ? '' : String(rawNumberCell);
        const m = rawStr.match(/^[ \t]*/);
        indent = m ? m[0].replace(/\t/g, '    ').length : 0;
      }
      rawIndents.push(indent);

      items.push({
        seq: items.length,
        number: number,
        title: cTitle >= 0 ? cellText(row[cTitle]) : '',
        description: cDesc >= 0 ? cellText(row[cDesc]) : '',
        qty: cQty >= 0 ? parseQty(row[cQty]) : null,
        level: level,
        isAssembly: null, // resolved below
        file: cFile >= 0 ? cellText(row[cFile]) : '',
        material: '',
        sourceRow: r + 1,
      });
    }
    if (!items.length) return null;

    // fall back to indentation-based levels when no explicit level/pos data
    let hasLevels = items.some(function (it) { return it.level !== null; });
    if (!hasLevels) {
      const lv = levelsFromIndents(rawIndents);
      if (lv) {
        items.forEach(function (it, i) { it.level = lv[i]; });
        hasLevels = true;
        warnings.push('Hierarchy inferred from row indentation.');
      }
    }
    // normalize levels to start at 1
    if (hasLevels) {
      let min = Infinity;
      for (const it of items) if (it.level !== null && it.level < min) min = it.level;
      if (min !== Infinity && min !== 1) for (const it of items) if (it.level !== null) it.level += 1 - min;
    }

    // isAssembly: file extension when available, otherwise "has children"
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (/\.iam$/i.test(it.file)) it.isAssembly = true;
      else if (/\.ipt$/i.test(it.file)) it.isAssembly = false;
      else if (hasLevels) {
        const next = items[i + 1];
        it.isAssembly = !!(next && next.level !== null && it.level !== null && next.level > it.level);
      }
    }

    const hasQty = cQty >= 0 && items.some(function (it) { return it.qty !== null; });
    if (!hasQty) warnings.push('No usable quantity column — quantity comparison unavailable for this file.');
    if (!hasLevels) warnings.push('No level/position information found — reference-assembly grouping will be inferred from the Item Master hierarchy.');

    return {
      kind: 'cad',
      source: opts.source || 'leveled-sheet',
      hasQty: hasQty,
      hasLevels: hasLevels,
      items: items,
      columns: cols,
      headerRow: headerRow,
      warnings: warnings,
    };
  }

  return { cadLeveledParser: { parse: parse, analyze: analyze, detectHeader: detectHeader, matchField: matchField } };
});
