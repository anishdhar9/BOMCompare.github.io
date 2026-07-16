/*
 * itemmaster.js — parser for the Vault "Item Master BOM" Excel export
 * (.xls or .xlsx). The export's visible columns are user-configurable in
 * Vault, so columns are located by header name; only 'Number' is mandatory.
 *
 * Produces: { kind:'itemmaster', rows:[{number,title,description,qty,path,
 *             rowType,sourceRow}], hasPaths, sheetName, columns, warnings }
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

  // '4 Each' -> 4, '1,5' -> 1.5, '16' -> 16, '-' -> null
  function parseQty(v) {
    const s = cellText(v);
    if (!s || s === '-') return null;
    const m = s.replace(',', '.').match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  }

  // '2.8.1' -> ['2','8','1'], '-' -> [] (the root row), '' -> null
  function parsePath(v) {
    const s = cellText(v);
    if (!s) return null;
    if (s === '-') return [];
    return s.split('.').map(function (p) { return p.trim(); });
  }

  function findHeader(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 15); r++) {
      const row = aoa[r] || [];
      const lower = row.map(function (c) { return cellText(c).toLowerCase(); });
      const numberCol = lower.indexOf('number');
      if (numberCol === -1) continue;
      // require one more Item-Master-ish header to avoid false positives
      const marker = lower.some(function (h) {
        return h === 'row order' || h === 'item qty' || h === 'quantity' ||
               h === 'row type' || h.indexOf('title') === 0 || h === 'category name';
      });
      if (!marker) continue;
      const find = function (pred) {
        for (let i = 0; i < lower.length; i++) if (pred(lower[i])) return i;
        return -1;
      };
      return {
        headerRow: r,
        cols: {
          number: numberCol,
          qty: find(function (h) { return h === 'item qty'; }),
          qtyFallback: find(function (h) { return h === 'quantity'; }),
          path: find(function (h) { return h === 'row order'; }),
          title: find(function (h) { return h.indexOf('title') === 0; }),
          description: find(function (h) { return h.indexOf('description') === 0; }),
          rowType: find(function (h) { return h === 'row type'; }),
        },
      };
    }
    return null;
  }

  // XLSX is the SheetJS namespace (injected so Node tests can pass their own).
  function parse(workbook, XLSX) {
    const warnings = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      // raw:false -> formatted text, keeps '6.10' distinct from '6.1' when the
      // Row Order cells are text (they are in Vault exports).
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
      const hdr = findHeader(aoa);
      if (!hdr) continue;

      const rows = [];
      for (let r = hdr.headerRow + 1; r < aoa.length; r++) {
        const row = aoa[r] || [];
        const number = cellText(row[hdr.cols.number]);
        if (!number) continue;
        let qty = hdr.cols.qty >= 0 ? parseQty(row[hdr.cols.qty]) : null;
        if (qty === null && hdr.cols.qtyFallback >= 0) qty = parseQty(row[hdr.cols.qtyFallback]);
        rows.push({
          number: number,
          title: hdr.cols.title >= 0 ? cellText(row[hdr.cols.title]) : '',
          description: hdr.cols.description >= 0 ? cellText(row[hdr.cols.description]) : '',
          qty: qty,
          path: hdr.cols.path >= 0 ? parsePath(row[hdr.cols.path]) : null,
          rowType: hdr.cols.rowType >= 0 ? cellText(row[hdr.cols.rowType]) : '',
          sourceRow: r + 1,
        });
      }
      if (!rows.length) continue;

      const hasPaths = hdr.cols.path >= 0;
      if (!hasPaths) warnings.push('No "Row Order" column found — quantity roll-up treats all rows as direct children.');
      if (hdr.cols.qty < 0 && hdr.cols.qtyFallback < 0) warnings.push('No quantity column found — quantity comparison unavailable.');

      return {
        kind: 'itemmaster',
        sheetName: sheetName,
        rows: rows,
        hasPaths: hasPaths,
        columns: hdr.cols,
        warnings: warnings,
      };
    }
    return null; // not an Item Master export
  }

  return { itemMasterParser: { parse: parse, parseQty: parseQty, parsePath: parsePath } };
});
