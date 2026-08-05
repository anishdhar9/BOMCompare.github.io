/*
 * ignorelist.js — parser for the "Ignore List" file: part numbers the user
 * has intentionally left in a non-standard BOM state (e.g. still Reference,
 * or a WIP/sketch placeholder) that should not be reported as findings.
 *
 * Real-sample layout: "S.No." / "Part Number" / "From" — "From" names which
 * check category to suppress the part from (e.g. "CAD vs Item compare").
 * The category -> check-key mapping lives in js/ignorelist-compare.js, which
 * also owns interpreting/validating the "From" text; this parser only reads
 * the raw cells.
 *
 * The file is always named "IgnoreList<code>" (e.g. "IgnoreListHSG.xlsx") —
 * this organization's own fixed convention. Detection (in js/folder.js and
 * the app's own dropzone) matches on that filename prefix, not sheet
 * content, so this parser doesn't need a content-based "looksLikeX" guard
 * the way lldbo.js does — the filename already disambiguates it before the
 * parser ever runs.
 *
 * Produces: { rows:[{number, from, sourceRow}], warnings }
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

  function findHeader(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 20); r++) {
      const row = aoa[r] || [];
      const lower = row.map(function (c) { return cellText(c).toLowerCase(); });
      const partCol = lower.findIndex(function (h) { return h === 'part number' || h === 'part no' || h === 'part no.'; });
      if (partCol === -1) continue;
      const fromCol = lower.findIndex(function (h) { return h === 'from'; });
      return { headerRow: r, cols: { partNumber: partCol, from: fromCol } };
    }
    return null;
  }

  // XLSX is the SheetJS namespace (injected — matches this codebase's style,
  // e.g. lldbo.js's parse(workbook, XLSX)).
  function parse(workbook, XLSX) {
    const warnings = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
      const hdr = findHeader(aoa);
      if (!hdr) continue;
      if (hdr.cols.from === -1) {
        warnings.push('No "From" column found — cannot tell which check(s) to ignore these parts from.');
      }

      const rows = [];
      for (let r = hdr.headerRow + 1; r < aoa.length; r++) {
        const row = aoa[r] || [];
        const number = cellText(row[hdr.cols.partNumber]);
        if (!number) continue;
        rows.push({
          number: number,
          from: hdr.cols.from >= 0 ? cellText(row[hdr.cols.from]) : '',
          sourceRow: r + 1,
        });
      }
      if (!rows.length) continue;

      return { rows: rows, warnings: warnings };
    }
    return null; // not an Ignore List document
  }

  return { ignoreListParser: { parse: parse, findHeader: findHeader } };
});
