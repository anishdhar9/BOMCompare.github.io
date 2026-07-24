/*
 * cad-flat-xlsx.js — parser for the flat (level-less) Vault CAD BOM paste,
 * as in the sample CAD_Bom.xlsx:
 *
 *   - headerless, depth-first pre-order listing
 *   - a boolean cell flags every record (true = CAD component, false = an
 *     attached export such as .stp — skipped)
 *   - fixed offsets from the boolean cell: +1 state, +2 filename,
 *     +3 revision, +4 material, +5 part number
 *   - title/description are the cells BEFORE the boolean; a few records are
 *     split across two physical rows (long text pushed the record one column
 *     left on the next row), so text-only "orphan" rows are buffered and
 *     merged into the following record
 *   - blank separator rows between records
 *
 * This export carries NO quantity column, so hasQty is false.
 *
 * Produces: { kind:'cad', source:'flat-xlsx', hasQty:false, hasLevels:false,
 *             items:[...], warnings }
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

  function isBlankRow(row) {
    return !row || row.every(function (c) { return c === null || c === undefined || cellText(c) === ''; });
  }

  function boolIndex(row) {
    for (let i = 0; i < row.length; i++) if (row[i] === true || row[i] === false) return i;
    return -1;
  }

  function parse(workbook, XLSX) {
    const warnings = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      // raw:true keeps the booleans as booleans
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

      const items = [];
      let orphan = null; // buffered text-only row(s) of a split record
      let skipped = 0;

      for (let r = 0; r < aoa.length; r++) {
        const row = aoa[r] || [];
        if (isBlankRow(row)) { orphan = null; continue; }
        const bi = boolIndex(row);
        if (bi === -1) {
          // text-only row: first part of a split record; keep the first one
          // (subsequent orphan rows of the same record add no new fields)
          if (!orphan) orphan = row;
          continue;
        }
        if (row[bi] === false) { skipped++; orphan = null; continue; } // attachment (.stp etc.)

        const file = cellText(row[bi + 2]);
        const number = cellText(row[bi + 5]);
        let title = '';
        let description = '';
        const before = row.slice(0, bi).map(cellText).filter(function (s) { return s !== ''; });
        if (before.length) {
          title = before[0];
          description = before.slice(1).join(' ');
        } else if (orphan) {
          const ob = orphan.map(cellText).filter(function (s) { return s !== ''; });
          title = ob[0] || '';
          description = ob.slice(1).join(' ');
        }
        orphan = null;
        if (!number) {
          warnings.push('Row ' + (r + 1) + ': component row without a part number — skipped.');
          continue;
        }
        items.push({
          seq: items.length,
          number: number,
          title: title,
          description: description,
          qty: null,
          level: null,
          isAssembly: /\.iam$/i.test(file),
          file: file,
          material: cellText(row[bi + 4]),
          revision: cellText(row[bi + 3]),
          state: cellText(row[bi + 1]),
          sourceRow: r + 1,
        });
      }

      if (items.length >= 3) {
        if (skipped) warnings.push(skipped + ' attachment rows (e.g. .stp exports) ignored.');
        return {
          kind: 'cad',
          source: 'flat-xlsx',
          sheetName: sheetName,
          hasQty: false,
          hasLevels: false,
          hasMaterial: items.some(function (it) { return it.material !== ''; }),
          hasRevision: items.some(function (it) { return it.revision !== ''; }),
          items: items,
          warnings: warnings,
        };
      }
    }
    return null; // not this format
  }

  return { cadFlatParser: { parse: parse } };
});
