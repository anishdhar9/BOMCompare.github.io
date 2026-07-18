/*
 * imqc-export.js — builds the styled "Item Master — data quality" worksheet
 * (real cell fills on flagged Title/Description/Material cells) from an
 * imqc.js runChecks() result. Requires a styling-capable XLSX namespace —
 * vendor/xlsx.full.min.js is xlsx-js-style (not community SheetJS) for
 * exactly this reason; community SheetJS silently drops `.s` cell styles
 * on write (verified empirically), so highlighting would be lost.
 *
 * Kept separate from app.js so it's requireable from the Node test suite,
 * which asserts the fills actually round-trip through a real .xlsx file
 * rather than just checking the cell text is present.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Mirrors the app's --red-*/--amber-* CSS custom properties (light-mode
  // values) so the exported cells read the same way the on-screen QC panel
  // does.
  const STYLE_RED = { fill: { fgColor: { rgb: 'FDECEC' } }, font: { color: { rgb: '9F1C21' }, bold: true } };
  const STYLE_AMBER = { fill: { fgColor: { rgb: 'FDF3E1' } }, font: { color: { rgb: '7D5807' }, bold: true } };

  function setCellStyle(XLSX, ws, colIdx, rowIdx, style) {
    const addr = XLSX.utils.encode_cell({ c: colIdx, r: rowIdx });
    if (!ws[addr]) ws[addr] = { t: 's', v: '' };
    ws[addr].s = style;
  }

  // The full Item Master, one physical row per BOM row, with actual cell
  // highlighting on the fields checks 5/6 (Title/Description, Material)
  // flagged — "highlighted in the export file" taken literally, in context,
  // rather than a separate flagged-only list.
  function buildStyledImSheet(XLSX, im, qc) {
    const cols = [
      ['number', 'Number'], ['sourceRow', 'Row #'], ['rowOrder', 'Row Order'],
      ['parentNumber', 'Parent Number'], ['parentTitle', 'Parent Title'],
      ['title', 'Title'], ['description', 'Description'], ['material', 'Material'],
      ['producer', 'Producer'], ['producerNumber', 'Producer Number'],
      ['itemQty', 'Item Qty'], ['quantityText', 'Quantity'],
    ];
    const rowOrderOf = function (row) { return Array.isArray(row.path) ? (row.path.join('.') || '-') : ''; };
    const pathIndex = new Map();
    for (const row of im.rows) {
      if (Array.isArray(row.path) && !pathIndex.has(row.path.join('.'))) pathIndex.set(row.path.join('.'), row);
    }
    const parentOf = function (row) {
      if (!Array.isArray(row.path) || row.path.length === 0) return null;
      const parent = pathIndex.get(row.path.slice(0, -1).join('.'));
      return parent ? { number: parent.number, title: parent.title || '' } : null;
    };
    const cellValue = function (c, row) {
      if (c[0] === 'rowOrder') return rowOrderOf(row);
      if (c[0] === 'parentNumber' || c[0] === 'parentTitle') {
        const p = parentOf(row);
        if (!p) return '';
        return c[0] === 'parentNumber' ? p.number : p.title;
      }
      return row[c[0]];
    };
    const aoa = [cols.map(function (c) { return c[1]; })];
    for (const row of im.rows) {
      aoa.push(cols.map(function (c) { return cellValue(c, row); }));
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    const titleDescBySourceRow = new Map();
    if (qc.c5.applicable) for (const f of qc.c5.fail) titleDescBySourceRow.set(f.sourceRow, f.kind);
    const materialFlagged = new Set(qc.c6.applicable ? qc.c6.fail.map(function (f) { return f.sourceRow; }) : []);
    const titleCol = cols.findIndex(function (c) { return c[0] === 'title'; });
    const descCol = cols.findIndex(function (c) { return c[0] === 'description'; });
    const materialCol = cols.findIndex(function (c) { return c[0] === 'material'; });

    im.rows.forEach(function (row, i) {
      const r = i + 1; // 0-based sheet row; row 0 is the header
      const kind = titleDescBySourceRow.get(row.sourceRow);
      if (kind === 'both-missing') {
        setCellStyle(XLSX, ws, titleCol, r, STYLE_RED);
        setCellStyle(XLSX, ws, descCol, r, STYLE_RED);
      } else if (kind === 'title-missing') {
        setCellStyle(XLSX, ws, titleCol, r, STYLE_AMBER);
      } else if (kind === 'description-missing') {
        setCellStyle(XLSX, ws, descCol, r, STYLE_AMBER);
      }
      if (materialFlagged.has(row.sourceRow)) setCellStyle(XLSX, ws, materialCol, r, STYLE_RED);
    });
    return ws;
  }

  return { imQcExport: { buildStyledImSheet: buildStyledImSheet, STYLE_RED: STYLE_RED, STYLE_AMBER: STYLE_AMBER } };
});
