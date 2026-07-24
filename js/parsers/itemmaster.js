/*
 * itemmaster.js — parser for the Vault "Item Master BOM" Excel export
 * (.xls or .xlsx). The export's visible columns are user-configurable in
 * Vault, so columns are located by header keyword rather than position —
 * only 'Number' (or an equivalent synonym) is mandatory.
 *
 * Produces: { kind:'itemmaster', rows:[{number,title,description,qty,
 *             itemQty,quantity,quantityText,producer,producerNumber,
 *             entityIcon,material,revision,path,rowType,sourceRow}],
 *             hasPaths, hasEntityIcon, hasProducer, hasMaterial,
 *             hasRevision, hasItemQty, hasQuantity,
 *             projectKey:{spn,pn}|null, sheetName, columns, warnings }
 *
 * `qty` is the resolved quantity used by compare.js's roll-up. Some exports
 * carry up to three quantity-ish columns -- "Item Quantity", "Quantity",
 * and "Quantity Per Unit" -- but "Quantity" (values like "1 Each") is the
 * one that reflects the actual as-released quantity, so it is preferred;
 * "Item Qty"/"Item Quantity" is only a fallback when "Quantity" is absent.
 * "Quantity Per Unit" is a distinct column (see `quantityPerUnit` keyword
 * below) and must never be confused for "Quantity" during header matching.
 * `itemQty`/`quantity` are kept separate (unresolved) so imqc.js can flag
 * when a manual edit to one doesn't match the other — a real, observed
 * failure mode: someone edits the displayed Quantity without updating Item
 * Qty (or vice versa).
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

  // Column sets vary across organizations' (and even the same organization's
  // different plants'/users') Vault exports -- e.g. "Part Number" or "Item
  // Number" instead of "Number" -- so columns are located by keyword rather
  // than a single exact header string, mirroring cad-leveled.js's
  // FIELD_KEYWORDS/matchField pattern.
  //
  // Deliberately does NOT include "PN" as a Number synonym: in this
  // organization's convention "PN" means Producer Number (the numeric half
  // of the project's SPN/PN key, e.g. "PN22426" -- see extractProjectKey()
  // below), never a part number. "PN" is instead a producerNumber synonym.
  const FIELD_KEYWORDS = {
    number: ['number', 'part number', 'item number'],
    qty: ['item qty', 'qty', 'qty.', 'item quantity'],
    qtyFallback: ['quantity'],
    // Per-unit quantity columns ("Quantity Per Unit", "QTY per Unit", "Unit
    // Qty", ...) are a DIFFERENT concept from the total/as-released
    // "Quantity" or the "Item Qty": they hold the quantity of this line per
    // single unit of its parent (usually 1, even when the total Quantity is
    // 4). They must never be captured as qty/qtyFallback, or Check 3 flags
    // every multi-qty row as a false mismatch. Enumerated here (matched
    // exactly, and via the longest-prefix rule in matchField) so a bare
    // 'qty'/'quantity' prefix can't grab them. Recognized only to block that
    // collision -- never captured into row data.
    quantityPerUnit: [
      'quantity per unit', 'qty per unit', 'qty. per unit',
      'quantity/unit', 'qty/unit', 'unit qty', 'unit quantity', 'qty per parent',
    ],
    path: ['row order', 'level', 'position', 'bom level'],
    title: ['title', 'name'],
    description: ['description', 'desc'],
    rowType: ['row type', 'type'],
    producer: ['producer'],
    producerNumber: ['producer number', 'pn'],
    entityIcon: ['entity icon', 'icon'],
    material: ['material'],
    revision: ['revision', 'rev'],
    // recognized only so it counts toward the header-row marker check below;
    // not a field this parser captures into row data.
    marker: ['category name'],
  };

  function matchField(headerText) {
    const h = headerText.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!h) return null;
    // 1. an exact header match wins outright.
    for (const field of Object.keys(FIELD_KEYWORDS)) {
      if (FIELD_KEYWORDS[field].indexOf(h) !== -1) return field;
    }
    // 2. otherwise, prefix match for compound headers like 'Title (Item,CO)'
    //    or 'QTY per Unit (Each)'. When several keywords are a prefix, the
    //    LONGEST (most specific) one wins -- so a short 'qty' can never beat a
    //    specific 'qty per unit', regardless of column/field order.
    let best = null, bestLen = 0;
    for (const field of Object.keys(FIELD_KEYWORDS)) {
      for (const kw of FIELD_KEYWORDS[field]) {
        if (kw.length >= 3 && h.indexOf(kw) === 0 && kw.length > bestLen) {
          best = field; bestLen = kw.length;
        }
      }
    }
    return best;
  }

  function findHeader(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 15); r++) {
      const row = aoa[r] || [];
      const cols = {};
      let score = 0;
      for (let c = 0; c < row.length; c++) {
        const f = matchField(cellText(row[c]));
        if (f && cols[f] === undefined) { cols[f] = c; score++; }
      }
      if (cols.number === undefined) continue;
      if (score < 2) continue; // require one more Item-Master-ish header to avoid false positives
      const at = function (field) { return cols[field] !== undefined ? cols[field] : -1; };
      return {
        headerRow: r,
        cols: {
          number: cols.number,
          qty: at('qty'),
          qtyFallback: at('qtyFallback'),
          path: at('path'),
          title: at('title'),
          description: at('description'),
          rowType: at('rowType'),
          producer: at('producer'),
          producerNumber: at('producerNumber'),
          entityIcon: at('entityIcon'),
          material: at('material'),
          revision: at('revision'),
        },
      };
    }
    return null;
  }

  // The root/FG-level row's Producer + Producer Number columns are this
  // organization's convention for the project's SPN/PN key (verified:
  // Producer='SPN016823', Producer Number='22426' on the root row only —
  // Producer is reused for real vendor names like 'SKF' on component rows).
  // Falls back to regex over Title+Description for exports without those
  // columns populated on the root row.
  function extractProjectKey(rootRow) {
    if (!rootRow) return null;
    let spn = null, pn = null;
    if (rootRow.producer && /^SPN\d+$/i.test(rootRow.producer)) spn = rootRow.producer.toUpperCase();
    if (rootRow.producerNumber && /^\d+$/.test(rootRow.producerNumber)) pn = 'PN' + rootRow.producerNumber;
    if (!spn || !pn) {
      const text = (rootRow.title || '') + ' ' + (rootRow.description || '');
      const m = text.match(/SPN(\d+)[_,\s]*PN(\d+)/i);
      if (m) {
        if (!spn) spn = 'SPN' + m[1];
        if (!pn) pn = 'PN' + m[2];
      }
    }
    if (!spn && !pn) return null;
    return { spn: spn, pn: pn };
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
        const itemQty = hdr.cols.qty >= 0 ? parseQty(row[hdr.cols.qty]) : null;
        const quantity = hdr.cols.qtyFallback >= 0 ? parseQty(row[hdr.cols.qtyFallback]) : null;
        // "Quantity" (e.g. "1 Each") is the as-released quantity; "Item
        // Qty"/"Item Quantity" is only a fallback when "Quantity" is absent.
        let qty = quantity;
        if (qty === null) qty = itemQty;
        rows.push({
          number: number,
          title: hdr.cols.title >= 0 ? cellText(row[hdr.cols.title]) : '',
          description: hdr.cols.description >= 0 ? cellText(row[hdr.cols.description]) : '',
          qty: qty,
          itemQty: itemQty,
          quantity: quantity,
          quantityText: hdr.cols.qtyFallback >= 0 ? cellText(row[hdr.cols.qtyFallback]) : '',
          producer: hdr.cols.producer >= 0 ? cellText(row[hdr.cols.producer]) : '',
          producerNumber: hdr.cols.producerNumber >= 0 ? cellText(row[hdr.cols.producerNumber]) : '',
          entityIcon: hdr.cols.entityIcon >= 0 ? cellText(row[hdr.cols.entityIcon]) : '',
          material: hdr.cols.material >= 0 ? cellText(row[hdr.cols.material]) : '',
          revision: hdr.cols.revision >= 0 ? cellText(row[hdr.cols.revision]) : '',
          path: hdr.cols.path >= 0 ? parsePath(row[hdr.cols.path]) : null,
          rowType: hdr.cols.rowType >= 0 ? cellText(row[hdr.cols.rowType]) : '',
          sourceRow: r + 1,
        });
      }
      if (!rows.length) continue;

      const hasPaths = hdr.cols.path >= 0;
      if (!hasPaths) warnings.push('No "Row Order" column found — quantity roll-up treats all rows as direct children.');
      if (hdr.cols.qty < 0 && hdr.cols.qtyFallback < 0) warnings.push('No quantity column found — quantity comparison unavailable.');

      const rootRow = rows.find(function (r) { return Array.isArray(r.path) && r.path.length === 0; }) || rows[0];

      return {
        kind: 'itemmaster',
        sheetName: sheetName,
        rows: rows,
        hasPaths: hasPaths,
        hasProducer: hdr.cols.producer >= 0 || hdr.cols.producerNumber >= 0,
        hasEntityIcon: hdr.cols.entityIcon >= 0,
        hasMaterial: hdr.cols.material >= 0,
        hasRevision: hdr.cols.revision >= 0,
        hasItemQty: hdr.cols.qty >= 0,
        hasQuantity: hdr.cols.qtyFallback >= 0,
        projectKey: extractProjectKey(rootRow),
        columns: hdr.cols,
        warnings: warnings,
      };
    }
    return null; // not an Item Master export
  }

  return { itemMasterParser: { parse: parse, parseQty: parseQty, parsePath: parsePath, extractProjectKey: extractProjectKey } };
});
