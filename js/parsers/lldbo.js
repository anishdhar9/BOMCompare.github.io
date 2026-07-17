/*
 * lldbo.js — parser for the "Long Lead Direct Bought Out (LLDBO)" list.
 * Long-lead-time items are released to procurement ahead of the normal BOM
 * release, to account for supplier lead times. This is the record of what
 * was pre-released; js/lldbo-compare.js checks it against the Item Master
 * to catch the process failure where a pre-released part never actually
 * made it into the released BOM (so it silently never gets ordered through
 * the normal channel either).
 *
 * Real-sample layout: a merged-cell document header (title, customer,
 * "DBO Doc No : SPN######_PN#####_<description>", issue date/doc no/date)
 * above a normal table with header "SR. No / PART NO / Item Description /
 * Specifications / Make / Qty. / Remarks". Column set isn't guaranteed
 * stable, so columns are located by header name; only PART NO is required.
 * Not every row has a part number yet — some are placeholders ("Pending"
 * specifications, no PN) for items not yet specified; those are counted
 * separately, not treated as findings.
 *
 * Deliberately NOT routed through detect.js's generic CAD leveled-table
 * detector: "PART NO"/"Qty." would false-match its keyword table and this
 * file would get misparsed as a CAD BOM.
 *
 * Produces: { kind:'lldbo', rows:[{srNo,partNo,description,specifications,
 *             make,qty,qtyText,remarks,sourceRow}], projectKey:{spn,pn}|null,
 *             customer, documentNo, sheetName, columns, warnings }
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

  // '1 Nos.' -> 1, '3 Nos.' -> 3, 'NA' -> null, '' -> null
  function parseQty(v) {
    const s = cellText(v);
    if (!s) return null;
    const m = s.match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  }

  function findHeader(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 20); r++) {
      const row = aoa[r] || [];
      const lower = row.map(function (c) { return cellText(c).toLowerCase(); });
      const partCol = lower.findIndex(function (h) { return h === 'part no' || h === 'part no.' || h === 'part number'; });
      if (partCol === -1) continue;
      const find = function (pred) {
        for (let i = 0; i < lower.length; i++) if (pred(lower[i])) return i;
        return -1;
      };
      return {
        headerRow: r,
        cols: {
          srNo: find(function (h) { return h === 'sr. no' || h === 'sr no' || h === 'sr.no'; }),
          partNo: partCol,
          description: find(function (h) { return h.indexOf('description') === 0 || h.indexOf('item description') === 0; }),
          specifications: find(function (h) { return h.indexOf('specification') === 0; }),
          make: find(function (h) { return h === 'make'; }),
          qty: find(function (h) { return h.indexOf('qty') === 0; }),
          remarks: find(function (h) { return h === 'remarks'; }),
        },
      };
    }
    return null;
  }

  // The document header (above the table) carries the same SPN/PN project
  // key convention used elsewhere in this organization's exports — scanned
  // from whichever cell contains it (e.g. "DBO Doc No : SPN016838_PN22260_
  // HSG PRO 800L"), plus a customer name for context.
  function extractDocInfo(aoa, headerRow) {
    let projectKey = null, customer = '';
    const limit = Math.min(headerRow, 10);
    for (let r = 0; r < limit; r++) {
      for (const cell of (aoa[r] || [])) {
        const s = cellText(cell);
        if (!s) continue;
        if (!projectKey) {
          const m = s.match(/SPN(\d+)[_,\s]*PN(\d+)/i);
          if (m) projectKey = { spn: 'SPN' + m[1], pn: 'PN' + m[2] };
        }
        if (!customer) {
          const cm = s.match(/customer\s*:\s*(.+)/i);
          if (cm) customer = cm[1].trim();
        }
      }
    }
    return { projectKey: projectKey, customer: customer };
  }

  // XLSX is the SheetJS namespace (injected so Node tests can pass their own).
  function parse(workbook, XLSX) {
    const warnings = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
      const hdr = findHeader(aoa);
      if (!hdr) continue;

      const rows = [];
      for (let r = hdr.headerRow + 1; r < aoa.length; r++) {
        const row = aoa[r] || [];
        const anyContent = row.some(function (c) { return cellText(c) !== ''; });
        if (!anyContent) continue; // spacer row (seen right after the header in the real sample)
        const qtyRaw = hdr.cols.qty >= 0 ? row[hdr.cols.qty] : null;
        rows.push({
          srNo: hdr.cols.srNo >= 0 ? cellText(row[hdr.cols.srNo]) : '',
          partNo: hdr.cols.partNo >= 0 ? cellText(row[hdr.cols.partNo]) : '',
          description: hdr.cols.description >= 0 ? cellText(row[hdr.cols.description]) : '',
          specifications: hdr.cols.specifications >= 0 ? cellText(row[hdr.cols.specifications]) : '',
          make: hdr.cols.make >= 0 ? cellText(row[hdr.cols.make]) : '',
          qty: parseQty(qtyRaw),
          qtyText: cellText(qtyRaw),
          remarks: hdr.cols.remarks >= 0 ? cellText(row[hdr.cols.remarks]) : '',
          sourceRow: r + 1,
        });
      }
      if (!rows.length) continue;

      const docInfo = extractDocInfo(aoa, hdr.headerRow);
      if (!docInfo.projectKey) {
        warnings.push('No SPN/PN project key found in the document header — cannot verify this LLDBO document matches the loaded Item Master\'s project.');
      }
      const noPartNumber = rows.filter(function (r) { return !r.partNo; }).length;
      if (noPartNumber) {
        warnings.push(noPartNumber + ' row(s) have no Part No yet (not-yet-specified placeholders) — not checked against the Item Master.');
      }

      return {
        kind: 'lldbo',
        sheetName: sheetName,
        rows: rows,
        projectKey: docInfo.projectKey,
        customer: docInfo.customer,
        columns: hdr.cols,
        warnings: warnings,
      };
    }
    return null; // not an LLDBO document
  }

  return { lldboParser: { parse: parse, parseQty: parseQty, findHeader: findHeader } };
});
