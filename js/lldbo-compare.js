/*
 * lldbo-compare.js — compares a parsed LLDBO (Long Lead Direct Bought Out)
 * list against the Item Master. Long-lead parts are released to
 * procurement ahead of the normal BOM release to account for supplier lead
 * times; this check verifies each one actually made it into the released
 * Item Master with a matching quantity — catching the process failure
 * where an early release never got captured, so the part quietly never
 * gets ordered through the normal channel either.
 *
 * Pure logic (no DOM). `indexItemMaster` is injected (it's compare.js's
 * function) rather than required directly, matching this codebase's
 * dependency-injection style (see itemmaster.js's `parse(workbook, XLSX)`,
 * cad-leveled.js's XLSX parameter) instead of a hard load-order coupling.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function normNumber(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim().toUpperCase();
  }

  function compareLldbo(lldbo, im, indexItemMaster) {
    const imIndex = indexItemMaster(im);

    // Aggregate LLDBO rows by part number — duplicates are summed (e.g. the
    // same catalog motor used in two different assemblies, each needing 1,
    // means 2 are expected in total; verified as a real, legitimate pattern
    // in the sample data, not a data error).
    const byPn = new Map(); // PN -> { number, rows:[...], totalQty, hasQty }
    let noPartNumber = 0;
    for (const row of lldbo.rows) {
      const pn = normNumber(row.partNo);
      if (!pn) { noPartNumber++; continue; }
      if (!byPn.has(pn)) byPn.set(pn, { number: row.partNo, rows: [], totalQty: 0, hasQty: true });
      const entry = byPn.get(pn);
      entry.rows.push(row);
      if (row.qty === null) entry.hasQty = false;
      else entry.totalQty += row.qty;
    }

    const missingFromIm = [];
    const qtyMismatches = [];
    for (const [pn, entry] of byPn) {
      const descriptions = entry.rows.map(function (r) { return r.description; }).filter(Boolean);
      const sourceRows = entry.rows.map(function (r) { return r.sourceRow; }).filter(Boolean).join(', ');
      if (!imIndex.byNumber.has(pn)) {
        missingFromIm.push({
          number: entry.number,
          description: descriptions.join(' / '),
          qtyText: entry.rows.map(function (r) { return r.qtyText; }).filter(Boolean).join(' + '),
          sourceRow: sourceRows,
          rows: entry.rows,
        });
        continue;
      }
      if (!entry.hasQty) continue; // present, but no comparable quantity ("NA")
      const imTotal = imIndex.totals.has(pn) ? imIndex.totals.get(pn) : null;
      if (imTotal === null) continue; // not computable on the Item Master side
      if (Math.abs(imTotal - entry.totalQty) > 1e-9) {
        const breakdown = imIndex.breakdowns.get(pn) || [];
        const foundUnder = breakdown
          .map(function (b) { return b.parentNumber ? (b.parentNumber + (b.parentTitle ? ' (' + b.parentTitle + ')' : '')) : ''; })
          .filter(Boolean).join(' + ');
        qtyMismatches.push({
          number: entry.number,
          description: descriptions.join(' / '),
          lldboQty: entry.totalQty,
          imQty: imTotal,
          sourceRow: sourceRows,
          foundUnder: foundUnder,
          rows: entry.rows,
        });
      }
    }

    return {
      totalLldboItems: byPn.size,
      noPartNumberCount: noPartNumber,
      missingFromIm: missingFromIm,
      qtyMismatches: qtyMismatches,
      projectKeyMismatch: (lldbo.projectKey && im.projectKey && lldbo.projectKey.pn !== im.projectKey.pn)
        ? { lldbo: lldbo.projectKey, im: im.projectKey }
        : null,
    };
  }

  return { lldboCompare: { compareLldbo: compareLldbo } };
});
