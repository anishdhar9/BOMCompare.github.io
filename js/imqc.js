/*
 * imqc.js — Item Master data-quality checks. Pure logic (no DOM), operating
 * on the structure produced by itemmaster.js. Runs on the Item Master alone
 * — no CAD file needed — catching manual-edit mistakes made directly in
 * Vault/ERP rather than CAD-vs-BOM drift (that's compare.js's job).
 *
 * Ported from a standalone "BOM QC Inspector" widget built for this
 * organization's Vault item BOM export; adapted to reuse itemmaster.js's
 * already-parsed rows and to degrade gracefully when an optional column
 * (Entity Icon) isn't present in a given export, rather than flagging every
 * row as failing.
 *
 * Produces: { c1, c2, c3, c4 }, each either
 *   { applicable: true, fail: [...] }  or
 *   { applicable: false, reason: string }
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const END_OF_LINE_NUMBER = '7-909-00001';

  function isRoot(row) {
    return Array.isArray(row.path) && row.path.length === 0;
  }

  function rowOrderText(row) {
    return Array.isArray(row.path) ? row.path.join('.') : '';
  }

  // Check 1: Producer / Producer Number vs Description, FG/top-level row(s) only.
  function checkProducerMatch(im) {
    if (!im.hasProducer) {
      return { applicable: false, reason: 'No "Producer" / "Producer Number" column found in this export.' };
    }
    const fail = [];
    let applicable = 0;
    for (const row of im.rows) {
      if (!isRoot(row)) continue;
      const producer = row.producer, prodNum = row.producerNumber;
      if (!producer && !prodNum) continue;
      applicable++;
      const desc = (row.description || '').toUpperCase();
      let ok = true;
      if (producer) ok = ok && desc.indexOf(producer.toUpperCase()) !== -1;
      if (prodNum) ok = ok && desc.indexOf(prodNum.toUpperCase()) !== -1;
      if (!ok) {
        fail.push({
          number: row.number,
          rowOrder: rowOrderText(row) || '-',
          producer: producer || '—',
          producerNumber: prodNum || '—',
          description: row.description || '(blank)',
        });
      }
    }
    if (applicable === 0) {
      return { applicable: false, reason: 'No top-level row has Producer / Producer Number set.' };
    }
    return { applicable: true, applicableCount: applicable, fail: fail };
  }

  // Check 2: the "END OF LINE" row must carry the org's fixed part number and
  // a whole-number Row Order (no decimal level).
  function checkEndOfLine(im) {
    const eolRows = im.rows.filter(function (row) {
      const t = ((row.title || '') + ' ' + (row.description || '')).toUpperCase();
      return t.indexOf('END OF LINE') !== -1;
    });
    const fail = [];
    if (!eolRows.length) {
      fail.push({ number: '—', rowOrder: '—', issue: 'No "END OF LINE" entry found in the BOM.' });
    } else {
      for (const row of eolRows) {
        const ro = rowOrderText(row);
        const issues = [];
        if (row.number !== END_OF_LINE_NUMBER) issues.push('Number is "' + row.number + '", expected "' + END_OF_LINE_NUMBER + '"');
        if (ro.indexOf('.') !== -1) issues.push('Row Order "' + ro + '" is a decimal level, expected a whole number');
        if (issues.length) fail.push({ number: row.number, rowOrder: ro || '-', issue: issues.join('; ') });
      }
    }
    return { applicable: true, found: eolRows.length, fail: fail };
  }

  // Check 3: Quantity (display, e.g. "2 Each") vs Item Qty (numeric) must agree.
  function checkQuantityVsItemQty(im) {
    if (im.rows.every(function (r) { return r.itemQty === null && r.quantity === null; })) {
      return { applicable: false, reason: 'Neither "Item Qty" nor "Quantity" column found in this export.' };
    }
    const fail = [];
    for (const row of im.rows) {
      if (row.quantity === null && row.itemQty === null) continue;
      if (row.quantity === null || row.itemQty === null || row.quantity !== row.itemQty) {
        fail.push({
          number: row.number,
          rowOrder: rowOrderText(row) || '-',
          title: row.title,
          quantity: row.quantityText || '(blank)',
          itemQty: row.itemQty === null ? '(blank)' : String(row.itemQty),
        });
      }
    }
    return { applicable: true, fail: fail };
  }

  // Check 4: Entity Icon should read "Normal" on every row (only meaningful
  // when the export actually carries this column).
  function checkEntityIcon(im) {
    if (!im.hasEntityIcon) {
      return { applicable: false, reason: 'No "Entity Icon" column found in this export.' };
    }
    const fail = [];
    for (const row of im.rows) {
      const icon = (row.entityIcon || '').toUpperCase();
      if (icon !== 'NORMAL') {
        fail.push({ number: row.number, rowOrder: rowOrderText(row) || '-', icon: row.entityIcon || '(blank)' });
      }
    }
    return { applicable: true, fail: fail };
  }

  function runChecks(im) {
    return {
      c1: checkProducerMatch(im),
      c2: checkEndOfLine(im),
      c3: checkQuantityVsItemQty(im),
      c4: checkEntityIcon(im),
      total: im.rows.length,
    };
  }

  return { imQc: { runChecks: runChecks, END_OF_LINE_NUMBER: END_OF_LINE_NUMBER } };
});
