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
 * Produces: { c1, c2, c3, c4, c5, c6 }, each either
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

  // Maps 'path.key' -> row, for O(1) parent lookup. Same key shape as
  // compare.js's indexItemMaster byPath (first occurrence at a position
  // wins), reimplemented locally to keep imqc.js dependency-free.
  function buildPathIndex(rows) {
    var byPath = new Map();
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (Array.isArray(row.path) && !byPath.has(row.path.join('.'))) byPath.set(row.path.join('.'), row);
    }
    return byPath;
  }

  // The immediate parent row (by Row Order path) for `row`, or null when
  // `row` is the root row, has no path, or its parent position is a gap in
  // this export.
  function parentOf(pathIndex, row) {
    if (!Array.isArray(row.path) || row.path.length === 0) return null;
    var parent = pathIndex.get(row.path.slice(0, -1).join('.'));
    return parent ? { number: parent.number, title: parent.title || '' } : null;
  }

  // Adds Row # (the row's position in the uploaded file) and Parent Number /
  // Parent Title (from Row Order) to a fail entry — every check's export
  // columns include these so a flagged row can actually be found and placed
  // in context without re-deriving it by hand.
  function withLocation(fields, pathIndex, row) {
    var p = pathIndex ? parentOf(pathIndex, row) : null;
    fields.sourceRow = row.sourceRow || '';
    fields.parentNumber = p ? p.number : '';
    fields.parentTitle = p ? p.title : '';
    return fields;
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
      const issues = [];
      if (producer && desc.indexOf(producer.toUpperCase()) === -1) issues.push('Producer "' + producer + '" not found in Description');
      if (prodNum && desc.indexOf(prodNum.toUpperCase()) === -1) issues.push('Producer Number "' + prodNum + '" not found in Description');
      if (issues.length) {
        fail.push(withLocation({
          number: row.number,
          rowOrder: rowOrderText(row) || '-',
          producer: producer || '—',
          producerNumber: prodNum || '—',
          description: row.description || '(blank)',
          issue: issues.join('; '),
        }, null, row));
      }
    }
    if (applicable === 0) {
      return { applicable: false, reason: 'No top-level row has Producer / Producer Number set.' };
    }
    return { applicable: true, applicableCount: applicable, fail: fail };
  }

  // Check 2: the "END OF LINE" row must carry the org's fixed part number and
  // a whole-number Row Order (no decimal level).
  function checkEndOfLine(im, pathIndex) {
    const eolRows = im.rows.filter(function (row) {
      const t = ((row.title || '') + ' ' + (row.description || '')).toUpperCase();
      return t.indexOf('END OF LINE') !== -1;
    });
    const fail = [];
    if (!eolRows.length) {
      fail.push({ number: '—', rowOrder: '—', issue: 'No "END OF LINE" entry found in the BOM.', sourceRow: '', parentNumber: '', parentTitle: '' });
    } else {
      for (const row of eolRows) {
        const ro = rowOrderText(row);
        const issues = [];
        if (row.number !== END_OF_LINE_NUMBER) issues.push('Number is "' + row.number + '", expected "' + END_OF_LINE_NUMBER + '"');
        if (ro.indexOf('.') !== -1) issues.push('Row Order "' + ro + '" is a decimal level, expected a whole number');
        if (issues.length) fail.push(withLocation({ number: row.number, rowOrder: ro || '-', issue: issues.join('; ') }, pathIndex, row));
      }
    }
    return { applicable: true, found: eolRows.length, fail: fail };
  }

  // Check 3: Quantity (display, e.g. "2 Each") vs Item Qty (numeric) must agree.
  function checkQuantityVsItemQty(im, pathIndex) {
    if (im.rows.every(function (r) { return r.itemQty === null && r.quantity === null; })) {
      return { applicable: false, reason: 'Neither "Item Qty" nor "Quantity" column found in this export.' };
    }
    const fail = [];
    for (const row of im.rows) {
      if (row.quantity === null && row.itemQty === null) continue;
      if (row.quantity === null || row.itemQty === null || row.quantity !== row.itemQty) {
        fail.push(withLocation({
          number: row.number,
          rowOrder: rowOrderText(row) || '-',
          title: row.title,
          quantity: row.quantityText || '(blank)',
          itemQty: row.itemQty === null ? '(blank)' : String(row.itemQty),
        }, pathIndex, row));
      }
    }
    return { applicable: true, fail: fail };
  }

  // Check 4: Entity Icon should read "Normal" on every row (only meaningful
  // when the export actually carries this column).
  function checkEntityIcon(im, pathIndex) {
    if (!im.hasEntityIcon) {
      return { applicable: false, reason: 'No "Entity Icon" column found in this export.' };
    }
    const fail = [];
    for (const row of im.rows) {
      const icon = (row.entityIcon || '').toUpperCase();
      if (icon !== 'NORMAL') {
        fail.push(withLocation({ number: row.number, rowOrder: rowOrderText(row) || '-', icon: row.entityIcon || '(blank)' }, pathIndex, row));
      }
    }
    return { applicable: true, fail: fail };
  }

  // Org convention (confirmed across every sample seen): a "X-999-nnnnn"
  // number is a purchased/catalog part (bearings, screws, o-rings...),
  // everything else is manufactured in-house.
  const PURCHASED_PART_RE = /^\d-999-/;

  function blank(v) {
    const s = (v || '').trim();
    return s === '' || s === '-';
  }

  // Check 5: Title / Description completeness. Purchased/catalog parts are
  // lenient (only flagged when BOTH are missing — e.g. "SCREW - HEX HEAD
  // DIN 933" needs no Description to be identifiable); every other part is
  // flagged when EITHER is missing, and which one is reported so the export
  // can highlight just that field.
  function checkTitleDescription(im, pathIndex) {
    const fail = [];
    for (const row of im.rows) {
      const titleBlank = blank(row.title);
      const descBlank = blank(row.description);
      if (!titleBlank && !descBlank) continue;
      const purchased = PURCHASED_PART_RE.test(row.number);
      let kind = null;
      if (titleBlank && descBlank) kind = 'both-missing';
      else if (!purchased) kind = titleBlank ? 'title-missing' : 'description-missing';
      if (!kind) continue; // purchased part, only one field blank -> allowed
      fail.push(withLocation({
        number: row.number,
        rowOrder: rowOrderText(row) || '-',
        title: row.title || '(blank)',
        description: row.description || '(blank)',
        kind: kind,
      }, pathIndex, row));
    }
    return { applicable: true, fail: fail };
  }

  // A Row Order path is an "assembly" if some other row's path is a strict
  // child of it (starts with path + '.'); the root ('-' -> []) always
  // counts. Same principle compare.js's indexItemMaster uses for childSets,
  // reimplemented locally to keep imqc.js dependency-free.
  function buildAssemblyPathSet(rows) {
    const paths = [];
    for (const row of rows) {
      if (Array.isArray(row.path)) paths.push(row.path.join('.'));
    }
    const assemblies = new Set(['']); // root
    for (const p of paths) {
      if (p === '' || assemblies.has(p)) continue;
      const prefix = p + '.';
      for (const q of paths) {
        if (q !== p && q.indexOf(prefix) === 0) { assemblies.add(p); break; }
      }
    }
    return assemblies;
  }

  // Check 6: Material should be populated on every non-assembly row.
  // Assemblies/weldments legitimately have no material of their own.
  // Purchased/catalog parts (7-999-*) are EXCLUDED here — verified on real
  // data that 105 of 111 flagged rows were purchased parts (bearings,
  // wheels, cylinders...) where a blank material is often not a real gap,
  // drowning out genuine manufactured-part gaps. They get their own
  // always-visible "Bought-Out Parts" reference panel instead (below),
  // which is not counted toward any flagged/action total.
  function checkMaterial(im, pathIndex) {
    if (!im.hasMaterial) {
      return { applicable: false, reason: 'No "Material" column found in this export.' };
    }
    if (!im.hasPaths) {
      return { applicable: false, reason: 'No "Row Order" column found — cannot tell assemblies from parts.' };
    }
    const assemblies = buildAssemblyPathSet(im.rows);
    const fail = [];
    for (const row of im.rows) {
      if (!Array.isArray(row.path)) continue;
      if (PURCHASED_PART_RE.test(row.number)) continue; // handled by the Bought-Out Parts panel instead
      if (assemblies.has(row.path.join('.'))) continue; // assembly, material not expected
      if (blank(row.material)) {
        fail.push(withLocation({ number: row.number, rowOrder: rowOrderText(row) || '-', title: row.title }, pathIndex, row));
      }
    }
    return { applicable: true, fail: fail };
  }

  // Every purchased/catalog part (7-999-*) in the Item Master, for the
  // always-visible reference panel — not a pass/fail check, just a listing
  // (js/material-compare.js fills in the CAD-side material when a CAD
  // source that carries material is also loaded).
  function boughtOutParts(im) {
    var seen = new Set(); // same part can occur at several BOM positions; list it once
    var out = [];
    var pathIndex = buildPathIndex(im.rows);
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      if (!PURCHASED_PART_RE.test(row.number)) continue;
      var key = String(row.number).trim().toUpperCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(withLocation({
        number: row.number,
        title: row.title,
        imMaterial: row.material || '',
        missingMaterial: im.hasMaterial ? blank(row.material) : false,
      }, pathIndex, row));
    }
    return out;
  }

  function runChecks(im) {
    var pathIndex = buildPathIndex(im.rows);
    return {
      c1: checkProducerMatch(im),
      c2: checkEndOfLine(im, pathIndex),
      c3: checkQuantityVsItemQty(im, pathIndex),
      c4: checkEntityIcon(im, pathIndex),
      c5: checkTitleDescription(im, pathIndex),
      c6: checkMaterial(im, pathIndex),
      total: im.rows.length,
    };
  }

  return {
    imQc: {
      runChecks: runChecks,
      boughtOutParts: boughtOutParts,
      END_OF_LINE_NUMBER: END_OF_LINE_NUMBER,
      PURCHASED_PART_RE: PURCHASED_PART_RE,
      buildAssemblyPathSet: buildAssemblyPathSet,
      buildPathIndex: buildPathIndex,
      parentOf: parentOf,
      blank: blank,
    },
  };
});
