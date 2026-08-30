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
 * Produces: { c1, c2, c3, c4, c5, c6, c7, c8, c9 }, each either
 *   { applicable: true, fail: [...] }  or
 *   { applicable: false, reason: string }
 * c9 additionally carries { errorCount, warnCount } so the UI can show amber
 * for "not yet certified" without calling it a failure.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const END_OF_LINE_NUMBER = '7-909-00001';

  // The "END OF LINE" sentinel that closes a released BOM. Matched on text
  // rather than on END_OF_LINE_NUMBER, because Check 2's whole job is to
  // catch the case where that row carries the WRONG number. It is a marker,
  // not a real part, so material checks skip it (it legitimately carries a
  // placeholder material such as "."). Title/Description must EQUAL "END OF
  // LINE" exactly (trimmed, case-insensitive) rather than merely contain it
  // — an unanchored substring test used to both false-positive (Check 2
  // failing a real part like "Sensor - End of Line Detector" for not
  // carrying the sentinel number/Row Order) and false-negative (that same
  // real part silently skipped by checks 3/5/6/7/9 as if it WERE the
  // marker).
  function isEndOfLine(row) {
    const title = String(row.title || '').trim().toUpperCase();
    const desc = String(row.description || '').trim().toUpperCase();
    return title === 'END OF LINE' || desc === 'END OF LINE';
  }

  function isRoot(row) {
    return Array.isArray(row.path) && row.path.length === 0;
  }

  function rowOrderText(row) {
    return Array.isArray(row.path) ? row.path.join('.') : '';
  }

  // Maps row -> its true parent row, resolved POSITIONALLY (file order plus
  // Row Order depth) rather than by Row Order string. Real exports reuse a
  // position for adjacent siblings, so a path-string lookup attributes a
  // child to whichever branch came first. Mirrors compare.js's
  // buildParentIndex — kept as its own copy to keep imqc.js dependency-free
  // (see the file header); change both together.
  function buildPathIndex(rows) {
    var parentOfRow = new Map();
    var stack = []; // {depth, row}
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!Array.isArray(row.path)) continue;
      var depth = row.path.length;
      while (stack.length && stack[stack.length - 1].depth >= depth) stack.pop();
      if (stack.length) parentOfRow.set(row, stack[stack.length - 1].row);
      stack.push({ depth: depth, row: row });
    }
    return parentOfRow;
  }

  // The immediate parent row for `row`, or null when `row` is the root row,
  // has no path, or its parent position is a gap in this export.
  function parentOf(pathIndex, row) {
    var parent = pathIndex.get(row);
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
    const eolRows = im.rows.filter(isEndOfLine);
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
  // Only meaningful when BOTH columns exist -- with just one of them there is
  // nothing to cross-check, so the check reports "not applicable" rather than
  // flagging every row (an Item-Qty-less export previously produced a storm of
  // false mismatches). Note a per-unit column ("QTY per Unit"/"Unit Qty") is
  // deliberately NOT treated as Item Qty by the parser, so it can't stand in
  // for a real Item Qty here.
  function checkQuantityVsItemQty(im, pathIndex) {
    const hasItemQty = im.hasItemQty !== undefined ? im.hasItemQty : im.rows.some(function (r) { return r.itemQty !== null; });
    const hasQuantity = im.hasQuantity !== undefined ? im.hasQuantity : im.rows.some(function (r) { return r.quantity !== null; });
    if (!hasItemQty || !hasQuantity) {
      const reason = !hasItemQty && !hasQuantity
        ? 'Needs both a "Quantity" and an "Item Qty" column to cross-check; this export has neither.'
        : !hasItemQty
          ? 'Needs an "Item Qty" column to cross-check against "Quantity"; this export has only "Quantity".'
          : 'Needs a "Quantity" column to cross-check against "Item Qty"; this export has only "Item Qty".';
      return { applicable: false, reason: reason };
    }
    const fail = [];
    for (const row of im.rows) {
      if (isEndOfLine(row)) continue; // ERP completeness marker, not a part
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
      if (isEndOfLine(row)) continue; // ERP completeness marker, not a part
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

  // Only "7-" numbers are made/controlled at this site. A leading 1-, 2-, 3-,
  // 5- ... is procured from another of the organization's locations, so a
  // missing or differing field on one is not something this site can act on;
  // those rows are excluded from the data-quality and CAD-comparison checks
  // rather than reported as unfixable noise.
  const MANUFACTURED_PART_RE = /^7-/;

  function isOwnPart(number) {
    return MANUFACTURED_PART_RE.test(String(number || '').trim());
  }

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
      if (isEndOfLine(row)) continue; // ERP completeness marker, not a part
      if (!isOwnPart(row.number)) continue; // procured at another site — not actionable here
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

  // A row is an "assembly" if it is some other row's parent. Derived from the
  // positional parent index rather than by matching Row Order path prefixes,
  // so a position reused by two different parts cannot make a leaf look like
  // an assembly (or vice versa). Returns a Set of row objects.
  function buildAssemblySet(rows, pathIndex) {
    const idx = pathIndex || buildPathIndex(rows);
    const assemblies = new Set();
    for (const row of rows) {
      const parent = idx.get(row);
      if (parent) assemblies.add(parent);
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
    const assemblies = buildAssemblySet(im.rows, pathIndex);
    const fail = [];
    for (const row of im.rows) {
      if (!Array.isArray(row.path)) continue;
      if (PURCHASED_PART_RE.test(row.number)) continue; // handled by the Bought-Out Parts panel instead
      if (!isOwnPart(row.number)) continue; // procured at another site — not actionable here
      if (isEndOfLine(row)) continue; // sentinel marker row, not a real part
      if (assemblies.has(row)) continue; // assembly, material not expected
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

  // Check 7: Revision Consistency. The same part number can legitimately
  // appear at multiple BOM positions (used in more than one assembly); when
  // it does, every occurrence should carry the same Revision value. A
  // mismatch usually means one assembly's occurrence was updated to a newer
  // released revision and the others weren't picked up. Values are compared
  // directly (trimmed/uppercased) — revisions are simple codes (e.g. "0",
  // "1", "A", "B"), not something needing a grade-equivalence lookup like
  // material.
  function checkRevisionConsistency(im, pathIndex) {
    if (!im.hasRevision) {
      return { applicable: false, reason: 'No "Revision" column found in this export.' };
    }
    var byPn = new Map(); // normalized PN -> [row, ...]
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      if (isEndOfLine(row)) continue; // ERP completeness marker, not a part
      var key = String(row.number).trim().toUpperCase();
      if (!key) continue;
      if (!byPn.has(key)) byPn.set(key, []);
      byPn.get(key).push(row);
    }
    var fail = [];
    byPn.forEach(function (rows) {
      if (rows.length < 2) return;
      var distinct = new Set();
      rows.forEach(function (r) {
        var rv = (r.revision || '').trim().toUpperCase();
        if (rv) distinct.add(rv);
      });
      if (distinct.size < 2) return; // all occurrences agree (or too few have a value to compare)
      rows.forEach(function (row2) {
        var others = rows.filter(function (x) { return x !== row2; })
          .map(function (x) { return (x.revision || '(blank)') + ' at Row Order ' + (rowOrderText(x) || '-'); })
          .join('; ');
        fail.push(withLocation({
          number: row2.number,
          rowOrder: rowOrderText(row2) || '-',
          revision: row2.revision || '(blank)',
          conflictsWith: others,
        }, pathIndex, row2));
      });
    });
    return { applicable: true, fail: fail };
  }

  // The full chain of ancestors down to `row`, from the top-level assembly
  // inwards, as "NUMBER (Title) › NUMBER (Title) › …". parentOf() gives only
  // the immediate parent; for a sketch part that must never have been released
  // you need the whole path to find how it got in.
  function ancestorTrail(pathIndex, row) {
    if (!Array.isArray(row.path) || row.path.length === 0) return '';
    var parts = [];
    // Walk up the real parent chain, then reverse: the trail reads top-level
    // assembly first. The root row itself (depth 0) is not part of the trail.
    for (var anc = pathIndex.get(row); anc; anc = pathIndex.get(anc)) {
      if (!Array.isArray(anc.path) || anc.path.length === 0) break;
      parts.push(anc.number + (anc.title ? ' (' + anc.title + ')' : ''));
    }
    parts.reverse();
    return parts.join(' › ');
  }

  // Check 8: sketch/placeholder parts that must never reach the Item Master.
  // 7-333-* numbers are the "rough sketch" equivalent in this organization's
  // 3D modelling: they may exist in CAD, but only ever as REFERENCE. One
  // reaching the released Item Master is a serious release error, so every hit
  // is reported with its Row # and full parent trail to trace how it got in.
  // Deliberately Item-Master-only — their presence in CAD is expected.
  //
  // Add further banned prefixes here if more families are ever introduced.
  var SKETCH_PART_RE = /^7-333-/i;

  function checkSketchParts(im, pathIndex) {
    var fail = [];
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      if (!SKETCH_PART_RE.test(String(row.number).trim())) continue;
      fail.push(withLocation({
        number: row.number,
        rowOrder: rowOrderText(row) || '-',
        title: row.title || '',
        state: row.state || '',
        trail: ancestorTrail(pathIndex, row) || '(top level)',
      }, pathIndex, row));
    }
    return { applicable: true, fail: fail };
  }

  // Check 9: lifecycle state. A released BOM should contain Certified (or
  // Released) items only. Obsolete / Invalid / Phased Out are hard errors —
  // the part was released against a dead revision. "New" is a warning: it
  // occurs in genuine released exports, so it is surfaced without being
  // treated as a failure.
  var BAD_STATES = ['obsolete', 'invalid', 'phased out', 'phased-out', 'phasedout'];
  var WARN_STATES = ['new', 'work in progress', 'in work'];

  function classifyState(value) {
    var s = String(value || '').trim().toLowerCase();
    if (!s) return 'blank';
    for (var i = 0; i < BAD_STATES.length; i++) if (s === BAD_STATES[i] || s.indexOf(BAD_STATES[i]) === 0) return 'error';
    for (var j = 0; j < WARN_STATES.length; j++) if (s === WARN_STATES[j] || s.indexOf(WARN_STATES[j]) === 0) return 'warn';
    return 'ok';
  }

  function checkItemState(im, pathIndex) {
    if (!im.hasState) {
      return { applicable: false, reason: 'No "State" column found in this export.' };
    }
    var fail = [];
    var errorCount = 0, warnCount = 0;
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      if (isEndOfLine(row)) continue; // ERP completeness marker, not a part
      var kind = classifyState(row.state);
      if (kind !== 'error' && kind !== 'warn') continue;
      if (kind === 'error') errorCount++; else warnCount++;
      fail.push(withLocation({
        number: row.number,
        rowOrder: rowOrderText(row) || '-',
        title: row.title || '',
        state: row.state || '(blank)',
        severity: kind === 'error' ? 'ERROR — must not be released' : 'warning — not yet certified',
      }, pathIndex, row));
    }
    // Sorted so the hard errors are read first.
    fail.sort(function (a, b) { return (a.severity < b.severity ? -1 : (a.severity > b.severity ? 1 : 0)); });
    return { applicable: true, fail: fail, errorCount: errorCount, warnCount: warnCount };
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
      c7: checkRevisionConsistency(im, pathIndex),
      c8: checkSketchParts(im, pathIndex),
      c9: checkItemState(im, pathIndex),
      total: im.rows.length,
    };
  }

  return {
    imQc: {
      runChecks: runChecks,
      boughtOutParts: boughtOutParts,
      END_OF_LINE_NUMBER: END_OF_LINE_NUMBER,
      PURCHASED_PART_RE: PURCHASED_PART_RE,
      MANUFACTURED_PART_RE: MANUFACTURED_PART_RE,
      SKETCH_PART_RE: SKETCH_PART_RE,
      isOwnPart: isOwnPart,
      isEndOfLine: isEndOfLine,
      buildAssemblySet: buildAssemblySet,
      buildPathIndex: buildPathIndex,
      parentOf: parentOf,
      ancestorTrail: ancestorTrail,
      classifyState: classifyState,
      blank: blank,
    },
  };
});
