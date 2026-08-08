/*
 * revision-compare.js — compares CAD BOM revision against Item Master
 * revision for the same part number.
 *
 * Not every CAD source carries revision: it's optional/user-configurable in
 * Vault exports, the same situation as material (see material-compare.js) —
 * this check merges revision data across every loaded CAD source that has
 * it (`hasRevision`), and is only applicable when at least one does and the
 * Item Master itself has a Revision column.
 *
 * Unlike material, there's no naming-convention ambiguity to normalize away
 * — revisions are simple short codes (verified on real Vault/Inventor
 * exports: plain integers "0".."8" on the Item Master side; other
 * organizations may use letters "A"/"B") — so this is a direct value
 * comparison (trimmed, case-insensitive), not a grade-equivalence lookup
 * like material's. The CAD side is sparser and also carries Inventor
 * Content-Center placeholder text on purchased parts ("ANY"/"NONE"/"-",
 * stable per-part metadata, not noise) — compareRevisionOrder() below
 * never attempts to order that, only clean integers.
 *
 * Pure logic (no DOM).
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(require('./imqc.js').imQc, require('./compare.js'));
  } else {
    const bc = root.BOMCompare || {};
    root.BOMCompare = Object.assign(bc, factory(bc.imQc, bc));
  }
})(typeof self !== 'undefined' ? self : this, function (imQc, compareLib) {
  'use strict';

  function normRevision(v) {
    return String(v || '').trim().toUpperCase();
  }

  function revisionsMatch(a, b) {
    var na = normRevision(a), nb = normRevision(b);
    if (!na || !nb) return false;
    return na === nb;
  }

  // Real Vault/Inventor exports show two very different kinds of revision
  // value on the CAD side: a genuine plain integer, or Inventor Content-
  // Center placeholder text on purchased parts ("ANY", "NONE", "-" — stable
  // per-part metadata, not noise, verified recurring identically for the
  // same PN across independent assemblies). Ordering is only ever asserted
  // between two clean non-negative integers; anything else returns null
  // rather than guess — this is what keeps placeholder text from ever being
  // mis-classified as "older" or "newer".
  function isCleanRevision(v) {
    return /^\d+$/.test(normRevision(v));
  }

  function compareRevisionOrder(a, b) {
    if (!isCleanRevision(a) || !isCleanRevision(b)) return null;
    var ia = parseInt(normRevision(a), 10), ib = parseInt(normRevision(b), 10);
    return ia < ib ? -1 : (ia > ib ? 1 : 0);
  }

  // Merges revision data across EVERY loaded CAD source with hasRevision,
  // not just the first one that has any — a single sparse source (e.g. an
  // Inventor export whose REV column is only populated on a couple percent
  // of items) used to silently win over a more complete one just by being
  // checked first. For a given PN, the first value found is kept unless it's
  // not a clean comparable integer (see compareRevisionOrder) and a later
  // source has one — then it's upgraded. Two sources reporting two
  // DIFFERENT comparable revisions for the same PN is left as first-found-
  // wins; that's a genuine CAD-source inconsistency worth its own check
  // some day, not something to silently resolve here.
  function cadRevisionByPn(cadSources) {
    var map = new Map();
    var contributing = [];
    for (var i = 0; i < cadSources.length; i++) {
      var src = cadSources[i];
      if (!src.hasRevision) continue;
      var any = false;
      for (var j = 0; j < src.items.length; j++) {
        var it = src.items[j];
        var pn = String(it.number || '').trim().toUpperCase();
        var rev = (it.revision || '').trim();
        if (!pn || !rev) continue;
        any = true;
        var existing = map.get(pn);
        if (!existing || (!isCleanRevision(existing) && isCleanRevision(rev))) {
          map.set(pn, rev);
        }
      }
      if (any) contributing.push(src);
    }
    return map.size ? { sources: contributing, byPn: map } : null;
  }

  // cadSources: the array passed to compareAll (0-2 CAD sources).
  // opts (optional): { isIgnored(pn, checkKey) } from js/ignorelist-compare.js
  // and/or js/app.js's bought-out-parts toggle — same predicate shape
  // js/compare.js's compareAll() uses, checked here against 'revision'.
  function compareRevision(cadSources, im, opts) {
    if (!im.hasRevision) {
      return { applicable: false, reason: 'No "Revision" column found in this Item Master export.' };
    }
    var cad = cadRevisionByPn(cadSources || []);
    if (!cad) {
      return { applicable: false, reason: 'No loaded CAD source carries revision data.' };
    }
    var isIgnored = (opts && opts.isIgnored) || function () { return false; };
    var ignoredFindings = [];

    var pathIndex = imQc.buildPathIndex(im.rows);
    var mismatches = [];
    var seenPn = new Set(); // same part can occur at several BOM positions; report it once
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      var pnKey = String(row.number).trim().toUpperCase();
      if (!pnKey || seenPn.has(pnKey)) continue;
      if (!row.revision) continue; // nothing to compare on the IM side
      var cadRev = cad.byPn.get(pnKey);
      if (!cadRev) continue; // part not in this CAD source, or CAD has no revision for it
      if (!revisionsMatch(row.revision, cadRev)) {
        seenPn.add(pnKey);
        var parent = imQc.parentOf(pathIndex, row);
        // imBehindCad: the design has moved on in CAD past what's recorded
        // in the released Item Master — only ever asserted when both sides
        // are clean integers (see compareRevisionOrder), never guessed at
        // for placeholder text like "ANY"/"NONE". staleDesign is a ready-
        // to-display label so the on-screen table and the export sheet
        // (which share the same column list) both get it for free.
        var behind = compareRevisionOrder(row.revision, cadRev) === -1;
        mismatches.push({
          number: row.number,
          title: row.title,
          imRevision: row.revision,
          cadRevision: cadRev,
          imBehindCad: behind,
          staleDesign: behind ? 'Yes — CAD is newer' : '',
          sourceRow: row.sourceRow,
          parentNumber: parent ? parent.number : '',
          parentTitle: parent ? parent.title : '',
        });
      }
    }

    mismatches = compareLib.filterIgnoredFlat(mismatches, isIgnored, 'revision', ignoredFindings);

    return {
      applicable: true,
      cadSourceFileName: cad.sources.map(function (s) { return s.fileName || ''; }).filter(Boolean).join(', '),
      mismatches: mismatches,
      ignoredFindings: ignoredFindings,
    };
  }

  return {
    revisionCompare: {
      compareRevision: compareRevision,
      revisionsMatch: revisionsMatch,
      compareRevisionOrder: compareRevisionOrder,
    },
  };
});
