/*
 * revision-compare.js — compares CAD BOM revision against Item Master
 * revision for the same part number.
 *
 * Not every CAD source carries revision: it's optional/user-configurable in
 * Vault exports, the same situation as material (see material-compare.js) —
 * this check uses whichever loaded CAD source actually has revision data
 * (`hasRevision`), and is only applicable when at least one does and the
 * Item Master itself has a Revision column.
 *
 * Unlike material, there's no naming-convention ambiguity to normalize away
 * — revisions are simple short codes (verified on a real Vault "Uses" PDF
 * export: plain integers "0", "1", "2"...; other organizations may use
 * letters "A"/"B") — so this is a direct value comparison (trimmed,
 * case-insensitive), not a grade-equivalence lookup like material's.
 *
 * Pure logic (no DOM).
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(require('./imqc.js').imQc);
  } else {
    const bc = root.BOMCompare || {};
    root.BOMCompare = Object.assign(bc, factory(bc.imQc));
  }
})(typeof self !== 'undefined' ? self : this, function (imQc) {
  'use strict';

  function normRevision(v) {
    return String(v || '').trim().toUpperCase();
  }

  function revisionsMatch(a, b) {
    var na = normRevision(a), nb = normRevision(b);
    if (!na || !nb) return false;
    return na === nb;
  }

  // Picks the first loaded CAD source that actually carries revision data
  // and a lookup of its first-seen revision per PN. Mirrors
  // material-compare.js's cadMaterialByPn.
  function cadRevisionByPn(cadSources) {
    for (var i = 0; i < cadSources.length; i++) {
      var src = cadSources[i];
      if (!src.hasRevision) continue;
      var map = new Map();
      var any = false;
      for (var j = 0; j < src.items.length; j++) {
        var it = src.items[j];
        var pn = String(it.number || '').trim().toUpperCase();
        var rev = (it.revision || '').trim();
        if (!pn || !rev) continue;
        any = true;
        if (!map.has(pn)) map.set(pn, rev);
      }
      if (any) return { source: src, byPn: map };
    }
    return null;
  }

  // cadSources: the array passed to compareAll (0-2 CAD sources).
  function compareRevision(cadSources, im) {
    if (!im.hasRevision) {
      return { applicable: false, reason: 'No "Revision" column found in this Item Master export.' };
    }
    var cad = cadRevisionByPn(cadSources || []);
    if (!cad) {
      return { applicable: false, reason: 'No loaded CAD source carries revision data.' };
    }

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
        mismatches.push({
          number: row.number,
          title: row.title,
          imRevision: row.revision,
          cadRevision: cadRev,
          sourceRow: row.sourceRow,
          parentNumber: parent ? parent.number : '',
          parentTitle: parent ? parent.title : '',
        });
      }
    }

    return {
      applicable: true,
      cadSourceFileName: cad.source.fileName || '',
      mismatches: mismatches,
    };
  }

  return {
    revisionCompare: {
      compareRevision: compareRevision,
      revisionsMatch: revisionsMatch,
    },
  };
});
