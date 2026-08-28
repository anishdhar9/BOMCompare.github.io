/*
 * titledesc-compare.js — compares the CAD BOM's Description against the Item
 * Master's Description for the same part number.
 *
 * Which Item Master field to compare against was settled from real data
 * rather than assumed: across a 2049-row Inventor export, the CAD Description
 * matched the Item Master's DESCRIPTION on 73% of shared parts but its TITLE
 * on only 1%. The Item Master's Title is the catalogue name ("Ventilation
 * Tube") while its Description carries the same free text the CAD model does
 * ("GCSmart 2010"), so Description-to-Description is the real pairing.
 *
 * Normalization is deliberately light — case and whitespace only, with
 * whitespace removed entirely rather than merely collapsed (the same thing
 * material-compare.js's stripFormatting does). The differences worth catching
 * sit in the punctuation and digits, which stay significant: "OD 539 X 4 THK."
 * vs "OD 539 X 3 THK.", "G 1/4" vs "G 1/8", "With 30MM SKIRTING" vs
 * "With 27MM SKIRTING". Removing whitespace outright absorbs the noise a mere
 * collapse would miss, where the stray space sits INSIDE a token:
 * "TANK - 300LTRS" vs "TANK - 300 LTRS", "GCPilot" vs "GC Pilot".
 *
 * Skipped rows: parts that are not this site's own "7-" numbers (procured
 * from other company locations, nothing actionable here), "X-999-*"
 * purchased/catalog parts (their descriptions are supplier text and differ
 * constantly — on the sample export 158 of them had no Item Master
 * description at all), the END OF LINE marker row, and any row where either
 * side is blank. Assemblies are NOT skipped: unlike material, an assembly
 * does carry a description.
 *
 * Pure logic (no DOM). `imQc` is injected, matching revision-compare.js.
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

  // Case-insensitive, whitespace removed. Everything else is significant.
  function normText(v) {
    return String(v || '').trim().toUpperCase().replace(/\s+/g, '');
  }

  function descriptionsMatch(a, b) {
    var na = normText(a), nb = normText(b);
    if (!na || !nb) return false;
    return na === nb;
  }

  // CAD conventionally appends extra text to some descriptions (most often
  // a material grade, e.g. "GSF PRO 180" -> "GSF PRO 180 AISI 304") that the
  // Item Master's Description never carries -- flooding the mismatch list
  // with one row per part that follows the convention. This is deliberately
  // an EXACT structural test, not a similarity/percentage threshold: the
  // shorter description must be a complete, unbroken prefix of the longer
  // one (CAD only ever ADDS text at the end; nothing in the shared portion
  // changed). A threshold-based "percentage of characters match" was
  // considered and rejected -- verified against real examples, "OD 539 X 4
  // THK." vs "OD 539 X 3 THK." (a genuine, must-catch dimension change)
  // shares MORE characters in a row (60%) than "GSF PRO 180" vs
  // "GSF PRO 180 AISI 304" (55%) does, so no single percentage threshold
  // can separate the two safely. An exact-prefix test has no such gap: any
  // edit inside the shared portion breaks it outright.
  function isAutoMatchedAppend(a, b) {
    var na = normText(a), nb = normText(b);
    if (!na || !nb || na === nb) return false; // exact matches are descriptionsMatch's job
    var shorter = na.length <= nb.length ? na : nb;
    var longer = na.length <= nb.length ? nb : na;
    if (longer.indexOf(shorter) !== 0) return false;
    // A digit immediately after the shared prefix means the "addition" is
    // extending a number's precision (e.g. "1.6" -> "1.63"), not appending
    // descriptive text -- that can be a genuine value change, not noise, so
    // it must stay flagged. Verified against real data: "DIA 21.3 X1.6" is a
    // literal prefix of "DIA 21.3 X1.63THK", and 1.6 vs 1.63 could be a real
    // dimension difference, not a benign convention like a material suffix.
    var nextChar = longer.charAt(shorter.length);
    if (nextChar >= '0' && nextChar <= '9') return false;
    return true;
  }

  // The Inventor BOM export specifically (source === 'leveled-sheet' &&
  // hasStructure -- the same signature app.js's cadSourceLabel() already
  // uses to identify it), and a lookup of its first-seen description per
  // PN. This is deliberately NOT "whichever loaded CAD source has any
  // non-empty description text": the Vault multi-level PDF has no real
  // Description column at all -- cad-leveled.js's PDF parsing path fills
  // `description` by echoing `title` -- so treating any non-empty text as
  // good enough let the PDF silently win whenever it loaded before the
  // Inventor export (both are technically "non-empty"), and "Load from
  // folder" always loads the PDF first. Verified on a real project: PDF
  // preferred by load order produced 498 spurious "differences" out of
  // ~620 shared parts; requiring the Inventor export produces the correct
  // ~24. The Inventor BOM export is the only source with genuine free-text
  // Description data (see the "not applicable" reason below, which already
  // documented this before this fix enforced it).
  function cadDescriptionByPn(cadSources) {
    var src = (cadSources || []).find(function (s) { return s.source === 'leveled-sheet' && s.hasStructure; });
    if (!src) return null;
    var map = new Map();
    var any = false;
    for (var j = 0; j < src.items.length; j++) {
      var it = src.items[j];
      var pn = String(it.number || '').trim().toUpperCase();
      var desc = (it.description || '').trim();
      if (!pn || !desc) continue;
      any = true;
      if (!map.has(pn)) map.set(pn, desc);
    }
    return any ? { source: src, byPn: map } : null;
  }

  // cadSources: the array passed to compareAll (0-2 CAD sources).
  function compareTitleDescription(cadSources, im) {
    var cad = cadDescriptionByPn(cadSources || []);
    if (!cad) {
      return {
        applicable: false,
        reason: 'No loaded CAD source carries description text. The Vault multi-level PDF does not; ' +
          'the Inventor BOM export does when the Description column is included.',
        mismatches: [],
        autoMatched: [],
      };
    }

    var pathIndex = imQc.buildPathIndex(im.rows);
    var mismatches = [];
    var autoMatched = []; // CAD-only-appends-extra-text cases -- see isAutoMatchedAppend
    var eligible = 0;
    var seenPn = new Set(); // same part can occur at several BOM positions; report it once
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      var pnKey = String(row.number).trim().toUpperCase();
      if (!pnKey || seenPn.has(pnKey)) continue;
      if (!imQc.isOwnPart(row.number)) continue;            // procured at another site
      if (imQc.PURCHASED_PART_RE.test(row.number)) continue; // supplier catalogue text
      if (imQc.isEndOfLine(row)) continue;                   // ERP completeness marker
      if (imQc.blank(row.description)) continue;             // Check 5 covers "IM description missing"
      var cadDesc = cad.byPn.get(pnKey);
      if (!cadDesc) continue; // part not in this CAD source, or CAD has no description for it
      eligible++;
      if (!descriptionsMatch(row.description, cadDesc)) {
        seenPn.add(pnKey);
        var parent = imQc.parentOf(pathIndex, row);
        var entry = {
          number: row.number,
          title: row.title || '',
          imDescription: row.description,
          cadDescription: cadDesc,
          sourceRow: row.sourceRow,
          parentNumber: parent ? parent.number : '',
          parentTitle: parent ? parent.title : '',
        };
        if (isAutoMatchedAppend(row.description, cadDesc)) autoMatched.push(entry);
        else mismatches.push(entry);
      }
    }

    return {
      applicable: true,
      cadSourceFileName: cad.source.fileName || '',
      autoMatched: autoMatched,
      eligibleCount: eligible,
      mismatches: mismatches,
    };
  }

  return {
    titleDescCompare: {
      compareTitleDescription: compareTitleDescription,
      descriptionsMatch: descriptionsMatch,
      isAutoMatchedAppend: isAutoMatchedAppend,
      normText: normText,
    },
  };
});
