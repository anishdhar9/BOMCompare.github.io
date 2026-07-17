/*
 * material-compare.js — compares CAD BOM material against Item Master
 * material for the same part number, and builds the always-visible
 * "Bought-Out Parts" (7-999-*) reference panel.
 *
 * Not every CAD source carries material on the CAD side: the Vault
 * multi-level PDF never has a material column, and it's optional in both
 * the flat Vault Excel export and the Inventor BOM export (Vault lets users
 * choose which columns are visible, so this varies by export) — this check
 * uses whichever loaded CAD source actually has material data
 * (`hasMaterial`), and is only applicable when at least one does.
 *
 * A raw string comparison is unusable: verified on real data that of 518
 * shared manufactured (non-purchased) part numbers between a flat Vault
 * export and its Item Master, 38 "differ" as plain strings and ALL 38 are
 * naming-convention variants of the same material (DIN vs AISI grade
 * designation, spacing, abbreviation, or CAD simply carrying more
 * descriptive detail) — not real errors. This module normalizes before
 * comparing, but deliberately keeps a steel grade's L-suffix (304 vs 304L,
 * 316 vs 316L) significant rather than silently equating it, since that
 * can be a genuine weldability/corrosion spec choice.
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

  function stripFormatting(s) {
    return String(s || '').toUpperCase().replace(/\s+/g, '');
  }

  // DIN/EN Werkstoffnummer <-> AISI/UNS grade equivalence for the stainless
  // designations actually seen in this organization's data. 304 and 304L
  // (low-carbon variant), 316 and 316L, are kept as DISTINCT groups on
  // purpose (see file header) -- verified request not to silently equate
  // an L-suffix difference.
  var GRADE_GROUPS = [
    ['AISI304', '1.4301', 'SS304', '304'],
    ['AISI304L', '1.4306', '1.4307', 'SS304L', '304L'],
    ['AISI316', '1.4401', 'SS316', '316'],
    ['AISI316L', '1.4404', 'SS316L', '316L'],
    ['AISI316TI', '1.4571', 'SS316TI', '316TI'],
  ];
  // Same substance, different-language spelling -- not a grade lookup.
  var SPELLING_GROUPS = [
    ['SILICON', 'SILIKON'],
    ['BOROSILICATE', 'BOROSILIKAT'],
  ];

  function buildTokenList(groups) {
    var tokens = [];
    groups.forEach(function (group) {
      var canonical = stripFormatting(group[0]);
      group.forEach(function (variant) { tokens.push({ variant: stripFormatting(variant), canonical: canonical }); });
    });
    // longest variant first, so e.g. 'AISI304L' is tried before the shorter
    // 'AISI304' (which is also a literal prefix of it).
    tokens.sort(function (a, b) { return b.variant.length - a.variant.length; });
    return tokens;
  }
  var GRADE_TOKENS = buildTokenList(GRADE_GROUPS);
  var SPELLING_TOKENS = buildTokenList(SPELLING_GROUPS);

  // Finds a known steel-grade token anywhere in the (already-stripped)
  // string -- e.g. 'AISI304' inside 'STAINLESSSTEELAISI304' -- but only
  // when it isn't itself a prefix of a different/longer grade at that
  // position (the char right after the match can't extend it into another
  // grade token, e.g. won't match '304' as a substring of '304L').
  function extractGradeToken(stripped) {
    for (var i = 0; i < GRADE_TOKENS.length; i++) {
      var tok = GRADE_TOKENS[i];
      var idx = stripped.indexOf(tok.variant);
      if (idx === -1) continue;
      var after = stripped.charAt(idx + tok.variant.length);
      if (after === 'L' || after === 'T' || (after >= '0' && after <= '9')) continue;
      return tok.canonical;
    }
    return null;
  }

  function applySpelling(stripped) {
    var s = stripped;
    for (var i = 0; i < SPELLING_TOKENS.length; i++) {
      var tok = SPELLING_TOKENS[i];
      if (s.indexOf(tok.variant) !== -1) s = s.split(tok.variant).join(tok.canonical);
    }
    return s;
  }

  // Exported for the test suite / debugging; not needed by callers of
  // materialsMatch, which is the actual comparison entry point.
  function normalizeMaterial(raw) {
    var s = stripFormatting(raw);
    if (!s) return '';
    var grade = extractGradeToken(s);
    return grade || applySpelling(s);
  }

  function materialsMatch(a, b) {
    var sa = stripFormatting(a), sb = stripFormatting(b);
    if (!sa || !sb) return false;
    var ga = extractGradeToken(sa), gb = extractGradeToken(sb);
    if (ga || gb) {
      // at least one side looks like a steel grade code -- compare grade
      // tokens exactly, never fuzzily, so an L-suffix difference (or any
      // other genuinely different grade) is never silently equated.
      return ga !== null && gb !== null && ga === gb;
    }
    var na = applySpelling(sa), nb = applySpelling(sb);
    if (na === nb) return true;
    // neither side is a recognized grade -- allow one to be a more
    // detailed qualifier of the other (e.g. 'Silikon' inside
    // 'Silikon/weiß/60°Shore', or 'EPDM' inside 'EPDM/Light Gray').
    var shorter = na.length <= nb.length ? na : nb;
    var longer = na.length <= nb.length ? nb : na;
    if (shorter.length < 3) return false; // avoid trivial/nonsense containment matches
    return longer.indexOf(shorter) !== -1;
  }

  // Picks the first loaded CAD source that actually carries material data
  // and a lookup of its first-seen material per PN. Not every CAD source
  // type does: the Vault multi-level PDF never has a material column, and
  // some Inventor BOM exports omit it (columns are user-configurable in
  // Vault) while others include it — so this checks actual content
  // (`hasMaterial`), not the source's format/kind.
  function cadMaterialByPn(cadSources) {
    for (var i = 0; i < cadSources.length; i++) {
      var src = cadSources[i];
      if (!src.hasMaterial) continue;
      var map = new Map();
      var any = false;
      for (var j = 0; j < src.items.length; j++) {
        var it = src.items[j];
        var pn = String(it.number || '').trim().toUpperCase();
        var mat = (it.material || '').trim();
        if (!pn || !mat) continue;
        any = true;
        if (!map.has(pn)) map.set(pn, mat);
      }
      if (any) return { source: src, byPn: map };
    }
    return null;
  }

  // cadSources: the array passed to compareAll (0-2 CAD sources).
  function compareMaterial(cadSources, im) {
    var cad = cadMaterialByPn(cadSources || []);
    if (!cad) {
      return {
        applicable: false,
        reason: 'No loaded CAD source carries material data — the Vault multi-level PDF never does, and it depends on which columns were included in a flat Vault export or Inventor BOM export.',
        boughtOut: imQc.boughtOutParts(im),
      };
    }

    var assemblies = imQc.buildAssemblyPathSet(im.rows);
    var mismatches = [];
    var seenPn = new Set(); // same part can occur at several BOM positions; report it once
    for (var i = 0; i < im.rows.length; i++) {
      var row = im.rows[i];
      var pnKey = String(row.number).trim().toUpperCase();
      if (imQc.PURCHASED_PART_RE.test(row.number)) continue; // handled by the Bought-Out Parts panel instead
      if (!Array.isArray(row.path) || assemblies.has(row.path.join('.'))) continue; // assemblies don't carry material
      if (imQc.blank(row.material)) continue; // Check 6 already covers "IM material missing"
      if (seenPn.has(pnKey)) continue;
      var cadMat = cad.byPn.get(pnKey);
      if (!cadMat) continue; // part not in this CAD source, or CAD has no material for it
      if (!materialsMatch(row.material, cadMat)) {
        seenPn.add(pnKey);
        mismatches.push({
          number: row.number,
          title: row.title,
          imMaterial: row.material,
          cadMaterial: cadMat,
          sourceRow: row.sourceRow,
        });
      }
    }

    var boughtOut = imQc.boughtOutParts(im).map(function (part) {
      var cadMat = cad.byPn.get(String(part.number).trim().toUpperCase()) || '';
      return Object.assign({}, part, {
        cadMaterial: cadMat,
        mismatch: !!(cadMat && part.imMaterial && !materialsMatch(part.imMaterial, cadMat)),
      });
    });

    return {
      applicable: true,
      cadSourceFileName: cad.source.fileName || '',
      mismatches: mismatches,
      boughtOut: boughtOut,
    };
  }

  return {
    materialCompare: {
      compareMaterial: compareMaterial,
      materialsMatch: materialsMatch,
      normalizeMaterial: normalizeMaterial,
    },
  };
});
