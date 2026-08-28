/*
 * ignorelist-compare.js — turns a parsed Ignore List (js/parsers/ignorelist.js)
 * into an isIgnored(pn, checkKey) predicate js/compare.js's compareAll() can
 * use to suppress specific parts from specific checks.
 *
 * "From" category text is matched case/whitespace-insensitively against a
 * small, explicit table — kept explicit rather than guessed at, so an
 * unrecognized "From" value is reported (see `unrecognized` below) instead
 * of silently doing nothing, or silently suppressing more than intended.
 *
 * Pure logic (no DOM), matching this codebase's other *-compare.js modules.
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
  function normCategory(v) {
    return String(v || '').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  // "From" text (as written in the Ignore List) -> which of compareAll()'s
  // check keys it suppresses. "cad vs item compare" is the broad, catch-all
  // category (covers all five); "quantity mismatch" is a narrower one for a
  // part that should still be flagged if genuinely missing/reference/
  // in-Item-Master-only, but has an accepted, known quantity discrepancy
  // that shouldn't keep getting reported — this also covers "qtyCascade"
  // (compareAll()'s grouped root-cause finding for a whole subtree released
  // at one uniform ratio), since a cascade root is fundamentally the same
  // kind of finding, just pre-grouped; "revision" suppresses a part from
  // the Revision: CAD vs Item Master check (js/revision-compare.js, a
  // separate module from compareAll() but filtered the same way); "lldbo
  // candidate" suppresses a part from the LLDBO-candidate check
  // (js/lldbo-compare.js's detectLldboCandidates(), also a separate module,
  // both confidence tiers) — for a keyword hit someone has confirmed is not
  // actually a long-lead part; "all" (below) is broader still — literally
  // everything the ignore list currently knows how to suppress. These are
  // independent — a part can be listed under any one or combination. Item
  // Master QC / Material / Description / Virtual Parts checks aren't wired
  // to the ignore list yet. Add a new category here (and to the
  // recognized-values text in app.js's warning message) if a "From" value
  // for one of those is needed later — "all" picks it up automatically.
  const CATEGORIES = {
    'cad vs item compare': ['missing', 'reference', 'qty', 'qtyCascade', 'imOnly'],
    'quantity mismatch': ['qty', 'qtyCascade'],
    'revision': ['revision'],
    'lldbo candidate': ['lldboCandidate'],
  };
  // "All" suppresses everything the ignore list can currently suppress for a
  // part — the union of every category above, computed rather than
  // hardcoded so it can never drift out of sync if a category is added.
  CATEGORIES.all = Array.from(Object.keys(CATEGORIES).reduce(function (set, cat) {
    for (const k of CATEGORIES[cat]) set.add(k);
    return set;
  }, new Set()));

  // Returns { isIgnored(pn, checkKey), byPn: Map<PN, Set<checkKey>>, unrecognized:[{number, from, sourceRow}] }.
  function buildIgnoreIndex(ignoreList) {
    const byPn = new Map();
    const unrecognized = [];
    for (const row of (ignoreList && ignoreList.rows) || []) {
      const pn = normNumber(row.number);
      if (!pn) continue;
      const keys = CATEGORIES[normCategory(row.from)];
      if (!keys) { unrecognized.push(row); continue; }
      if (!byPn.has(pn)) byPn.set(pn, new Set());
      for (const k of keys) byPn.get(pn).add(k);
    }
    return {
      byPn: byPn,
      unrecognized: unrecognized,
      isIgnored: function (pn, checkKey) {
        const keys = byPn.get(normNumber(pn));
        return !!(keys && keys.has(checkKey));
      },
    };
  }

  return { ignoreListCompare: { buildIgnoreIndex: buildIgnoreIndex, CATEGORIES: CATEGORIES } };
});
