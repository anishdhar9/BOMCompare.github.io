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
  // check keys it suppresses. Only the CAD-vs-Item-Master comparison keys
  // are covered today, since that's the only category seen in real Ignore
  // List files so far — Item Master QC / Material / Revision / Description
  // checks aren't wired to the ignore list yet. Add a new category here (and
  // to the recognized-values text in app.js's warning message) if a "From"
  // value for one of those is needed later.
  const CATEGORIES = {
    'cad vs item compare': ['missing', 'reference', 'qty', 'imOnly'],
  };

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
