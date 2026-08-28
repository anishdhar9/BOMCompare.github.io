/*
 * qty-parse.js — shared, locale-aware quantity-cell parser used by
 * cad-leveled.js, itemmaster.js, and lldbo.js. Previously each of those
 * files had its own copy (cad-leveled.js and itemmaster.js were byte-for-
 * byte identical; lldbo.js's had no comma handling at all), all sharing the
 * same underlying bug: `s.replace(',', '.')` unconditionally treated the
 * FIRST comma as a decimal point. That is correct for a lone European
 * decimal comma ("1,5" -> 1.5) but silently wrong for a thousands-grouped
 * value in EITHER convention: "1.234,5" (EU, meaning 1234.5) became
 * "1.234.5" -> parsed as 1.234, off by ~1000x; even a plain dot-thousands
 * integer like "1.234" (meaning 1234, no comma present at all) was never
 * touched by the replace and parsed as 1.234, also off by >1000x — with no
 * warning either way. This codebase clearly deals with non-English-locale
 * exports (German column synonyms throughout the other parsers), so this is
 * a real, plausible-to-trigger case, not a hypothetical one.
 *
 * Approach: extract the numeric-with-separators run, then decide what the
 * separators mean from their own shape rather than guessing at a fixed
 * locale:
 *   - both ',' and '.' present: whichever comes LAST is the decimal point;
 *     the other is thousands-grouping and is stripped ("1,234.5" US-style,
 *     "1.234,5" EU-style both resolve to 1234.5).
 *   - only commas, or only dots: thousands-grouping only when every group
 *     after the first is exactly 3 digits ("1,200" / "1.234" -> 1200 /
 *     1234); otherwise it's a lone decimal separator ("1,5" -> 1.5, and an
 *     ordinary "4.5" is untouched since dot is already the decimal point).
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const THOUSANDS_COMMA = /^-?\d{1,3}(,\d{3})+$/;
  const THOUSANDS_DOT = /^-?\d{1,3}(\.\d{3})+$/;

  function parseQty(v) {
    if (v === null || v === undefined) return null;
    const s = String(v).trim();
    if (!s || s === '-') return null;
    const m = s.match(/-?[\d.,]*\d/);
    if (!m) return null;
    const raw = m[0];

    const hasComma = raw.indexOf(',') !== -1;
    const hasDot = raw.indexOf('.') !== -1;
    let normalized;
    if (hasComma && hasDot) {
      normalized = raw.lastIndexOf(',') > raw.lastIndexOf('.')
        ? raw.replace(/\./g, '').replace(',', '.') // EU: 1.234,5 -> 1234.5
        : raw.replace(/,/g, '');                    // US: 1,234.5 -> 1234.5
    } else if (hasComma) {
      normalized = THOUSANDS_COMMA.test(raw) ? raw.replace(/,/g, '') : raw.replace(',', '.');
    } else if (hasDot) {
      normalized = THOUSANDS_DOT.test(raw) ? raw.replace(/\./g, '') : raw;
    } else {
      normalized = raw;
    }
    const n = parseFloat(normalized);
    return isNaN(n) ? null : n;
  }

  return { qtyParse: { parseQty: parseQty } };
});
