/*
 * virtual-parts.js — finds "virtual" subassemblies: a BOM entry that exists
 * in the CAD BOM but has no CAD file behind it, whose whole child BOM lives
 * only in the Item Master.
 *
 * Why these slip through every other check: the part itself IS in both BOMs,
 * so "Missing from Item Master" and "In Item Master only" both pass it. Its
 * children are only in the Item Master, so they land in "In Item Master
 * only" — but scattered as one finding per child, because their parent is
 * not itself an Item-Master-only row and so cannot group them. The result is
 * a placeholder assembly nobody notices and a pile of orphan rows nobody can
 * connect back to it.
 *
 * Detection has two tiers:
 *
 *   confirmed — the Inventor export includes the Thumbnail column and this
 *     row's value is the literal string "(NULL)" (Inventor writes an empty
 *     cell when a CAD file exists, "(NULL)" when none does), it has at least
 *     one Item Master child, and every one of those children is absent from
 *     the CAD BOM.
 *
 *   suspected — no Thumbnail column in any loaded CAD source, so the same
 *     shape is inferred structurally: a CAD part with at least
 *     SUSPECT_MIN_CHILDREN Item Master children, all absent from CAD.
 *     The threshold is 3 because 1- and 2-child matches are dominated by
 *     hose/cable cut-length rows (verified across three real project pairs:
 *     a >=3 threshold produced zero false positives while still finding
 *     every part the Thumbnail column confirms).
 *
 * Deliberately NOT filtered to this site's own "7-" parts: unlike a missing
 * material or description, a virtual subassembly hides its entire child BOM
 * from every other check, so it is worth reporting wherever it occurs.
 *
 * Pure logic (no DOM). `indexItemMaster` is injected (compare.js's function)
 * rather than required directly, matching this codebase's style.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SUSPECT_MIN_CHILDREN = 3;

  function normNumber(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim().toUpperCase();
  }

  // cadSources: the array passed to compareAll (1-2 parsed CAD results).
  function detectVirtualParts(cadSources, im, indexItemMaster) {
    const sources = cadSources || [];
    const imIndex = indexItemMaster(im);

    if (!im.rows.some(function (r) { return Array.isArray(r.path); })) {
      return {
        applicable: false,
        reason: 'No "Row Order" column in the Item Master — cannot tell which parts have children.',
        confirmed: [], suspected: [], anchorRows: new Map(),
      };
    }

    // Every part number anywhere in CAD, plus the thumbnail verdict per part.
    const cadPNs = new Set();
    const thumbnailMissing = new Map(); // PN -> true when some source says "(NULL)"
    let hasThumbnailColumn = false;
    for (const src of sources) {
      if (!src) continue;
      if (src.hasThumbnail) hasThumbnailColumn = true;
      for (const it of src.items) {
        const pn = normNumber(it.number);
        if (!pn) continue;
        cadPNs.add(pn);
        if (it.thumbnailMissing) thumbnailMissing.set(pn, true);
      }
    }
    if (!cadPNs.size) {
      return {
        applicable: false,
        reason: 'No CAD BOM source is loaded.',
        confirmed: [], suspected: [], anchorRows: new Map(),
      };
    }

    const confirmed = [];
    const suspected = [];
    const anchorRows = new Map(); // PN -> the Item Master row, for grouping

    // childRows is keyed by parent PN and built from the positional parent
    // index, so a Row Order position reused by two different parts cannot
    // graft one assembly's children onto another.
    imIndex.childRows.forEach(function (children, parentPn) {
      if (!cadPNs.has(parentPn)) return; // not in CAD at all -> a "missing"/imOnly finding, not virtual
      const seen = new Set();
      const orphans = [];
      for (const child of children) {
        const cpn = normNumber(child.number);
        if (!cpn || cpn === parentPn || seen.has(cpn)) continue;
        seen.add(cpn);
        if (!cadPNs.has(cpn)) orphans.push(child);
      }
      if (!seen.size || orphans.length !== seen.size) return; // some children ARE in CAD -> a real modelled assembly

      const parentRows = imIndex.byNumber.get(parentPn);
      if (!parentRows || !parentRows.length) return;
      const parentRow = parentRows[0];
      const entry = {
        number: parentRow.number,
        title: parentRow.title || '',
        description: parentRow.description || '',
        childCount: orphans.length,
        children: orphans.map(function (c) {
          return { number: c.number, title: c.title || '', sourceRow: c.sourceRow || '' };
        }),
        sourceRow: parentRow.sourceRow || '',
      };

      if (hasThumbnailColumn) {
        if (!thumbnailMissing.get(parentPn)) return; // has a CAD file -> genuinely modelled
        entry.confidence = 'confirmed';
        confirmed.push(entry);
        anchorRows.set(parentPn, parentRow);
      } else if (orphans.length >= SUSPECT_MIN_CHILDREN) {
        entry.confidence = 'suspected';
        suspected.push(entry);
        anchorRows.set(parentPn, parentRow);
      }
    });

    const byNumber = function (a, b) { return normNumber(a.number) < normNumber(b.number) ? -1 : 1; };
    confirmed.sort(byNumber);
    suspected.sort(byNumber);

    return {
      applicable: true,
      hasThumbnailColumn: hasThumbnailColumn,
      confirmed: confirmed,
      suspected: suspected,
      anchorRows: anchorRows,
      reason: hasThumbnailColumn ? '' :
        'No "Thumbnail" column in the CAD export, so these are inferred from BOM structure. ' +
        'Include the Thumbnail column in the Inventor BOM export for an exact answer.',
    };
  }

  return {
    virtualParts: {
      detectVirtualParts: detectVirtualParts,
      SUSPECT_MIN_CHILDREN: SUSPECT_MIN_CHILDREN,
    },
  };
});
