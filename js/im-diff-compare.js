/*
 * im-diff-compare.js — diffs two exports of the SAME Item Master BOM taken at
 * different points in time (e.g. before/after a release), to show exactly
 * what changed: parts added, parts removed, and per-field changes (Quantity,
 * Revision, Material, Title, Description, State) on parts present in both.
 *
 * This is Item-Master-vs-itself across time, distinct from the main app's
 * CAD-vs-Item-Master comparison (js/compare.js's compareAll). It reuses
 * itemMasterParser.parse() and indexItemMaster() from that same pipeline —
 * both injected, matching this codebase's dependency-injection style (see
 * js/lldbo-compare.js) — rather than reimplementing Item Master parsing.
 *
 * Diffing key: part number only. Row Order (the dotted position path) is a
 * same-file positional index — it shifts whenever a sibling is added,
 * removed, or reordered between releases, so it is not a stable identity
 * across two different exports. Part number is the only safe key here.
 *
 * Pure logic (no DOM).
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

  function normStr(v) {
    return v === null || v === undefined ? '' : String(v).trim();
  }

  function revEqual(a, b) {
    return normStr(a).toUpperCase() === normStr(b).toUpperCase();
  }

  function qtyEqual(a, b) {
    if (a === null || a === undefined) return b === null || b === undefined;
    if (b === null || b === undefined) return false;
    return Math.abs(a - b) < 1e-9;
  }

  function fmtQty(v) {
    if (v === null || v === undefined) return '(blank)';
    return Math.abs(v - Math.round(v)) < 1e-9 ? String(Math.round(v)) : String(Math.round(v * 1000) / 1000);
  }

  // indexItemMaster/materialsMatch are injected (compare.js's / material-compare.js's
  // functions) rather than required directly. materialsMatch is optional — when not
  // given, Material falls back to a plain trimmed string comparison.
  function diffItemMasters(imOld, imNew, indexItemMaster, materialsMatch) {
    const oldIndex = indexItemMaster(imOld);
    const newIndex = indexItemMaster(imNew);
    const oldPNs = new Set(oldIndex.byNumber.keys());
    const newPNs = new Set(newIndex.byNumber.keys());

    const added = [];
    const removed = [];
    const changed = [];
    let unchangedCount = 0;

    for (const pn of newPNs) {
      if (oldPNs.has(pn)) continue;
      const row = newIndex.byNumber.get(pn)[0];
      added.push({ number: row.number, title: row.title || '', description: row.description || '', sourceRow: row.sourceRow || '' });
    }
    for (const pn of oldPNs) {
      if (newPNs.has(pn)) continue;
      const row = oldIndex.byNumber.get(pn)[0];
      removed.push({ number: row.number, title: row.title || '', description: row.description || '', sourceRow: row.sourceRow || '' });
    }
    for (const pn of oldPNs) {
      if (!newPNs.has(pn)) continue;
      const oldRow = oldIndex.byNumber.get(pn)[0];
      const newRow = newIndex.byNumber.get(pn)[0];
      const fields = [];

      if (!qtyEqual(oldRow.qty, newRow.qty)) {
        fields.push({ field: 'Quantity', old: fmtQty(oldRow.qty), new: fmtQty(newRow.qty) });
      }
      if (!revEqual(oldRow.revision, newRow.revision)) {
        fields.push({ field: 'Revision', old: normStr(oldRow.revision), new: normStr(newRow.revision) });
      }
      // Raw equality is checked FIRST and short-circuits materialsMatch: that
      // function treats a blank as never matching anything (by design, for
      // the main CAD-vs-IM check, where a blank should never silently pass),
      // which would otherwise misreport two identical placeholder values
      // (e.g. both sides literally "-") as "changed" once stripped to "".
      const oldMat = normStr(oldRow.material);
      const newMat = normStr(newRow.material);
      let materialsDiffer = oldMat !== newMat;
      if (materialsDiffer && materialsMatch) materialsDiffer = !materialsMatch(oldMat, newMat);
      if (materialsDiffer) {
        fields.push({ field: 'Material', old: oldMat, new: newMat });
      }
      if (normStr(oldRow.title) !== normStr(newRow.title)) {
        fields.push({ field: 'Title', old: normStr(oldRow.title), new: normStr(newRow.title) });
      }
      if (normStr(oldRow.description) !== normStr(newRow.description)) {
        fields.push({ field: 'Description', old: normStr(oldRow.description), new: normStr(newRow.description) });
      }
      if (normStr(oldRow.state) !== normStr(newRow.state)) {
        fields.push({ field: 'State', old: normStr(oldRow.state), new: normStr(newRow.state) });
      }

      if (fields.length) {
        changed.push({
          number: newRow.number, title: newRow.title || '', description: newRow.description || '',
          oldSourceRow: oldRow.sourceRow || '', newSourceRow: newRow.sourceRow || '',
          fields: fields,
        });
      } else {
        unchangedCount++;
      }
    }

    const byNumber = function (a, b) { return normNumber(a.number) < normNumber(b.number) ? -1 : 1; };
    added.sort(byNumber);
    removed.sort(byNumber);
    changed.sort(byNumber);

    return {
      added: added, removed: removed, changed: changed, unchangedCount: unchangedCount,
      oldUniqueCount: oldPNs.size, newUniqueCount: newPNs.size,
    };
  }

  return { imDiffCompare: { diffItemMasters: diffItemMasters } };
});
