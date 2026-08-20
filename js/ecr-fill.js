/*
 * ecr-fill.js — builds this organization's ECR (Engineering Change Request)
 * sheet rows from a two-version Item Master diff, and writes them into the
 * company's fixed ECR Excel template.
 *
 * Verified against a real filled ECR (SPN017059/PN22124, Zydus Lifesciences)
 * and its two source Item Master exports (the same BOM at revision B and at
 * revision C):
 *
 *   Sheet1 table (header row 11, data from row 12): Line No. | Sub Assly
 *   Number With Rev | Item No. With Rev | Description | Old Quantity |
 *   New Quantity | Action | Reason Code | Remarks.
 *
 *   "Item No. With Rev" fuses the part number and its RAW revision value,
 *   unpadded ("7-332-20950-1", "7-330-20008-C" — never zero-padded).
 *
 *   A revision bump is an old-row/new-row PAIR, not one changed row, and the
 *   pair is always emitted **revised row first, then obsolete row**: the new
 *   number+rev gets Old Qty=0, New Qty=1, Action="Drg. Revised"; the old
 *   number+rev gets Old Qty=1, New Qty=0, Action="Drg. Obsolete" — a binary
 *   0/1, not the part's real BOM quantity, regardless of what that quantity
 *   is. Both halves of the pair reference the same "Sub Assly Number With
 *   Rev": the parent's CURRENT (new-side) revision, even for the obsolete
 *   half — the pair describes what the BOM is becoming, not a snapshot of
 *   what it used to be.
 *
 *   This is a genuinely hierarchical diff, not a flat "which part numbers
 *   differ" comparison: a part is matched against its counterpart at the
 *   SAME position in the tree (same immediate parent), not against the
 *   first same-numbered row anywhere in the file. Two consequences that a
 *   flat, first-occurrence-only diff gets wrong: (1) a part reused under
 *   several different parents needs one row per affected parent, not one
 *   row total; (2) an unrelated part elsewhere in the BOM that happens to
 *   share revision-bump timing with the part being diffed must not be
 *   conflated with it just because the two share a part number.
 *
 *   Every matched node's children are walked, regardless of whether the
 *   node's own revision changed, since a subassembly's own revision can
 *   stay put while its children still gained/lost/re-quantified parts.
 *   Whether a wholly new/removed subtree's own children are also walked
 *   (each getting their own Added/Deleted row) or the subtree collapses to
 *   one row is the `cascadeIntoNewSubtrees` option below — this org's real
 *   sample doesn't resolve which behavior is "right" in general, so it's
 *   left to the caller instead of hardcoded.
 *
 *   Action is mechanically derivable from the diff (confirmed against
 *   Sheet2's exact dropdown list: Part Added / Part Deleted / Qty Changed /
 *   Drg. Obsolete / Drg. Revised). Reason Code (a 16-value business-
 *   justification list) is NOT derivable — it needs human judgment about
 *   WHY the BOM changed, so it is left blank for the person filing the ECR
 *   to fill in. Likewise, real BOMs can carry revision bumps or structural
 *   changes unrelated to the specific engineering change an ECR is being
 *   filed for (confirmed in the real sample: two parts in an unrelated
 *   branch of the BOM had genuine revision bumps that the real ECR did not
 *   include) — this diff reports everything it finds; excluding
 *   out-of-scope rows before filing is left to the person generating the
 *   sheet.
 *
 * Pure logic (no DOM) except fillEcrTemplate, which needs the XLSX
 * namespace to write into a loaded template workbook (injected, matching
 * this codebase's dependency-injection style).
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

  function qtyEqual(a, b) {
    if (a === null || a === undefined) return b === null || b === undefined;
    if (b === null || b === undefined) return false;
    return Math.abs(a - b) < 1e-9;
  }

  // Numeric (not text-with-unit) quantity for an ECR cell — the reference
  // sheet stores Old/New Quantity as plain numbers (1), never raw source
  // text ("1 Each").
  function roundQty(v) {
    if (v === null || v === undefined) return 0;
    return Math.abs(v - Math.round(v)) < 1e-9 ? Math.round(v) : Math.round(v * 1000) / 1000;
  }

  // Raw revision suffix, unpadded. '' revision -> no suffix at all. Used for
  // the "Sub Assly Number With Rev" column (always) and for the "Item No.
  // With Rev" column on a Drg. Revised/Drg. Obsolete row, where the suffix
  // is the entire point (distinguishing the old identity from the new one).
  function itemNoWithRev(row) {
    const rev = normStr(row.revision);
    return rev ? row.number + '-' + rev : row.number;
  }

  // Bare part number, no revision suffix. Used for the "Item No. With Rev"
  // column on Part Added/Part Deleted/Qty Changed rows: confirmed against
  // the real reference sheet, which never appends a revision suffix there —
  // only revision-swap rows need one, since revision isn't what changed.
  function bareNumber(row) {
    return normStr(row.number);
  }

  function descriptionOf(row) {
    return row.title || row.description || '';
  }

  function findRoot(im) {
    if (!im || !Array.isArray(im.rows) || !im.rows.length) return null;
    return im.rows.find(function (r) { return Array.isArray(r.path) && r.path.length === 0; }) || im.rows[0];
  }

  // row -> [child rows] in file order, built from the positional parent
  // index (indexItemMaster's parentOf) rather than the PN-keyed childRows
  // map, since a PN-keyed map conflates every occurrence of a repeated
  // parent part number — this needs children of ONE specific row/position.
  function buildChildrenOfRow(im, parentOf) {
    const map = new Map();
    im.rows.forEach(function (row) {
      const parent = parentOf.get(row);
      if (!parent) return;
      if (!map.has(parent)) map.set(parent, []);
      map.get(parent).push(row);
    });
    return map;
  }

  // Matches one parent's old children against its new children BY PART
  // NUMBER WITHIN THIS PARENT ONLY (never globally) — duplicate part
  // numbers under the same parent pair up in file order via a queue, so
  // they don't collide. Returns one ordered list of {oldRow,newRow} pairs:
  // matched/added children in new-file order (interleaved as siblings
  // actually appear), with any leftover removed-only children appended in
  // old-file order.
  function matchChildren(oldChildren, newChildren) {
    const oldQueues = new Map();
    oldChildren.forEach(function (r, i) {
      const pn = normNumber(r.number);
      if (!oldQueues.has(pn)) oldQueues.set(pn, []);
      oldQueues.get(pn).push(i);
    });
    const consumed = new Array(oldChildren.length).fill(false);
    const ordered = [];
    newChildren.forEach(function (nr) {
      const pn = normNumber(nr.number);
      const q = oldQueues.get(pn);
      if (q && q.length) {
        const idx = q.shift();
        consumed[idx] = true;
        ordered.push({ oldRow: oldChildren[idx], newRow: nr });
      } else {
        ordered.push({ oldRow: null, newRow: nr });
      }
    });
    oldChildren.forEach(function (r, i) {
      if (!consumed[i]) ordered.push({ oldRow: r, newRow: null });
    });
    return ordered;
  }

  // Emits the ECR row(s), if any, for a node present on BOTH sides at a
  // matched position: a revision-bump pair, a Qty Changed row, or (no
  // Action code exists for these) an "other change" report.
  function emitMatchedNodeRow(oldRow, newRow, parentComposite, ctx) {
    if (normStr(oldRow.revision).toUpperCase() !== normStr(newRow.revision).toUpperCase()) {
      ctx.rows.push({
        subAssyNumberWithRev: parentComposite, itemNoWithRev: itemNoWithRev(newRow),
        description: descriptionOf(newRow), oldQty: 0, newQty: 1, action: 'Drg. Revised',
      });
      ctx.rows.push({
        subAssyNumberWithRev: parentComposite, itemNoWithRev: itemNoWithRev(oldRow),
        description: descriptionOf(oldRow), oldQty: 1, newQty: 0, action: 'Drg. Obsolete',
      });
      return;
    }
    if (!qtyEqual(oldRow.qty, newRow.qty)) {
      ctx.rows.push({
        subAssyNumberWithRev: parentComposite, itemNoWithRev: bareNumber(newRow),
        description: descriptionOf(newRow), oldQty: roundQty(oldRow.qty), newQty: roundQty(newRow.qty),
        action: 'Qty Changed',
      });
      return;
    }
    const fields = [];
    if (normStr(oldRow.material) !== normStr(newRow.material)) fields.push('Material');
    if (normStr(oldRow.title) !== normStr(newRow.title)) fields.push('Title');
    if (normStr(oldRow.description) !== normStr(newRow.description)) fields.push('Description');
    if (normStr(oldRow.state) !== normStr(newRow.state)) fields.push('State');
    if (fields.length) ctx.otherChanges.push({ number: newRow.number, title: newRow.title || '', fields: fields });
  }

  // Recursive tree walk. Exactly one of oldRow/newRow may be null (a wholly
  // added or wholly removed node); both present means a matched position.
  // parentComposite is always the CURRENT identity already established for
  // the containing node — the new-side composite when the containing node
  // still exists in the new tree, the old-side composite when it doesn't
  // (a removed subtree has no new-side identity to point at).
  function walk(oldRow, newRow, parentComposite, opts, ctx) {
    if (oldRow && newRow) {
      emitMatchedNodeRow(oldRow, newRow, parentComposite, ctx);
      const thisComposite = itemNoWithRev(newRow);
      const oldChildren = ctx.oldChildrenOfRow.get(oldRow) || [];
      const newChildren = ctx.newChildrenOfRow.get(newRow) || [];
      matchChildren(oldChildren, newChildren).forEach(function (pair) {
        walk(pair.oldRow, pair.newRow, thisComposite, opts, ctx);
      });
      return;
    }
    if (newRow) {
      ctx.rows.push({
        subAssyNumberWithRev: parentComposite, itemNoWithRev: bareNumber(newRow),
        description: descriptionOf(newRow), oldQty: 0, newQty: roundQty(newRow.qty), action: 'Part Added',
      });
      if (opts.cascadeIntoNewSubtrees) {
        // Children still get the parent's REVISIONED composite (the "Sub
        // Assly Number With Rev" column always carries the suffix), even
        // though this node's own Item No. column just above did not.
        const thisComposite = itemNoWithRev(newRow);
        (ctx.newChildrenOfRow.get(newRow) || []).forEach(function (c) { walk(null, c, thisComposite, opts, ctx); });
      }
      return;
    }
    if (oldRow) {
      ctx.rows.push({
        subAssyNumberWithRev: parentComposite, itemNoWithRev: bareNumber(oldRow),
        description: descriptionOf(oldRow), oldQty: roundQty(oldRow.qty), newQty: 0, action: 'Part Deleted',
      });
      if (opts.cascadeIntoNewSubtrees) {
        const thisComposite = itemNoWithRev(oldRow);
        (ctx.oldChildrenOfRow.get(oldRow) || []).forEach(function (c) { walk(c, null, thisComposite, opts, ctx); });
      }
    }
  }

  // imOld/imNew: parsed Item Master results. indexItemMaster: js/compare.js's
  // function (injected), used here only for its positional parentOf map.
  // opts.cascadeIntoNewSubtrees (default true): whether a wholly new/removed
  // subassembly's own children each get their own Added/Deleted row, or the
  // whole subtree collapses to one row for the subassembly itself.
  // Returns { rows: [...ECR table rows, in hierarchical order], otherChanges }.
  function diffForEcr(imOld, imNew, indexItemMaster, opts) {
    const options = { cascadeIntoNewSubtrees: !opts || opts.cascadeIntoNewSubtrees !== false };
    const oldIndex = indexItemMaster(imOld);
    const newIndex = indexItemMaster(imNew);
    const ctx = {
      rows: [], otherChanges: [],
      oldChildrenOfRow: buildChildrenOfRow(imOld, oldIndex.parentOf),
      newChildrenOfRow: buildChildrenOfRow(imNew, newIndex.parentOf),
    };
    const oldRoot = findRoot(imOld);
    const newRoot = findRoot(imNew);
    if (oldRoot || newRoot) walk(oldRoot, newRoot, '', options, ctx);

    ctx.otherChanges.sort(function (a, b) { return normNumber(a.number) < normNumber(b.number) ? -1 : 1; });
    return { rows: ctx.rows, otherChanges: ctx.otherChanges };
  }

  // templateWb: an XLSX workbook already read from vendor/ECR_template.xlsx
  // (XLSX.read/readFile — injected as the `XLSX` param). Mutates and returns
  // the same workbook with the table (from row 12) filled in. projectKey
  // ({spn,pn}|null, e.g. from imNew.projectKey) fills the "PN" header field
  // (B3) when available — mechanically derivable, unlike the rest of the
  // header block (Project Status, Details Of Change, Engg. Comment, the
  // document/department checkboxes, the 5-Why block), which is left blank
  // since those are narrative/judgment fields a two-file diff cannot fill.
  function fillEcrTemplate(templateWb, ecrRows, XLSX, projectKey) {
    const sheet = templateWb.Sheets['Sheet1'];
    const aoa = ecrRows.map(function (r, i) {
      return [i + 1, r.subAssyNumberWithRev, r.itemNoWithRev, r.description, r.oldQty, r.newQty, r.action, '', ''];
    });
    if (aoa.length) XLSX.utils.sheet_add_aoa(sheet, aoa, { origin: 'A12' });
    if (projectKey && projectKey.pn) {
      const m = String(projectKey.pn).match(/\d+/);
      if (m) XLSX.utils.sheet_add_aoa(sheet, [[Number(m[0])]], { origin: 'B3' });
    }
    return templateWb;
  }

  return { ecrFill: { diffForEcr: diffForEcr, fillEcrTemplate: fillEcrTemplate, itemNoWithRev: itemNoWithRev } };
});
