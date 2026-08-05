/*
 * compare.js — pure comparison logic (no DOM). Used by the browser app and by
 * the Node test runner.
 *
 * Inputs are the normalized structures produced by the parsers:
 *
 * CAD BOM: {
 *   source: 'flat-xlsx' | 'pdf' | 'leveled-sheet',
 *   hasQty: boolean,
 *   hasLevels: boolean,
 *   items: [{ seq, number, title, description, qty, level, isAssembly,
 *             file, material, sourceRow }],
 * }
 *
 * Item Master: {
 *   rows: [{ number, title, description, qty, path, rowType, sourceRow }],
 *   // path: array of segments from the 'Row Order' column ('2.8.1' -> ['2','8','1']),
 *   // [] for the root row ('-'). null when the export has no Row Order column.
 * }
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

  /* ------------------------------------------------------------------ *
   * Item Master indexing
   * ------------------------------------------------------------------ */

  // Resolves each row's true parent POSITIONALLY — by file order plus Row
  // Order depth — instead of by Row Order string. Real exports reuse a
  // position for adjacent siblings (verified on a real 2081-row export: 40
  // duplicate positions, 249 rows left with an ambiguous ancestor), so a
  // path-string lookup silently attributes a child to whichever branch
  // happened to come first. Sequence + depth cannot collide that way. This
  // is the same stack walk cadTotals() and groupMissingLeveled() already use
  // for CAD levels.
  //
  // Returns Map<row, parentRow>; a row with no entry is a root (or carries
  // no Row Order at all).
  function buildParentIndex(rows) {
    const parentOf = new Map();
    const stack = []; // {depth, row}
    for (const row of rows) {
      if (!Array.isArray(row.path)) continue;
      const depth = row.path.length;
      while (stack.length && stack[stack.length - 1].depth >= depth) stack.pop();
      if (stack.length) parentOf.set(row, stack[stack.length - 1].row);
      stack.push({ depth: depth, row: row });
    }
    return parentOf;
  }

  function pathDepth(row) {
    return Array.isArray(row.path) ? row.path.length : -1;
  }

  // Builds lookup structures from the Item Master rows.
  function indexItemMaster(im) {
    const byNumber = new Map();       // PN -> [rows]
    const byPath = new Map();         // 'path.key' -> row (first wins)
    const parentOf = buildParentIndex(im.rows); // row -> parent row
    const warnings = [];
    let dupPaths = 0;

    for (const row of im.rows) {
      const pn = normNumber(row.number);
      if (!pn) continue;
      if (!byNumber.has(pn)) byNumber.set(pn, []);
      byNumber.get(pn).push(row);
      if (Array.isArray(row.path)) {
        const key = row.path.join('.');
        if (byPath.has(key) && key !== '') {
          dupPaths++;
        } else {
          byPath.set(key, row);
        }
      }
    }
    if (dupPaths) {
      warnings.push(dupPaths + ' Item Master rows share a "Row Order" position with another row; all rows are compared, but the quantity roll-up uses the first row at each position.');
    }

    // childSets: PN -> Set of direct-child PNs (union over all occurrences),
    // and childRows: PN -> [child rows]. Both derived from the positional
    // parent index, so a duplicated Row Order position can't graft one
    // assembly's children onto another.
    const childSets = new Map();
    const childRows = new Map();
    for (const row of im.rows) {
      const parent = parentOf.get(row);
      if (!parent) continue;
      const ppn = normNumber(parent.number);
      if (!ppn) continue;
      if (!childSets.has(ppn)) { childSets.set(ppn, new Set()); childRows.set(ppn, []); }
      childSets.get(ppn).add(normNumber(row.number));
      childRows.get(ppn).push(row);
    }

    // Rolled-up total quantity per PN: sum over occurrences of
    // own qty x product of ancestor quantities (via path prefixes).
    const totals = new Map();         // PN -> number|null (null = not computable)
    const breakdowns = new Map();     // PN -> [{parentNumber, parentTitle, qty, effQty}]
    const hasPaths = im.rows.some(function (r) { return Array.isArray(r.path); });
    for (const row of im.rows) {
      const pn = normNumber(row.number);
      if (!pn) continue;
      if (Array.isArray(row.path) && row.path.length === 0) continue; // root row: qty '-'
      let eff = row.qty;
      let parentRow = null;
      if (hasPaths && Array.isArray(row.path)) {
        // Walk the real parent chain, multiplying by every ancestor below
        // the root (the root row carries qty '-' and is not a multiplier) —
        // same set of ancestors the previous path-prefix loop covered, but
        // resolved positionally so duplicated positions can't mis-multiply.
        const immediate = parentOf.get(row);
        if (immediate && pathDepth(immediate) >= 1) parentRow = immediate;
        for (let anc = immediate; anc && pathDepth(anc) >= 1 && eff !== null; anc = parentOf.get(anc)) {
          if (anc.qty !== null) eff *= anc.qty;
        }
      }
      if (!breakdowns.has(pn)) breakdowns.set(pn, []);
      breakdowns.get(pn).push({
        parentNumber: parentRow ? normNumber(parentRow.number) : '',
        parentTitle: parentRow ? (parentRow.title || '') : '',
        qty: row.qty,
        effQty: eff,
        sourceRow: row.sourceRow || '',
        rowOrder: Array.isArray(row.path) ? row.path.join('.') : '',
      });
      if (row.qty === null) { totals.set(pn, null); continue; }
      if (totals.get(pn) !== null || !totals.has(pn)) {
        totals.set(pn, (totals.get(pn) || 0) + (eff === null ? row.qty : eff));
      }
    }

    return {
      byNumber: byNumber, byPath: byPath, parentOf: parentOf,
      childSets: childSets, childRows: childRows,
      totals: totals, breakdowns: breakdowns, warnings: warnings,
    };
  }

  /* ------------------------------------------------------------------ *
   * CAD quantity roll-up
   * ------------------------------------------------------------------ */

  // Per-PN rolled-up totals for a leveled CAD BOM (levels + qty per row).
  // Builds breakdowns unconditionally, even when the source has no quantity
  // column — qty/effQty simply stay null then (mirrors indexItemMaster's
  // breakdowns, which are never qty-gated either) — so callers that only
  // want "every place this PN occurs" (e.g. compareAll's cadOccurrences,
  // built for on-screen traceability) don't need a quantity column to work.
  function cadTotals(cad) {
    const totals = new Map();
    const breakdowns = new Map();
    const stack = []; // {level, number, title, qty, effQty}
    for (const it of cad.items) {
      const pn = normNumber(it.number);
      if (!pn) continue;
      const level = cad.hasLevels && it.level !== null ? it.level : 1;
      while (stack.length && stack[stack.length - 1].level >= level) stack.pop();
      const parent = stack.length ? stack[stack.length - 1] : null;
      const mult = parent ? parent.effQty : 1;
      const eff = it.qty === null || mult === null ? null : it.qty * mult;
      if (!breakdowns.has(pn)) breakdowns.set(pn, []);
      breakdowns.get(pn).push({
        parentNumber: parent ? parent.number : '',
        parentTitle: parent ? parent.title : '',
        qty: it.qty,
        effQty: eff,
        sourceRow: it.sourceRow || '',
        page: it.page || null,
        file: it.file || '',
      });
      if (it.qty === null || eff === null) totals.set(pn, null);
      else if (totals.get(pn) !== null || !totals.has(pn)) totals.set(pn, (totals.get(pn) || 0) + eff);
      stack.push({ level: level, number: pn, title: it.title || '', qty: it.qty, effQty: eff === null ? null : eff });
    }
    return { totals: totals, breakdowns: breakdowns };
  }

  /* ------------------------------------------------------------------ *
   * Missing-item grouping
   * ------------------------------------------------------------------ */

  function makeNode(item) {
    return { item: item, children: [], childPNs: new Set() };
  }

  function attachChild(parentNode, node) {
    const pn = normNumber(node.item.number);
    if (parentNode.childPNs.has(pn)) return false;
    parentNode.childPNs.add(pn);
    parentNode.children.push(node);
    return true;
  }

  // Exact grouping when the CAD source has levels: a missing item whose
  // ancestor (by level) is also missing goes under that ancestor.
  // `hasFn(pn)` says whether a PN counts as present (e.g. in the Item Master,
  // or — for reference detection — in the intended-BOM export).
  function groupMissingLeveled(cad, hasFn) {
    const rootNodes = [];
    const seen = new Map(); // PN -> node (first occurrence wins)
    const stack = [];       // {level, missing, node|null}
    for (const it of cad.items) {
      const pn = normNumber(it.number);
      if (!pn) continue;
      const level = it.level !== null ? it.level : 1;
      while (stack.length && stack[stack.length - 1].level >= level) stack.pop();
      const missing = !hasFn(pn);
      let node = null;
      if (missing) {
        node = seen.get(pn) || null;
        const isNew = !node;
        if (isNew) { node = makeNode(it); seen.set(pn, node); }
        // nearest missing ancestor on the stack
        let anc = null;
        for (let i = stack.length - 1; i >= 0; i--) {
          if (stack[i].missing) { anc = stack[i]; break; }
        }
        if (isNew) {
          if (anc) attachChild(anc.node, node);
          else rootNodes.push(node);
        }
      }
      stack.push({ level: level, missing: missing, node: node });
    }
    return rootNodes;
  }

  // Inference-based grouping for the flat (pre-order, level-less) Vault
  // export. Walk the CAD sequence with a stack of "open" assemblies. Present
  // assemblies carry the set of child PNs the Item Master expects under them;
  // seeing a present item that belongs to such a set closes everything opened
  // above that assembly — which is what bounds a missing (reference)
  // assembly's subtree.
  function groupMissingFlat(cad, hasFn, childSets) {
    const rootNodes = [];
    const seen = new Map(); // PN -> node
    const stack = [];       // {number, missing, expected:Set|null, node|null}
    for (const it of cad.items) {
      const pn = normNumber(it.number);
      if (!pn) continue;
      const present = hasFn(pn);
      if (present) {
        // resync: deepest open assembly that expects this PN as a child
        let idx = -1;
        for (let i = stack.length - 1; i >= 0; i--) {
          if (stack[i].expected && stack[i].expected.has(pn)) { idx = i; break; }
        }
        if (idx >= 0) stack.length = idx + 1;
        if (it.isAssembly) {
          stack.push({ number: pn, missing: false, expected: childSets.get(pn) || new Set(), node: null });
        }
      } else {
        let node = seen.get(pn) || null;
        const isNew = !node;
        if (isNew) { node = makeNode(it); seen.set(pn, node); }
        // The shallowest open missing assembly absorbs this finding. (The
        // inferred structure INSIDE a missing subtree is unreliable — nothing
        // in the Item Master bounds it — so grouped children are kept flat
        // under the group root instead of pretending to know their nesting.)
        let anc = null;
        for (let i = 0; i < stack.length; i++) {
          if (stack[i].missing) { anc = stack[i]; break; }
        }
        if (isNew) {
          if (anc) attachChild(anc.node, node);
          else rootNodes.push(node);
        }
        if (it.isAssembly) {
          stack.push({ number: pn, missing: true, expected: null, node: node });
        }
      }
    }
    return rootNodes;
  }

  function countDescendants(node) {
    let n = 0;
    for (const c of node.children) n += 1 + countDescendants(c);
    return n;
  }

  // Same idea as groupMissingLeveled, but over the Item Master's own Row Order
  // hierarchy instead of CAD levels: an "in Item Master only" row whose nearest
  // ancestor is itself Item-Master-only goes under that ancestor. When a whole
  // subassembly is absent from the CAD BOM every one of its parts is flagged,
  // which is one finding (the subassembly), not hundreds — on a real sample this
  // collapses 1033 flat rows to 11 roots.
  //
  // `rows` are the flat imOnly entries (in file order); `parentOf` is
  // indexItemMaster's positional parent index. `anchorRows` is an optional
  // Map<PN, imRow> of extra parts allowed to act as grouping parents even
  // though they are not themselves Item-Master-only — used for virtual parts,
  // which ARE in the CAD BOM but whose whole child BOM is not, so without this
  // their children scatter as one root each. Falls back to one root per row
  // when the export has no Row Order column (nothing to group by).
  //
  // Runs in two passes: every node is created first, then attached. A
  // single-pass walk had to give up whenever the flagged ancestor's node did
  // not exist yet (it appears later in row order) and silently emitted the
  // child as a root.
  function groupImOnly(rows, parentOf, anchorRows) {
    const rootNodes = [];
    const nodes = new Map();   // PN -> node (first occurrence wins)
    const rowFor = new Map();  // PN -> the row that created the node
    const flagged = new Set();
    const anchors = anchorRows || new Map();
    for (const r of rows) {
      const pn = normNumber(r.number);
      if (pn) flagged.add(pn);
    }
    // pass 1 — one node per unique Item-Master-only PN
    for (const row of rows) {
      const pn = normNumber(row.number);
      if (!pn || nodes.has(pn)) continue;
      nodes.set(pn, makeNode(row));
      rowFor.set(pn, row);
    }
    // Anchor nodes are created lazily, so an anchor that never absorbs a child
    // does not appear as an empty finding.
    const anchorNodes = new Map();
    const anchorNodeFor = function (pn) {
      if (anchorNodes.has(pn)) return anchorNodes.get(pn);
      const r = anchors.get(pn);
      if (!r) return null;
      const n = makeNode(r);
      n.isAnchor = true;
      anchorNodes.set(pn, n);
      return n;
    };
    // Entries handed in may be copies of the Item Master rows; the parent
    // index is keyed by row identity, so hop through __srcRow when present.
    const parentRowOf = function (r) { return parentOf.get(r && r.__srcRow ? r.__srcRow : r); };
    // pass 2 — attach each node under its nearest groupable ancestor
    for (const [pn, node] of nodes) {
      const row = rowFor.get(pn);
      let anc = null;
      for (let cand = parentRowOf(row); cand; cand = parentRowOf(cand)) {
        const cpn = normNumber(cand.number);
        if (cpn === pn) continue;
        if (flagged.has(cpn)) { anc = nodes.get(cpn) || null; break; }
        if (anchors.has(cpn)) { anc = anchorNodeFor(cpn); break; }
      }
      if (anc && anc !== node) attachChild(anc, node);
      else rootNodes.push(node);
    }
    for (const node of anchorNodes.values()) {
      if (node.children.length) rootNodes.push(node);
    }
    return rootNodes;
  }

  /* ------------------------------------------------------------------ *
   * Quantity-cascade detection
   * ------------------------------------------------------------------ *
   * A whole Item Master subtree released at one clean, uniform ratio of its
   * CAD-required quantity (e.g. every direct child of one assembly entered
   * at exactly 2x) usually traces back to ONE data-entry error, not many
   * independent ones. The per-part "Quantity mismatches" check can't see
   * this on its own — it only compares each part's total ROLLED-UP
   * quantity, so a single doubled assembly node can surface as hundreds of
   * separately-flagged descendants. This groups those back into one
   * root-cause finding, the same way "Missing"/"In Item Master only"
   * already group a missing subtree under its topmost missing ancestor.
   */

  const CASCADE_MIN_CHILDREN = 2;    // fewer than this is just one ordinary mismatch
  const CASCADE_RATIO_PRECISION = 6; // decimal places used to group "the same ratio"

  // Map<childPN, [{parentNumber, qty, ...}]> (a *Totals breakdowns map) ->
  // Map<parentPN, Map<childPN, summed local qty>>. Both cadTotals() and
  // indexItemMaster() already return breakdowns in this per-child-occurrence
  // shape (used for the qty-mismatch expander's "where is this used" rows);
  // this just re-groups the same data by parent instead of by child.
  function buildParentChildQty(breakdowns) {
    const map = new Map();
    for (const [childPn, occurrences] of breakdowns) {
      for (const occ of occurrences) {
        if (!occ.parentNumber || occ.qty === null) continue;
        if (!map.has(occ.parentNumber)) map.set(occ.parentNumber, new Map());
        const kids = map.get(occ.parentNumber);
        kids.set(childPn, (kids.get(childPn) || 0) + occ.qty);
      }
    }
    return map;
  }

  // Full Item Master descendant subtree under `parentPn`, built from
  // indexItemMaster's childSets (direct-child PN sets) and byNumber (PN ->
  // representative row) — so every downstream part inherits the cascade
  // grouping, not just the direct children the ratio was measured on.
  function buildImSubtree(parentPn, imIndex, rootItem) {
    const root = makeNode(rootItem);
    const seen = new Set([parentPn]);
    const walk = function (node, pn) {
      const kids = imIndex.childSets.get(pn);
      if (!kids) return;
      for (const childPn of kids) {
        if (seen.has(childPn)) continue;
        seen.add(childPn);
        const rows = imIndex.byNumber.get(childPn);
        if (!rows || !rows.length) continue;
        const childNode = makeNode(rows[0]);
        attachChild(node, childNode);
        walk(childNode, childPn);
      }
    };
    walk(root, parentPn);
    return root;
  }

  // ct: cadTotals() result for the qty-carrying CAD source (needs .breakdowns).
  // imIndex: indexItemMaster() result for the loaded Item Master.
  function detectQuantityCascades(ct, imIndex) {
    if (!ct || !ct.breakdowns) return { applicable: false, roots: [] };
    const cadMap = buildParentChildQty(ct.breakdowns);
    const imMap = buildParentChildQty(imIndex.breakdowns);
    if (!cadMap.size) return { applicable: false, roots: [] };

    const candidates = []; // {parentPn, ratio, childCount, mismatchedChildCount}
    for (const [parentPn, imKids] of imMap) {
      const cadKids = cadMap.get(parentPn);
      if (!cadKids) continue;
      const ratioBuckets = new Map(); // rounded ratio -> count of children at that ratio
      let comparable = 0;
      for (const [childPn, imQty] of imKids) {
        const cadQty = cadKids.get(childPn);
        if (cadQty === undefined || cadQty === 0) continue;
        comparable++;
        if (Math.abs(cadQty - imQty) < 1e-9) continue; // matches — not mismatched
        const ratio = Number((imQty / cadQty).toFixed(CASCADE_RATIO_PRECISION));
        ratioBuckets.set(ratio, (ratioBuckets.get(ratio) || 0) + 1);
      }
      if (!ratioBuckets.size) continue;
      let bestRatio = null, bestCount = 0;
      for (const [ratio, count] of ratioBuckets) {
        if (count > bestCount) { bestRatio = ratio; bestCount = count; }
      }
      // Every mismatched child under this parent must share the one ratio —
      // a mixed bag of different ratios isn't one clean cause, so it's left
      // for the ordinary per-part quantity-mismatch check instead.
      const totalMismatched = Array.from(ratioBuckets.values()).reduce(function (a, b) { return a + b; }, 0);
      if (bestCount < CASCADE_MIN_CHILDREN || bestCount !== totalMismatched) continue;
      candidates.push({ parentPn: parentPn, ratio: bestRatio, childCount: comparable, mismatchedChildCount: bestCount });
    }
    if (!candidates.length) return { applicable: true, roots: [] };

    // Drop any candidate that is itself a descendant of another candidate —
    // it is already covered by the ancestor's subtree grouping below.
    const pathOf = function (pn) {
      const rows = imIndex.byNumber.get(pn);
      return rows && rows.length && Array.isArray(rows[0].path) ? rows[0].path : null;
    };
    const isDescendantOf = function (path, ancestorPath) {
      if (!path || !ancestorPath || path.length <= ancestorPath.length) return false;
      for (let i = 0; i < ancestorPath.length; i++) if (path[i] !== ancestorPath[i]) return false;
      return true;
    };
    const candidatePaths = candidates.map(function (c) { return pathOf(c.parentPn); });
    const roots = [];
    for (let i = 0; i < candidates.length; i++) {
      const c = candidates[i];
      const path = candidatePaths[i];
      let nested = false;
      for (let j = 0; j < candidates.length; j++) {
        if (i !== j && isDescendantOf(path, candidatePaths[j])) { nested = true; break; }
      }
      if (nested) continue;
      const rows = imIndex.byNumber.get(c.parentPn);
      if (!rows || !rows.length) continue;
      const rootItem = Object.assign({}, rows[0], {
        cascadeRatio: c.ratio,
        cascadeChildCount: c.childCount,
        cascadeMismatchedChildCount: c.mismatchedChildCount,
      });
      roots.push(buildImSubtree(c.parentPn, imIndex, rootItem));
    }
    roots.sort(function (a, b) { return b.item.cascadeMismatchedChildCount - a.item.cascadeMismatchedChildCount; });
    return { applicable: true, roots: roots };
  }

  /* ------------------------------------------------------------------ *
   * Dual-source helpers
   * ------------------------------------------------------------------ */

  // Direct-child PN sets derived from a leveled CAD BOM (parent level -> its
  // children), used to bound missing subtrees when the structure source is
  // the level-less flat export but a leveled intended-BOM export is present.
  function cadChildSets(cad) {
    const sets = new Map();
    const stack = []; // {level, pn}
    for (const it of cad.items) {
      const pn = normNumber(it.number);
      if (!pn) continue;
      const level = it.level !== null ? it.level : 1;
      while (stack.length && stack[stack.length - 1].level >= level) stack.pop();
      if (stack.length) {
        const parent = stack[stack.length - 1].pn;
        if (!sets.has(parent)) sets.set(parent, new Set());
        sets.get(parent).add(pn);
      }
      stack.push({ level: level, pn: pn });
    }
    return sets;
  }

  // The record for the machine itself: the single item at the shallowest
  // level (the Vault PDF starts with the root assembly; exports of the BOM
  // grid usually don't include it).
  function rootPNOf(cad) {
    if (!cad.hasLevels) return null;
    let min = Infinity;
    for (const it of cad.items) if (it.level !== null && it.level < min) min = it.level;
    const atMin = cad.items.filter(function (it) { return it.level === min; });
    return atMin.length === 1 ? normNumber(atMin[0].number) : null;
  }

  // Pick which uploaded CAD source plays which role.
  //  - structure: the full CAD structure incl. reference components
  //    (Vault "Uses" PDF or the flat Vault export); falls back to the first.
  //  - bom: the intended-BOM export (leveled sheet, ideally with quantities);
  //    only distinct from `structure` when two sources are given.
  function pickRoles(sources) {
    let structure = sources.find(function (s) { return s.source === 'pdf'; }) ||
                    sources.find(function (s) { return s.source === 'flat-xlsx'; }) ||
                    sources[0];
    let bom = sources.find(function (s) { return s !== structure && s.source === 'leveled-sheet'; }) ||
              sources.find(function (s) { return s !== structure; }) || null;
    return { structure: structure, bom: bom };
  }

  /* ------------------------------------------------------------------ *
   * Ignore list
   * ------------------------------------------------------------------ *
   * A user-supplied Ignore List (js/parsers/ignorelist.js, resolved to a
   * predicate by js/ignorelist-compare.js) names parts intentionally left in
   * a non-standard BOM state that should not be reported by specific checks.
   * Applied here — not in app.js's renderers or findings.js's registry — so
   * every consumer of compareAll()'s output (results tabs, tile counts,
   * "Parts needing attention") already sees ignored parts removed, with
   * nothing to special-case downstream.
   */

  // Removes any node `isIgnored(pn, checkKey)` flags, promoting its children
  // to its own parent's level (or the root list) so they remain independent
  // findings instead of vanishing along with a suppressed grouping parent —
  // they may not themselves be ignored. Pushes a flat, UI-ready record for
  // each removed node onto `collector`, so the "Ignored findings" section
  // can show exactly what was suppressed and from where.
  function filterIgnoredTree(roots, isIgnored, checkKey, collector) {
    const out = [];
    for (const node of roots || []) {
      const pn = normNumber(node.item.number);
      if (isIgnored(pn, checkKey)) {
        collector.push({
          checkKey: checkKey, number: node.item.number, title: node.item.title || '',
          description: node.item.description || '', sourceRow: node.item.sourceRow || '',
          parentNumber: '', parentTitle: '',
        });
        for (const promoted of filterIgnoredTree(node.children || [], isIgnored, checkKey, collector)) out.push(promoted);
      } else {
        node.children = filterIgnoredTree(node.children || [], isIgnored, checkKey, collector);
        out.push(node);
      }
    }
    return out;
  }

  // Same idea for a flat (non-tree) findings list, e.g. qtyMismatches.
  function filterIgnoredFlat(list, isIgnored, checkKey, collector) {
    const kept = [];
    for (const r of list || []) {
      const pn = normNumber(r.number);
      if (isIgnored(pn, checkKey)) {
        collector.push({
          checkKey: checkKey, number: r.number, title: r.title || '', description: r.description || '',
          sourceRow: r.sourceRow || '', parentNumber: r.parentNumber || '', parentTitle: r.parentTitle || '',
        });
      } else {
        kept.push(r);
      }
    }
    return kept;
  }

  /* ------------------------------------------------------------------ *
   * Main entry
   * ------------------------------------------------------------------ */

  // cadSources: one or two parsed CAD results (e.g. Vault PDF + Inventor xlsx)
  // opts (optional): { virtualAnchorRows: Map<PN, imRow> } from
  // js/virtual-parts.js — see the imOnly grouping below; { isIgnored(pn,
  // checkKey) } from js/ignorelist-compare.js — see "Ignore list" above.
  function compareAll(cadSources, im, opts) {
    const imIndex = indexItemMaster(im);
    const roles = pickRoles(cadSources);
    const structure = roles.structure;
    const bom = roles.bom;
    const inIM = function (pn) { return imIndex.byNumber.has(pn); };
    const isIgnored = (opts && opts.isIgnored) || function () { return false; };
    const ignoredFindings = []; // flat, UI-ready — collected below as each check is filtered

    const cadPNs = new Set();
    const firstCadItem = new Map(); // PN -> item (first occurrence, structure source wins)
    for (const src of [structure, bom]) {
      if (!src) continue;
      for (const it of src.items) {
        const pn = normNumber(it.number);
        if (!pn) continue;
        cadPNs.add(pn);
        if (!firstCadItem.has(pn)) firstCadItem.set(pn, it);
      }
    }

    // 1) missing from Item Master, grouped on the structure source; parts
    // that only exist in the intended-BOM export (virtual components have no
    // CAD file, so they never appear in the Vault PDF) are appended as
    // standalone findings.
    let missingRoots = structure.hasLevels
      ? groupMissingLeveled(structure, inIM)
      : groupMissingFlat(structure, inIM, imIndex.childSets);
    if (bom) {
      const structPNs = new Set();
      for (const it of structure.items) structPNs.add(normNumber(it.number));
      const added = new Set();
      for (const it of bom.items) {
        const pn = normNumber(it.number);
        if (!pn || structPNs.has(pn) || inIM(pn) || added.has(pn)) continue;
        added.add(pn);
        missingRoots.push(makeNode(it));
      }
    }
    missingRoots = filterIgnoredTree(missingRoots, isIgnored, 'missing', ignoredFindings);
    let missingTotal = 0;
    for (const pn of cadPNs) if (!inIM(pn) && !isIgnored(pn, 'missing')) missingTotal++;

    // 2) quantity mismatches — from whichever source carries quantities
    const qtySource = (bom && bom.hasQty) ? bom : (structure.hasQty ? structure : null);
    let qtyMismatches = null;
    let qtyCascades = { applicable: false, roots: [] };
    if (qtySource) {
      const ct = cadTotals(qtySource);
      qtyCascades = detectQuantityCascades(ct, imIndex);
      qtyMismatches = [];
      for (const [pn, cadTotal] of ct.totals) {
        if (!inIM(pn)) continue; // covered by "missing"
        const imTotal = imIndex.totals.has(pn) ? imIndex.totals.get(pn) : null;
        if (cadTotal === null || imTotal === null) continue; // not computable
        if (Math.abs(cadTotal - imTotal) > 1e-9) {
          const item = firstCadItem.get(pn);
          qtyMismatches.push({
            number: pn,
            title: item ? item.title : '',
            description: item ? item.description : '',
            cadQty: cadTotal,
            imQty: imTotal,
            cadBreakdown: ct.breakdowns.get(pn) || [],
            imBreakdown: imIndex.breakdowns.get(pn) || [],
          });
        }
      }
      qtyMismatches.sort(function (a, b) { return a.number < b.number ? -1 : 1; });
      qtyMismatches = filterIgnoredFlat(qtyMismatches, isIgnored, 'qty', ignoredFindings);
    }

    // 3) in Item Master only
    const imOnlyRaw = [];
    const seenImOnly = new Set();
    for (const row of im.rows) {
      const pn = normNumber(row.number);
      if (!pn || cadPNs.has(pn) || seenImOnly.has(pn)) continue;
      // The Item Master's own root row is the machine's top-level record; it
      // is never a line in the CAD BOM, so it is not a finding.
      if (Array.isArray(row.path) && row.path.length === 0) continue;
      seenImOnly.add(pn);
      const parentRow = imIndex.parentOf.get(row) || null;
      const entry = Object.assign({}, row, {
        parentNumber: parentRow ? parentRow.number : '',
        parentTitle: parentRow ? (parentRow.title || '') : '',
      });
      // These entries are COPIES, so they are not keys in the parent index
      // (which is keyed by row identity). Keep a link back to the source row
      // so grouping can still resolve the real ancestor chain. Non-enumerable
      // so it never leaks into an export sheet.
      Object.defineProperty(entry, '__srcRow', { value: row, enumerable: false });
      imOnlyRaw.push(entry);
    }
    // Filtered before grouping (not after) so groupImOnly() builds its tree
    // fresh from the already-ignored-free flat list — an ignored row's
    // non-ignored children then attach to the next real ancestor on their
    // own, the same promotion groupImOnly already does for any child whose
    // immediate parent isn't itself flagged.
    const imOnly = filterIgnoredFlat(imOnlyRaw, isIgnored, 'imOnly', ignoredFindings);
    // Virtual parts (js/virtual-parts.js, computed by the caller since it is
    // a separate module) act as grouping anchors: they ARE in the CAD BOM, so
    // they are not Item-Master-only rows themselves, but their whole child BOM
    // is — without this those children scatter as one root each.
    const imOnlyRoots = groupImOnly(imOnly, imIndex.parentOf, opts && opts.virtualAnchorRows);

    // 4) reference components: in the full CAD structure but not in the
    // intended-BOM export — the direct review list for "was this meant to be
    // reference?". Needs both sources. The machine's own root record is not a
    // component; treat it as present.
    let referenceRoots = null;
    let referenceTotal = 0;
    if (bom && bom !== structure) {
      const bomPNs = new Set();
      for (const it of bom.items) bomPNs.add(normNumber(it.number));
      const rootPN = rootPNOf(structure);
      const inBom = function (pn) { return bomPNs.has(pn) || pn === rootPN; };
      referenceRoots = structure.hasLevels
        ? groupMissingLeveled(structure, inBom)
        : groupMissingFlat(structure, inBom, cadChildSets(bom));
      (function annotate(nodes) {
        for (const n of nodes) {
          n.inItemMaster = inIM(normNumber(n.item.number));
          annotate(n.children);
        }
      })(referenceRoots);
      referenceRoots = filterIgnoredTree(referenceRoots, isIgnored, 'reference', ignoredFindings);
      const seenRef = new Set();
      for (const it of structure.items) {
        const pn = normNumber(it.number);
        if (pn && !inBom(pn) && !seenRef.has(pn) && !isIgnored(pn, 'reference')) seenRef.add(pn);
      }
      referenceTotal = seenRef.size;
    }

    // 5) cross-source occurrence index, for on-screen traceability — every
    // place a PN was found in each loaded CAD source. Built over every
    // source (not just qtySource) and regardless of whether any source has
    // quantities, so a flagged row in any results list can point back at its
    // own file/row/page (imOccurrences below covers the Item Master side)
    // without the user reopening any of the uploaded files by hand. An
    // absent PN in either map means "not found in that source" — the same
    // signal "missing"/"in Item Master only" are already built from.
    const cadOccurrences = new Map(); // PN -> [{source, fileName, sourceRow, page, file, parentNumber, parentTitle, qty, effQty}]
    for (const src of [structure, bom]) {
      if (!src) continue;
      const ctSrc = cadTotals(src);
      for (const [pn, occs] of ctSrc.breakdowns) {
        const tagged = occs.map(function (o) {
          return Object.assign({}, o, { source: src.source, fileName: src.fileName || '' });
        });
        cadOccurrences.set(pn, (cadOccurrences.get(pn) || []).concat(tagged));
      }
    }

    return {
      cadUniqueCount: cadPNs.size,
      imUniqueCount: imIndex.byNumber.size,
      missingTotal: missingTotal,             // unique missing PNs, incl. grouped children
      missingRoots: missingRoots,             // actionable top-level findings (tree)
      actionableCount: missingRoots.length,
      qtyMismatches: qtyMismatches,           // null when no CAD source has quantities
      qtyCascades: qtyCascades,               // {applicable, roots} — subtrees released at one uniform ratio
      imOnly: imOnly,                         // flat list (exports + findings registry)
      imOnlyRoots: imOnlyRoots,               // same entries grouped under their IM parent
      imOnlyActionable: imOnlyRoots.length,
      referenceRoots: referenceRoots,         // null unless structure + intended-BOM sources present
      referenceTotal: referenceTotal,
      ignoredFindings: ignoredFindings,       // [{checkKey, number, title, description, sourceRow, parentNumber, parentTitle}] — suppressed by the Ignore List
      cadOccurrences: cadOccurrences,         // PN -> [{source, fileName, sourceRow, page, file, parentNumber, parentTitle, qty, effQty}]
      imOccurrences: imIndex.breakdowns,      // same shape, Item Master side
      hasQty: !!qtySource,
      hasLevels: structure.hasLevels,
      roles: {
        structure: { source: structure.source, fileName: structure.fileName || '' },
        bom: bom ? { source: bom.source, fileName: bom.fileName || '' } : null,
      },
      warnings: imIndex.warnings,
    };
  }

  // single-source compatibility wrapper
  function compare(cad, im) {
    return compareAll([cad], im);
  }

  return {
    normNumber: normNumber,
    indexItemMaster: indexItemMaster,
    cadTotals: cadTotals,
    cadChildSets: cadChildSets,
    compare: compare,
    compareAll: compareAll,
    countDescendants: countDescendants,
    groupImOnly: groupImOnly,
    detectQuantityCascades: detectQuantityCascades,
    filterIgnoredTree: filterIgnoredTree,
    filterIgnoredFlat: filterIgnoredFlat,
  };
});
