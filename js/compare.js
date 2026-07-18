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

  // Builds lookup structures from the Item Master rows.
  function indexItemMaster(im) {
    const byNumber = new Map();       // PN -> [rows]
    const byPath = new Map();         // 'path.key' -> row (first wins)
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

    // childSets: PN -> Set of direct-child PNs (union over all occurrences).
    // Needs paths. children of path p are rows whose path is p + one segment.
    const childSets = new Map();
    const childrenOfPath = new Map(); // parent path key -> [rows]
    for (const row of im.rows) {
      if (!Array.isArray(row.path) || row.path.length === 0) continue;
      const parentKey = row.path.slice(0, -1).join('.');
      if (!childrenOfPath.has(parentKey)) childrenOfPath.set(parentKey, []);
      childrenOfPath.get(parentKey).push(row);
    }
    for (const [key, row] of byPath) {
      const kids = childrenOfPath.get(key);
      if (!kids) continue;
      const pn = normNumber(row.number);
      if (!childSets.has(pn)) childSets.set(pn, new Set());
      const set = childSets.get(pn);
      for (const k of kids) set.add(normNumber(k.number));
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
      if (eff !== null && hasPaths && Array.isArray(row.path)) {
        for (let i = row.path.length - 1; i >= 1 && eff !== null; i--) {
          const anc = byPath.get(row.path.slice(0, i).join('.'));
          if (!anc) continue; // gap in export; treat multiplier as 1
          if (i === row.path.length - 1) parentRow = anc;
          if (anc.qty !== null) eff *= anc.qty;
        }
      }
      if (!breakdowns.has(pn)) breakdowns.set(pn, []);
      breakdowns.get(pn).push({
        parentNumber: parentRow ? normNumber(parentRow.number) : '',
        parentTitle: parentRow ? (parentRow.title || '') : '',
        qty: row.qty,
        effQty: eff,
      });
      if (row.qty === null) { totals.set(pn, null); continue; }
      if (totals.get(pn) !== null || !totals.has(pn)) {
        totals.set(pn, (totals.get(pn) || 0) + (eff === null ? row.qty : eff));
      }
    }

    return { byNumber: byNumber, byPath: byPath, childSets: childSets, totals: totals, breakdowns: breakdowns, warnings: warnings };
  }

  /* ------------------------------------------------------------------ *
   * CAD quantity roll-up
   * ------------------------------------------------------------------ */

  // Per-PN rolled-up totals for a leveled CAD BOM (levels + qty per row).
  function cadTotals(cad) {
    if (!cad.hasQty) return { totals: null, breakdowns: null };
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
   * Main entry
   * ------------------------------------------------------------------ */

  // cadSources: one or two parsed CAD results (e.g. Vault PDF + Inventor xlsx)
  function compareAll(cadSources, im) {
    const imIndex = indexItemMaster(im);
    const roles = pickRoles(cadSources);
    const structure = roles.structure;
    const bom = roles.bom;
    const inIM = function (pn) { return imIndex.byNumber.has(pn); };

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
    const missingRoots = structure.hasLevels
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
    let missingTotal = 0;
    for (const pn of cadPNs) if (!inIM(pn)) missingTotal++;

    // 2) quantity mismatches — from whichever source carries quantities
    const qtySource = (bom && bom.hasQty) ? bom : (structure.hasQty ? structure : null);
    let qtyMismatches = null;
    if (qtySource) {
      const ct = cadTotals(qtySource);
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
    }

    // 3) in Item Master only
    const imOnly = [];
    const seenImOnly = new Set();
    for (const row of im.rows) {
      const pn = normNumber(row.number);
      if (!pn || cadPNs.has(pn) || seenImOnly.has(pn)) continue;
      seenImOnly.add(pn);
      const parentRow = Array.isArray(row.path) && row.path.length
        ? imIndex.byPath.get(row.path.slice(0, -1).join('.'))
        : null;
      imOnly.push(Object.assign({}, row, {
        parentNumber: parentRow ? parentRow.number : '',
        parentTitle: parentRow ? (parentRow.title || '') : '',
      }));
    }

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
      const seenRef = new Set();
      for (const it of structure.items) {
        const pn = normNumber(it.number);
        if (pn && !inBom(pn) && !seenRef.has(pn)) seenRef.add(pn);
      }
      referenceTotal = seenRef.size;
    }

    return {
      cadUniqueCount: cadPNs.size,
      imUniqueCount: imIndex.byNumber.size,
      missingTotal: missingTotal,             // unique missing PNs, incl. grouped children
      missingRoots: missingRoots,             // actionable top-level findings (tree)
      actionableCount: missingRoots.length,
      qtyMismatches: qtyMismatches,           // null when no CAD source has quantities
      imOnly: imOnly,
      referenceRoots: referenceRoots,         // null unless structure + intended-BOM sources present
      referenceTotal: referenceTotal,
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
  };
});
