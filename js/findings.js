/*
 * findings.js — cross-check findings registry. Pure logic (no DOM).
 *
 * Every check runs independently, so one part can legitimately be flagged by
 * several of them at once: a part that is in the Item Master but not in CAD
 * shows up under "In Item Master only" AND, if its Quantity and Item Qty
 * disagree, under Check 3 as well. Reported side by side that reads as several
 * different problems when it is one part needing one look.
 *
 * This module collects every finding, keys them by part number, and ranks them
 * by severity so each part gets exactly ONE primary finding — the most serious
 * check that flagged it — while the rest become secondary cross-references.
 * Nothing is dropped: the detail sections still show every row, they just
 * render the secondary ones muted and pointing at the primary.
 *
 * Produces: { parts, byPn, isSecondary(checkKey, pn), counts }
 *   parts:  [{ number, title, description, primary, issues[], grouped,
 *              groupedUnder }] sorted by severity (worst first), then number
 *   issues: [{ key, label, severity, section, tab, detail, sourceRow,
 *              parentNumber, parentTitle, primary:bool }]
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Same normalization as compare.js — findings arrive with `number` in mixed
  // form (only compare.js's qtyMismatches stores it already-normalized), so
  // everything is normalized on the way in.
  function normNumber(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim().toUpperCase();
  }

  // Severity order decides which check "owns" a part when several flag it.
  // Rough principle: a part that will never be ordered outranks a part whose
  // quantity is wrong, which outranks a field being inconsistent inside the
  // Item Master. `section`/`tab` are where the UI should jump to.
  const CHECKS = [
    // Above "missing": a sketch part that reached the released Item Master is
    // the most serious thing this tool can find — it should never be there at
    // all, so it always owns the part rather than being demoted behind
    // another finding.
    { key: 'c8', severity: 120, label: 'Sketch part (7-333-) in Item Master', section: 'im-qc' },
    { key: 'c9', severity: 110, label: 'Obsolete / invalid state', section: 'im-qc' },
    // "New" (not yet certified) is a warning, not a release-blocking error, so
    // it ranks below the real findings instead of dominating the list — a
    // genuine export can carry several legitimately-New rows.
    { key: 'c9warn', severity: 47, label: 'Not yet certified', section: 'im-qc' },
    { key: 'missing', severity: 100, label: 'Missing from Item Master', section: 'results', tab: 'missing' },
    { key: 'lldboMissing', severity: 95, label: 'Long-lead part missing from Item Master', section: 'lldbo-panel' },
    // A virtual subassembly hides its whole child BOM from every other check:
    // the part itself passes both presence checks, while its children scatter
    // as unexplained "In Item Master only" rows. Ranked above `reference` so
    // it owns those children rather than competing with them.
    { key: 'virtualPart', severity: 92, label: 'Virtual part (no CAD file behind it)', section: 'results', tab: 'virtual' },
    { key: 'reference', severity: 90, label: 'Reference item', section: 'results', tab: 'ref' },
    // A whole subtree released at one uniform ratio (e.g. every direct child
    // of one assembly entered at exactly 2x) is a single root-cause error,
    // not N independent ones — it outranks the ordinary per-part quantity
    // check so every descendant it explains is demoted under it instead of
    // flooding the list with its own separate findings.
    { key: 'qtyCascade', severity: 85, label: 'Quantity cascade (whole subtree released at one ratio)', section: 'results', tab: 'cascade' },
    { key: 'qty', severity: 80, label: 'Quantity mismatch (CAD vs Item Master)', section: 'results', tab: 'qty' },
    { key: 'lldboQty', severity: 75, label: 'Long-lead quantity mismatch', section: 'lldbo-panel' },
    // Below lldboQty (a confirmed mismatch on an already-tracked part) since
    // this is a heuristic inference on an untracked part; above the
    // data-quality checks since an untracked long-lead part risks never
    // getting ordered, not just a records nit.
    { key: 'lldboCandidate', severity: 72, label: 'Long-lead candidate (motor/actuator/seal/nozzle) not tracked in LLDBO', section: 'lldbo-panel' },
    { key: 'revision', severity: 70, label: 'Revision mismatch vs CAD', section: 'revision-sections' },
    { key: 'material', severity: 60, label: 'Material mismatch vs CAD', section: 'material-sections' },
    { key: 'titleDesc', severity: 55, label: 'Description mismatch vs CAD', section: 'titledesc-sections' },
    { key: 'imOnly', severity: 50, label: 'In Item Master only', section: 'results', tab: 'imonly' },
    { key: 'c3', severity: 45, label: 'Quantity vs Item Qty', section: 'im-qc' },
    { key: 'c7', severity: 44, label: 'Revision inconsistent across positions', section: 'im-qc' },
    { key: 'c6', severity: 43, label: 'Material missing', section: 'im-qc' },
    { key: 'c5', severity: 42, label: 'Title / Description missing', section: 'im-qc' },
    { key: 'c4', severity: 41, label: 'Entity Icon not "Normal"', section: 'im-qc' },
    { key: 'c1', severity: 40, label: 'Producer not in Description', section: 'im-qc' },
    { key: 'c2', severity: 39, label: 'End of Line integrity', section: 'im-qc' },
  ];

  const CHECK_BY_KEY = {};
  for (const c of CHECKS) CHECK_BY_KEY[c.key] = c;

  function fmtQty(v) {
    if (v === null || v === undefined) return '?';
    return Math.abs(v - Math.round(v)) < 1e-9 ? String(Math.round(v)) : String(Math.round(v * 1000) / 1000);
  }

  function buildRegistry(sources) {
    sources = sources || {};
    const byPn = new Map();

    // One entry per (part, check). A part legitimately appears in several
    // checks; within a single check it is recorded once (checks that emit one
    // row per BOM position — imqc c3..c7 — collapse here, with the extra
    // positions counted in `occurrences`).
    function record(checkKey, rawNumber, fields) {
      const pn = normNumber(rawNumber);
      // Check 2 emits a synthetic '—' row when the BOM has no END OF LINE
      // entry at all; that is a whole-BOM assertion, not a part.
      if (!pn || pn === '—' || pn === '-') return;
      const meta = CHECK_BY_KEY[checkKey];
      if (!meta) return;
      let part = byPn.get(pn);
      if (!part) {
        part = {
          number: fields.number || rawNumber,
          title: fields.title || '',
          description: fields.description || '',
          issues: [],
          primary: null,
          grouped: false,
          groupedUnder: null,
        };
        byPn.set(pn, part);
      }
      if (!part.title && fields.title) part.title = fields.title;
      if (!part.description && fields.description) part.description = fields.description;

      const existing = part.issues.find(function (i) { return i.key === checkKey; });
      if (existing) { existing.occurrences++; return; }
      part.issues.push({
        key: checkKey,
        label: meta.label,
        severity: meta.severity,
        section: meta.section,
        tab: meta.tab || null,
        detail: fields.detail || '',
        sourceRow: fields.sourceRow === undefined ? '' : fields.sourceRow,
        parentNumber: fields.parentNumber || '',
        parentTitle: fields.parentTitle || '',
        occurrences: 1,
        primary: false,
      });
    }

    // Tree-shaped findings (missing / reference / imOnly): depth 0 is the
    // actionable finding; deeper nodes are children explained by their parent
    // and are marked `grouped` so they don't count as separate problems.
    function recordTree(checkKey, roots, detailFor) {
      const walk = function (node, depth, rootPn) {
        const it = node.item;
        const pn = normNumber(it.number);
        record(checkKey, it.number, {
          number: it.number,
          title: it.title || '',
          description: it.description || '',
          detail: detailFor ? detailFor(node, depth) : '',
          sourceRow: it.sourceRow || '',
          parentNumber: depth === 0 ? '' : rootPn,
        });
        const part = byPn.get(pn);
        if (part && depth > 0 && !part.grouped) {
          part.grouped = true;
          part.groupedUnder = rootPn;
        }
        for (const c of node.children || []) walk(c, depth + 1, depth === 0 ? pn : rootPn);
      };
      for (const n of roots || []) walk(n, 0, null);
    }

    const res = sources.result;
    if (res) {
      recordTree('missing', res.missingRoots);
      if (res.referenceRoots) recordTree('reference', res.referenceRoots);
      if (res.qtyCascades && res.qtyCascades.applicable && res.qtyCascades.roots.length) {
        recordTree('qtyCascade', res.qtyCascades.roots, function (node, depth) {
          if (depth !== 0) return '';
          const it = node.item;
          return it.cascadeMismatchedChildCount + ' of ' + it.cascadeChildCount +
            ' children released at ' + it.cascadeRatio + '×';
        });
      }
      for (const m of res.qtyMismatches || []) {
        record('qty', m.number, {
          number: m.number, title: m.title, description: m.description,
          detail: 'CAD ' + fmtQty(m.cadQty) + ' vs Item Master ' + fmtQty(m.imQty),
        });
      }
      // Prefer the grouped tree so a whole subassembly missing from the CAD BOM
      // is one finding, not one per part; fall back to the flat list.
      if (res.imOnlyRoots) recordTree('imOnly', res.imOnlyRoots);
      else for (const r of res.imOnly || []) {
        record('imOnly', r.number, {
          number: r.number, title: r.title, description: r.description,
          sourceRow: r.sourceRow, parentNumber: r.parentNumber, parentTitle: r.parentTitle,
        });
      }
    }

    const mat = sources.materialResult;
    if (mat && mat.applicable) {
      for (const m of mat.mismatches || []) {
        record('material', m.number, {
          number: m.number, title: m.title,
          detail: 'Item Master "' + m.imMaterial + '" vs CAD "' + m.cadMaterial + '"',
          sourceRow: m.sourceRow, parentNumber: m.parentNumber, parentTitle: m.parentTitle,
        });
      }
    }

    const td = sources.titleDescResult;
    if (td && td.applicable) {
      for (const m of td.mismatches || []) {
        record('titleDesc', m.number, {
          number: m.number, title: m.title,
          detail: 'Item Master "' + m.imDescription + '" vs CAD "' + m.cadDescription + '"',
          sourceRow: m.sourceRow, parentNumber: m.parentNumber, parentTitle: m.parentTitle,
        });
      }
    }

    // Virtual parts own the orphaned children they explain, so each one is a
    // single finding rather than N unexplained "In Item Master only" rows.
    const virt = sources.virtualResult;
    if (virt && virt.applicable) {
      for (const v of (virt.confirmed || []).concat(virt.suspected || [])) {
        record('virtualPart', v.number, {
          number: v.number, title: v.title, description: v.description,
          detail: (v.confidence === 'suspected' ? 'Suspected: ' : '') +
            v.childCount + ' child part(s) exist only in the Item Master',
          sourceRow: v.sourceRow,
        });
        for (const c of v.children || []) {
          record('virtualPart', c.number, {
            number: c.number, title: c.title, sourceRow: c.sourceRow,
            detail: 'Child of virtual part ' + v.number,
            parentNumber: v.number, parentTitle: v.title,
          });
          const part = byPn.get(normNumber(c.number));
          if (part && !part.grouped) {
            part.grouped = true;
            part.groupedUnder = normNumber(v.number);
          }
        }
      }
    }

    const rev = sources.revisionResult;
    if (rev && rev.applicable) {
      for (const m of rev.mismatches || []) {
        record('revision', m.number, {
          number: m.number, title: m.title,
          detail: 'Item Master "' + m.imRevision + '" vs CAD "' + m.cadRevision + '"',
          sourceRow: m.sourceRow, parentNumber: m.parentNumber, parentTitle: m.parentTitle,
        });
      }
    }

    const qc = sources.imQc;
    if (qc) {
      const detailFor = {
        c3: function (f) { return 'Quantity ' + f.quantity + ' vs Item Qty ' + f.itemQty; },
        c7: function (f) { return 'Revision ' + f.revision + '; also ' + f.conflictsWith; },
        c5: function (f) { return f.kind === 'both-missing' ? 'Title and Description blank' : (f.kind === 'title-missing' ? 'Title blank' : 'Description blank'); },
        c4: function (f) { return 'Entity Icon "' + f.icon + '"'; },
        c1: function (f) { return f.issue || ''; },
        c2: function (f) { return f.issue || ''; },
        c8: function (f) { return 'Found at ' + (f.trail || 'top level') + (f.state ? ' — state ' + f.state : ''); },
        c9: function (f) { return 'State "' + f.state + '" — ' + f.severity; },
      };
      for (const key of ['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9']) {
        const r = qc[key];
        if (!r || !r.applicable) continue;
        for (const f of r.fail || []) {
          // Check 9 mixes release-blocking states with merely-uncertified
          // ones; they are ranked separately so a legitimately "New" row
          // doesn't outrank a part that is missing outright.
          const k = (key === 'c9' && !/^ERROR/.test(f.severity || '')) ? 'c9warn' : key;
          record(k, f.number, {
            number: f.number, title: f.title || '',
            detail: detailFor[key] ? detailFor[key](f) : '',
            sourceRow: f.sourceRow, parentNumber: f.parentNumber, parentTitle: f.parentTitle,
          });
        }
      }
    }

    const lld = sources.lldboResult;
    if (lld) {
      for (const m of lld.missingFromIm || []) {
        record('lldboMissing', m.number, { number: m.number, description: m.description, sourceRow: m.sourceRow });
      }
      for (const m of lld.qtyMismatches || []) {
        record('lldboQty', m.number, {
          number: m.number, description: m.description,
          detail: 'LLDBO ' + fmtQty(m.lldboQty) + ' vs Item Master ' + fmtQty(m.imQty),
          sourceRow: m.sourceRow,
        });
      }
    }

    // Only the confident tier, and only once cross-checked against a loaded
    // LLDBO file — the review tier (keyword match outside the X-999-*
    // convention) is deliberately never recorded here (needs a human look,
    // not an automatic finding), and the Item-Master-only preview state
    // (crossChecked:false) is a "these might need tracking" list, not yet a
    // confirmed gap, so it must not feed "Parts needing attention" either.
    const lldCand = sources.lldboCandidatesResult;
    if (lldCand && lldCand.crossChecked) {
      for (const c of lldCand.confident || []) {
        record('lldboCandidate', c.number, {
          number: c.number, title: c.title, description: c.description,
          detail: 'Matched keyword "' + c.keyword + '" — purchased-part-numbered (X-999-*), not found in the loaded LLDBO file',
          sourceRow: c.sourceRow, parentNumber: c.parentNumber, parentTitle: c.parentTitle,
        });
      }
    }

    // Rank: the most severe issue on each part becomes its primary.
    const parts = [];
    byPn.forEach(function (part) {
      part.issues.sort(function (a, b) { return b.severity - a.severity; });
      part.primary = part.issues[0] || null;
      if (part.primary) part.primary.primary = true;
      parts.push(part);
    });
    parts.sort(function (a, b) {
      const sa = a.primary ? a.primary.severity : 0;
      const sb = b.primary ? b.primary.severity : 0;
      if (sa !== sb) return sb - sa;
      return normNumber(a.number) < normNumber(b.number) ? -1 : 1;
    });

    const primaryByCheck = {};
    let secondaryTotal = 0;
    let actionable = 0;
    for (const p of parts) {
      if (!p.grouped) actionable++;
      for (const i of p.issues) {
        if (i.primary) primaryByCheck[i.key] = (primaryByCheck[i.key] || 0) + 1;
        else secondaryTotal++;
      }
    }

    return {
      parts: parts,
      byPn: byPn,
      // True when this check is NOT the part's primary owner — the UI uses it
      // to render the row muted with a pointer at the primary finding.
      isSecondary: function (checkKey, number) {
        const part = byPn.get(normNumber(number));
        if (!part || !part.primary) return false;
        return part.primary.key !== checkKey;
      },
      primaryFor: function (number) {
        const part = byPn.get(normNumber(number));
        return part ? part.primary : null;
      },
      counts: {
        parts: parts.length,
        actionable: actionable,
        grouped: parts.length - actionable,
        primaryByCheck: primaryByCheck,
        secondaryTotal: secondaryTotal,
      },
    };
  }

  return { findings: { buildRegistry: buildRegistry, CHECKS: CHECKS, CHECK_BY_KEY: CHECK_BY_KEY, normNumber: normNumber } };
});
