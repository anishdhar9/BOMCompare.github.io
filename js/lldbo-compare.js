/*
 * lldbo-compare.js — compares a parsed LLDBO (Long Lead Direct Bought Out)
 * list against the Item Master. Long-lead parts are released to
 * procurement ahead of the normal BOM release to account for supplier lead
 * times; compareLldbo() verifies each one actually made it into the
 * released Item Master with a matching quantity — catching the process
 * failure where an early release never got captured, so the part quietly
 * never gets ordered through the normal channel either.
 *
 * detectLldboCandidates() is the complementary, opposite-direction check:
 * it starts from the Item Master instead, looking for parts that LOOK LIKE
 * they should be tracked as long-lead (motors, actuators, specialty seals,
 * nozzles) and flags ones not present in the LLDBO file — catching a
 * genuine long-lead part that never got added to LLDBO tracking in the
 * first place.
 *
 * Pure logic (no DOM). `indexItemMaster` is injected into compareLldbo()
 * (it's compare.js's function) rather than required directly, matching
 * this codebase's dependency-injection style (see itemmaster.js's
 * `parse(workbook, XLSX)`, cad-leveled.js's XLSX parameter) instead of a
 * hard load-order coupling. detectLldboCandidates() instead injects `imQc`
 * and `compare.js` wholesale through the UMD factory (same pattern as
 * js/revision-compare.js), since it needs several of their helpers
 * (PURCHASED_PART_RE, buildPathIndex/parentOf, filterIgnoredFlat).
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(require('./imqc.js').imQc, require('./compare.js'));
  } else {
    const bc = root.BOMCompare || {};
    root.BOMCompare = Object.assign(bc, factory(bc.imQc, bc));
  }
})(typeof self !== 'undefined' ? self : this, function (imQc, compareLib) {
  'use strict';

  function normNumber(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim().toUpperCase();
  }

  function compareLldbo(lldbo, im, indexItemMaster) {
    const imIndex = indexItemMaster(im);

    // Aggregate LLDBO rows by part number — duplicates are summed (e.g. the
    // same catalog motor used in two different assemblies, each needing 1,
    // means 2 are expected in total; verified as a real, legitimate pattern
    // in the sample data, not a data error).
    const byPn = new Map(); // PN -> { number, rows:[...], totalQty, hasQty }
    let noPartNumber = 0;
    for (const row of lldbo.rows) {
      const pn = normNumber(row.partNo);
      if (!pn) { noPartNumber++; continue; }
      if (!byPn.has(pn)) byPn.set(pn, { number: row.partNo, rows: [], totalQty: 0, hasQty: true });
      const entry = byPn.get(pn);
      entry.rows.push(row);
      if (row.qty === null) entry.hasQty = false;
      else entry.totalQty += row.qty;
    }

    const missingFromIm = [];
    const qtyMismatches = [];
    for (const [pn, entry] of byPn) {
      const descriptions = entry.rows.map(function (r) { return r.description; }).filter(Boolean);
      const sourceRows = entry.rows.map(function (r) { return r.sourceRow; }).filter(Boolean).join(', ');
      if (!imIndex.byNumber.has(pn)) {
        missingFromIm.push({
          number: entry.number,
          description: descriptions.join(' / '),
          qtyText: entry.rows.map(function (r) { return r.qtyText; }).filter(Boolean).join(' + '),
          sourceRow: sourceRows,
          rows: entry.rows,
        });
        continue;
      }
      if (!entry.hasQty) continue; // present, but no comparable quantity ("NA")
      const imTotal = imIndex.totals.has(pn) ? imIndex.totals.get(pn) : null;
      if (imTotal === null) continue; // not computable on the Item Master side
      if (Math.abs(imTotal - entry.totalQty) > 1e-9) {
        const breakdown = imIndex.breakdowns.get(pn) || [];
        const foundUnder = breakdown
          .map(function (b) { return b.parentNumber ? (b.parentNumber + (b.parentTitle ? ' (' + b.parentTitle + ')' : '')) : ''; })
          .filter(Boolean).join(' + ');
        qtyMismatches.push({
          number: entry.number,
          description: descriptions.join(' / '),
          lldboQty: entry.totalQty,
          imQty: imTotal,
          sourceRow: sourceRows,
          foundUnder: foundUnder,
          rows: entry.rows,
        });
      }
    }

    return {
      totalLldboItems: byPn.size,
      noPartNumberCount: noPartNumber,
      missingFromIm: missingFromIm,
      qtyMismatches: qtyMismatches,
      projectKeyMismatch: (lldbo.projectKey && im.projectKey && lldbo.projectKey.pn !== im.projectKey.pn)
        ? { lldbo: lldbo.projectKey, im: im.projectKey }
        : null,
    };
  }

  // Keyword list validated against real Item Master samples: English terms
  // for the named part families (motor/actuator/seal/nozzle), plus "gasket"
  // and "servo"/"cylinder" as adjacent long-lead-shaped hardware, plus the
  // German terms actually seen on real bought parts ("dichtung" = seal,
  // "düse" = nozzle — this organization sources some of these from Germany).
  // Substring match, case-insensitive, over Title+Description — verified
  // against a real ~1250-row Item Master to produce zero false substring
  // hits (e.g. "cylindrical" does not contain "cylinder").
  const LLDBO_CANDIDATE_KEYWORDS = ['motor', 'actuator', 'seal', 'gasket', 'servo', 'cylinder', 'dichtung', 'nozzle', 'düse'];

  function matchCandidateKeyword(row) {
    const text = ((row.title || '') + ' ' + (row.description || '')).toLowerCase();
    for (const kw of LLDBO_CANDIDATE_KEYWORDS) {
      if (text.indexOf(kw) !== -1) return kw;
    }
    return null;
  }

  // Scans the Item Master itself for parts that look like they should be
  // tracked as long-lead (motors, actuators, specialty seals, nozzles) and
  // flags ones not present in the loaded LLDBO file — the complementary
  // check to compareLldbo() above, which starts from the LLDBO file instead.
  //
  // Two confidence tiers, since a keyword match alone is unreliable (a
  // custom bracket or cover that merely MENTIONS "motor"/"seal" in its title
  // is not itself a bought long-lead part): "confident" additionally
  // requires the part number to follow the existing purchased-part
  // convention (imQc.PURCHASED_PART_RE, "X-999-*"); "review" is a keyword
  // match outside that convention — shown separately for a human to check,
  // never treated as an actionable finding on its own (see js/findings.js).
  //
  // Assemblies are skipped even when their own title/description matches a
  // keyword (e.g. "Motor Cover Assembly"): an assembly is a container, not
  // itself the purchasable long-lead item — the real candidate is one of
  // its children (e.g. the actual motor inside it), which gets its own,
  // separate entry when the loop reaches that row. "Has a child" is read
  // straight from the Item Master's own Row Order hierarchy
  // (imQc.buildAssemblySet, the same Item-Master-only helper
  // material-compare.js already uses to exclude assemblies from the
  // Material Completeness check) — this never looks at CAD data.
  //
  // Works with the Item Master alone: `lldbo` may be null/undefined (no
  // LLDBO file loaded yet). In that case every candidate is returned
  // unfiltered as an uncross-checked preview (crossChecked:false) — nothing
  // here is a confirmed gap yet, just "these parts look like they might need
  // LLDBO tracking." Once `lldbo` is supplied, candidates already present in
  // its Part No set are removed from both tiers (crossChecked:true) — only
  // genuinely-absent ones remain.
  //
  // opts.isIgnored(pn, checkKey) — same shape as revisionCompare.compareRevision.
  function detectLldboCandidates(im, lldbo, opts) {
    const isIgnored = (opts && opts.isIgnored) || function () { return false; };
    const pathIndex = imQc.buildPathIndex(im.rows);
    const assemblies = imQc.buildAssemblySet(im.rows, pathIndex);

    const confidentAll = [];
    const reviewAll = [];
    const seen = new Set(); // same part can occur at several BOM positions; report it once
    for (const row of im.rows) {
      if (imQc.isEndOfLine(row)) continue; // ERP completeness marker, not a part
      if (Array.isArray(row.path) && row.path.length === 0) continue; // root row: the machine itself
      if (assemblies.has(row)) continue; // container, not itself a purchasable part — see comment above
      const pn = normNumber(row.number);
      if (!pn || seen.has(pn)) continue;
      const kw = matchCandidateKeyword(row);
      if (!kw) continue;
      seen.add(pn);
      const parent = imQc.parentOf(pathIndex, row);
      const entry = {
        number: row.number, title: row.title || '', description: row.description || '',
        keyword: kw, quantity: row.quantityText || (row.qty === null || row.qty === undefined ? '' : String(row.qty)),
        sourceRow: row.sourceRow || '',
        parentNumber: parent ? parent.number : '', parentTitle: parent ? (parent.title || '') : '',
      };
      (imQc.PURCHASED_PART_RE.test(row.number) ? confidentAll : reviewAll).push(entry);
    }

    const crossChecked = !!(lldbo && Array.isArray(lldbo.rows));
    let confident = confidentAll, review = reviewAll, trackedCount = 0;
    if (crossChecked) {
      const tracked = new Set();
      for (const r of lldbo.rows) {
        const pn = normNumber(r.partNo);
        if (pn) tracked.add(pn);
      }
      const notTracked = function (c) { return !tracked.has(normNumber(c.number)); };
      confident = confidentAll.filter(notTracked);
      review = reviewAll.filter(notTracked);
      trackedCount = (confidentAll.length - confident.length) + (reviewAll.length - review.length);
    }

    const ignoredFindings = [];
    confident = compareLib.filterIgnoredFlat(confident, isIgnored, 'lldboCandidate', ignoredFindings);
    review = compareLib.filterIgnoredFlat(review, isIgnored, 'lldboCandidate', ignoredFindings);

    return {
      crossChecked: crossChecked,
      confident: confident,
      review: review,
      trackedCount: trackedCount,
      ignoredFindings: ignoredFindings,
    };
  }

  return {
    lldboCompare: {
      compareLldbo: compareLldbo,
      detectLldboCandidates: detectLldboCandidates,
      LLDBO_CANDIDATE_KEYWORDS: LLDBO_CANDIDATE_KEYWORDS,
    },
  };
});
