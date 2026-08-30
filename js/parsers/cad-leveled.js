/*
 * cad-leveled.js — parser for multi-level CAD BOM tables with a header row:
 * a leveled Excel/CSV export, or the row grid reconstructed from the Vault
 * multi-level BOM PDF by pdf-extract.js.
 *
 * Column sets vary (Vault lets users pick visible fields), so columns are
 * auto-detected by header keywords and can be overridden by an explicit
 * `mapping` (from the app's column-mapping UI).
 *
 * Hierarchy level per row, in order of preference:
 *   1. an explicit Level column (numeric)
 *   2. a dotted item/position column ('1.2.3' -> level 3)
 *   3. row indentation (leading spaces, or x-offsets supplied by pdf-extract)
 *
 * Produces: { kind:'cad', source, hasQty, hasLevels, hasMaterial, hasRevision,
 *             items:[...], columns, headerRow, warnings }
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(require('./qty-parse.js').qtyParse);
  } else {
    const bc = root.BOMCompare || {};
    root.BOMCompare = Object.assign(bc, factory(bc.qtyParse));
  }
})(typeof self !== 'undefined' ? self : this, function (qtyParse) {
  'use strict';

  function cellText(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  // Locale-aware ('1,5' -> 1.5, '1.234,5' -> 1234.5, '1,200' -> 1200) --
  // shared with itemmaster.js and lldbo.js, see js/parsers/qty-parse.js.
  const parseQty = qtyParse.parseQty;

  const FIELD_KEYWORDS = {
    number: ['number', 'part number', 'part no', 'part no.', 'item number', 'document number', 'artikelnummer', 'teilenummer', 'sachnummer'],
    // 'unit qty' is deliberately NOT a qty keyword: Inventor's structured
    // export has both 'Unit QTY' (text such as 'Each') and 'QTY' (the count).
    // The other per-unit-quantity phrasings below are the same decoy,
    // matching itemmaster.js's quantityPerUnit list -- a per-unit column
    // holds the quantity of one single unit of the parent (usually 1, even
    // when the real total Quantity is 4), and must never be captured as qty.
    qty: ['qty', 'qty.', 'quantity', 'item qty', 'menge', 'anzahl', 'stück', 'stck'],
    unit: [
      'unit', 'unit qty', 'unit quantity', 'bom unit', 'base unit', 'einheit',
      'quantity per unit', 'qty per unit', 'qty. per unit', 'quantity/unit', 'qty/unit', 'qty per parent',
    ],
    level: ['level', 'ebene', 'stufe', 'depth'],
    pos: ['item', 'pos', 'pos.', 'position', 'row order', 'bom structure position'],
    structure: ['bom structure', 'bomstructure', 'structure'],
    title: ['title', 'name', 'bezeichnung', 'benennung'],
    // 'item description' is listed explicitly (not just inferred from
    // 'description') because a bare 'item' prefix would otherwise win it for
    // `pos` above -- 'description' alone is not a prefix of "item
    // description", so without this synonym the longest-prefix rule in
    // matchField() has nothing to prefer it over 'item'.
    description: ['description', 'beschreibung', 'item description'],
    file: ['file', 'file name', 'filename', 'dateiname', 'document'],
    material: ['material', 'werkstoff'],
    revision: ['revision', 'rev', 'rev.'],
    // Inventor exports the thumbnail column as an EMPTY cell when a CAD file
    // backs the row, and as the literal text "(NULL)" when none does. That
    // makes "(NULL)" the marker for a virtual component (a BOM entry with no
    // model behind it). The column is user-configurable, so it is often absent.
    thumbnail: ['thumbnail', 'preview', 'image'],
  };

  function matchField(headerText) {
    const h = headerText.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!h) return null;
    // 1. an exact header match wins outright.
    for (const field of Object.keys(FIELD_KEYWORDS)) {
      if (FIELD_KEYWORDS[field].indexOf(h) !== -1) return field;
    }
    // 2. otherwise, prefix match for compound headers like 'Title (Item,CO)'
    //    or 'Qty per Unit (Each)'. When several keywords are a prefix, the
    //    LONGEST (most specific) one wins -- so a short 'item'/'pos' can
    //    never beat a more specific 'item description'/'qty per unit',
    //    regardless of field/column order. Mirrors itemmaster.js's
    //    matchField, which fixed this identical bug class first.
    let best = null, bestLen = 0;
    for (const field of Object.keys(FIELD_KEYWORDS)) {
      for (const kw of FIELD_KEYWORDS[field]) {
        if (kw.length >= 3 && h.indexOf(kw) === 0 && kw.length > bestLen) {
          best = field; bestLen = kw.length;
        }
      }
    }
    return best;
  }

  function detectHeader(aoa) {
    let best = null;
    const limit = Math.min(aoa.length, 12);
    for (let r = 0; r < limit; r++) {
      const row = aoa[r] || [];
      const cols = {};
      let score = 0;
      for (let c = 0; c < row.length; c++) {
        const f = matchField(cellText(row[c]));
        if (f && cols[f] === undefined) { cols[f] = c; score++; }
      }
      if (cols.number === undefined) continue;
      if (score >= 2 && (!best || score > best.score)) best = { headerRow: r, cols: cols, score: score };
    }
    return best;
  }

  // Quantize indentation offsets (leading-space counts or PDF x-offsets)
  // into 1-based levels.
  function levelsFromIndents(indents) {
    const uniq = Array.from(new Set(indents.filter(function (v) { return v !== null; }))).sort(function (a, b) { return a - b; });
    if (uniq.length < 2) return null;
    // merge offsets closer than half the median step (PDF x jitter)
    const steps = [];
    for (let i = 1; i < uniq.length; i++) steps.push(uniq[i] - uniq[i - 1]);
    steps.sort(function (a, b) { return a - b; });
    const tol = steps[Math.floor(steps.length / 2)] / 2;
    const buckets = [uniq[0]];
    for (const v of uniq.slice(1)) {
      if (v - buckets[buckets.length - 1] > tol) buckets.push(v);
    }
    return indents.map(function (v) {
      if (v === null) return null;
      let lvl = 1;
      for (let i = 0; i < buckets.length; i++) if (v >= buckets[i] - tol) lvl = i + 1;
      return lvl;
    });
  }

  // For the mapping UI: header + sample values per column.
  function analyze(aoa) {
    const hdr = detectHeader(aoa);
    const headerRow = hdr ? hdr.headerRow : 0;
    const header = (aoa[headerRow] || []).map(cellText);
    const width = Math.max.apply(null, aoa.slice(0, 30).map(function (r) { return (r || []).length; }).concat([header.length]));
    const columns = [];
    for (let c = 0; c < width; c++) {
      const samples = [];
      for (let r = headerRow + 1; r < aoa.length && samples.length < 5; r++) {
        const t = cellText((aoa[r] || [])[c]);
        if (t) samples.push(t);
      }
      columns.push({ index: c, header: header[c] || '(column ' + (c + 1) + ')', samples: samples });
    }
    return { headerRow: headerRow, columns: columns, detected: hdr ? hdr.cols : { }, score: hdr ? hdr.score : 0 };
  }

  /**
   * @param aoa      array-of-arrays including the header row
   * @param opts     { mapping?: {number,qty,level,pos,title,description,file},
   *                   headerRow?: number,
   *                   indents?: (number|null)[] aligned to aoa rows (from pdf-extract),
   *                   source?: string }
   */
  function parse(aoa, opts) {
    opts = opts || {};
    const warnings = [];
    let cols, headerRow;
    if (opts.mapping) {
      cols = opts.mapping;
      headerRow = opts.headerRow !== undefined ? opts.headerRow : 0;
    } else {
      const hdr = detectHeader(aoa);
      if (!hdr) return null;
      cols = hdr.cols;
      headerRow = hdr.headerRow;
    }
    if (cols.number === undefined || cols.number === null || cols.number < 0) return null;

    const col = function (name) {
      return cols[name] !== undefined && cols[name] !== null && cols[name] >= 0 ? cols[name] : -1;
    };
    const cNumber = col('number'), cQty = col('qty'), cLevel = col('level'),
          cPos = col('pos'), cTitle = col('title'), cDesc = col('description'),
          cFile = col('file'), cStructure = col('structure'), cMaterial = col('material'),
          cRevision = col('revision'), cThumbnail = col('thumbnail');

    const items = [];
    const rawIndents = [];
    for (let r = headerRow + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const rawNumberCell = row[cNumber];
      const number = cellText(rawNumberCell);
      if (!number) continue;
      // skip repeated header rows (multi-page PDFs)
      if (matchField(number) === 'number') continue;

      let level = null;
      if (cLevel >= 0) {
        const lv = parseQty(row[cLevel]);
        if (lv !== null) level = Math.round(lv);
      }
      if (level === null && cPos >= 0) {
        const pos = cellText(row[cPos]);
        if (/^\d+(\.\d+)*$/.test(pos)) level = pos.split('.').length;
      }
      let indent = null;
      if (opts.indents && opts.indents[r] !== undefined && opts.indents[r] !== null) {
        indent = opts.indents[r];
      } else {
        const rawStr = rawNumberCell === null || rawNumberCell === undefined ? '' : String(rawNumberCell);
        const m = rawStr.match(/^[ \t]*/);
        indent = m ? m[0].replace(/\t/g, '    ').length : 0;
      }
      rawIndents.push(indent);

      // 1-based PDF page this row came from (pdf-extract.js's pageOf), null
      // for every non-PDF source — lets the UI point back at the exact page
      // to verify a flagged part on, instead of just a reconstructed row #.
      const page = opts.pageOf && opts.pageOf[r] !== undefined && opts.pageOf[r] !== null ? opts.pageOf[r] : null;

      const bomStructure = cStructure >= 0 ? cellText(row[cStructure]) : '';
      items.push({
        seq: items.length,
        number: number,
        title: cTitle >= 0 ? cellText(row[cTitle]) : '',
        description: cDesc >= 0 ? cellText(row[cDesc]) : '',
        qty: cQty >= 0 ? parseQty(row[cQty]) : null,
        level: level,
        isAssembly: null, // resolved below
        file: cFile >= 0 ? cellText(row[cFile]) : '',
        material: cMaterial >= 0 ? cellText(row[cMaterial]) : '',
        revision: cRevision >= 0 ? cellText(row[cRevision]) : '',
        bomStructure: bomStructure,
        isReference: /reference/i.test(bomStructure),
        // "(NULL)" means Inventor has no CAD file behind this row; an empty
        // cell means it does. Only the explicit "(NULL)" counts, so a export
        // without the column never looks like every row is virtual.
        thumbnailMissing: cThumbnail >= 0 && /^\(NULL\)$/i.test(cellText(row[cThumbnail])),
        sourceRow: r + 1,
        page: page,
      });
    }
    if (!items.length) return null;

    // fall back to indentation-based levels when no explicit level/pos data
    let hasLevels = items.some(function (it) { return it.level !== null; });
    if (!hasLevels) {
      const lv = levelsFromIndents(rawIndents);
      if (lv) {
        items.forEach(function (it, i) { it.level = lv[i]; });
        hasLevels = true;
        warnings.push('Hierarchy inferred from row indentation.');
      }
    }
    // A blank Level/Position cell mid-file (an ordinary data gap) is carried
    // forward from the row above rather than left null. compare.js's
    // consumers (cadTotals, groupMissingLeveled, cadChildSets) all fall a
    // null level back to 1 (top of tree) as a last-resort defensive default
    // — since hasLevels=true here does not guarantee every row parsed one,
    // that default would silently re-anchor everything after the gap to the
    // root, corrupting grouping/roll-ups for the rest of that branch. A
    // blank cell almost always means "same level as the row above," not
    // "back to the top." This only fires on individual gaps in an
    // otherwise-leveled file — the indentation fallback above already
    // handles the separate "no level data anywhere" case.
    if (hasLevels) {
      let lastLevel = null, filled = 0;
      for (const it of items) {
        if (it.level === null) {
          if (lastLevel !== null) { it.level = lastLevel; filled++; }
        } else {
          lastLevel = it.level;
        }
      }
      if (filled) warnings.push(filled + ' row(s) had a blank Level/Position — assumed the same level as the row above.');
    }
    // normalize levels to start at 1
    if (hasLevels) {
      let min = Infinity;
      for (const it of items) if (it.level !== null && it.level < min) min = it.level;
      if (min !== Infinity && min !== 1) for (const it of items) if (it.level !== null) it.level += 1 - min;
    }

    // isAssembly: file extension when available, otherwise "has children"
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (/\.iam$/i.test(it.file)) it.isAssembly = true;
      else if (/\.ipt$/i.test(it.file)) it.isAssembly = false;
      else if (hasLevels) {
        const next = items[i + 1];
        it.isAssembly = !!(next && next.level !== null && it.level !== null && next.level > it.level);
      }
    }

    const hasQty = cQty >= 0 && items.some(function (it) { return it.qty !== null; });
    if (!hasQty) warnings.push('No usable quantity column — quantity comparison unavailable for this file.');
    if (!hasLevels) warnings.push('No level/position information found — reference-assembly grouping will be inferred from the Item Master hierarchy.');
    const hasMaterial = cMaterial >= 0 && items.some(function (it) { return it.material !== ''; });
    const hasRevision = cRevision >= 0 && items.some(function (it) { return it.revision !== ''; });

    return {
      kind: 'cad',
      source: opts.source || 'leveled-sheet',
      hasQty: hasQty,
      hasLevels: hasLevels,
      hasStructure: cStructure >= 0,
      hasMaterial: hasMaterial,
      hasRevision: hasRevision,
      // Header presence alone, unlike hasMaterial/hasRevision which also
      // require a non-empty value: a clean export with no virtual components
      // has an entirely empty Thumbnail column, and that must read as
      // "checked, none found" rather than "column missing".
      hasThumbnail: cThumbnail >= 0,
      items: items,
      columns: cols,
      headerRow: headerRow,
      warnings: warnings,
    };
  }

  return { cadLeveledParser: { parse: parse, analyze: analyze, detectHeader: detectHeader, matchField: matchField } };
});
