/*
 * detect.js — decides which parser handles an uploaded file and validates
 * that the file matches the dropzone it was dropped on.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(
      require('./itemmaster.js').itemMasterParser,
      require('./cad-flat-xlsx.js').cadFlatParser,
      require('./cad-leveled.js').cadLeveledParser,
      require('./lldbo.js').lldboParser
    );
  } else {
    const bc = root.BOMCompare || {};
    root.BOMCompare = Object.assign(bc, factory(bc.itemMasterParser, bc.cadFlatParser, bc.cadLeveledParser, bc.lldboParser));
  }
})(typeof self !== 'undefined' ? self : this, function (itemMasterParser, cadFlatParser, cadLeveledParser, lldboParser) {
  'use strict';

  // Headers that only appear in the Item Master (Vault item/BOM grid) export.
  const IM_SIGNATURE = ['row order', 'row type', 'item qty', 'vault status', 'category name',
    'file link state', 'component type item', 'on/off bom row', 'source item'];

  // Headers distinctive of the LLDBO list, to redirect a misdropped file
  // rather than silently misparsing it as a generic leveled CAD table
  // ("PART NO"/"Qty." alone would match the CAD detector's keywords too).
  const LLDBO_SIGNATURE = ['sr. no', 'sr no', 'specifications', 'remarks', 'item description'];

  function cellText(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function looksLikeItemMaster(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 15); r++) {
      const lower = (aoa[r] || []).map(function (c) { return cellText(c).toLowerCase(); });
      if (lower.indexOf('number') === -1) continue;
      let hits = 0;
      for (const sig of IM_SIGNATURE) if (lower.indexOf(sig) !== -1) hits++;
      if (hits >= 2) return true;
    }
    return false;
  }

  function looksLikeLldbo(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 20); r++) {
      const lower = (aoa[r] || []).map(function (c) { return cellText(c).toLowerCase(); });
      if (lower.indexOf('part no') === -1 && lower.indexOf('part no.') === -1) continue;
      let hits = 0;
      for (const sig of LLDBO_SIGNATURE) if (lower.indexOf(sig) !== -1) hits++;
      if (hits >= 1) return true;
    }
    return false;
  }

  function parseItemMasterFromWorkbook(workbook, XLSX) {
    return itemMasterParser.parse(workbook, XLSX);
  }

  function parseLldboFromWorkbook(workbook, XLSX) {
    return lldboParser.parse(workbook, XLSX);
  }

  // Try the flat Vault paste first, then a leveled table. Returns:
  //   { ok: result } | { needsMapping: {analysis, aoa, sheetName} } | null
  function parseCadFromWorkbook(workbook, XLSX) {
    const flat = cadFlatParser.parse(workbook, XLSX);
    if (flat) return { ok: flat };

    for (const sheetName of workbook.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: null });
      if (looksLikeLldbo(aoa)) continue; // "PART NO"/"Qty." would otherwise false-match the CAD keyword table
      const leveled = cadLeveledParser.parse(aoa, { source: 'leveled-sheet' });
      if (leveled) {
        if (looksLikeItemMaster(aoa)) leveled.imShaped = true; // app suggests swapping zones
        return { ok: leveled };
      }
    }

    // nothing auto-detected: offer manual mapping over the densest sheet
    let best = null;
    for (const sheetName of workbook.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: null });
      const filled = aoa.reduce(function (n, r) {
        return n + (r || []).filter(function (c) { return cellText(c) !== ''; }).length;
      }, 0);
      if (!best || filled > best.filled) best = { sheetName: sheetName, aoa: aoa, filled: filled };
    }
    if (best && best.filled > 0) {
      return { needsMapping: { analysis: cadLeveledParser.analyze(best.aoa), aoa: best.aoa, sheetName: best.sheetName } };
    }
    return null;
  }

  return {
    detect: {
      looksLikeItemMaster: looksLikeItemMaster,
      looksLikeLldbo: looksLikeLldbo,
      parseItemMasterFromWorkbook: parseItemMasterFromWorkbook,
      parseLldboFromWorkbook: parseLldboFromWorkbook,
      parseCadFromWorkbook: parseCadFromWorkbook,
    },
  };
});
