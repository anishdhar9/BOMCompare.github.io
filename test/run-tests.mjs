/*
 * Node smoke tests for the parsers + comparison logic.
 *
 * Usage:
 *   node test/run-tests.mjs [CAD_Bom.xlsx Item_Master_BOM.xls [Vault_723.pdf Vault_732.pdf Inventor_732.xlsx]]
 *
 * The real sample exports are NOT committed (BOM data may be sensitive).
 * Without arguments only the synthetic tests run; with the sample files the
 * full baseline assertions run as well. The PDF tests additionally need
 * `npm install` (pdfjs-dist, pinned to the vendored pdf.js version).
 */
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';

const require = createRequire(import.meta.url);
const rootDir = path.dirname(path.dirname(fileURLToPath(import.meta.url)));

const XLSX = require(path.join(rootDir, 'vendor/xlsx.full.min.js'));
const { compare, compareAll, countDescendants, indexItemMaster, normNumber, groupImOnly } = require(path.join(rootDir, 'js/compare.js'));
const { itemMasterParser } = require(path.join(rootDir, 'js/parsers/itemmaster.js'));
const { cadFlatParser } = require(path.join(rootDir, 'js/parsers/cad-flat-xlsx.js'));
const { cadLeveledParser } = require(path.join(rootDir, 'js/parsers/cad-leveled.js'));
const { detect } = require(path.join(rootDir, 'js/parsers/detect.js'));
const { imQc } = require(path.join(rootDir, 'js/imqc.js'));
const { materialCompare } = require(path.join(rootDir, 'js/material-compare.js'));
const { revisionCompare } = require(path.join(rootDir, 'js/revision-compare.js'));
const { findings } = require(path.join(rootDir, 'js/findings.js'));
const { folder } = require(path.join(rootDir, 'js/folder.js'));
const { lldboParser } = require(path.join(rootDir, 'js/parsers/lldbo.js'));
const { lldboCompare } = require(path.join(rootDir, 'js/lldbo-compare.js'));
const { imDiffCompare } = require(path.join(rootDir, 'js/im-diff-compare.js'));

let failures = 0;
function check(name, cond, extra) {
  if (cond) console.log('  ok  ' + name);
  else { failures++; console.error('FAIL  ' + name + (extra !== undefined ? ' — got: ' + JSON.stringify(extra) : '')); }
}

/* ---------------- synthetic leveled-CAD tests ---------------- */

console.log('\n== synthetic: leveled CAD parsing, exact grouping, qty roll-up ==');
{
  // Machine (1) -> REF-ASSY (2, missing w/ children) + PART-A (2) + PART-B (2, qty differs)
  const aoa = [
    ['Item', 'Number', 'Title', 'QTY', 'File'],
    ['1', 'MACH-01', 'Machine', '1', 'mach.iam'],
    ['1.1', 'REF-ASSY', 'Reference assembly', '1', 'ref.iam'],
    ['1.1.1', 'CHILD-1', 'Child part 1', '2', 'c1.ipt'],
    ['1.1.2', 'CHILD-2', 'Child part 2', '4', 'c2.ipt'],
    ['1.2', 'PART-A', 'Part A', '3', 'pa.ipt'],
    ['1.3', 'PART-B', 'Part B', '5', 'pb.ipt'],
    ['1.4', 'PART-C', 'Part C standalone missing', '1', 'pc.ipt'],
  ];
  const cad = cadLeveledParser.parse(aoa, { source: 'leveled-sheet' });
  check('leveled parse found items', cad && cad.items.length === 7, cad && cad.items.length);
  check('leveled parse hasQty', cad.hasQty === true);
  check('leveled parse hasLevels', cad.hasLevels === true);
  check('assembly detection via file ext', cad.items[1].isAssembly === true && cad.items[2].isAssembly === false);

  const im = {
    rows: [
      { number: 'MACH-01', title: 'Machine', qty: null, path: [] },
      { number: 'PART-A', title: 'Part A', qty: 3, path: ['1'] },
      { number: 'PART-B', title: 'Part B', qty: 2, path: ['2'] }, // CAD says 5
    ],
  };
  const res = compare(cad, im);
  check('missing roots = REF-ASSY + PART-C', res.missingRoots.length === 2, res.missingRoots.map(n => n.item.number));
  const ref = res.missingRoots.find(n => n.item.number === 'REF-ASSY');
  check('REF-ASSY children grouped', ref && countDescendants(ref) === 2, ref && countDescendants(ref));
  const pc = res.missingRoots.find(n => n.item.number === 'PART-C');
  check('PART-C standalone', !!pc && pc.children.length === 0);
  check('qty mismatch found for PART-B only', res.qtyMismatches.length === 1 && res.qtyMismatches[0].number === 'PART-B',
    res.qtyMismatches.map(m => m.number));
  check('qty mismatch values', res.qtyMismatches[0].cadQty === 5 && res.qtyMismatches[0].imQty === 2, res.qtyMismatches[0]);
  check('missingTotal counts unique PNs', res.missingTotal === 4, res.missingTotal);
  check('qty mismatch cadBreakdown carries the CAD row #', res.qtyMismatches[0].cadBreakdown[0].sourceRow === 7,
    res.qtyMismatches[0].cadBreakdown);
  check('qty mismatch imBreakdown carries the Item Master row # and Row Order (undefined here -> "" is still a valid, present field)',
    'sourceRow' in res.qtyMismatches[0].imBreakdown[0] && 'rowOrder' in res.qtyMismatches[0].imBreakdown[0],
    res.qtyMismatches[0].imBreakdown[0]);
}

console.log('\n== synthetic: Qty mismatch breakdown carries Item Master Row # ==');
{
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Item Qty'],
    ['MACH-03', '-', 'Machine 3', 'desc', '-'],
    ['ASSY-1', '1', 'Sub-assembly', 'desc', '1'],
    ['PART-X', '1.1', 'Part X', 'desc', '2'], // CAD will say 5
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => aoa },
  });
  const partXRow = im.rows.find(r => r.number === 'PART-X');
  check('PART-X sourceRow is its real row position in the aoa (row 4, 1-based)', partXRow.sourceRow === 4, partXRow.sourceRow);

  const cadAoa = [
    ['Item', 'Number', 'Title', 'QTY'],
    ['1', 'MACH-03', 'Machine 3', '1'],
    ['1.1', 'ASSY-1', 'Sub-assembly', '1'],
    ['1.1.1', 'PART-X', 'Part X', '5'],
  ];
  const cad = cadLeveledParser.parse(cadAoa, { source: 'leveled-sheet' });
  const res = compare(cad, im);
  const mismatch = res.qtyMismatches.find(m => m.number === 'PART-X');
  check('PART-X flagged as a qty mismatch (5 vs 2)', !!mismatch && mismatch.cadQty === 5 && mismatch.imQty === 2, mismatch);
  check('PART-X imBreakdown carries the real Item Master Row # (4) and Row Order ("1.1")',
    mismatch.imBreakdown.length === 1 && mismatch.imBreakdown[0].sourceRow === 4 && mismatch.imBreakdown[0].rowOrder === '1.1',
    mismatch.imBreakdown);
}

console.log('\n== synthetic: Item Master diff (diffItemMasters) ==');
{
  const header = ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity', 'Revision', 'Material', 'State'];
  const imOldAoa = [
    header,
    ['MACH-01', '-', 'Machine', 'desc', '-', '0', '', 'Certified'],
    ['PART-REMOVED', '1', 'Removed Part', 'desc', '2 Each', '0', 'AISI 304', 'Certified'],
    ['PART-QTY', '2', 'Qty Part', 'desc', '2 Each', '0', 'AISI 304', 'Certified'],
    ['PART-REV', '3', 'Rev Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified'],
    ['PART-SAME', '4', 'Same Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified'],
    ['PART-MAT', '5', 'Mat Part', 'desc', '1 Each', '0', '1.4301', 'Certified'],
  ];
  const imNewAoa = [
    header,
    ['MACH-01', '-', 'Machine', 'desc', '-', '0', '', 'Certified'],
    ['PART-ADDED', '1', 'Added Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified'],
    ['PART-QTY', '2', 'Qty Part', 'desc', '4 Each', '0', 'AISI 304', 'Certified'],
    ['PART-REV', '3', 'Rev Part', 'desc', '1 Each', '1', 'AISI 304', 'Certified'],
    ['PART-SAME', '4', 'Same Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified'],
    ['PART-MAT', '5', 'Mat Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified'],
  ];
  const parse = (aoa) => itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, { utils: { sheet_to_json: () => aoa } });
  const imOld = parse(imOldAoa);
  const imNew = parse(imNewAoa);

  const diffRaw = imDiffCompare.diffItemMasters(imOld, imNew, indexItemMaster);
  check('one part added', diffRaw.added.length === 1 && diffRaw.added[0].number === 'PART-ADDED', diffRaw.added);
  check('one part removed', diffRaw.removed.length === 1 && diffRaw.removed[0].number === 'PART-REMOVED', diffRaw.removed);
  check('without materialsMatch: 3 changed (Qty, Rev, and raw-text Material difference)',
    diffRaw.changed.length === 3, diffRaw.changed.map(c => c.number));
  check('PART-SAME never appears as changed', !diffRaw.changed.some(c => c.number === 'PART-SAME'));

  const qtyChange = diffRaw.changed.find(c => c.number === 'PART-QTY');
  check('PART-QTY field diff names Quantity, old 2 -> new 4',
    qtyChange && qtyChange.fields.length === 1 && qtyChange.fields[0].field === 'Quantity' &&
    qtyChange.fields[0].old === '2' && qtyChange.fields[0].new === '4', qtyChange);

  const revChange = diffRaw.changed.find(c => c.number === 'PART-REV');
  check('PART-REV field diff names Revision, old 0 -> new 1',
    revChange && revChange.fields.length === 1 && revChange.fields[0].field === 'Revision' &&
    revChange.fields[0].old === '0' && revChange.fields[0].new === '1', revChange);

  const matChangeRaw = diffRaw.changed.find(c => c.number === 'PART-MAT');
  check('without materialsMatch, PART-MAT (1.4301 -> AISI 304, same grade) IS flagged as changed (raw text differs)',
    !!matChangeRaw, diffRaw.changed.map(c => c.number));

  const diffNormalized = imDiffCompare.diffItemMasters(imOld, imNew, indexItemMaster, materialCompare.materialsMatch);
  check('with materialsMatch injected, PART-MAT (same grade, different spelling) is NOT flagged',
    !diffNormalized.changed.some(c => c.number === 'PART-MAT'), diffNormalized.changed.map(c => c.number));
  check('with materialsMatch injected, only 2 changed remain (Qty, Rev)',
    diffNormalized.changed.length === 2, diffNormalized.changed.map(c => c.number));

  check('unique counts carried through', diffRaw.oldUniqueCount === 6 && diffRaw.newUniqueCount === 6,
    { old: diffRaw.oldUniqueCount, new: diffRaw.newUniqueCount });
  check('unchangedCount excludes added/removed/changed (MACH-01 + PART-SAME = 2)',
    diffRaw.unchangedCount === 2, diffRaw.unchangedCount);

  // Regression: materialsMatch treats a blank as never matching anything (by
  // design, for the main CAD-vs-IM check). A literal "-" placeholder material
  // strips down to "" under its normalization, so naively calling
  // materialsMatch(oldMat, newMat) on two IDENTICAL "-" values would
  // misreport them as "changed" — raw equality must short-circuit first.
  const placeholderAoa = (mat) => [
    header,
    ['MACH-01', '-', 'Machine', 'desc', '-', '0', '', 'Certified'],
    ['PART-DASH', '1', 'Dash Part', 'desc', '1 Each', '0', mat, 'Certified'],
  ];
  const imDashOld = parse(placeholderAoa('-'));
  const imDashNew = parse(placeholderAoa('-'));
  const dashDiff = imDiffCompare.diffItemMasters(imDashOld, imDashNew, indexItemMaster, materialCompare.materialsMatch);
  check('identical "-" placeholder material on both sides is NOT flagged as changed',
    !dashDiff.changed.some(c => c.number === 'PART-DASH'), dashDiff.changed);
}

console.log('\n== synthetic: quantity-cascade detection (detectQuantityCascades) ==');
{
  // ASSY-P: all 5 direct children released at a clean 2x in the Item Master
  // -> one cascade finding. ASSY-Q: a single unrelated mismatched child
  // (below the 2-child minimum) -> stays an ordinary qty mismatch, not a
  // cascade.
  const cadAoa = [
    ['Item', 'Number', 'Title', 'QTY'],
    ['1', 'MACH-01', 'Machine', '1'],
    ['1.1', 'ASSY-P', 'Assembly P', '1'],
    ['1.1.1', 'CHILD-1', 'Child 1', '1'],
    ['1.1.2', 'CHILD-2', 'Child 2', '1'],
    ['1.1.3', 'CHILD-3', 'Child 3', '1'],
    ['1.1.4', 'CHILD-4', 'Child 4', '1'],
    ['1.1.5', 'CHILD-5', 'Child 5', '1'],
    ['1.2', 'ASSY-Q', 'Assembly Q', '1'],
    ['1.2.1', 'CHILD-X', 'Child X', '1'],
  ];
  const cad = cadLeveledParser.parse(cadAoa, { source: 'leveled-sheet' });

  const imAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity'],
    ['MACH-01', '-', 'Machine', 'desc', '-'],
    ['ASSY-P', '1', 'Assembly P', 'desc', '1 Each'],
    ['CHILD-1', '1.1', 'Child 1', 'desc', '2 Each'],
    ['CHILD-2', '1.2', 'Child 2', 'desc', '2 Each'],
    ['CHILD-3', '1.3', 'Child 3', 'desc', '2 Each'],
    ['CHILD-4', '1.4', 'Child 4', 'desc', '2 Each'],
    ['CHILD-5', '1.5', 'Child 5', 'desc', '2 Each'],
    ['ASSY-Q', '2', 'Assembly Q', 'desc', '1 Each'],
    ['CHILD-X', '2.1', 'Child X', 'desc', '3 Each'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => imAoa },
  });

  const res = compare(cad, im);
  check('qtyCascades applicable', res.qtyCascades.applicable === true);
  check('exactly one cascade root (ASSY-P)', res.qtyCascades.roots.length === 1,
    res.qtyCascades.roots.map(r => r.item.number));
  const root = res.qtyCascades.roots[0];
  check('cascade root is ASSY-P', root.item.number === 'ASSY-P', root.item.number);
  check('cascade ratio is 2', root.item.cascadeRatio === 2, root.item.cascadeRatio);
  check('cascade counts 5 of 5 children mismatched', root.item.cascadeChildCount === 5 &&
    root.item.cascadeMismatchedChildCount === 5, root.item);
  check('cascade subtree has all 5 children, no deeper nesting', countDescendants(root) === 5,
    countDescendants(root));
  check('CHILD-X (single unrelated mismatch) is not part of any cascade', !res.qtyCascades.roots.some(r =>
    (function contains(n) { return n.item.number === 'CHILD-X' || n.children.some(contains); })(r)));
  check('CHILD-X still reported as an ordinary qty mismatch (1 vs 3)',
    res.qtyMismatches.some(m => m.number === 'CHILD-X' && m.cadQty === 1 && m.imQty === 3),
    res.qtyMismatches.map(m => m.number));
  check('CHILD-1..5 also still reported as ordinary rolled-up qty mismatches (1 vs 2)',
    ['CHILD-1', 'CHILD-2', 'CHILD-3', 'CHILD-4', 'CHILD-5'].every(n =>
      res.qtyMismatches.some(m => m.number === n && m.cadQty === 1 && m.imQty === 2)));
}

console.log('\n== synthetic: leveled CAD parsing captures Material column ==');
{
  const aoaWithMaterial = [
    ['Item', 'Part Number', 'BOM Structure', 'QTY', 'Description', 'Material'],
    ['1', 'PART-A', 'Normal', '1', 'desc', 'AISI 304'],
    ['1.1', 'PART-B', 'Purchased', '2', 'desc', ''],
  ];
  const cadWithMat = cadLeveledParser.parse(aoaWithMaterial, { source: 'leveled-sheet' });
  check('material column detected', cadWithMat.hasMaterial === true, cadWithMat.hasMaterial);
  check('material value captured per item', cadWithMat.items[0].material === 'AISI 304', cadWithMat.items[0].material);
  check('blank material stays empty string, not null', cadWithMat.items[1].material === '', cadWithMat.items[1].material);

  const aoaNoMaterial = [
    ['Item', 'Part Number', 'QTY'],
    ['1', 'PART-A', '1'],
  ];
  const cadNoMat = cadLeveledParser.parse(aoaNoMaterial, { source: 'leveled-sheet' });
  check('hasMaterial false when no Material column exists', cadNoMat.hasMaterial === false, cadNoMat.hasMaterial);
  check('werkstoff (German) recognized as a material header',
    cadLeveledParser.parse([['Number', 'Werkstoff'], ['PART-A', 'AISI 316']], { source: 'leveled-sheet' }).hasMaterial === true);
}

console.log('\n== synthetic: leveled CAD parsing captures Revision column ==');
{
  const aoaWithRev = [
    ['Item', 'Part Number', 'QTY', 'Description', 'Revision'],
    ['1', 'PART-A', '1', 'desc', '2'],
    ['1.1', 'PART-B', '2', 'desc', ''],
  ];
  const cadWithRev = cadLeveledParser.parse(aoaWithRev, { source: 'leveled-sheet' });
  check('revision column detected', cadWithRev.hasRevision === true, cadWithRev.hasRevision);
  check('revision value captured per item', cadWithRev.items[0].revision === '2', cadWithRev.items[0].revision);
  check('blank revision stays empty string, not null', cadWithRev.items[1].revision === '', cadWithRev.items[1].revision);

  const cadNoRev = cadLeveledParser.parse([['Item', 'Part Number', 'QTY'], ['1', 'PART-A', '1']], { source: 'leveled-sheet' });
  check('hasRevision false when no Revision column exists', cadNoRev.hasRevision === false, cadNoRev.hasRevision);
  check('"Rev" recognized as a revision header',
    cadLeveledParser.parse([['Number', 'Rev'], ['PART-A', 'B']], { source: 'leveled-sheet' }).hasRevision === true);

  // this is the same grid shape pdf-extract.js hands to cad-leveled.js after
  // reconstructing the Vault "Uses" PDF table (see js/parsers/pdf-extract.js's
  // HEADER_KEYS) -- confirms a PDF-sourced Revision column flows through.
  const pdfLikeGrid = [
    ['File', 'Revision', 'State', 'Title', 'Description', 'Part Number'],
    ['mach.iam', '1', 'Released', 'Machine', 'desc', 'MACH-01'],
  ];
  const fromPdfGrid = cadLeveledParser.parse(pdfLikeGrid, { source: 'pdf' });
  check('Revision column from a PDF-shaped grid is captured', fromPdfGrid.hasRevision === true && fromPdfGrid.items[0].revision === '1',
    fromPdfGrid.items[0]);
}

console.log('\n== synthetic: flat Vault Excel paste (cad-flat-xlsx.js) captures Revision + Material ==');
{
  const aoa = [
    ['Part X title', 'Part X desc', true, 'Released', 'partx.ipt', '2', 'Steel', 'PART-X'],
    ['Part Y title', 'Part Y desc', true, 'Released', 'party.ipt', '', 'Aluminum', 'PART-Y'],
    ['Assy Z title', 'Assy Z desc', true, 'Released', 'assyz.iam', '0', '', 'ASSY-Z'],
  ];
  const res = cadFlatParser.parse({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } }, { utils: { sheet_to_json: () => aoa } });
  check('flat-xlsx parsed', !!res, res);
  check('flat-xlsx hasRevision true', res.hasRevision === true, res.hasRevision);
  check('flat-xlsx revision captured per item', res.items[0].revision === '2' && res.items[0].material === 'Steel', res.items[0]);
  check('flat-xlsx blank revision stays empty string', res.items[1].revision === '', res.items[1].revision);
}

console.log('\n== synthetic: IM quantity roll-up through ancestors ==');
{
  const im = {
    rows: [
      { number: 'ROOT', title: '', qty: null, path: [] },
      { number: 'ASSY', title: '', qty: 2, path: ['1'] },
      { number: 'SCREW', title: '', qty: 4, path: ['1', '1'] },   // 4 x 2 = 8
      { number: 'SCREW', title: '', qty: 3, path: ['2'] },        // + 3 top-level
    ],
  };
  const idx = indexItemMaster(im);
  check('rolled-up SCREW total = 11', idx.totals.get('SCREW') === 11, idx.totals.get('SCREW'));
  check('ASSY child set contains SCREW', idx.childSets.get('ASSY') && idx.childSets.get('ASSY').has('SCREW'));
}

console.log('\n== synthetic: indentation-based levels ==');
{
  const aoa = [
    ['Number', 'Title', 'Qty'],
    ['A-1', 'top', '1'],
    ['  B-1', 'child', '2'],
    ['  B-2', 'child2', '1'],
    ['A-2', 'top2', '1'],
  ];
  const cad = cadLeveledParser.parse(aoa, { source: 'leveled-sheet' });
  check('indent levels inferred', cad.hasLevels === true);
  check('indent level values', cad.items.map(i => i.level).join(',') === '1,2,2,1', cad.items.map(i => i.level));
}


console.log('\n== synthetic: Vault PDF table reconstruction ==');
{
  const { pdfExtract } = require(path.join(rootDir, 'js/parsers/pdf-extract.js'));
  global.BOMCompare = { cadLeveledParser };
  const mk = (str, x, y, w = 30, h = 8) => ({ str, transform: [1, 0, 0, h, x, y], width: w });
  const items = [
    mk('Name', 33, 700), mk('Revision', 292, 700), mk('State', 355, 700), mk('Title', 438, 700), mk('Description', 642, 700), mk('Part', 825, 700),
    mk('Number', 825, 690),
    mk('7-230-20509.iam', 64, 650, 120), mk('1', 292, 650), mk('Released', 371, 650), mk('MAIN GRANULATOR_HSG PRO', 438, 650, 165), mk('SPN016823_PN22426_SUN', 642, 650, 170), mk('7-230-', 825, 650, 45),
    mk('PILOT', 438, 640, 40), mk('PHARMACEUTICALS LTD,', 642, 640, 150), mk('20509', 825, 640, 40),
    mk('7-099-200063.iam', 78, 610, 120), mk('0', 292, 610), mk('Released', 371, 610), mk('REDUCER 4” TO 2”', 438, 610, 120), mk('WITH CLAMP ASSEMBLY', 642, 610, 150), mk('7-099-', 825, 610, 45),
    mk('200063', 825, 600, 45),
    mk('7-999-00044I00.ipt', 91, 570, 130), mk('0', 292, 570), mk('Released', 371, 570), mk('TC CLAMP - 4”', 438, 570, 100), mk('BS 4825 - 211057 - CLAMP', 642, 570, 170), mk('7-999-', 825, 570, 45),
    mk('00044', 825, 560, 40),
  ];
  const pdfjsLib = { getDocument: () => ({ promise: Promise.resolve({ numPages: 1, getPage: async () => ({ getTextContent: async () => ({ items }) }) }) }) };
  const grid = await pdfExtract.extractGrid(new ArrayBuffer(0), { pdfjsLib });
  check('PDF grid recognizes wrapped Part Number header', grid.rows[0][5] === 'Part Number', grid.rows[0]);
  check('PDF grid merges wrapped part numbers', grid.rows[1][5] === '7-230-20509' && grid.rows[2][5] === '7-099-200063', grid.rows.slice(1));
  const cad = cadLeveledParser.parse(grid.rows, { indents: grid.indents, source: 'pdf' });
  check('PDF CAD parses rows', cad && cad.items.length === 3, cad && cad.items.map(i => i.number));
  check('PDF Name column becomes file column', cad.items[0].file === '7-230-20509.iam' && cad.items[0].isAssembly === true, cad.items[0]);
  check('PDF indentation infers levels', cad.hasLevels === true && cad.items.map(i => i.level).join(',') === '1,2,3', cad && cad.items.map(i => i.level));
}

console.log('\n== synthetic: Item Master column-name robustness (itemmaster.js) ==');
{
  // Different organizations' (or the same organization's different
  // plants'/users') Vault exports don't all spell headers the same way, or
  // put them in the same order. This block proves both: header synonyms are
  // recognized, and column ORDER doesn't matter (it never did -- this parser
  // is header-name-based -- but there was no dedicated test proving it).
  const synonymAoa = [
    ['Row Order', 'Part Number', 'Description (Item,CO)', 'Title (Item,CO)', 'Qty'],
    ['-', 'MACH-01', 'desc', 'Machine', '-'],
    ['1', 'PART-A', 'desc', 'Part A', '3'],
  ];
  const imSyn = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => synonymAoa } });
  check('"Part Number" recognized as the Number column', !!imSyn && imSyn.rows[1].number === 'PART-A', imSyn && imSyn.rows);
  check('"Qty" recognized as the Item Qty column (numeric slot)', imSyn && imSyn.rows[1].itemQty === 3, imSyn && imSyn.rows[1]);
  check('columns in a shuffled, non-canonical order still parse (Row Order first, Number second)',
    imSyn && imSyn.hasPaths === true && imSyn.rows[1].path.join('.') === '1', imSyn && imSyn.rows[1]);

  // Critical negative test: "PN" must NOT be treated as a Number synonym --
  // in this organization's convention "PN" means Producer Number (part of
  // the project's SPN/PN key), never a part number.
  const pnAoa = [
    ['PN', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)'],
    ['MACH-01', '-', 'Machine', 'desc'],
  ];
  const imPn = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => pnAoa } });
  check('a bare "PN" header does NOT get picked up as the Number column (no header row found at all here)', imPn === null, imPn);

  // "PN" as its own column DOES mean Producer Number when Number is also present.
  const producerNumberAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Producer', 'PN'],
    ['MACH-01', '-', 'Machine', 'PN22426_SPN016823_ACME', 'SPN016823', '22426'],
  ];
  const imProdNum = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => producerNumberAoa } });
  check('"PN" recognized as the Producer Number column (not Number)',
    imProdNum && imProdNum.rows[0].number === 'MACH-01' && imProdNum.rows[0].producerNumber === '22426', imProdNum && imProdNum.rows[0]);
  check('projectKey still resolves correctly with PN-as-header', imProdNum && imProdNum.projectKey &&
    imProdNum.projectKey.spn === 'SPN016823' && imProdNum.projectKey.pn === 'PN22426', imProdNum && imProdNum.projectKey);

  // "Level"/"Position" as Row Order synonyms, "Item Number" as a Number synonym.
  const levelAoa = [
    ['Item Number', 'Level', 'Title (Item,CO)', 'Description (Item,CO)'],
    ['MACH-01', '-', 'Machine', 'desc'],
    ['PART-A', '1', 'Part A', 'desc'],
  ];
  const imLevel = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => levelAoa } });
  check('"Item Number" recognized as the Number column and "Level" as Row Order',
    imLevel && imLevel.rows[1].number === 'PART-A' && imLevel.hasPaths === true, imLevel && imLevel.rows);

  // Revision column + hasRevision flag.
  const revAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Revision'],
    ['MACH-01', '-', 'Machine', 'desc', '1'],
  ];
  const imRev = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => revAoa } });
  check('Item Master Revision column captured', imRev && imRev.hasRevision === true && imRev.rows[0].revision === '1', imRev && imRev.rows[0]);
  const imNoRev = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => revAoa.map(r => r.slice(0, 4)) } });
  check('hasRevision false when no Revision column exists', imNoRev && imNoRev.hasRevision === false, imNoRev && imNoRev.hasRevision);

  // Some exports carry THREE quantity-ish columns: "Item Quantity",
  // "Quantity", and "Quantity Per Unit". The resolved `qty` used for
  // comparison must come from "Quantity" (as-released qty, e.g. "1 Each"),
  // never "Item Quantity" and never "Quantity Per Unit".
  const threeQtyAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Item Quantity', 'Quantity', 'Quantity Per Unit'],
    ['MACH-01', '-', 'Machine', 'desc', '1', '1 Each', '1 Each'],
    ['PART-A', '1', 'Part A', 'desc', '9', '4 Each', '2 Each'],
  ];
  const imThreeQty = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => threeQtyAoa } });
  check('"Item Quantity" recognized as the Item Qty column', imThreeQty && imThreeQty.rows[1].itemQty === 9, imThreeQty && imThreeQty.rows[1]);
  check('"Quantity" (not "Quantity Per Unit") recognized as the Quantity column', imThreeQty && imThreeQty.rows[1].quantity === 4, imThreeQty && imThreeQty.rows[1]);
  check('resolved qty prefers "Quantity" over "Item Quantity"', imThreeQty && imThreeQty.rows[1].qty === 4, imThreeQty && imThreeQty.rows[1]);

  // "Quantity" falls back to "Item Qty" only when "Quantity" itself is absent.
  const onlyItemQtyAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Item Qty'],
    ['PART-A', '1', 'Part A', 'desc', '7'],
  ];
  const imOnlyItemQty = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => onlyItemQtyAoa } });
  check('resolved qty falls back to Item Qty when no Quantity column exists', imOnlyItemQty && imOnlyItemQty.rows[0].qty === 7, imOnlyItemQty && imOnlyItemQty.rows[0]);

  // Column order shouldn't matter: "Quantity Per Unit" listed BEFORE
  // "Quantity" must still not be mistaken for it.
  const reorderedAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity Per Unit', 'Quantity'],
    ['PART-A', '1', 'Part A', 'desc', '2 Each', '5 Each'],
  ];
  const imReordered = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => reorderedAoa } });
  check('"Quantity Per Unit" before "Quantity" in column order still resolves to "Quantity"',
    imReordered && imReordered.rows[0].qty === 5, imReordered && imReordered.rows[0]);

  // Real-world regression (the two 726020775 EBOM exports): "QTY per Unit"
  // must NOT be captured as the Item Qty column. Before the longest-prefix
  // fix, the 'qty' prefix grabbed "QTY per Unit" into the Item Qty slot,
  // producing a storm of false Check-3 mismatches on every multi-qty row.
  const qtyPerUnitAoa = [
    ['Number', 'Revision', 'Row Order', 'Position Number', 'Quantity', 'Title (Item,CO)', 'Description (Item,CO)', 'Material', 'QTY per Unit'],
    ['MACH-01', '1', '-', '', '1 Each', 'Machine', 'desc', '', '1'],
    ['PART-A', '1', '1', '1', '4 Each', 'Part A', 'desc', 'AISI 304', '1'],
  ];
  const imQpu = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => qtyPerUnitAoa } });
  check('"QTY per Unit" is NOT captured as the Item Qty column', imQpu && imQpu.hasItemQty === false && imQpu.columns.qty === -1, imQpu && imQpu.columns);
  check('with only "Quantity" present, hasQuantity true / hasItemQty false', imQpu && imQpu.hasQuantity === true && imQpu.hasItemQty === false, imQpu && { q: imQpu.hasQuantity, iq: imQpu.hasItemQty });
  check('"QTY per Unit" export resolves qty from "Quantity" (4), not per-unit (1)', imQpu && imQpu.rows[1].qty === 4, imQpu && imQpu.rows[1]);
  // Check 3 must be NOT applicable here (no genuine Item Qty column to
  // cross-check against), never a mass-flag of every row.
  const qpuQc = imQc.runChecks(imQpu);
  check('Check 3 not-applicable when there is a Quantity but no Item Qty column (no false mismatch storm)',
    qpuQc.c3.applicable === false && /Item Qty/.test(qpuQc.c3.reason), qpuQc.c3);

  // "Unit Qty" is likewise per-unit and must not displace a genuine "Item
  // Qty" that sits in the same export (as the 726020768 EBOM does).
  const unitQtyAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity', 'Unit Qty', 'Item Qty'],
    ['PART-A', '1', 'Part A', 'desc', '4 Each', '1 Each', '4'],
  ];
  const imUnitQty = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => unitQtyAoa } });
  check('"Unit Qty" does not displace a genuine "Item Qty" column', imUnitQty && imUnitQty.rows[0].itemQty === 4 && imUnitQty.hasItemQty === true, imUnitQty && imUnitQty.rows[0]);
  check('with a genuine Item Qty present, Check 3 is applicable', imQc.runChecks(imUnitQty).c3.applicable === true, imQc.runChecks(imUnitQty).c3);

  // "State" sits in wildly different positions across real exports (column 1
  // in one 726020775 EBOM, column 9 in the other) and shares a prefix with
  // two decoy columns — the match must be positional-agnostic and exact.
  const stateEarly = [
    ['Number', 'State', 'Quantity', 'File Link State', 'Row Order', 'Title (Item,CO)'],
    ['PART-A', 'Obsolete', '1 Each', 'Out of Date', '1', 'A'],
  ];
  const stateLate = [
    ['Number', 'Revision', 'Row Order', 'Quantity', 'Title (Item,CO)', 'Material', 'State', 'State (Historical)'],
    ['PART-A', '1', '1', '1 Each', 'A', 'AISI 304', 'Obsolete', 'Certified'],
  ];
  for (const [label, aoa] of [['State early (col 1)', stateEarly], ['State late (col 6)', stateLate]]) {
    const parsed = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => aoa } });
    check(label + ': State resolved to the item state, decoys ignored',
      parsed && parsed.hasState === true && parsed.rows[0].state === 'Obsolete', parsed && parsed.rows[0]);
  }
}

/* ---------------- synthetic: dual-source reference detection ---------------- */

console.log('\n== synthetic: reference items (structure vs intended BOM) ==');
{
  // structure = full CAD incl. reference; bom = intended BOM with qty
  const structureAoa = [
    ['Item', 'Number', 'Title', 'File'],
    ['1', 'MACH-01', 'Machine', 'mach.iam'],
    ['1.1', 'REF-1', 'Reference part', 'r1.ipt'],
    ['1.2', 'REF-ASSY', 'Reference assembly', 'ra.iam'],
    ['1.2.1', 'REF-CHILD', 'Its child', 'rc.ipt'],
    ['1.3', 'PART-A', 'Part A', 'pa.ipt'],
  ];
  const bomAoa = [
    ['Item', 'Number', 'Title', 'QTY', 'BOM Structure'],
    ['1', 'PART-A', 'Part A', '3', 'Normal'],
    ['2', 'VIRT-1', 'Virtual part (no CAD file)', '1', 'Normal'],
  ];
  const structure = cadLeveledParser.parse(structureAoa, { source: 'pdf' });
  const bom = cadLeveledParser.parse(bomAoa, { source: 'leveled-sheet' });
  check('bom export captures BOM Structure', bom.hasStructure === true && bom.items[0].bomStructure === 'Normal');
  const im = {
    rows: [
      { number: 'MACH-01', title: '', qty: null, path: [] },
      { number: 'PART-A', title: '', qty: 3, path: ['1'] },
      { number: 'REF-1', title: '', qty: 1, path: ['2'] }, // reference that DID reach the IM
    ],
  };
  const res = compareAll([structure, bom], im);
  check('reference roots = REF-1 + REF-ASSY (root machine excluded)', res.referenceRoots.length === 2,
    res.referenceRoots.map(n => n.item.number));
  const refAssy = res.referenceRoots.find(n => n.item.number === 'REF-ASSY');
  check('REF-ASSY groups its child', refAssy && countDescendants(refAssy) === 1);
  const ref1 = res.referenceRoots.find(n => n.item.number === 'REF-1');
  check('reference annotated with IM presence', ref1 && ref1.inItemMaster === true && refAssy.inItemMaster === false);
  check('referenceTotal = 3 unique PNs', res.referenceTotal === 3, res.referenceTotal);
  check('virtual part missing from IM appended standalone', res.missingRoots.some(n => n.item.number === 'VIRT-1'),
    res.missingRoots.map(n => n.item.number));
  check('qty taken from bom source', res.hasQty === true);
}

console.log('\n== synthetic: Item Master QC checks ==');
{
  // header includes Producer / Producer Number / Entity Icon so all 4 checks
  // are applicable; one deliberate failure planted per check.
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity', 'Item Qty', 'Producer', 'Producer Number', 'Entity Icon'],
    ['MACH-01', '-', 'Machine', 'SPN000111_PN00222_ACME CORP', '-', '-', 'SPN000111', '00222', 'Normal'],
    ['7-909-00001', '1', 'END OF LINE', 'END OF LINE', '1 Each', '1', '', '', 'Normal'],
    ['PART-A', '1.1', 'Part A', 'desc', '2 Each', '2', '', '', 'Normal'],
    ['PART-B', '1.2', 'Part B (qty edited, Item Qty stale)', 'desc', '5 Each', '3', '', '', 'Normal'],
    ['PART-C', '1.3', 'Part C (bad icon)', 'desc', '1 Each', '1', '', '', 'Reference'],
    ['7-909-00002', '1.4', 'not really end of line but matches text END OF LINE', 'desc', '1 Each', '1', '', '', 'Normal'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => aoa },
  });
  check('IM parsed with QC columns', !!im && im.hasProducer === true && im.hasEntityIcon === true);
  check('projectKey from root Producer/Producer Number', im.projectKey && im.projectKey.spn === 'SPN000111' && im.projectKey.pn === 'PN00222', im.projectKey);

  const qc = imQc.runChecks(im);
  check('c1 producer match passes (SPN000111 + 00222 both in description)', qc.c1.applicable === true && qc.c1.fail.length === 0, qc.c1);

  // Reported false-positive-looking case: Producer Number IS present in the
  // Description (as a substring of a longer token, "PN22759"), but Producer
  // itself is not -- the fail entry must say exactly which field is missing,
  // not just "flagged", so this isn't mistaken for the Producer Number being
  // the (absent) one.
  const aoaPartialMatch = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Producer', 'Producer Number'],
    ['MACH-02', '-', 'Machine 2', 'PN22759_SPN017160_WALTER BUSHNELL LIFE SCIENCES PVT LTD, UK, India', 'GLATT', '22759'],
  ];
  const imPartial = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => aoaPartialMatch },
  });
  const qcPartial = imQc.runChecks(imPartial);
  check('c1: Producer Number found in Description (substring of "PN22759") but Producer "GLATT" is not -> flagged',
    qcPartial.c1.fail.length === 1, qcPartial.c1.fail);
  check('c1: fail entry\'s issue names Producer specifically, not Producer Number',
    qcPartial.c1.fail.length === 1 &&
    qcPartial.c1.fail[0].issue.indexOf('Producer "GLATT" not found') !== -1 &&
    qcPartial.c1.fail[0].issue.indexOf('Producer Number') === -1,
    qcPartial.c1.fail[0] && qcPartial.c1.fail[0].issue);
  check('c2 flags the second "END OF LINE"-text row with wrong number', qc.c2.found === 2 && qc.c2.fail.length === 1 && qc.c2.fail[0].number === '7-909-00002', qc.c2);
  check('c3 flags PART-B only', qc.c3.applicable === true && qc.c3.fail.length === 1 && qc.c3.fail[0].number === 'PART-B', qc.c3.fail);
  check('c4 flags PART-C only', qc.c4.applicable === true && qc.c4.fail.length === 1 && qc.c4.fail[0].number === 'PART-C', qc.c4.fail);

  // no Producer/Entity Icon columns at all -> both checks report not-applicable, not mass-fail
  const bareAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Item Qty'],
    ['MACH-02', '-', 'Machine 2', 'SPN000333_PN00444, some customer', '-'],
    ['PART-X', '1', 'Part X', 'desc', '1'],
  ];
  const bareIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => bareAoa },
  });
  check('bare export: hasProducer/hasEntityIcon false', bareIm.hasProducer === false && bareIm.hasEntityIcon === false);
  check('projectKey falls back to description regex', bareIm.projectKey && bareIm.projectKey.spn === 'SPN000333' && bareIm.projectKey.pn === 'PN00444', bareIm.projectKey);
  const bareQc = imQc.runChecks(bareIm);
  check('c1 not-applicable without Producer column (not mass-fail)', bareQc.c1.applicable === false, bareQc.c1);
  check('c4 not-applicable without Entity Icon column (not mass-fail)', bareQc.c4.applicable === false, bareQc.c4);
}

console.log('\n== synthetic: Title/Description completeness (c5) + Material completeness (c6) ==');
{
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Material'],
    ['MACH-01', '-', 'Machine', 'desc', ''],
    ['7-999-PURCH-1', '1', 'Purchased part 1', '', 'AISI 304'],       // purchased, desc blank -> leniency, no flag
    ['7-999-PURCH-2', '2', '', '', 'AISI 304'],                       // purchased, both blank -> flagged
    ['MFG-PART-1', '3', 'Manufactured part 1', '', 'AISI 304'],       // non-purchased, desc blank -> flagged
    ['MFG-PART-2', '4', '', 'Manufactured part 2 desc', 'AISI 304'],  // non-purchased, title blank -> flagged
    ['ASSY-1', '5', 'Assembly 1', 'desc', ''],                        // has a child below -> assembly, material excluded
    ['ASSY-1-CHILD', '5.1', 'Assy child', 'desc', ''],                // leaf, material blank -> flagged
    ['LEAF-2', '6', 'Leaf 2', 'desc', 'AISI 316L'],                   // leaf, material present -> not flagged
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => aoa },
  });
  check('IM parsed with Material column', im.hasMaterial === true);
  const qc = imQc.runChecks(im);

  check('c5 flags exactly 3 rows', qc.c5.fail.length === 3, qc.c5.fail.map(f => f.number));
  check('c5: purchased part with one blank field is NOT flagged (leniency)',
    !qc.c5.fail.some(f => f.number === '7-999-PURCH-1'));
  check('c5: purchased part with both blank IS flagged as both-missing',
    qc.c5.fail.some(f => f.number === '7-999-PURCH-2' && f.kind === 'both-missing'));
  check('c5: non-purchased part missing description flagged correctly',
    qc.c5.fail.some(f => f.number === 'MFG-PART-1' && f.kind === 'description-missing'));
  check('c5: non-purchased part missing title flagged correctly',
    qc.c5.fail.some(f => f.number === 'MFG-PART-2' && f.kind === 'title-missing'));

  check('c6 flags exactly 1 row', qc.c6.applicable === true && qc.c6.fail.length === 1, qc.c6.fail);
  check('c6: assembly with children is excluded despite blank material',
    !qc.c6.fail.some(f => f.number === 'ASSY-1'));
  check('c6: leaf part with blank material IS flagged',
    qc.c6.fail.some(f => f.number === 'ASSY-1-CHILD'));
  check('c6: root row excluded (always counts as assembly)',
    !qc.c6.fail.some(f => f.number === 'MACH-01'));

  // no Material column at all -> not-applicable, not mass-fail
  const noMatAoa = aoa.map(r => r.slice(0, 4)); // drop the Material column
  const noMatIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => noMatAoa },
  });
  check('no Material column -> hasMaterial false', noMatIm.hasMaterial === false);
  const noMatQc = imQc.runChecks(noMatIm);
  check('c6 not-applicable without Material column (not mass-fail)', noMatQc.c6.applicable === false, noMatQc.c6);
}

console.log('\n== synthetic: Revision Consistency (c7) + Revision: CAD vs Item Master ==');
{
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Revision'],
    ['MACH-01', '-', 'Machine', 'desc', '2'],
    ['ASSY-A', '1', 'Assy A', 'desc', '1'],
    ['PART-X', '1.1', 'Part X', 'desc', '0'],
    ['ASSY-B', '2', 'Assy B', 'desc', '1'],
    ['PART-X', '2.1', 'Part X', 'desc', '1'],   // same PN as above, conflicting revision
    ['PART-Y', '1.2', 'Part Y', 'desc', '3'],
    ['PART-Y', '2.2', 'Part Y', 'desc', '3'],   // same PN, same revision -> not flagged
    ['PART-Z', '1.3', 'Part Z', 'desc', ''],    // blank revision, no other occurrence -> not flagged
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, { utils: { sheet_to_json: () => aoa } });
  check('IM parsed with Revision column', im.hasRevision === true);
  const qc = imQc.runChecks(im);

  check('c7 flags exactly 2 rows (both PART-X occurrences)', qc.c7.applicable === true && qc.c7.fail.length === 2,
    qc.c7.fail.map(f => f.rowOrder));
  check('c7: both PART-X occurrences flagged, PART-Y (same revision) is not',
    qc.c7.fail.every(f => f.number === 'PART-X') && !qc.c7.fail.some(f => f.number === 'PART-Y'), qc.c7.fail);
  check('c7: PART-Z (single occurrence, blank revision) is not flagged',
    !qc.c7.fail.some(f => f.number === 'PART-Z'));
  check('c7 fail entries carry Row #, parent, and a conflictsWith summary',
    qc.c7.fail.every(f => f.sourceRow > 1 && f.parentNumber && f.conflictsWith), qc.c7.fail);

  const noRevAoa = aoa.map(r => r.slice(0, 4));
  const noRevIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, { utils: { sheet_to_json: () => noRevAoa } });
  check('c7 not-applicable without Revision column (not mass-fail)', imQc.runChecks(noRevIm).c7.applicable === false);

  // CAD vs Item Master revision comparison (js/revision-compare.js)
  const cadSource = {
    kind: 'cad', hasRevision: true,
    fileName: 'inventor.xlsx',
    items: [
      { number: 'ASSY-A', revision: '1' },   // matches IM
      { number: 'ASSY-B', revision: '1' },   // matches IM
      { number: 'PART-Y', revision: '9' },   // genuinely differs from IM's '3'
    ],
  };
  const rev = revisionCompare.compareRevision([cadSource], im);
  check('revision check applicable with a hasRevision CAD source', rev.applicable === true, rev);
  check('revision mismatch found for PART-Y only (ASSY-A and ASSY-B agree with CAD)',
    rev.mismatches.length === 1 && rev.mismatches[0].number === 'PART-Y', rev.mismatches);
  check('revision mismatch entry carries Row #, parent, and both revision values',
    rev.mismatches[0].sourceRow > 1 && rev.mismatches[0].parentNumber &&
    rev.mismatches[0].imRevision === '3' && rev.mismatches[0].cadRevision === '9', rev.mismatches[0]);

  check('revision check not-applicable without a hasRevision CAD source',
    revisionCompare.compareRevision([{ kind: 'cad', hasRevision: false, items: [] }], im).applicable === false);
  check('revision check not-applicable without an Item Master Revision column',
    revisionCompare.compareRevision([cadSource], noRevIm).applicable === false);

  check('revisionsMatch: exact values match, case/whitespace-insensitive', revisionCompare.revisionsMatch(' b ', 'B'));
  check('revisionsMatch: different values do not match', !revisionCompare.revisionsMatch('A', 'B'));
  check('revisionsMatch: blank never matches', !revisionCompare.revisionsMatch('', 'A') && !revisionCompare.revisionsMatch('A', ''));
}

console.log('\n== synthetic: sketch parts in Item Master (c8) + item state (c9) ==');
{
  // 7-333-* are rough-sketch models. They may exist in CAD, but only ever as
  // Reference -- one reaching the released Item Master is a serious release
  // error, so each hit must be traceable to exactly where it got in.
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'State', 'Quantity', 'Item Qty'],
    ['MACH-01', '-', 'Main Machine', 'desc', 'Certified', '1 Each', '1'],
    ['ASSY-A', '1', 'Top Assembly', 'desc', 'Certified', '1 Each', '1'],
    ['SUB-B', '1.2', 'Sub Assembly', 'desc', 'Certified', '1 Each', '1'],
    ['7-333-29220', '1.2.3', 'ROUGH SKETCH BRACKET', 'sketch', 'Obsolete', '2 Each', '2'],
    ['7-333-10074', '2', 'SKETCH PLATE', 'sketch', 'Certified', '1 Each', '1'],
    ['PART-OK', '1.2.4', 'Fine part', 'desc', 'Certified', '1 Each', '1'],
    ['PART-INV', '1.5', 'Invalid part', 'desc', 'Invalid', '1 Each', '1'],
    ['PART-PO', '1.6', 'Phased part', 'desc', 'Phased Out', '1 Each', '1'],
    ['PART-NEW', '1.7', 'New part', 'desc', 'New', '1 Each', '1'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => aoa } });
  check('Item Master State column captured', im.hasState === true && im.rows[3].state === 'Obsolete', im.rows[3]);
  const qc = imQc.runChecks(im);

  check('c8 flags both 7-333-* parts and nothing else',
    qc.c8.applicable === true && qc.c8.fail.length === 2 &&
    qc.c8.fail.every(f => /^7-333-/.test(f.number)), qc.c8.fail.map(f => f.number));
  check('c8 records the Row # of each sketch part', qc.c8.fail[0].sourceRow === 5 && qc.c8.fail[1].sourceRow === 6,
    qc.c8.fail.map(f => f.sourceRow));
  // The whole point: not just the immediate parent, the full chain in.
  check('c8 maps the FULL parent trail, not just the immediate parent',
    qc.c8.fail[0].trail === 'ASSY-A (Top Assembly) › SUB-B (Sub Assembly)', qc.c8.fail[0].trail);
  check('c8 still reports a top-level sketch part, with no trail',
    qc.c8.fail[1].rowOrder === '2' && qc.c8.fail[1].trail === '(top level)', qc.c8.fail[1]);
  check('c8 is always applicable (a clean BOM simply has no hits)',
    imQc.runChecks(itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } },
      { utils: { sheet_to_json: () => aoa.filter(r => !/^7-333-/.test(String(r[0]))) } })).c8.fail.length === 0);

  check('c9 counts Obsolete/Invalid/Phased Out as errors and New as a warning',
    qc.c9.applicable === true && qc.c9.errorCount === 3 && qc.c9.warnCount === 1,
    { err: qc.c9.errorCount, warn: qc.c9.warnCount });
  check('c9 leaves Certified rows alone', !qc.c9.fail.some(f => f.state === 'Certified'), qc.c9.fail.map(f => f.state));
  check('c9 sorts hard errors above warnings', /^ERROR/.test(qc.c9.fail[0].severity) &&
    /^warning/.test(qc.c9.fail[qc.c9.fail.length - 1].severity), qc.c9.fail.map(f => f.severity));
  check('c9 not-applicable when the export has no State column',
    imQc.runChecks(itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } },
      { utils: { sheet_to_json: () => aoa.map(r => r.slice(0, 4)) } })).c9.applicable === false);

  check('classifyState: Certified/Released pass', imQc.classifyState('Certified') === 'ok' && imQc.classifyState('Released') === 'ok');
  check('classifyState: Obsolete/Invalid/Phased Out are errors, case- and hyphen-insensitive',
    ['Obsolete', 'invalid', 'Phased Out', 'phased-out'].every(s => imQc.classifyState(s) === 'error'));
  check('classifyState: New is a warning, not an error', imQc.classifyState('New') === 'warn');
  check('classifyState: blank is neither', imQc.classifyState('') === 'blank');

  // "File Link State" (Current/Out of Date) and "State (Historical)" describe
  // the CAD file link, not the item -- neither must be read as the item state.
  const decoyAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'File Link State', 'State (Historical)'],
    ['PART-A', '1', 'A', 'Out of Date', 'Obsolete'],
  ];
  const decoyIm = itemMasterParser.parse({ SheetNames: ['S'], Sheets: { S: {} } }, { utils: { sheet_to_json: () => decoyAoa } });
  check('"File Link State" / "State (Historical)" are NOT read as the item State',
    decoyIm.hasState === false && imQc.runChecks(decoyIm).c9.applicable === false, decoyIm.columns);

  // Registry ranking: a sketch part outranks everything; "New" ranks low so a
  // legitimately-uncertified row can't bury a genuinely missing part.
  const reg = findings.buildRegistry({ imQc: qc });
  const sketch = reg.byPn.get('7-333-29220');
  check('a sketch part is owned by c8, above its own obsolete state',
    sketch.primary.key === 'c8' && sketch.issues.some(i => i.key === 'c9'), sketch.issues.map(i => i.key));
  // Both sketch parts tie at the top severity, so they sort by number between
  // themselves — what matters is that they outrank every other finding.
  check('sketch parts are the worst findings in the whole registry',
    reg.parts[0].primary.key === 'c8' && reg.parts[1].primary.key === 'c8' &&
    reg.parts.slice(2).every(p => p.primary.key !== 'c8'), reg.parts.slice(0, 3).map(p => p.number + ':' + p.primary.key));
  check('an obsolete part outranks a merely-new one',
    reg.byPn.get('PART-INV').primary.severity > reg.byPn.get('PART-NEW').primary.severity);
  check('"New" is ranked under its own low-severity key, not as an error',
    reg.byPn.get('PART-NEW').primary.key === 'c9warn', reg.byPn.get('PART-NEW').primary);
}

console.log('\n== synthetic: findings registry (one primary finding per part) ==');
{
  // A part flagged by several checks should be reported once, owned by the
  // most serious check, with the rest demoted to cross-references.
  const reg = findings.buildRegistry({
    result: {
      missingRoots: [],
      referenceRoots: null,
      qtyMismatches: [{ number: 'PART-Q', title: 'Q', description: '', cadQty: 7, imQty: 1, cadBreakdown: [], imBreakdown: [] }],
      imOnly: [{ number: 'PART-IO', title: 'IO', description: '', sourceRow: 5, parentNumber: 'ASSY-1', parentTitle: 'A1' }],
    },
    imQc: {
      c3: { applicable: true, fail: [{ number: 'PART-Q', rowOrder: '1.1', title: 'Q', quantity: '7 Each', itemQty: '1', sourceRow: 9 },
                                     { number: 'PART-IO', rowOrder: '1.2', title: 'IO', quantity: '2 Each', itemQty: '1', sourceRow: 5 }] },
      c5: { applicable: true, fail: [{ number: 'PART-IO', rowOrder: '1.2', title: '', description: '', kind: 'title-missing', sourceRow: 5 }] },
      c2: { applicable: true, fail: [{ number: '—', rowOrder: '—', issue: 'No "END OF LINE" entry found in the BOM.' }] },
    },
  });

  const q = reg.byPn.get('PART-Q');
  check('part in both qty-mismatch and c3 is registered once', reg.parts.filter(p => p.number === 'PART-Q').length === 1, reg.parts.map(p => p.number));
  check('the more serious check (qty, sev 80) owns it over c3 (sev 45)', q.primary.key === 'qty', q.primary);
  check('the losing check is still recorded, as a secondary issue', q.issues.length === 2 && q.issues[1].key === 'c3' && q.issues[1].primary === false, q.issues);
  check('isSecondary() true for the demoted check, false for the primary',
    reg.isSecondary('c3', 'PART-Q') === true && reg.isSecondary('qty', 'PART-Q') === false);
  check('qty issue carries a human detail string', /7/.test(q.primary.detail) && /1/.test(q.primary.detail), q.primary.detail);

  const io = reg.byPn.get('PART-IO');
  check('a part flagged by three checks still has one primary, the most serious',
    io.issues.length === 3 && io.primary.key === 'imOnly', io.issues.map(i => i.key));
  check('isSecondary() true for both losing checks on that part',
    reg.isSecondary('c3', 'PART-IO') && reg.isSecondary('c5', 'PART-IO'));

  // Check 2's "no END OF LINE row at all" entry is a whole-BOM assertion, not
  // a part — it must never become a row in the parts registry.
  check('Check 2 synthetic "—" row is not registered as a part',
    !reg.byPn.has('—') && !reg.parts.some(p => p.number === '—'), reg.parts.map(p => p.number));

  check('parts are ordered worst-first', reg.parts[0].number === 'PART-Q', reg.parts.map(p => p.number));
  check('counts: 2 unique parts, 3 secondary issues', reg.counts.parts === 2 && reg.counts.secondaryTotal === 3, reg.counts);
  check('unknown part numbers are simply not secondary', reg.isSecondary('c3', 'NOPE') === false);

  // Grouped children: a part explained by a flagged parent is not its own finding.
  const treeReg = findings.buildRegistry({
    result: {
      missingRoots: [{ item: { number: 'ASSY-1', title: 'Assembly' }, children: [{ item: { number: 'CHILD-1', title: 'Child' }, children: [] }] }],
      referenceRoots: null, qtyMismatches: [], imOnly: [],
    },
  });
  check('a missing assembly is actionable, its child is grouped',
    treeReg.counts.actionable === 1 && treeReg.counts.grouped === 1, treeReg.counts);
  check('the grouped child records which parent explains it',
    treeReg.byPn.get('CHILD-1').grouped === true && treeReg.byPn.get('CHILD-1').groupedUnder === 'ASSY-1', treeReg.byPn.get('CHILD-1'));
  check('the root itself is not marked grouped', treeReg.byPn.get('ASSY-1').grouped === false);

  // A quantity cascade should outrank the ordinary per-part qty finding it
  // explains: the cascade root becomes the one actionable finding, and its
  // descendants (including one that ALSO independently shows up in the flat
  // qtyMismatches list) are demoted/grouped under it instead of competing.
  const cascadeReg = findings.buildRegistry({
    result: {
      missingRoots: [], referenceRoots: null, imOnly: [],
      qtyMismatches: [{ number: 'CHILD-A', title: 'A', description: '', cadQty: 1, imQty: 2, cadBreakdown: [], imBreakdown: [] }],
      qtyCascades: {
        applicable: true,
        roots: [{
          item: { number: 'ASSY-CASCADE', title: 'Cascade Assy', cascadeRatio: 2, cascadeChildCount: 3, cascadeMismatchedChildCount: 3 },
          children: [
            { item: { number: 'CHILD-A', title: 'A' }, children: [] },
            { item: { number: 'CHILD-B', title: 'B' }, children: [] },
            { item: { number: 'CHILD-C', title: 'C' }, children: [] },
          ],
        }],
      },
    },
  });
  const cascadeRootPart = cascadeReg.byPn.get('ASSY-CASCADE');
  check('the cascade root is its own actionable finding, not grouped',
    !!cascadeRootPart && cascadeRootPart.primary.key === 'qtyCascade' && cascadeRootPart.grouped === false, cascadeRootPart);
  check('the cascade root detail names the ratio and child counts', /3 of 3/.test(cascadeRootPart.primary.detail) &&
    /2×/.test(cascadeRootPart.primary.detail), cascadeRootPart.primary.detail);
  check('CHILD-A is owned by the cascade (sev 85), not the plain qty finding (sev 80)',
    cascadeReg.byPn.get('CHILD-A').primary.key === 'qtyCascade', cascadeReg.byPn.get('CHILD-A'));
  check('CHILD-A\'s own qty finding is still recorded, just demoted',
    cascadeReg.isSecondary('qty', 'CHILD-A') === true);
  check('CHILD-B/CHILD-C (grouped, no independent qty finding) are still registered and grouped',
    cascadeReg.byPn.get('CHILD-B').grouped === true && cascadeReg.byPn.get('CHILD-C').grouped === true);
}

console.log('\n== synthetic: "In Item Master only" parent rollup (groupImOnly) ==');
{
  // Reuses the same shape as the missing-item tree so renderTree/countDescendants work.
  const rows = [
    { number: 'ASSY-1', title: 'Assembly', path: ['1'], sourceRow: 2 },
    { number: 'CHILD-A', title: 'Child A', path: ['1', '1'], sourceRow: 3 },
    { number: 'CHILD-B', title: 'Child B', path: ['1', '2'], sourceRow: 4 },
    { number: 'GRAND-C', title: 'Grandchild', path: ['1', '2', '1'], sourceRow: 5 },
    { number: 'LONE-1', title: 'Unrelated', path: ['2'], sourceRow: 6 },
  ];
  const byPath = new Map(rows.map(r => [r.path.join('.'), r]));
  const roots = groupImOnly(rows, byPath);
  check('a whole flagged subassembly collapses to one root', roots.length === 2, roots.map(r => r.item.number));
  check('every row is still present in the tree (nothing dropped)',
    roots.reduce((n, r) => n + 1 + countDescendants(r), 0) === rows.length,
    roots.map(r => r.item.number + ':' + countDescendants(r)));
  check('nesting follows the Row Order hierarchy, not just the direct parent',
    countDescendants(roots[0]) === 3 && roots[1].item.number === 'LONE-1', roots);

  // No Row Order column -> nothing to group by; must degrade, not crash.
  const flat = groupImOnly(rows.map(r => ({ number: r.number, title: r.title, path: null })), new Map());
  check('degrades to one root per part when the export has no Row Order', flat.length === rows.length, flat.length);

  // A parent that is NOT itself flagged must not swallow its children.
  const orphanRows = [{ number: 'CHILD-X', title: 'X', path: ['9', '1'], sourceRow: 7 }];
  const orphanByPath = new Map([['9', { number: 'PRESENT-PARENT', title: 'in CAD' }], ['9.1', orphanRows[0]]]);
  check('a child whose parent is not flagged stays top-level',
    groupImOnly(orphanRows, orphanByPath).length === 1);
}

console.log('\n== synthetic: folder auto-load file classification ==');
{
  const cases = [
    // [filename, expected classification]
    ['Autodesk Vault- 723020509.pdf', 'cad-pdf'],                 // this org's stated naming
    ['Autodesk_Vault__723020509.iam.pdf', 'cad-pdf'],              // Vault web client's own default naming (real sample seen)
    ['autodesk vault - 733020013.pdf', 'cad-pdf'],                 // case-insensitive, extra spacing
    ['EBOM_723020509.xlsx', 'item-master'],
    ['ebom-723020509.xls', 'item-master'],
    ['EBOM.xlsx', 'item-master'],
    ['HSG_Item_Master_BOM.xls', null],                             // old naming convention, not auto-matched
    ['PN22426_LLDBO.xlsx', 'lldbo'],
    ['PN22260_LLDBO.xlsx', 'lldbo'],                               // real sample naming
    ['PN22260_LLDBO_rev2.xlsx', 'lldbo'],                          // tolerant of a suffix
    ['LLDBO_PN22260.xlsx', null],                                  // wrong order, not this org's convention
    ['PN22260_LLDBO.docx', null],                                  // right prefix, wrong extension
    ['Autodesk Vault- 723020509.dwg', null],                       // right prefix, wrong extension
    ['readme.txt', null],
    ['', null],
    ['INVENTOR_BOM_726020768.xlsx', 'inventor-bom'],               // real sample naming
    ['Inventor BOM - 726020768.xls', 'inventor-bom'],              // case-insensitive, spacing variant
    ['inventor-bom-726020768.xlsx', 'inventor-bom'],
    ['INVENTOR_BOM_726020768.docx', null],                         // right prefix, wrong extension
  ];
  for (const [name, expected] of cases) {
    check('classifyFolderFile(' + JSON.stringify(name) + ') = ' + expected,
      folder.classifyFolderFile(name) === expected, folder.classifyFolderFile(name));
  }

  // scanFolder against a mock FileSystemDirectoryHandle-shaped object —
  // proves the traversal/bucketing logic works without a real browser
  // picker (only window.showDirectoryPicker() itself, in app.js, needs one).
  function mockDir(entries) {
    return {
      values: async function* () {
        for (const e of entries) yield e;
      },
    };
  }
  const mockEntries = [
    { kind: 'file', name: 'Autodesk Vault- 723020509.pdf', getFile: async () => ({ name: 'Autodesk Vault- 723020509.pdf' }) },
    { kind: 'file', name: 'EBOM_723020509.xlsx', getFile: async () => ({ name: 'EBOM_723020509.xlsx' }) },
    { kind: 'file', name: 'INVENTOR_BOM_723020509.xlsx', getFile: async () => ({ name: 'INVENTOR_BOM_723020509.xlsx' }) },
    { kind: 'file', name: 'notes.txt', getFile: async () => ({ name: 'notes.txt' }) },
    { kind: 'directory', name: 'subfolder' },
  ];
  const found = await folder.scanFolder(mockDir(mockEntries));
  check('scanFolder finds exactly 1 cad-pdf', found['cad-pdf'].length === 1 && found['cad-pdf'][0].name === 'Autodesk Vault- 723020509.pdf', found['cad-pdf']);
  check('scanFolder finds exactly 1 item-master', found['item-master'].length === 1 && found['item-master'][0].name === 'EBOM_723020509.xlsx', found['item-master']);
  check('scanFolder finds exactly 1 inventor-bom', found['inventor-bom'].length === 1 && found['inventor-bom'][0].name === 'INVENTOR_BOM_723020509.xlsx', found['inventor-bom']);
  check('scanFolder ignores directories and unmatched files', found['cad-pdf'].length + found['item-master'].length + found['inventor-bom'].length === 3);

  // ambiguous folder (two EBOM files) -> both bucketed, caller decides what to do
  const ambiguousEntries = mockEntries.concat([
    { kind: 'file', name: 'EBOM_old_version.xlsx', getFile: async () => ({}) },
  ]);
  const ambiguousFound = await folder.scanFolder(mockDir(ambiguousEntries));
  check('scanFolder reports ambiguous matches rather than picking one', ambiguousFound['item-master'].length === 2);
}

console.log('\n== synthetic: LLDBO parsing + comparison against Item Master ==');
{
  // mirrors the real sample's layout: merged-cell document header above a
  // "SR. No / PART NO / Item Description / Specifications / Make / Qty. / Remarks" table
  const lldboAoa = [
    ['', '', 'LONG LEAD DIRECT BOUGHT OUT (LLDBO) LIST', '', '', 'ISSUE DATE', '', ''],
    ['', '', '', '', '', 'DOCUMENT NO', '', ''],
    ['', '', 'CUSTOMER: ACME CORP', '', '', 'DATE', '', ''],
    ['Glatt Systems Pvt Ltd.', '', 'DBO Doc No : SPN000999_PN33445_TEST MACHINE', '', '', '', '', ''],
    ['SR. No', 'PART NO', 'Item Description', 'Specifications', 'Make', 'Qty.', 'Remarks'],
    [],
    ['', 'PART-A', 'Present, qty matches', 'spec', 'MAKE', '1 Nos.', ''],
    ['', 'PART-B', 'Missing from IM', 'spec', 'MAKE', '1 Nos.', ''],
    ['', 'PART-C', 'Used twice, same PN', 'spec', 'MAKE', '1 Nos.', ''],
    ['', 'PART-C', 'Used twice, same PN (2nd)', 'spec', 'MAKE', '1 Nos.', ''],
    ['', '', 'Not yet specified', 'Pending', '', 'NA', ''],
  ];
  const lldbo = lldboParser.parse({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } }, {
    utils: { sheet_to_json: () => lldboAoa },
  });
  check('LLDBO parsed', !!lldbo && lldbo.rows.length === 5, lldbo && lldbo.rows.length);
  check('LLDBO projectKey extracted from document header', lldbo.projectKey && lldbo.projectKey.spn === 'SPN000999' && lldbo.projectKey.pn === 'PN33445', lldbo.projectKey);
  check('LLDBO customer extracted', lldbo.customer === 'ACME CORP', lldbo.customer);
  check('LLDBO warns about the 1 no-part-number row', lldbo.warnings.some(w => w.indexOf('1 row') === 0), lldbo.warnings);

  // matching project: PART-A present+correct qty, PART-B missing, PART-C
  // summed to 2 across its two LLDBO rows but IM only has 1 -> mismatch
  const imAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Item Qty', 'Producer', 'Producer Number'],
    ['MACH-01', '-', 'Test Machine', 'SPN000999_PN33445_TEST MACHINE', '-', 'SPN000999', '33445'],
    ['PART-A', '1', 'Part A', 'desc', '1', '', ''],
    ['PART-C', '2', 'Part C', 'desc', '1', '', ''],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => imAoa },
  });
  const res = lldboCompare.compareLldbo(lldbo, im, indexItemMaster);
  check('LLDBO vs matching-project IM: no project key mismatch', res.projectKeyMismatch === null, res.projectKeyMismatch);
  check('LLDBO: 3 unique part numbers, 1 without a PN yet', res.totalLldboItems === 3 && res.noPartNumberCount === 1, res);
  check('LLDBO: PART-B correctly flagged missing from IM', res.missingFromIm.length === 1 && res.missingFromIm[0].number === 'PART-B', res.missingFromIm);
  check('LLDBO: PART-B missing-from-IM entry carries an LLDBO Row #', !!res.missingFromIm[0].sourceRow, res.missingFromIm[0]);
  check('LLDBO: PART-C flagged with summed qty 2 vs IM qty 1', res.qtyMismatches.length === 1 &&
    res.qtyMismatches[0].number === 'PART-C' && res.qtyMismatches[0].lldboQty === 2 && res.qtyMismatches[0].imQty === 1,
    res.qtyMismatches);
  check('LLDBO: PART-C qty-mismatch entry carries an LLDBO Row # and a Found Under (Item Master) string',
    !!res.qtyMismatches[0].sourceRow && typeof res.qtyMismatches[0].foundUnder === 'string',
    res.qtyMismatches[0]);
  check('LLDBO: PART-A (present, qty matches) not flagged anywhere',
    !res.missingFromIm.some(m => m.number === 'PART-A') && !res.qtyMismatches.some(m => m.number === 'PART-A'));

  // mismatched project: same LLDBO, IM for a different PN -> project key warning
  const otherImAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Item Qty', 'Producer', 'Producer Number'],
    ['MACH-02', '-', 'Other Machine', 'SPN000111_PN99999_OTHER MACHINE', '-', 'SPN000111', '99999'],
    ['PART-A', '1', 'Part A', 'desc', '1', '', ''],
  ];
  const otherIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => otherImAoa },
  });
  const crossRes = lldboCompare.compareLldbo(lldbo, otherIm, indexItemMaster);
  check('LLDBO vs wrong-project IM: project key mismatch flagged',
    crossRes.projectKeyMismatch && crossRes.projectKeyMismatch.lldbo.pn === 'PN33445' && crossRes.projectKeyMismatch.im.pn === 'PN99999',
    crossRes.projectKeyMismatch);

  // detect.js routing: LLDBO must not be swallowed by the generic CAD leveled-table detector
  const looksLikeLldbo = detect.looksLikeLldbo(lldboAoa);
  check('detect.looksLikeLldbo recognizes the real layout', looksLikeLldbo === true);
}

console.log('\n== synthetic: material normalization (materialsMatch) ==');
{
  const M = materialCompare.materialsMatch;
  // naming-convention variants that must NOT be flagged
  check('DIN vs AISI grade match: 1.4301 == AISI 304', M('1.4301', 'AISI 304') === true);
  check('spacing-only variant match: AISI 316L == AISI 316 L', M('AISI 316L', 'AISI 316 L') === true);
  check('abbreviation variant match: SS316L == AISI 316L', M('SS316L', 'AISI 316L') === true);
  check('bare grade number match: 316Ti == 1.4571', M('316Ti', '1.4571') === true);
  check('qualifier-detail containment match: Silikon == Silikon/weiß/60°Shore', M('Silikon', 'Silikon/weiß/60°Shore') === true);
  check('embedded-in-longer-string match: 1.4301 == Stainless Steel AISI 304', M('1.4301', 'Stainless Steel AISI 304') === true);
  check('language-spelling match: Silicon == Silikon', M('Silicon', 'Silikon') === true);
  check('language-spelling match: Borosilicate == Borosilikat', M('Borosilicate', 'Borosilikat') === true);
  // genuine differences that MUST still be flagged
  check('304 vs 304L stays flagged (not equated)', M('AISI 304', 'AISI 304L') === false);
  check('316 vs 316L stays flagged (not equated)', M('AISI 316', 'AISI 316 L') === false);
  check('304 vs 316 (different grade entirely) stays flagged', M('AISI 304', '1.4404') === false);
  check('genuinely different materials stay flagged', M('Aluminium', 'AISI 304') === false);
  check('blank never matches', M('', 'AISI 304') === false && M('AISI 304', '') === false);
  check('placeholder "." never matches', M('.', 'AISI 316 L') === false);
  // separator-only variants (real false positives reported against the app)
  check('slash vs space separator: Silikon/transparent == Silikon transparent', M('Silikon/transparent', 'Silikon transparent') === true);
  check('hyphen vs space separator: St-37 == St 37', M('St-37', 'St 37') === true);
  check('German eszett vs ascii: PTFE/weiß == PTFE weiss', M('PTFE/weiß', 'PTFE weiss') === true);
  check('eszett + qualifier-detail containment: PTFE/weiß == PTFE-weiß/Edelstahl', M('PTFE/weiß', 'PTFE-weiß/Edelstahl') === true);
  // ISO 3506 fastener grades A2/A4 as synonyms of AISI 304/316
  check('fastener grade match: A2 == AISI 304', M('A2', 'AISI 304') === true);
  check('fastener grade + property class match: A4-70 == AISI 316', M('A4-70', 'AISI 316') === true);
  check('fastener grade + property class match: A2-70 == 1.4301', M('A2-70', '1.4301') === true);
  // regressions: A2/A4 additions must not blur the still-distinct grades
  check('304 vs 316L stays flagged (not equated)', M('AISI 304', 'AISI 316 L') === false);
  check('316L vs 316 stays flagged (not equated)', M('AISI 316L', 'AISI 316') === false);
  check('unrelated materials with no separator overlap stay flagged', M('Brass, Soft Yellow', 'Gun Metal') === false);
}

console.log('\n== synthetic: material comparison (CAD vs Item Master) + bought-out parts ==');
{
  const imAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Material'],
    ['MACH-01', '-', 'Machine', 'desc', ''],
    ['ASSY-1', '1', 'An assembly', 'desc', ''],           // assembly (has children below) -> material not expected
    ['PART-A', '1.1', 'Matches (naming variant)', 'desc', '1.4301'],
    ['PART-B', '1.2', 'Genuine mismatch', 'desc', 'AISI 304'],
    ['7-999-00001', '1.3', 'Purchased, missing material', 'desc', ''],
    ['7-999-00002', '1.4', 'Purchased, mismatch vs CAD', 'desc', 'AISI 304'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => imAoa },
  });

  const cadSource = {
    kind: 'cad', source: 'flat-xlsx', hasQty: false, hasMaterial: true, items: [
      { number: 'PART-A', title: 'Part A', material: 'AISI 304', isAssembly: false },
      { number: 'PART-B', title: 'Part B', material: 'AISI 304L', isAssembly: false }, // genuine grade difference
      { number: '7-999-00002', title: 'Purchased', material: 'AISI 316', isAssembly: false },
    ],
  };

  const noCadRes = materialCompare.compareMaterial([], im);
  check('not applicable with no CAD source carrying material', noCadRes.applicable === false, noCadRes.reason);
  check('bought-out list still populated when not applicable', noCadRes.boughtOut.length === 2, noCadRes.boughtOut.length);

  const res = materialCompare.compareMaterial([cadSource], im);
  check('applicable with a flat-xlsx CAD source', res.applicable === true);
  check('PART-A naming variant not flagged, only PART-B (genuine mismatch)',
    res.mismatches.length === 1 && res.mismatches[0].number === 'PART-B', res.mismatches.map(m => m.number));
  check('purchased parts excluded from the mismatches list', !res.mismatches.some(m => /^\d-999-/.test(m.number)));

  check('bought-out panel lists both purchased parts', res.boughtOut.length === 2, res.boughtOut.map(b => b.number));
  const bo1 = res.boughtOut.find(b => b.number === '7-999-00001');
  check('bought-out: missing IM material flagged, no CAD data', bo1 && bo1.missingMaterial === true && bo1.cadMaterial === '', bo1);
  const bo2 = res.boughtOut.find(b => b.number === '7-999-00002');
  check('bought-out: CAD/IM mismatch flagged for purchased part', bo2 && bo2.mismatch === true && bo2.cadMaterial === 'AISI 316', bo2);
}

/* ---------------- real-sample baseline tests ---------------- */

const [cadPath, imPath, pdf723Path, pdf732Path, inv732Path, pdf733Path, im733Path, lldboPath, invBomPath, imBomMatPath, pdf726Path, invBom22819Path, imBom22819Path, imDiffOldPath, imDiffNewPath] = process.argv.slice(2);
let pdfjsLib = null;
try { pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js'); } catch (e) { /* npm install to enable PDF tests */ }

async function parsePdf(file) {
  const { pdfExtract } = require(path.join(rootDir, 'js/parsers/pdf-extract.js'));
  const buf = new Uint8Array(fs.readFileSync(file)).buffer;
  const grid = await pdfExtract.extractGrid(buf, { pdfjsLib });
  const parsed = cadLeveledParser.parse(grid.rows, { indents: grid.indents, source: 'pdf' });
  return { grid, parsed };
}

if (cadPath && imPath) {
  console.log('\n== real samples: parsing ==');
  const cadWb = XLSX.read(fs.readFileSync(cadPath), { type: 'buffer' });
  const imWb = XLSX.read(fs.readFileSync(imPath), { type: 'buffer' });

  const im = detect.parseItemMasterFromWorkbook(imWb, XLSX);
  check('IM parsed', !!im);
  check('IM rows = 1431', im.rows.length === 1431, im.rows.length);
  const imUnique = new Set(im.rows.map(r => r.number.toUpperCase()));
  check('IM unique PNs = 1076', imUnique.size === 1076, imUnique.size);
  check('IM has paths', im.hasPaths === true);
  check('IM projectKey = SPN016823 / PN22426', im.projectKey && im.projectKey.spn === 'SPN016823' && im.projectKey.pn === 'PN22426', im.projectKey);
  // Real-world confirmation: this Item Master export has no Revision column
  // at all (verified across all three real Item Master samples this
  // project has) -- the revision checks must degrade gracefully, not crash.
  check('HSG Item Master has no Revision column', im.hasRevision === false, im.hasRevision);

  console.log('\n== real samples: HSG Item Master QC ==');
  const hsgQc = imQc.runChecks(im);
  check('HSG c1 producer match: no failures', hsgQc.c1.applicable === true && hsgQc.c1.fail.length === 0, hsgQc.c1.fail);
  check('HSG c2 end of line: found and clean', hsgQc.c2.found === 1 && hsgQc.c2.fail.length === 0, hsgQc.c2);
  check('HSG c3 qty vs item qty: exactly 4 real mismatches', hsgQc.c3.applicable === true &&
    JSON.stringify(hsgQc.c3.fail.map(f => f.number).sort()) === JSON.stringify(['2-999-06110', '2-999-97034', '7-238-23791', '7-999-01282']),
    hsgQc.c3.fail.map(f => f.number));
  check('HSG c4 entity icon: not applicable (column absent)', hsgQc.c4.applicable === false, hsgQc.c4);
  check('HSG c7 revision consistency: not applicable (no Revision column)', hsgQc.c7.applicable === false, hsgQc.c7);
  check('HSG c5 title/desc: 49 flagged (all description-missing on non-purchased parts)',
    hsgQc.c5.fail.length === 49 && hsgQc.c5.fail.every(f => f.kind === 'description-missing'), hsgQc.c5.fail.length);
  check('HSG c6 material: 6 non-assembly, non-purchased parts flagged (purchased parts excluded)',
    hsgQc.c6.applicable === true && hsgQc.c6.fail.length === 6 && hsgQc.c6.fail.every(f => !/^\d-999-/.test(f.number)),
    hsgQc.c6.fail.length);
  check('HSG c6 fail rows carry a real Row # and a non-empty parent (all are nested under some assembly)',
    hsgQc.c6.fail.every(f => f.sourceRow > 1 && f.parentNumber && f.parentTitle),
    hsgQc.c6.fail.map(f => ({ number: f.number, sourceRow: f.sourceRow, parentNumber: f.parentNumber, parentTitle: f.parentTitle })));
  check('HSG c5 fail rows carry a real Row # (parent may be blank only for a top-level row)',
    hsgQc.c5.fail.every(f => f.sourceRow > 1), hsgQc.c5.fail.length);
  // A clean release: no sketch part should ever have reached the Item Master.
  check('HSG c8: no 7-333-* sketch part in this Item Master (as it should be)',
    hsgQc.c8.applicable === true && hsgQc.c8.fail.length === 0, hsgQc.c8.fail);
  check('HSG c9: State column read, every row Certified',
    hsgQc.c9.applicable === true && hsgQc.c9.errorCount === 0 && hsgQc.c9.warnCount === 0, hsgQc.c9);
  check('HSG State column resolved to "State", not "File Link State"',
    im.hasState === true && im.rows.every(r => r.state === 'Certified'),
    Array.from(new Set(im.rows.map(r => r.state))));

  console.log('\n== real samples: material — bought-out parts + CAD vs Item Master ==');
  const boughtOut = imQc.boughtOutParts(im);
  check('HSG bought-out parts: 375 unique 7-999-* parts', boughtOut.length === 375, boughtOut.length);
  check('bought-out parts are all 7-999-* (deduplicated)',
    boughtOut.every(b => /^\d-999-/.test(b.number)) && new Set(boughtOut.map(b => b.number)).size === boughtOut.length,
    boughtOut.length);

  const cadRes0 = detect.parseCadFromWorkbook(cadWb, XLSX);
  const cad0 = cadRes0.ok;
  const matRes = materialCompare.compareMaterial([cad0], im);
  check('material check applicable with flat-xlsx CAD source', matRes.applicable === true, matRes);
  check('7 genuine (deduplicated, normalized) material mismatches on manufactured parts',
    matRes.mismatches.length === 7 && new Set(matRes.mismatches.map(m => m.number)).size === 7, matRes.mismatches.map(m => m.number));
  check('none of the mismatches are naming-convention noise (spot check: no bare 1.4301-vs-AISI304-style pair)',
    !matRes.mismatches.some(m => m.number === '7-240-21292'), matRes.mismatches.map(m => m.number));
  check('304-vs-304L genuine difference still flagged', matRes.mismatches.some(m => m.number === '7-238-27981'), matRes.mismatches.map(m => m.number));
  check('material mismatches carry Row # + parent info', matRes.mismatches.every(m => m.sourceRow > 1 && m.parentNumber),
    matRes.mismatches.map(m => ({ number: m.number, sourceRow: m.sourceRow, parentNumber: m.parentNumber })));
  check('bought-out parts carry Row # (all purchased parts are nested, none is the root row)',
    boughtOut.every(b => b.sourceRow > 1), boughtOut.length);
  check('material check not-applicable without a flat-xlsx CAD source (e.g. PDF-only)',
    materialCompare.compareMaterial([], im).applicable === false);

  const cadRes = detect.parseCadFromWorkbook(cadWb, XLSX);
  check('CAD parsed via flat parser', !!(cadRes && cadRes.ok && cadRes.ok.source === 'flat-xlsx'), cadRes && cadRes.ok && cadRes.ok.source);
  const cad = cadRes.ok;
  check('CAD records = 1723 (incl. split rows)', cad.items.length === 1723, cad.items.length);
  const cadUnique = new Set(cad.items.map(i => i.number.toUpperCase()));
  check('CAD unique PNs = 1231', cadUnique.size === 1231, cadUnique.size);
  check('split-row record recovered (7-236-20259 has title)', (function () {
    const it = cad.items.find(i => i.number === '7-236-20259');
    return it && it.title === 'Bearing Housing';
  })(), cad.items.find(i => i.number === '7-236-20259'));
  check('CAD flat export has no qty', cad.hasQty === false);
  // Real-world confirmation: the flat Vault export DOES carry Revision (a
  // documented-but-previously-unread fixed offset — see cad-flat-xlsx.js) —
  // real values are plain integers ("0", "1"...), not letters.
  check('CAD flat export hasRevision true', cad.hasRevision === true, cad.hasRevision);
  check('CAD flat export revision values are simple non-blank codes',
    cad.items.slice(0, 20).every(it => typeof it.revision === 'string' && it.revision !== ''),
    cad.items.slice(0, 5).map(it => it.revision));
  // The Item Master side has no Revision column (checked above), so the
  // cross-source comparison must degrade gracefully rather than crash.
  const revNotApplicable = revisionCompare.compareRevision([cad], im);
  check('revision check not-applicable on real data (Item Master has no Revision column)',
    revNotApplicable.applicable === false, revNotApplicable);

  console.log('\n== real samples: comparison baseline ==');
  const res = compare(cad, im);
  check('CAD unique count', res.cadUniqueCount === 1231, res.cadUniqueCount);
  check('IM unique count', res.imUniqueCount === 1076, res.imUniqueCount);
  check('missing unique PNs = 183', res.missingTotal === 183, res.missingTotal);
  check('qty comparison disabled', res.qtyMismatches === null);
  check('IM-only = 28', res.imOnly.length === 28, res.imOnly.length);
  check('IM-only rows carry sourceRow + parentNumber/parentTitle fields',
    res.imOnly.every(r => typeof r.sourceRow === 'number' && r.sourceRow > 0 &&
      typeof r.parentNumber === 'string' && typeof r.parentTitle === 'string'),
    res.imOnly.slice(0, 3));
  check('IM-only entries are also grouped under their flagged parent, losing nothing',
    res.imOnlyRoots.length <= res.imOnly.length &&
    res.imOnlyRoots.reduce((n, r) => n + 1 + countDescendants(r), 0) === res.imOnly.length,
    { roots: res.imOnlyRoots.length, flat: res.imOnly.length });

  // Cross-check de-duplication on real data. In THIS pairing (flat CAD_Bom
  // export vs the HSG Item Master) three parts are genuinely flagged by two
  // checks at once; each must collapse to one part owned by the more serious
  // check, with the loser demoted to a cross-reference.
  console.log('\n== real samples: findings registry de-duplication ==');
  const hsgReg = findings.buildRegistry({
    result: res, imQc: hsgQc, materialResult: matRes,
    revisionResult: revisionCompare.compareRevision([cad], im),
  });
  const overlaps = hsgReg.parts.filter(p => p.issues.length > 1);
  check('exactly 3 real parts are flagged by more than one check',
    overlaps.length === 3, overlaps.map(p => p.number + ':' + p.issues.map(i => i.key).join('+')));
  check('every one of them is owned by "In Item Master only", the more serious check',
    overlaps.every(p => p.primary.key === 'imOnly'), overlaps.map(p => p.primary.key));
  check('2-999-06110 (imOnly + c3): Check-3 row demoted, imOnly row kept primary',
    hsgReg.isSecondary('c3', '2-999-06110') === true && hsgReg.isSecondary('imOnly', '2-999-06110') === false);
  check('5-233-20286 (imOnly + c5): Check-5 row demoted', hsgReg.isSecondary('c5', '5-233-20286') === true);
  check('no part is registered twice',
    new Set(hsgReg.parts.map(p => normNumber(p.number))).size === hsgReg.parts.length, hsgReg.parts.length);
  check('every registered part has exactly one primary issue',
    hsgReg.parts.every(p => p.issues.filter(i => i.primary).length === 1), hsgReg.counts);
  check('actionable + grouped accounts for every registered part',
    hsgReg.counts.actionable + hsgReg.counts.grouped === hsgReg.counts.parts, hsgReg.counts);
  check('grouped children (explained by a flagged parent) are excluded from actionable',
    hsgReg.counts.grouped > 0 && hsgReg.parts.filter(p => p.grouped).every(p => !!p.groupedUnder),
    hsgReg.counts.grouped);
  // 7-238-23791 (LTB-4 PANEL MOUNTING FRAME) is present in this CAD export, so
  // here it is only an internal Check-3 finding — it is the OTHER pairing
  // (against the 732020066 export) where it is also "In Item Master only".
  const ltb = hsgReg.byPn.get('7-238-23791');
  check('LTB-4 7-238-23791 registered once, owned by Check 3 in this pairing',
    !!ltb && ltb.issues.length === 1 && ltb.primary.key === 'c3', ltb && ltb.issues.map(i => i.key));

  // The pairing the user actually reported: the 732020066 Inventor export vs
  // this same Item Master. There LTB-4 IS absent from the CAD BOM, so it is
  // flagged by "In Item Master only" AND Check 3 — the exact double-report
  // this whole change exists to collapse.
  if (inv732Path) {
    const inv732 = detect.parseCadFromWorkbook(XLSX.read(fs.readFileSync(inv732Path), { type: 'buffer' }), XLSX).ok;
    const res732 = compareAll([inv732], im);
    const reg732 = findings.buildRegistry({
      result: res732, imQc: hsgQc,
      materialResult: materialCompare.compareMaterial([inv732], im),
      revisionResult: revisionCompare.compareRevision([inv732], im),
    });
    const ltb732 = reg732.byPn.get('7-238-23791');
    check('LTB-4 vs 732020066: flagged by both "In Item Master only" and Check 3',
      !!ltb732 && ltb732.issues.length === 2 &&
      ltb732.issues.some(i => i.key === 'imOnly') && ltb732.issues.some(i => i.key === 'c3'),
      ltb732 && ltb732.issues.map(i => i.key));
    check('LTB-4 vs 732020066: reported once, owned by "In Item Master only"',
      !!ltb732 && ltb732.primary.key === 'imOnly' &&
      reg732.parts.filter(p => normNumber(p.number) === '7-238-23791').length === 1,
      ltb732 && ltb732.primary.key);
    check('LTB-4 vs 732020066: its Check-3 row renders as a cross-reference',
      reg732.isSecondary('c3', '7-238-23791') === true && reg732.isSecondary('imOnly', '7-238-23791') === false);
    check('vs 732020066: 1033 flat "In Item Master only" rows collapse to 11 roots',
      res732.imOnly.length === 1033 && res732.imOnlyRoots.length === 11,
      { flat: res732.imOnly.length, roots: res732.imOnlyRoots.length });
    check('vs 732020066: the rollup loses nothing (tree still holds all 1033)',
      res732.imOnlyRoots.reduce((n, r) => n + 1 + countDescendants(r), 0) === res732.imOnly.length);
  }

  const roots = res.missingRoots;
  check('actionable findings = 18', roots.length === 18, roots.length);
  const wetMill = roots.find(n => n.item.number === '7-260-20736');
  check('WET MILL is a top-level finding', !!wetMill);
  const wetMillDesc = wetMill ? countDescendants(wetMill) : 0;
  check('WET MILL absorbed its subtree (165 descendants)', wetMillDesc === 165, wetMillDesc);
  check('7-305-21355 (Top plate for bracket) grouped under WET MILL, not standalone',
    !roots.some(n => n.item.number === '7-305-21355'));

  // every missing PN appears exactly once in the result tree
  const treePNs = [];
  (function walk(nodes) { for (const n of nodes) { treePNs.push(n.item.number.toUpperCase()); walk(n.children); } })(roots);
  check('tree covers all missing PNs exactly once', treePNs.length === 183 && new Set(treePNs).size === 183, treePNs.length);

  console.log('\nActionable top-level findings (' + roots.length + '):');
  for (const n of roots) {
    const d = countDescendants(n);
    console.log('  ' + n.item.number.padEnd(18) + (n.item.isAssembly ? '[ASM] ' : '      ') +
      String(n.item.title).slice(0, 45).padEnd(46) + (d ? ' +' + d + ' grouped children' : ''));
  }
} else {
  console.log('\n(no sample file paths given — skipped real-sample baseline tests)');
}

if (pdf723Path && imPath && cadPath) {
  if (!pdfjsLib) {
    console.log('\n(pdfjs-dist not installed — run `npm install` to enable the PDF baseline tests)');
  } else {
    console.log('\n== real samples: Vault PDF 7-230-20509 (64 pages) vs Item Master ==');
    const { parsed: p723 } = await parsePdf(pdf723Path);
    check('723 PDF parsed 1820 records', p723.items.length === 1820, p723.items.length);
    check('723 PDF has levels, no qty', p723.hasLevels === true && p723.hasQty === false);
    check('723 PDF hasRevision true — the Vault "Uses" report header includes Revision', p723.hasRevision === true, p723.hasRevision);
    const badPn = p723.items.filter(i => !/^\d-\d{3}-\S+$/.test(i.number));
    check('723 PDF: every record has a clean part number', badPn.length === 0, badPn.slice(0, 3).map(i => i.number));

    const imWb2 = XLSX.read(fs.readFileSync(imPath), { type: 'buffer' });
    const im2 = detect.parseItemMasterFromWorkbook(imWb2, XLSX);
    const cadWb2 = XLSX.read(fs.readFileSync(cadPath), { type: 'buffer' });
    const flat2 = detect.parseCadFromWorkbook(cadWb2, XLSX).ok;
    const flatPNs = new Set(flat2.items.map(i => normNumber(i.number)));
    const pdfPNs = new Set(p723.items.map(i => normNumber(i.number)));
    const flatNotInPdf = [...flatPNs].filter(pn => !pdfPNs.has(pn));
    check('723 PDF covers every PN of the flat export', flatNotInPdf.length === 0, flatNotInPdf.slice(0, 5));

    const res723 = compareAll([p723], im2);
    check('723 PDF vs IM: 183 missing PNs', res723.missingTotal === 183, res723.missingTotal);
    check('723 PDF vs IM: 18 actionable findings', res723.actionableCount === 18, res723.actionableCount);
    const wetMill = res723.missingRoots.find(n => n.item.number === '7-260-20736');
    check('WET MILL is one grouped finding with 165 descendants', wetMill && countDescendants(wetMill) === 165,
      wetMill && countDescendants(wetMill));

    if (pdf732Path && inv732Path) {
      console.log('\n== real samples: Vault PDF 7-320-20066 + Inventor BOM export (reference detection) ==');
      const { parsed: p732 } = await parsePdf(pdf732Path);
      check('732 PDF parsed 639 records', p732.items.length === 639, p732.items.length);
      const invWb = XLSX.read(fs.readFileSync(inv732Path), { type: 'buffer' });
      const inv = detect.parseCadFromWorkbook(invWb, XLSX).ok;
      check('Inventor export parsed as leveled sheet', !!inv && inv.source === 'leveled-sheet', inv && inv.source);
      check('Inventor export: 608 items, qty + levels + BOM Structure',
        inv.items.length === 608 && inv.hasQty === true && inv.hasLevels === true && inv.hasStructure === true,
        inv && { n: inv.items.length, q: inv.hasQty, l: inv.hasLevels, s: inv.hasStructure });

      // no Item Master exists for this machine in the samples; reference
      // detection is independent of the IM, so a stub built from the
      // Inventor export is enough to drive compareAll.
      const stubIm = { rows: inv.items.map(it => ({ number: it.number, title: it.title, description: '', qty: it.qty, path: null, rowType: '', sourceRow: it.sourceRow })) };
      const res732 = compareAll([p732, inv], stubIm);
      check('732: 19 unique reference components', res732.referenceTotal === 19, res732.referenceTotal);
      check('732: reference findings grouped into 12 roots', res732.referenceRoots.length === 12, res732.referenceRoots.length);
      check('732: HUMAN mannequin detected as reference', res732.referenceRoots.some(n => n.item.number === '7-240-00000'),
        res732.referenceRoots.map(n => n.item.number));
      check('732: qty comparison active via Inventor export', res732.hasQty === true);
      console.log('\nReference components (' + res732.referenceTotal + ' in ' + res732.referenceRoots.length + ' findings):');
      for (const n of res732.referenceRoots) {
        const d = countDescendants(n);
        console.log('  L' + n.item.level + ' ' + n.item.number.padEnd(16) + String(n.item.title).slice(0, 42).padEnd(44) +
          (d ? ' +' + d + ' children' : ''));
      }
    }
  }
} else if (pdf723Path || pdf732Path) {
  console.log('\n(PDF tests need the flat CAD xlsx + Item Master paths as the first two arguments)');
}

if (pdf733Path && im733Path) {
  if (!pdfjsLib) {
    console.log('\n(pdfjs-dist not installed — run `npm install` to enable the PDF baseline tests)');
  } else {
    console.log('\n== real samples: Vault PDF 7-330-20013 (lab machine, 13 pages) vs Item Master ==');
    const { grid: g733, parsed: p733 } = await parsePdf(pdf733Path);
    check('733 PDF: no extraction warnings', (g733.warnings || []).length === 0, g733.warnings);
    check('733 PDF parsed 226 records', p733.items.length === 226, p733.items.length);
    check('733 PDF has levels, no qty', p733.hasLevels === true && p733.hasQty === false);
    const badPn733 = p733.items.filter(i => !/^\d-\d{3}-\S+$/.test(i.number));
    check('733 PDF: every record has a clean part number', badPn733.length === 0, badPn733.slice(0, 5).map(i => i.number));
    const lvl733 = {};
    for (const it of p733.items) lvl733[it.level] = (lvl733[it.level] || 0) + 1;
    check('733 PDF level histogram matches indentation depth 1-7',
      JSON.stringify(lvl733) === JSON.stringify({ 1: 1, 2: 44, 3: 64, 4: 63, 5: 25, 6: 13, 7: 16 }), lvl733);

    const im733Wb = XLSX.read(fs.readFileSync(im733Path), { type: 'buffer' });
    const im733 = detect.parseItemMasterFromWorkbook(im733Wb, XLSX);
    check('733 IM parsed with 194 rows and dotted Row Order paths', !!im733 && im733.rows.length === 194 && im733.hasPaths === true,
      im733 && { rows: im733.rows.length, hasPaths: im733.hasPaths });
    check('733 IM projectKey = SPN016808 / PN22752', im733.projectKey && im733.projectKey.spn === 'SPN016808' && im733.projectKey.pn === 'PN22752', im733.projectKey);

    const labQc = imQc.runChecks(im733);
    check('733 c3 qty vs item qty: 0 mismatches', labQc.c3.applicable === true && labQc.c3.fail.length === 0, labQc.c3.fail);
    check('733 c4 entity icon: not applicable (column absent)', labQc.c4.applicable === false, labQc.c4);

    const res733 = compareAll([p733], im733);
    check('733: 24 missing PNs / 22 actionable findings', res733.missingTotal === 24 && res733.actionableCount === 22,
      { missingTotal: res733.missingTotal, actionableCount: res733.actionableCount });
    check('733: TENTE castor is a standalone finding', res733.missingRoots.some(n => n.item.number === '7-999-11840' && n.children.length === 0),
      res733.missingRoots.map(n => n.item.number));
    const knob = res733.missingRoots.find(n => n.item.number === '7-331-20014');
    check('733: KNOB groups 1 child', knob && countDescendants(knob) === 1, knob && countDescendants(knob));
    const pulley = res733.missingRoots.find(n => n.item.number === '7-331-20005');
    check('733: PULLEY MACHINING groups 1 child', pulley && countDescendants(pulley) === 1, pulley && countDescendants(pulley));
  }
} else if (pdf733Path || im733Path) {
  console.log('\n(the lab-machine PDF test needs both the PDF and its Item Master path)');
}

if (lldboPath) {
  console.log('\n== real sample: PN22260 LLDBO parsing ==');
  const lldboWb = XLSX.read(fs.readFileSync(lldboPath), { type: 'buffer' });
  const lldbo = detect.parseLldboFromWorkbook(lldboWb, XLSX);
  check('LLDBO parsed with 16 rows', !!lldbo && lldbo.rows.length === 16, lldbo && lldbo.rows.length);
  check('LLDBO projectKey = SPN016838 / PN22260', lldbo.projectKey && lldbo.projectKey.spn === 'SPN016838' && lldbo.projectKey.pn === 'PN22260', lldbo.projectKey);
  check('LLDBO customer extracted', lldbo.customer === 'RADIANT NUTRACEUTICALS LTD, Bangladesh', lldbo.customer);
  const withPn = lldbo.rows.filter(r => r.partNo).length;
  check('LLDBO has 9 rows with a Part No, 7 without', withPn === 9 && (lldbo.rows.length - withPn) === 7, { withPn, total: lldbo.rows.length });
  check('LLDBO duplicate PN 7-999-07921 (wet + dry mill motor) both captured',
    lldbo.rows.filter(r => r.partNo === '7-999-07921').length === 2,
    lldbo.rows.filter(r => r.partNo === '7-999-07921').map(r => r.description));

  // dropped in the CAD box by mistake: must not silently misparse as a leveled CAD BOM
  const asCad = detect.parseCadFromWorkbook(lldboWb, XLSX);
  check('LLDBO dropped as CAD does not silently succeed', !(asCad && asCad.ok), asCad);

  if (imPath) {
    console.log('\n== real samples: LLDBO(PN22260) vs HSG Item Master(PN22426) — cross-project sanity check ==');
    const imWb2 = XLSX.read(fs.readFileSync(imPath), { type: 'buffer' });
    const im2 = detect.parseItemMasterFromWorkbook(imWb2, XLSX);
    const res = lldboCompare.compareLldbo(lldbo, im2, indexItemMaster);
    check('cross-project mismatch correctly flagged (different PN)',
      res.projectKeyMismatch && res.projectKeyMismatch.lldbo.pn === 'PN22260' && res.projectKeyMismatch.im.pn === 'PN22426',
      res.projectKeyMismatch);
    check('7 of 8 unique LLDBO part numbers absent from the unrelated Item Master',
      res.totalLldboItems === 8 && res.missingFromIm.length === 7, { total: res.totalLldboItems, missing: res.missingFromIm.length });
  }
}

if (invBomPath && imBomMatPath) {
  console.log('\n== real samples: INVENTOR_BOM_726020768.xlsx (material column) vs its Item Master ==');
  const invWb = XLSX.read(fs.readFileSync(invBomPath), { type: 'buffer' });
  const invRes = detect.parseCadFromWorkbook(invWb, XLSX);
  check('Inventor BOM export parsed as leveled sheet', !!(invRes && invRes.ok && invRes.ok.source === 'leveled-sheet'), invRes && invRes.ok && invRes.ok.source);
  const invCad = invRes.ok;
  check('726020768 Inventor BOM: 281 items, qty + material', invCad.items.length === 281 && invCad.hasQty === true && invCad.hasMaterial === true,
    invCad && { n: invCad.items.length, q: invCad.hasQty, m: invCad.hasMaterial });
  const withMat = invCad.items.filter(i => i.material).length;
  check('726020768 Inventor BOM: 237 items have a material value', withMat === 237, withMat);

  const imBomMatWb = XLSX.read(fs.readFileSync(imBomMatPath), { type: 'buffer' });
  const imBomMat = detect.parseItemMasterFromWorkbook(imBomMatWb, XLSX);
  check('726020768 Item Master: 302 rows, projectKey SPN016704/PN22829',
    !!imBomMat && imBomMat.rows.length === 302 && imBomMat.projectKey && imBomMat.projectKey.spn === 'SPN016704' && imBomMat.projectKey.pn === 'PN22829',
    imBomMat && { rows: imBomMat.rows.length, projectKey: imBomMat.projectKey });

  const qc726 = imQc.runChecks(imBomMat);
  check('726020768 c8: no sketch parts in this Item Master', qc726.c8.fail.length === 0, qc726.c8.fail);
  check('726020768 c9: State read, all Certified', qc726.c9.applicable === true &&
    qc726.c9.errorCount === 0 && qc726.c9.warnCount === 0, qc726.c9);

  const matRes726 = materialCompare.compareMaterial([invCad], imBomMat);
  check('material check applicable via the Inventor BOM export (not flat-xlsx)', matRes726.applicable === true, matRes726);
  check('726020768: 6 genuine material mismatches',
    matRes726.mismatches.length === 6 && new Set(matRes726.mismatches.map(m => m.number)).size === 6,
    matRes726.mismatches.map(m => m.number));
  check('726020768: CAD-side modeling gap caught (Generic.1 placeholder material)',
    matRes726.mismatches.some(m => m.cadMaterial === 'Generic.1'), matRes726.mismatches.map(m => m.cadMaterial));

  const folderKind = folder.classifyFolderFile('INVENTOR_BOM_726020768.xlsx'); // this org's real naming, not the upload system's temp path
  check('folder auto-load recognizes this exact real filename', folderKind === 'inventor-bom', folderKind);

  if (pdf726Path && pdfjsLib) {
    console.log('\n== real samples: 726020768 three-way (Vault PDF + Inventor BOM + Item Master) ==');
    const { parsed: pdfCad726 } = await parsePdf(pdf726Path);
    check('Vault PDF has no material column (as expected)', pdfCad726.hasMaterial === false, pdfCad726.hasMaterial);

    const combined = compareAll([pdfCad726, invCad], imBomMat);
    check('three-way compareAll does not crash and produces a result', !!combined && typeof combined.missingTotal === 'number', combined);
    check('material check still finds the Inventor BOM source when PDF is listed first',
      materialCompare.compareMaterial([pdfCad726, invCad], imBomMat).applicable === true);
  } else if (pdf726Path) {
    console.log('\n(pdfjs-dist not installed — skipped the 726020768 three-way PDF test)');
  }
} else if (invBomPath || imBomMatPath) {
  console.log('\n(the Inventor BOM material test needs both the Inventor BOM export and its Item Master path)');
}

if (invBom22819Path && imBom22819Path) {
  console.log('\n== real samples: PN22819 quantity-cascade regression (Housing subtree released at 2x) ==');
  const invWb = XLSX.read(fs.readFileSync(invBom22819Path), { type: 'buffer' });
  const invRes = detect.parseCadFromWorkbook(invWb, XLSX);
  const invCad = invRes && invRes.ok;
  check('PN22819 Inventor BOM export parsed as leveled sheet', !!(invCad && invCad.source === 'leveled-sheet'), invCad && invCad.source);

  const imWb = XLSX.read(fs.readFileSync(imBom22819Path), { type: 'buffer' });
  const im22819 = detect.parseItemMasterFromWorkbook(imWb, XLSX);
  check('PN22819 Item Master parsed', !!im22819 && im22819.rows.length > 0, im22819 && im22819.rows.length);

  const res22819 = compareAll([invCad], im22819);
  check('380 flat quantity mismatches (unchanged — the flat list is untouched by cascade detection)',
    res22819.qtyMismatches.length === 380, res22819.qtyMismatches.length);
  check('exactly one quantity cascade found', res22819.qtyCascades.applicable === true &&
    res22819.qtyCascades.roots.length === 1, res22819.qtyCascades.roots.map(r => r.item.number));
  const cascadeRoot = res22819.qtyCascades.roots[0];
  check('cascade root is the Housing assembly (7-705-23863), all 92 direct children at a clean 2x',
    cascadeRoot.item.number === '7-705-23863' && cascadeRoot.item.cascadeRatio === 2 &&
    cascadeRoot.item.cascadeChildCount === 92 && cascadeRoot.item.cascadeMismatchedChildCount === 92,
    cascadeRoot.item);
  check('cascade subtree covers a large share of the flat mismatch list (>= 300 descendants)',
    countDescendants(cascadeRoot) >= 300, countDescendants(cascadeRoot));
} else if (invBom22819Path || imBom22819Path) {
  console.log('\n(the PN22819 cascade regression needs both the Inventor BOM export and its Item Master path)');
}

if (imDiffOldPath && imDiffNewPath) {
  console.log('\n== real samples: Item Master diff (two PN22819 EBOM exports from this session) ==');
  const oldWb = XLSX.read(fs.readFileSync(imDiffOldPath), { type: 'buffer' });
  const imOld = itemMasterParser.parse(oldWb, XLSX);
  const newWb = XLSX.read(fs.readFileSync(imDiffNewPath), { type: 'buffer' });
  const imNew = itemMasterParser.parse(newWb, XLSX);
  check('both real Item Master files parsed', !!imOld && !!imNew, { old: !!imOld, new: !!imNew });

  const realDiff = imDiffCompare.diffItemMasters(imOld, imNew, indexItemMaster, materialCompare.materialsMatch);
  check('no false-positive "-" placeholder changes (regression: raw-equality must short-circuit materialsMatch)',
    !realDiff.changed.some(c => c.fields.some(f => f.field === 'Material' && f.old === f.new)),
    realDiff.changed.filter(c => c.fields.some(f => f.field === 'Material' && f.old === f.new)));
  check('every reported change actually differs (old !== new for every field)',
    realDiff.changed.every(c => c.fields.every(f => f.old !== f.new)),
    realDiff.changed.filter(c => c.fields.some(f => f.old === f.new)));
} else if (imDiffOldPath || imDiffNewPath) {
  console.log('\n(the Item Master diff regression needs both the older and newer Item Master paths)');
}

console.log(failures ? '\n' + failures + ' FAILURE(S)' : '\nall tests passed');
process.exit(failures ? 1 : 0);
