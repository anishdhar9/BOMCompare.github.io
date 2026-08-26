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
const { ignoreListParser } = require(path.join(rootDir, 'js/parsers/ignorelist.js'));
const { ignoreListCompare } = require(path.join(rootDir, 'js/ignorelist-compare.js'));
const { imDiffCompare } = require(path.join(rootDir, 'js/im-diff-compare.js'));
const { titleDescCompare } = require(path.join(rootDir, 'js/titledesc-compare.js'));
const { virtualParts } = require(path.join(rootDir, 'js/virtual-parts.js'));
const { ecrFill } = require(path.join(rootDir, 'js/ecr-fill.js'));

let failures = 0;
function check(name, cond, extra) {
  if (cond) console.log('  ok  ' + name);
  else { failures++; console.error('FAIL  ' + name + (extra !== undefined ? ' — got: ' + JSON.stringify(extra) : '')); }
}

// Shape + whitespace check for parsed CAD part numbers, shared by every real
// PDF sample below. Shape alone (no embedded-whitespace requirement) would
// still pass a hyphen-glued corruption like "7-999-00732-1/64" undetected —
// exactly the failure mode js/parsers/pdf-extract.js's footer floor + Part
// Number concatenation guard exist to prevent.
function checkCleanPns(label, items) {
  const badShape = items.filter(i => !/^\d-\d{3}-\S+$/.test(i.number));
  check(label + ': every record has a clean-shaped part number', badShape.length === 0, badShape.slice(0, 5).map(i => i.number));
  const withSpace = items.filter(i => /\s/.test(i.number));
  check(label + ': no part number contains embedded whitespace', withSpace.length === 0, withSpace.slice(0, 5).map(i => i.number));
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

console.log('\n== synthetic: ECR sheet generation (diffForEcr / fillEcrTemplate) ==');
{
  // Tree: MACH-01 (root, unchanged) -> ASSY-1 (revision bump 0->1) -> a
  // removed part, a qty-changed part, a material-only part, an unchanged
  // part, and a wholly new assembly (PART-ADDED-ASSY) with its own nested
  // child (PART-ADDED-CHILD); plus PARENT-A/PARENT-B, two SIBLINGS of
  // ASSY-1 that both use the SAME part number (SHARED-PART) but only the
  // occurrence under PARENT-B actually changes quantity — this is the
  // position-aware-matching regression case (a flat "first occurrence of
  // this PN anywhere" diff gets this wrong).
  const header = [
    'Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity', 'Revision', 'Material', 'State',
    'Producer', 'Producer Number',
  ];
  const imOldAoa = [
    header,
    ['MACH-01', '-', 'Machine', 'desc', '-', '0', '', 'Certified', 'SPN99999', '12345'],
    ['ASSY-1', '1', 'Assembly One', 'desc', '1 Each', '0', '', 'Certified', '', ''],
    ['PART-REMOVED', '1.1', 'Removed Part', 'desc', '2 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PART-QTYCHANGE', '1.2', 'Qty Change Part', 'desc', '2 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PART-MATONLY', '1.3', 'Mat Only Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PART-SAME', '1.4', 'Same Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PARENT-A', '2', 'Parent A', 'desc', '1 Each', '0', '', 'Certified', '', ''],
    ['SHARED-PART', '2.1', 'Shared Part', 'desc', '3 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PARENT-B', '3', 'Parent B', 'desc', '1 Each', '0', '', 'Certified', '', ''],
    ['SHARED-PART', '3.1', 'Shared Part', 'desc', '5 Each', '0', 'AISI 304', 'Certified', '', ''],
  ];
  const imNewAoa = [
    header,
    ['MACH-01', '-', 'Machine', 'desc', '-', '0', '', 'Certified', 'SPN99999', '12345'],
    ['ASSY-1', '1', 'Assembly One', 'desc', '1 Each', '1', '', 'Certified', '', ''],
    ['PART-QTYCHANGE', '1.1', 'Qty Change Part', 'desc', '4 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PART-MATONLY', '1.2', 'Mat Only Part', 'desc', '1 Each', '0', 'AISI 316', 'Certified', '', ''],
    ['PART-SAME', '1.3', 'Same Part', 'desc', '1 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PART-ADDED-ASSY', '1.4', 'Added Assembly', 'desc', '1 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PART-ADDED-CHILD', '1.4.1', 'Added Child', 'desc', '2 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PARENT-A', '2', 'Parent A', 'desc', '1 Each', '0', '', 'Certified', '', ''],
    ['SHARED-PART', '2.1', 'Shared Part', 'desc', '3 Each', '0', 'AISI 304', 'Certified', '', ''],
    ['PARENT-B', '3', 'Parent B', 'desc', '1 Each', '0', '', 'Certified', '', ''],
    ['SHARED-PART', '3.1', 'Shared Part', 'desc', '7 Each', '0', 'AISI 304', 'Certified', '', ''],
  ];
  const parseEcr = (aoa) => itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, { utils: { sheet_to_json: () => aoa } });
  const imOld = parseEcr(imOldAoa);
  const imNew = parseEcr(imNewAoa);

  const ecr = ecrFill.diffForEcr(imOld, imNew, indexItemMaster);
  check('7 ECR rows (default cascadeIntoNewSubtrees: true): revision-bump pair, qty changed, ' +
    'added assembly + its nested child, removed part, and the ONE occurrence of SHARED-PART that actually changed',
    ecr.rows.length === 7, ecr.rows.map(r => r.itemNoWithRev + ':' + r.action));

  check('revision bump: Revised row (new rev, 0->1) comes BEFORE Obsolete row (old rev, 1->0), ' +
    'both under the parent\'s NEW-side composite MACH-01-0 (not the old side)',
    ecr.rows[0].itemNoWithRev === 'ASSY-1-1' && ecr.rows[0].action === 'Drg. Revised' &&
    ecr.rows[0].oldQty === 0 && ecr.rows[0].newQty === 1 && ecr.rows[0].subAssyNumberWithRev === 'MACH-01-0' &&
    ecr.rows[1].itemNoWithRev === 'ASSY-1-0' && ecr.rows[1].action === 'Drg. Obsolete' &&
    ecr.rows[1].oldQty === 1 && ecr.rows[1].newQty === 0 && ecr.rows[1].subAssyNumberWithRev === 'MACH-01-0',
    ecr.rows.slice(0, 2));

  const qtyRow = ecr.rows.find(r => r.action === 'Qty Changed' && r.subAssyNumberWithRev === 'ASSY-1-1');
  check('qty-only change under the (now-revised) parent: bare item number (no revision suffix), numeric 2 -> 4',
    !!qtyRow && qtyRow.itemNoWithRev === 'PART-QTYCHANGE' && qtyRow.oldQty === 2 && qtyRow.newQty === 4, qtyRow);

  const addedAssyRow = ecr.rows.find(r => r.action === 'Part Added' && r.itemNoWithRev === 'PART-ADDED-ASSY');
  check('added assembly: bare item number, parent ASSY-1-1, Old 0 -> New 1 (numeric)',
    !!addedAssyRow && addedAssyRow.subAssyNumberWithRev === 'ASSY-1-1' &&
    addedAssyRow.oldQty === 0 && addedAssyRow.newQty === 1, addedAssyRow);

  const addedChildRow = ecr.rows.find(r => r.itemNoWithRev === 'PART-ADDED-CHILD');
  check('cascade into a wholly new subtree: nested child gets its own Part Added row, ' +
    'parented to the new assembly\'s OWN revisioned composite (PART-ADDED-ASSY-0)',
    !!addedChildRow && addedChildRow.action === 'Part Added' &&
    addedChildRow.subAssyNumberWithRev === 'PART-ADDED-ASSY-0' && addedChildRow.newQty === 2, addedChildRow);

  const removedRow = ecr.rows.find(r => r.action === 'Part Deleted');
  check('removed row: bare item number, parent ASSY-1-1 (still exists in the new tree), Old 2 -> New 0',
    !!removedRow && removedRow.itemNoWithRev === 'PART-REMOVED' &&
    removedRow.subAssyNumberWithRev === 'ASSY-1-1' && removedRow.oldQty === 2 && removedRow.newQty === 0, removedRow);

  const sharedRows = ecr.rows.filter(r => r.itemNoWithRev === 'SHARED-PART');
  check('multi-occurrence: SHARED-PART changes ONLY under PARENT-B (5->7) — the unchanged occurrence under ' +
    'PARENT-A produces no row at all, not a false positive or a merged/first-occurrence result',
    sharedRows.length === 1 && sharedRows[0].subAssyNumberWithRev === 'PARENT-B-0' &&
    sharedRows[0].oldQty === 5 && sharedRows[0].newQty === 7, sharedRows);

  check('material-only change (PART-MATONLY) is reported as an "other change", not an ECR row; ' +
    'PART-SAME never appears anywhere',
    ecr.otherChanges.length === 1 && ecr.otherChanges[0].number === 'PART-MATONLY' &&
    ecr.otherChanges[0].fields.includes('Material') &&
    !ecr.rows.some(r => r.itemNoWithRev === 'PART-SAME'), ecr.otherChanges);

  const ecrNoCascade = ecrFill.diffForEcr(imOld, imNew, indexItemMaster, { cascadeIntoNewSubtrees: false });
  check('cascadeIntoNewSubtrees: false collapses the new assembly to ONE row — its nested child is not listed',
    ecrNoCascade.rows.length === 6 && !ecrNoCascade.rows.some(r => r.itemNoWithRev === 'PART-ADDED-CHILD') &&
    ecrNoCascade.rows.some(r => r.itemNoWithRev === 'PART-ADDED-ASSY'), ecrNoCascade.rows.map(r => r.itemNoWithRev));

  // fillEcrTemplate against the real vendored company template.
  const templatePath = path.join(rootDir, 'vendor/ECR_template.xlsx');
  if (fs.existsSync(templatePath)) {
    const templateWb = XLSX.readFile(templatePath);
    ecrFill.fillEcrTemplate(templateWb, ecr.rows, XLSX, imNew.projectKey);
    const sheet = templateWb.Sheets['Sheet1'];
    const filledAoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
    check('template header row (row 11) is untouched', filledAoa[10][0] === 'Line No.', filledAoa[10]);
    check('template data starts at row 12 with Line No. 1', String(filledAoa[11][0]) === '1', filledAoa[11]);
    check('template row count matches the number of ECR rows generated',
      filledAoa.slice(11, 11 + ecr.rows.length).every(r => r[0] !== null), filledAoa.slice(11, 16));
    check('a filled row carries the right composite item number in column C',
      filledAoa[11][2] === ecr.rows[0].itemNoWithRev, { got: filledAoa[11][2], expected: ecr.rows[0].itemNoWithRev });
    check('PN header field (B3) is auto-filled from the parsed project key',
      String(filledAoa[2][1]) === '12345', filledAoa[2]);
  } else {
    console.log('\n(vendor/ECR_template.xlsx not found — skipped the template-fill check)');
  }
}

console.log('\n== synthetic: positional ancestor resolution (duplicate Row Order positions) ==');
{
  // Real exports reuse a Row Order position for adjacent siblings. Here
  // "1.1" appears twice: the second one (PART-SECOND) is the true parent of
  // the row at "1.1.1" that follows it. A path-string lookup returns the
  // FIRST "1.1" row instead, mis-parenting the child and, through the
  // ancestor multipliers, inflating its rolled-up quantity.
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity'],
    ['7-000-ROOT', '-', 'Machine', 'desc', '-'],
    ['7-100-BRANCH', '1', 'Branch', 'desc', '1 Each'],
    ['7-100-FIRST', '1.1', 'First at this position', 'desc', '5 Each'],
    ['7-100-FIRST-KID', '1.1.1', 'Child of FIRST', 'desc', '1 Each'],
    ['7-100-SECOND', '1.1', 'Second at the SAME position', 'desc', '2 Each'],
    ['7-100-SECOND-KID', '1.1.1', 'Child of SECOND', 'desc', '3 Each'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => aoa },
  });
  const idx = indexItemMaster(im);
  const rowOf = (n) => im.rows.find(r => r.number === n);

  check('duplicate Row Order positions are still reported as a warning',
    idx.warnings.some(w => /share a "Row Order" position/.test(w)), idx.warnings);
  check('the second child resolves to the SECOND row at the reused position, not the first',
    idx.parentOf.get(rowOf('7-100-SECOND-KID')) === rowOf('7-100-SECOND'),
    idx.parentOf.get(rowOf('7-100-SECOND-KID')) && idx.parentOf.get(rowOf('7-100-SECOND-KID')).number);
  check('the first child still resolves to the first row at that position',
    idx.parentOf.get(rowOf('7-100-FIRST-KID')) === rowOf('7-100-FIRST'),
    idx.parentOf.get(rowOf('7-100-FIRST-KID')) && idx.parentOf.get(rowOf('7-100-FIRST-KID')).number);
  check('a depth-1 row resolves to the root row', idx.parentOf.get(rowOf('7-100-BRANCH')) === rowOf('7-000-ROOT'));
  check('the root row itself has no parent', idx.parentOf.get(rowOf('7-000-ROOT')) === undefined);
  // SECOND-KID: own 3 x SECOND 2 x BRANCH 1 = 6. Under the old path-string
  // lookup its ancestor would have been FIRST (qty 5), giving 15.
  check('rolled-up quantity uses the correct ancestor chain (3 x 2 x 1 = 6, not 3 x 5 x 1 = 15)',
    idx.totals.get('7-100-SECOND-KID') === 6, idx.totals.get('7-100-SECOND-KID'));
  check('childRows attributes each child to its real parent',
    (idx.childRows.get('7-100-SECOND') || []).length === 1 &&
    idx.childRows.get('7-100-SECOND')[0].number === '7-100-SECOND-KID',
    (idx.childRows.get('7-100-SECOND') || []).map(r => r.number));
}

console.log('\n== synthetic: Description CAD vs Item Master (titledesc-compare) ==');
{
  const M = titleDescCompare.descriptionsMatch;
  check('case is ignored', M('gcpilot', 'GCPilot') === true);
  check('spacing inside a token is ignored: "TANK - 300LTRS" == "TANK - 300 LTRS"',
    M('TANK - 300LTRS', 'TANK - 300 LTRS') === true);
  check('spacing is ignored: "GCPilot" == "GC Pilot"', M('GCPilot', 'GC Pilot') === true);
  check('a changed dimension is NOT absorbed: "OD 539 X 4 THK." != "OD 539 X 3 THK."',
    M('OD 539 X 4 THK.', 'OD 539 X 3 THK.') === false);
  check('a changed thread size is NOT absorbed: "G 1/4" != "G 1/8"',
    M('G 1/4, Dia. 7.5', 'G 1/8, Dia. 7.5') === false);
  check('punctuation stays significant: "A - B" != "A + B"', M('A - B', 'A + B') === false);
  check('blank never matches', M('', 'x') === false && M('x', '') === false);

  const A = titleDescCompare.isAutoMatchedAppend;
  check('CAD-appended material suffix is an auto-match: "GSF PRO 180" -> "GSF PRO 180 AISI 304"',
    A('GSF PRO 180', 'GSF PRO 180 AISI 304') === true);
  check('auto-match is direction-agnostic (IM longer than CAD too)',
    A('GSF PRO 180 AISI 304', 'GSF PRO 180') === true);
  check('a genuine mid-string digit change is NOT an auto-match (must stay flagged)',
    A('OD 539 X 4 THK.', 'OD 539 X 3 THK.') === false);
  check('a genuine thread-size change is NOT an auto-match (must stay flagged)',
    A('G 1/4, Dia. 7.5', 'G 1/8, Dia. 7.5') === false);
  check('already-equal strings are not "auto-matched" (that is descriptionsMatch\'s job)',
    A('Same Text', 'Same Text') === false);
  check('blank never auto-matches', A('', 'x') === false && A('x', '') === false);

  const imAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)'],
    ['7-000-ROOT', '-', 'Machine', 'Machine desc'],
    ['7-100-SAME', '1', 'Same', 'PIPE OD38.1X1.6'],
    ['7-100-DIFF', '2', 'Differs', 'OD 539 X 3 THK.'],
    ['7-100-SPACING', '3', 'Spacing only', 'TANK - 300 LTRS'],
    ['7-999-00001', '4', 'Purchased', 'supplier text A'],
    ['2-100-PROCURED', '5', 'Procured elsewhere', 'something else'],
    ['7-909-00001', '6', 'END OF LINE', 'marker text'],
    ['7-100-IMBLANK', '7', 'IM blank', ''],
    ['7-100-SUFFIX', '8', 'Material suffix', 'GSF PRO 180'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => imAoa },
  });
  const cadSource = {
    kind: 'cad', source: 'leveled-sheet', hasQty: false, items: [
      { number: '7-100-SAME', description: 'PIPE OD38.1X1.6' },
      { number: '7-100-DIFF', description: 'OD 539 X 4 THK.' },
      { number: '7-100-SPACING', description: 'TANK - 300LTRS' },
      { number: '7-999-00001', description: 'supplier text B' },
      { number: '2-100-PROCURED', description: 'a different thing' },
      { number: '7-909-00001', description: 'anything at all' },
      { number: '7-100-IMBLANK', description: 'CAD has one' },
      { number: '7-100-SUFFIX', description: 'GSF PRO 180 AISI 304' },
    ],
  };
  const res = titleDescCompare.compareTitleDescription([cadSource], im);
  check('applicable when a CAD source carries descriptions', res.applicable === true, res.reason);
  check('only the genuine difference is flagged',
    res.mismatches.length === 1 && res.mismatches[0].number === '7-100-DIFF',
    res.mismatches.map(m => m.number));
  check('a CAD-appended material suffix lands in autoMatched, not mismatches',
    !res.mismatches.some(m => m.number === '7-100-SUFFIX') &&
    res.autoMatched.length === 1 && res.autoMatched[0].number === '7-100-SUFFIX',
    { mismatches: res.mismatches.map(m => m.number), autoMatched: res.autoMatched.map(m => m.number) });
  check('spacing-only difference is not flagged', !res.mismatches.some(m => m.number === '7-100-SPACING'));
  check('purchased X-999 part is excluded', !res.mismatches.some(m => m.number === '7-999-00001'));
  check('non-7 procured part is excluded', !res.mismatches.some(m => m.number === '2-100-PROCURED'));
  check('END OF LINE marker is excluded', !res.mismatches.some(m => m.number === '7-909-00001'));
  check('a blank Item Master description is left to Check 5, not flagged here',
    !res.mismatches.some(m => m.number === '7-100-IMBLANK'));
  check('mismatch carries both values and location', res.mismatches[0].imDescription === 'OD 539 X 3 THK.' &&
    res.mismatches[0].cadDescription === 'OD 539 X 4 THK.' && res.mismatches[0].sourceRow > 1,
    res.mismatches[0]);

  const noCad = titleDescCompare.compareTitleDescription([{ items: [{ number: 'X', description: '' }] }], im);
  check('not applicable when no CAD source carries description text', noCad.applicable === false, noCad.reason);
}

console.log('\n== synthetic: virtual parts (no CAD file behind a subassembly) ==');
{
  // VIRT-A is in CAD with thumbnail "(NULL)" and all 2 of its Item Master
  // children absent from CAD -> confirmed virtual.
  // REAL-B is in CAD with a thumbnail and its children in CAD -> not virtual.
  // PARTIAL-C has children but one of them IS in CAD -> a real assembly.
  const imAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)'],
    ['7-000-ROOT', '-', 'Machine', 'desc'],
    ['7-100-VIRT-A', '1', 'Virtual assembly', 'desc'],
    ['7-100-VA-KID1', '1.1', 'Kid 1', 'desc'],
    ['7-100-VA-KID2', '1.2', 'Kid 2', 'desc'],
    ['7-100-REAL-B', '2', 'Real assembly', 'desc'],
    ['7-100-RB-KID1', '2.1', 'Kid 1', 'desc'],
    ['7-100-PARTIAL-C', '3', 'Partly modelled', 'desc'],
    ['7-100-PC-KID1', '3.1', 'In CAD', 'desc'],
    ['7-100-PC-KID2', '3.2', 'Not in CAD', 'desc'],
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => imAoa },
  });
  const withThumb = {
    kind: 'cad', source: 'leveled-sheet', hasThumbnail: true, items: [
      { number: '7-100-VIRT-A', thumbnailMissing: true },
      { number: '7-100-REAL-B', thumbnailMissing: false },
      { number: '7-100-RB-KID1', thumbnailMissing: false },
      { number: '7-100-PARTIAL-C', thumbnailMissing: true },
      { number: '7-100-PC-KID1', thumbnailMissing: false },
    ],
  };
  const res = virtualParts.detectVirtualParts([withThumb], im, indexItemMaster);
  check('applicable with a Thumbnail column present', res.applicable === true && res.hasThumbnailColumn === true);
  check('exactly one confirmed virtual part (VIRT-A)',
    res.confirmed.length === 1 && res.confirmed[0].number === '7-100-VIRT-A', res.confirmed.map(v => v.number));
  check('it reports both orphaned children', res.confirmed[0].childCount === 2 &&
    res.confirmed[0].children.map(c => c.number).join(',') === '7-100-VA-KID1,7-100-VA-KID2',
    res.confirmed[0].children);
  check('a modelled assembly whose children are in CAD is not virtual',
    !res.confirmed.some(v => v.number === '7-100-REAL-B'));
  check('an assembly with SOME children in CAD is not virtual, even with a (NULL) thumbnail',
    !res.confirmed.some(v => v.number === '7-100-PARTIAL-C'));
  check('nothing is "suspected" when the Thumbnail column is available', res.suspected.length === 0);
  check('anchorRows exposes the virtual part for imOnly grouping', res.anchorRows.has('7-100-VIRT-A'));

  // Without the Thumbnail column, fall back to the >=3-orphan-child rule.
  const noThumb = {
    kind: 'cad', source: 'leveled-sheet', hasThumbnail: false, items: [
      { number: '7-100-VIRT-A' }, { number: '7-100-REAL-B' }, { number: '7-100-RB-KID1' },
      { number: '7-100-PARTIAL-C' }, { number: '7-100-PC-KID1' },
    ],
  };
  const fb = virtualParts.detectVirtualParts([noThumb], im, indexItemMaster);
  check('fallback: 2 orphan children is below the threshold, so nothing is reported',
    fb.applicable === true && fb.hasThumbnailColumn === false &&
    fb.confirmed.length === 0 && fb.suspected.length === 0,
    { confirmed: fb.confirmed.length, suspected: fb.suspected.length });
  check('fallback explains that the Thumbnail column would give an exact answer',
    /Thumbnail/.test(fb.reason), fb.reason);

  const imBig = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: {
      sheet_to_json: () => [
        ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)'],
        ['7-000-ROOT', '-', 'Machine', 'desc'],
        ['7-100-VIRT-A', '1', 'Virtual assembly', 'desc'],
        ['7-100-K1', '1.1', 'K1', 'desc'], ['7-100-K2', '1.2', 'K2', 'desc'], ['7-100-K3', '1.3', 'K3', 'desc'],
      ],
    },
  });
  const fb3 = virtualParts.detectVirtualParts(
    [{ kind: 'cad', hasThumbnail: false, items: [{ number: '7-100-VIRT-A' }] }], imBig, indexItemMaster);
  check('fallback: 3 orphan children meets the threshold and is reported as suspected',
    fb3.suspected.length === 1 && fb3.suspected[0].number === '7-100-VIRT-A' &&
    fb3.suspected[0].confidence === 'suspected', fb3.suspected.map(v => v.number));

  const noPaths = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => [['Number', 'Title (Item,CO)'], ['7-100-A', 'A']] },
  });
  const na = virtualParts.detectVirtualParts([withThumb], noPaths, indexItemMaster);
  check('not applicable without a Row Order column (degrades, does not crash)',
    na.applicable === false && /Row Order/.test(na.reason), na.reason);
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

console.log('\n== synthetic: leveled CAD parsing captures Thumbnail column ==');
{
  // Inventor writes an EMPTY cell when a CAD file backs the row and the
  // literal text "(NULL)" when none does, so only "(NULL)" counts.
  const aoa = [
    ['Item', 'Part Number', 'Thumbnail', 'BOM Structure', 'QTY'],
    ['1', '7-100-HASFILE', '', 'Normal', '1'],
    ['2', '7-100-VIRTUAL', '(NULL)', 'Normal', '1'],
  ];
  const cad = cadLeveledParser.parse(aoa, { source: 'leveled-sheet' });
  check('Thumbnail column detected', cad.hasThumbnail === true);
  check('empty thumbnail cell means a CAD file exists',
    cad.items.find(i => i.number === '7-100-HASFILE').thumbnailMissing === false);
  check('"(NULL)" thumbnail marks a virtual component',
    cad.items.find(i => i.number === '7-100-VIRTUAL').thumbnailMissing === true);

  const noCol = cadLeveledParser.parse([['Item', 'Part Number', 'QTY'], ['1', '7-100-A', '1']], { source: 'leveled-sheet' });
  check('hasThumbnail false when the column is absent', noCol.hasThumbnail === false);
  check('no row looks virtual when the column is absent',
    noCol.items.every(i => i.thumbnailMissing === false));

  // Header presence alone drives hasThumbnail: an export with no virtual
  // parts has an entirely empty column and must still read as "checked".
  const allEmpty = cadLeveledParser.parse(
    [['Item', 'Part Number', 'Thumbnail', 'QTY'], ['1', '7-100-A', '', '1']], { source: 'leveled-sheet' });
  check('an all-empty Thumbnail column still counts as present (checked, none found)',
    allEmpty.hasThumbnail === true && allEmpty.items.every(i => i.thumbnailMissing === false));
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

console.log('\n== synthetic: Vault PDF footer/pagination exclusion (2 pages) ==');
{
  // Reproduces the reported bug: a page-number footer ("N / total") sitting
  // below the last real row on a page gets glued onto that row's Part
  // Number, e.g. "7-999-00732" + " 1/64" -> "7-999-00732 1/64". This is
  // covered by two independent layers in extractGrid (see js/parsers/
  // pdf-extract.js) and this fixture is built to exercise both:
  //  - page 1's footer sits well below the page's established row pitch, so
  //    the footer-floor filter should drop it before it ever reaches a row.
  //  - page 2's footer sits close enough to its nearest row that the floor
  //    alone might not catch it — the Part Number concatenation guard (only
  //    a fragment after a trailing "-" may extend a Part Number) is what
  //    must reject it there instead.
  const { pdfExtract } = require(path.join(rootDir, 'js/parsers/pdf-extract.js'));
  global.BOMCompare = { cadLeveledParser };
  const mk = (str, x, y, w = 30, h = 8) => ({ str, transform: [1, 0, 0, h, x, y], width: w });
  const header = [
    mk('Name', 33, 700), mk('Revision', 292, 700), mk('State', 355, 700), mk('Title', 438, 700), mk('Description', 642, 700), mk('Part', 825, 700),
    mk('Number', 825, 690),
  ];
  const bomRow = (num, x, y, title, desc) => [
    mk(num + '.ipt', x, y, 120), mk('0', 292, y), mk('Released', 371, y), mk(title, 438, y, 80), mk(desc, 642, y, 100), mk(num, 825, y, 70),
  ];
  const page1 = header
    .concat(bomRow('7-111-11111', 64, 650, 'Part One', 'First part'))
    .concat(bomRow('7-111-22222', 64, 610, 'Part Two', 'Second part'))
    .concat([mk('1 / 2', 825, 430, 40)]); // far below the row pitch -> floor should drop it
  const page2 = bomRow('7-111-33333', 64, 650, 'Part Three', 'Third part')
    .concat(bomRow('7-111-44444', 64, 610, 'Part Four', 'Fourth part'))
    .concat([mk('2 / 2', 825, 590, 40)]); // close to its row -> the concatenation guard must catch it
  const pagesItems = [page1, page2];
  const pdfjsLib = {
    getDocument: () => ({
      promise: Promise.resolve({
        numPages: pagesItems.length,
        getPage: async (p) => ({ getTextContent: async () => ({ items: pagesItems[p - 1] }) }),
      }),
    }),
  };
  const grid = await pdfExtract.extractGrid(new ArrayBuffer(0), { pdfjsLib });
  const numbers = grid.rows.slice(1).map(r => r[5]);
  check('all 4 part numbers recovered clean, no footer text attached',
    JSON.stringify(numbers) === JSON.stringify(['7-111-11111', '7-111-22222', '7-111-33333', '7-111-44444']), numbers);
  check('pageOf tracks each row\'s source page', JSON.stringify(grid.pageOf.slice(1)) === JSON.stringify([1, 1, 2, 2]), grid.pageOf);
  check('the page-2 footer (too close for the floor alone) was caught and logged',
    (grid.warnings || []).some(w => /discarded unexpected text/.test(w)), grid.warnings);
}

console.log('\n== synthetic: Ignore List — parsing and category mapping ==');
{
  const aoa = [
    ['S.No.', 'Part Number', 'From'],
    ['1', 'IGNORED-MISSING', 'CAD vs Item compare'],
    ['2', 'ignored-qty', ' CAD VS ITEM COMPARE '], // case/whitespace-insensitive match
    ['3', 'IGNORED-UNKNOWN', 'Some Other Category'], // unrecognized -> reported, not applied
    ['4', '', 'CAD vs Item compare'], // blank part number -> skipped
  ];
  const wb = { SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } };
  const XLSXStub = { utils: { sheet_to_json: () => aoa } };
  const parsed = ignoreListParser.parse(wb, XLSXStub);
  check('ignore list parsed', !!parsed, parsed);
  check('ignore list: 3 rows (blank Part Number row skipped)', parsed.rows.length === 3, parsed.rows.length);
  check('ignore list: sourceRow is the real spreadsheet row', parsed.rows[0].sourceRow === 2, parsed.rows[0]);
  check('non-Ignore-List sheet (no "Part Number" header) returns null',
    ignoreListParser.parse(wb, { utils: { sheet_to_json: () => [['Number', 'Title'], ['X', 'Y']] } }) === null);

  const idx = ignoreListCompare.buildIgnoreIndex(parsed);
  check('recognized category resolves case/whitespace-insensitively',
    idx.isIgnored('IGNORED-MISSING', 'missing') && idx.isIgnored('ignored-qty', 'qty'));
  check('recognized category covers all 4 CAD-vs-Item-Master keys',
    ['missing', 'reference', 'qty', 'imOnly'].every(k => idx.isIgnored('IGNORED-MISSING', k)));
  check('recognized category does not cover an unrelated key', !idx.isIgnored('IGNORED-MISSING', 'material'));
  check('unrecognized "From" value is reported, not silently applied',
    idx.unrecognized.length === 1 && idx.unrecognized[0].number === 'IGNORED-UNKNOWN' && !idx.isIgnored('IGNORED-UNKNOWN', 'missing'),
    idx.unrecognized);
  check('buildIgnoreIndex(null) is a safe no-op', ignoreListCompare.buildIgnoreIndex(null).isIgnored('ANY', 'missing') === false);

  // "Quantity Mismatch" is the narrower category: a part still gets flagged
  // if genuinely missing/reference/in-Item-Master-only, only qty is suppressed.
  const idxQty = ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'IGNORED-QTY-ONLY', from: 'Quantity Mismatch', sourceRow: 5 }] });
  check('"Quantity Mismatch" category resolves case/whitespace-insensitively too',
    ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'X', from: ' quantity  MISMATCH ', sourceRow: 1 }] }).isIgnored('X', 'qty'));
  check('"Quantity Mismatch" category suppresses qty', idxQty.isIgnored('IGNORED-QTY-ONLY', 'qty'));
  check('"Quantity Mismatch" category does NOT suppress missing/reference/imOnly',
    !idxQty.isIgnored('IGNORED-QTY-ONLY', 'missing') && !idxQty.isIgnored('IGNORED-QTY-ONLY', 'reference') && !idxQty.isIgnored('IGNORED-QTY-ONLY', 'imOnly'));

  // "Revision" suppresses js/revision-compare.js's compareRevision() — a
  // separate module from compareAll(), reusing the same checkKey mechanism.
  const idxRevision = ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'IGNORED-REV', from: 'Revision', sourceRow: 7 }] });
  check('"Revision" category resolves case/whitespace-insensitively too',
    ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'X', from: ' revision ', sourceRow: 1 }] }).isIgnored('X', 'revision'));
  check('"Revision" category suppresses revision', idxRevision.isIgnored('IGNORED-REV', 'revision'));
  check('"Revision" category does NOT suppress missing/reference/qty/imOnly',
    !idxRevision.isIgnored('IGNORED-REV', 'missing') && !idxRevision.isIgnored('IGNORED-REV', 'qty') && !idxRevision.isIgnored('IGNORED-REV', 'imOnly'));

  // "LLDBO Candidate" suppresses js/lldbo-compare.js's detectLldboCandidates()
  // — a separate module from compareAll(), reusing the same checkKey mechanism.
  const idxLldboCand = ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'IGNORED-CAND', from: 'LLDBO Candidate', sourceRow: 8 }] });
  check('"LLDBO Candidate" category resolves case/whitespace-insensitively too',
    ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'X', from: ' lldbo   candidate ', sourceRow: 1 }] }).isIgnored('X', 'lldboCandidate'));
  check('"LLDBO Candidate" category suppresses lldboCandidate', idxLldboCand.isIgnored('IGNORED-CAND', 'lldboCandidate'));
  check('"LLDBO Candidate" category does NOT suppress missing/reference/qty/imOnly/revision',
    !idxLldboCand.isIgnored('IGNORED-CAND', 'missing') && !idxLldboCand.isIgnored('IGNORED-CAND', 'qty') &&
    !idxLldboCand.isIgnored('IGNORED-CAND', 'imOnly') && !idxLldboCand.isIgnored('IGNORED-CAND', 'revision'));

  // "All" is the broadest category: the union of every other category,
  // computed rather than hardcoded — so it stays complete even as more
  // categories get added later, not just the ones known at the time it
  // was written.
  const idxAll = ignoreListCompare.buildIgnoreIndex({ rows: [{ number: 'IGNORED-ALL', from: ' ALL ', sourceRow: 6 }] });
  check('"All" category (mixed case/whitespace in the file) resolves and applies',
    idxAll.isIgnored('IGNORED-ALL', 'missing'), idxAll.byPn.get('IGNORED-ALL'));
  check('"All" category suppresses every key any other category maps to',
    Object.keys(ignoreListCompare.CATEGORIES).filter(c => c !== 'all')
      .every(cat => ignoreListCompare.CATEGORIES[cat].every(k => idxAll.isIgnored('IGNORED-ALL', k))),
    ignoreListCompare.CATEGORIES);
  check('"All" specifically covers missing/reference/qty/imOnly today',
    ['missing', 'reference', 'qty', 'imOnly'].every(k => idxAll.isIgnored('IGNORED-ALL', k)));
  check('"All" also covers revision — proves the computed union stayed in sync when the category was added',
    idxAll.isIgnored('IGNORED-ALL', 'revision'));
  check('"All" also covers lldboCandidate — proves the computed union stayed in sync when the category was added',
    idxAll.isIgnored('IGNORED-ALL', 'lldboCandidate'));
}

console.log('\n== synthetic: Ignore List suppresses compareAll() findings, with child promotion ==');
{
  // Machine -> IGNORED-ASSY (missing, ignored; children CHILD-1 kept, IGNORED-CHILD also ignored)
  //         -> PART-A (missing, not ignored)
  //         -> PART-B (qty mismatch, ignored)
  //         -> PART-C (qty mismatch, not ignored)
  const aoa = [
    ['Item', 'Number', 'Title', 'QTY', 'File'],
    ['1', 'MACH-01', 'Machine', '1', 'mach.iam'],
    ['1.1', 'IGNORED-ASSY', 'Ignored assembly', '1', 'ia.iam'],
    ['1.1.1', 'CHILD-1', 'Kept child', '2', 'c1.ipt'],
    ['1.1.2', 'IGNORED-CHILD', 'Also ignored child', '1', 'c2.ipt'],
    ['1.2', 'PART-A', 'Not ignored, missing', '1', 'pa.ipt'],
    ['1.3', 'PART-B', 'Ignored qty mismatch', '5', 'pb.ipt'],
    ['1.4', 'PART-C', 'Not ignored qty mismatch', '5', 'pc.ipt'],
  ];
  const cad = cadLeveledParser.parse(aoa, { source: 'leveled-sheet' });
  const im = {
    rows: [
      { number: 'MACH-01', title: 'Machine', qty: null, path: [] },
      { number: 'PART-B', title: '', qty: 3, path: ['1'] },
      { number: 'PART-C', title: '', qty: 3, path: ['2'] },
    ],
  };
  const ignoreList = {
    rows: [
      { number: 'IGNORED-ASSY', from: 'CAD vs Item compare', sourceRow: 2 },
      { number: 'IGNORED-CHILD', from: 'CAD vs Item compare', sourceRow: 3 },
      { number: 'PART-B', from: 'CAD vs Item compare', sourceRow: 4 },
    ],
  };
  const idx = ignoreListCompare.buildIgnoreIndex(ignoreList);
  const res = compareAll([cad], im, { isIgnored: idx.isIgnored });

  check('IGNORED-ASSY removed from the missing tree',
    !res.missingRoots.some(n => n.item.number === 'IGNORED-ASSY'), res.missingRoots.map(n => n.item.number));
  check('CHILD-1 (not ignored) promoted to its own root finding instead of vanishing with its ignored parent',
    res.missingRoots.some(n => n.item.number === 'CHILD-1' && n.children.length === 0), res.missingRoots.map(n => n.item.number));
  check('IGNORED-CHILD (also ignored) does not reappear anywhere',
    !res.missingRoots.some(n => n.item.number === 'IGNORED-CHILD'));
  check('PART-A (not ignored) still reported missing', res.missingRoots.some(n => n.item.number === 'PART-A'));
  check('missingTotal excludes both ignored parts (counts CHILD-1 + PART-A only)', res.missingTotal === 2, res.missingTotal);

  check('PART-B qty mismatch suppressed', !res.qtyMismatches.some(m => m.number === 'PART-B'), res.qtyMismatches.map(m => m.number));
  check('PART-C qty mismatch still reported', res.qtyMismatches.some(m => m.number === 'PART-C'), res.qtyMismatches.map(m => m.number));

  check('ignoredFindings records all 3 suppressed occurrences', res.ignoredFindings.length === 3, res.ignoredFindings);
  check('ignoredFindings tags each occurrence with the right checkKey',
    res.ignoredFindings.find(f => f.number === 'IGNORED-ASSY').checkKey === 'missing' &&
    res.ignoredFindings.find(f => f.number === 'IGNORED-CHILD').checkKey === 'missing' &&
    res.ignoredFindings.find(f => f.number === 'PART-B').checkKey === 'qty',
    res.ignoredFindings);

  // Findings registry: ignored parts are filtered upstream (in compareAll),
  // so buildRegistry() needs no ignore-list awareness of its own — verifying
  // that here doubles as regression coverage for that design assumption.
  const reg = findings.buildRegistry({ result: res });
  check('findings registry never sees the ignored parts at all',
    !reg.byPn.has('IGNORED-ASSY') && !reg.byPn.has('IGNORED-CHILD') && !reg.byPn.has('PART-B'),
    Array.from(reg.byPn.keys()));
  check('findings registry still reports the non-ignored parts',
    reg.byPn.has('CHILD-1') && reg.byPn.has('PART-A') && reg.byPn.has('PART-C'),
    Array.from(reg.byPn.keys()));

  // No ignore predicate at all (the ordinary case) must reproduce the exact
  // same tree compareAll() already produced before this feature existed.
  const resNoIgnore = compareAll([cad], im, {});
  check('without an ignore predicate, nothing is suppressed (backward compatible)',
    resNoIgnore.missingRoots.some(n => n.item.number === 'IGNORED-ASSY') && resNoIgnore.ignoredFindings.length === 0,
    resNoIgnore.missingRoots.map(n => n.item.number));
}

console.log('\n== synthetic: folder classification recognizes the Ignore List filename ==');
{
  check('classifyFolderFile("IgnoreListHSG.xlsx") = ignore-list', folder.classifyFolderFile('IgnoreListHSG.xlsx') === 'ignore-list');
  check('classifyFolderFile("IgnoreListGFB.xls") = ignore-list', folder.classifyFolderFile('IgnoreListGFB.xls') === 'ignore-list');
  check('classifyFolderFile("ignorelist.xlsx") = ignore-list (case-insensitive)', folder.classifyFolderFile('ignorelist.xlsx') === 'ignore-list');
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
  // Part numbers are 7-* here because only this site's own "7-" parts are
  // checked; 1-/2-/3-/5- numbers are procured from other company locations
  // and are deliberately excluded (see MANUFACTURED_PART_RE).
  const aoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Material'],
    ['7-000-MACH-01', '-', 'Machine', 'desc', ''],
    ['7-999-PURCH-1', '1', 'Purchased part 1', '', 'AISI 304'],       // purchased, desc blank -> leniency, no flag
    ['7-999-PURCH-2', '2', '', '', 'AISI 304'],                       // purchased, both blank -> flagged
    ['7-100-MFG-1', '3', 'Manufactured part 1', '', 'AISI 304'],      // non-purchased, desc blank -> flagged
    ['7-100-MFG-2', '4', '', 'Manufactured part 2 desc', 'AISI 304'], // non-purchased, title blank -> flagged
    ['7-200-ASSY-1', '5', 'Assembly 1', 'desc', ''],                  // has a child below -> assembly, material excluded
    ['7-200-ASSY-CHILD', '5.1', 'Assy child', 'desc', ''],            // leaf, material blank -> flagged
    ['7-300-LEAF-2', '6', 'Leaf 2', 'desc', 'AISI 316L'],             // leaf, material present -> not flagged
    ['2-400-PROCURED', '7', '', '', ''],                              // non-7 prefix -> excluded from BOTH c5 and c6
    ['7-909-00001', '8', 'END OF LINE', '', ''],                      // EOL sentinel -> excluded from c6
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => aoa },
  });
  check('IM parsed with Material column', im.hasMaterial === true);
  const qc = imQc.runChecks(im);

  check('c5 flags exactly 3 rows', qc.c5.fail.length === 3, qc.c5.fail.map(f => f.number));
  // The END OF LINE row is just an ERP marker saying "this BOM is complete",
  // not a part, so every part-level check skips it (c3, c4, c5, c6, c7, c9).
  // Only Check 2, whose whole job is to validate that marker, looks at it.
  check('c5: END OF LINE marker row excluded despite a blank Description',
    !qc.c5.fail.some(f => f.number === '7-909-00001'), qc.c5.fail.map(f => f.number));
  check('c5: purchased part with one blank field is NOT flagged (leniency)',
    !qc.c5.fail.some(f => f.number === '7-999-PURCH-1'));
  check('c5: purchased part with both blank IS flagged as both-missing',
    qc.c5.fail.some(f => f.number === '7-999-PURCH-2' && f.kind === 'both-missing'));
  check('c5: non-purchased part missing description flagged correctly',
    qc.c5.fail.some(f => f.number === '7-100-MFG-1' && f.kind === 'description-missing'));
  check('c5: non-purchased part missing title flagged correctly',
    qc.c5.fail.some(f => f.number === '7-100-MFG-2' && f.kind === 'title-missing'));
  check('c5: non-7 procured part is excluded even with both fields blank',
    !qc.c5.fail.some(f => f.number === '2-400-PROCURED'));

  check('c6 flags exactly 1 row', qc.c6.applicable === true && qc.c6.fail.length === 1, qc.c6.fail);
  check('c6: assembly with children is excluded despite blank material',
    !qc.c6.fail.some(f => f.number === '7-200-ASSY-1'));
  check('c6: leaf part with blank material IS flagged',
    qc.c6.fail.some(f => f.number === '7-200-ASSY-CHILD'));
  check('c6: root row excluded (always counts as assembly)',
    !qc.c6.fail.some(f => f.number === '7-000-MACH-01'));
  check('c6: non-7 procured part excluded despite blank material',
    !qc.c6.fail.some(f => f.number === '2-400-PROCURED'));
  check('c6: END OF LINE sentinel row excluded despite blank material',
    !qc.c6.fail.some(f => f.number === '7-909-00001'));

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

  // Ignore List support: opts.isIgnored(pn, 'revision') suppresses a
  // mismatch and records it in ignoredFindings, mirroring compareAll()'s
  // filterIgnoredFlat (js/compare.js) — reused directly by revision-compare.js.
  const revIgnorePartY = revisionCompare.compareRevision([cadSource], im,
    { isIgnored: (pn, key) => key === 'revision' && pn === 'PART-Y' });
  check('opts.isIgnored suppresses the matching mismatch', revIgnorePartY.mismatches.length === 0, revIgnorePartY.mismatches);
  check('opts.isIgnored records the suppression in ignoredFindings',
    revIgnorePartY.ignoredFindings.length === 1 && revIgnorePartY.ignoredFindings[0].checkKey === 'revision' &&
    revIgnorePartY.ignoredFindings[0].number === 'PART-Y', revIgnorePartY.ignoredFindings);
  check('opts.isIgnored is only checked against the "revision" key',
    revisionCompare.compareRevision([cadSource], im, { isIgnored: (pn, key) => key === 'missing' }).mismatches.length === 1);
  check('without opts.isIgnored, behaves exactly as before (backward compatible)',
    rev.mismatches.length === 1 && rev.ignoredFindings.length === 0, rev.ignoredFindings);

  // js/app.js's "ignore bought-out parts (X-999-*)" button builds a predicate
  // combining imQc.PURCHASED_PART_RE with the file-based Ignore List, exactly
  // like this — verified directly against that shared regex so a change to
  // the convention is caught here too.
  const boughtOutCad = {
    kind: 'cad', hasRevision: true, fileName: 'inventor.xlsx',
    items: [{ number: '7-999-00001', revision: '9' }],
  };
  const boughtOutIm = { hasRevision: true, rows: [{ number: '7-999-00001', title: 'Bearing', revision: '5', sourceRow: 2, path: [] }] };
  const boughtOutIsIgnored = (pn, key) => key === 'revision' && imQc.PURCHASED_PART_RE.test(pn);
  const revBoughtOut = revisionCompare.compareRevision([boughtOutCad], boughtOutIm, { isIgnored: boughtOutIsIgnored });
  check('bought-out (X-999-*) revision mismatch suppressed by the PURCHASED_PART_RE-based predicate',
    revBoughtOut.mismatches.length === 0 && revBoughtOut.ignoredFindings.length === 1, revBoughtOut);
  check('same bought-out part, toggle off (no isIgnored), is reported normally',
    revisionCompare.compareRevision([boughtOutCad], boughtOutIm).mismatches.length === 1);

  // compareRevisionOrder: only ever asserted between two clean non-negative
  // integers — never guesses at placeholder text (Inventor Content-Center
  // defaults like "ANY"/"NONE" on purchased parts, verified real values).
  const CRO = revisionCompare.compareRevisionOrder;
  check('compareRevisionOrder: 3 < 9 -> -1', CRO('3', '9') === -1);
  check('compareRevisionOrder: 9 > 3 -> 1', CRO('9', '3') === 1);
  check('compareRevisionOrder: equal -> 0', CRO('4', '4') === 0);
  check('compareRevisionOrder: double-digit compares numerically, not lexicographically', CRO('10', '2') === 1);
  check('compareRevisionOrder: placeholder text on either side -> null',
    CRO('ANY', '3') === null && CRO('3', 'NONE') === null && CRO('-', '-') === null);
  check('compareRevisionOrder: blank on either side -> null', CRO('', '3') === null && CRO('3', '') === null);

  // rev.mismatches[0] is PART-Y: IM '3' vs CAD '9' -- IM is genuinely behind.
  check('imBehindCad set true when Item Master revision is numerically lower than CAD\'s',
    rev.mismatches[0].imBehindCad === true && rev.mismatches[0].staleDesign === 'Yes — CAD is newer', rev.mismatches[0]);

  // Opposite direction (IM ahead of CAD) and a non-comparable mismatch
  // (placeholder CAD text) must both stay unflagged -- this is the shape
  // most real mismatches actually take (verified: 8 of 11 real mismatches
  // found in this org's sample data are exactly the placeholder-text case).
  const orderIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: {
      sheet_to_json: () => [
        ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Revision'],
        ['MACH-01', '-', 'Machine', 'desc', '1'],
        ['PART-AHEAD', '1', 'IM ahead of CAD', 'desc', '5'],
        ['PART-NA', '2', 'CAD side not comparable', 'desc', '2'],
      ],
    },
  });
  const orderCad = {
    kind: 'cad', hasRevision: true, fileName: 'inventor.xlsx',
    items: [
      { number: 'PART-AHEAD', revision: '2' },  // IM '5' > CAD '2'
      { number: 'PART-NA', revision: 'ANY' },   // placeholder text, not comparable
    ],
  };
  const orderRes = revisionCompare.compareRevision([orderCad], orderIm);
  check('IM ahead of CAD: mismatch flagged but imBehindCad is false',
    orderRes.mismatches.find(m => m.number === 'PART-AHEAD').imBehindCad === false,
    orderRes.mismatches.find(m => m.number === 'PART-AHEAD'));
  check('non-comparable (placeholder CAD text): mismatch flagged but imBehindCad is false',
    orderRes.mismatches.find(m => m.number === 'PART-NA').imBehindCad === false &&
    orderRes.mismatches.find(m => m.number === 'PART-NA').staleDesign === '',
    orderRes.mismatches.find(m => m.number === 'PART-NA'));

  // cadRevisionByPn multi-source merge: a sparse/placeholder-only first
  // source must not shadow a genuinely comparable value on a second source.
  const sparseFirst = {
    kind: 'cad', hasRevision: true, fileName: 'sparse-first.xlsx',
    items: [{ number: 'PART-MERGE', revision: 'NONE' }],
  };
  const cleanSecond = {
    kind: 'cad', hasRevision: true, fileName: 'clean-second.xlsx',
    items: [{ number: 'PART-MERGE', revision: '7' }],
  };
  const mergeIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => [
      ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Revision'],
      ['MACH-01', '-', 'Machine', 'desc', '1'],
      ['PART-MERGE', '1', 'Merged across sources', 'desc', '3'],
    ] },
  });
  const mergedRes = revisionCompare.compareRevision([sparseFirst, cleanSecond], mergeIm);
  check('multi-source merge: placeholder-only first source upgraded by a comparable second source (IM 3 < merged CAD 7 -> genuinely stale)',
    mergedRes.mismatches.length === 1 && mergedRes.mismatches[0].cadRevision === '7' && mergedRes.mismatches[0].imBehindCad === true,
    mergedRes.mismatches);
  check('multi-source merge: cadSourceFileName reports every contributing source',
    mergedRes.cadSourceFileName === 'sparse-first.xlsx, clean-second.xlsx', mergedRes.cadSourceFileName);

  // A source that already has a comparable value first is NOT overridden by
  // a second source (first-comparable-wins, not last-wins).
  const cleanFirst = { kind: 'cad', hasRevision: true, fileName: 'clean-first.xlsx', items: [{ number: 'PART-MERGE2', revision: '1' }] };
  const otherSecond = { kind: 'cad', hasRevision: true, fileName: 'other-second.xlsx', items: [{ number: 'PART-MERGE2', revision: '9' }] };
  const mergeIm2 = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => [
      ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Revision'],
      ['MACH-01', '-', 'Machine', 'desc', '1'],
      ['PART-MERGE2', '1', 'First comparable value wins', 'desc', '1'],
    ] },
  });
  const mergedRes2 = revisionCompare.compareRevision([cleanFirst, otherSecond], mergeIm2);
  check('multi-source merge: first comparable value is not overridden by a later source',
    mergedRes2.mismatches.length === 0, mergedRes2.mismatches);

  // Single-source case is unchanged (existing tested behavior): cadSourceFileName
  // is just that one file's name, no join artifacts.
  check('single CAD source: cadSourceFileName has no comma/join artifact',
    rev.cadSourceFileName === 'inventor.xlsx', rev.cadSourceFileName);
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
  // groupImOnly now takes a positional parent index (Map<row, parentRow>).
  const rows = [
    { number: 'ASSY-1', title: 'Assembly', path: ['1'], sourceRow: 2 },
    { number: 'CHILD-A', title: 'Child A', path: ['1', '1'], sourceRow: 3 },
    { number: 'CHILD-B', title: 'Child B', path: ['1', '2'], sourceRow: 4 },
    { number: 'GRAND-C', title: 'Grandchild', path: ['1', '2', '1'], sourceRow: 5 },
    { number: 'LONE-1', title: 'Unrelated', path: ['2'], sourceRow: 6 },
  ];
  const parentOf = new Map([
    [rows[1], rows[0]], [rows[2], rows[0]], [rows[3], rows[2]],
  ]);
  const roots = groupImOnly(rows, parentOf);
  check('a whole flagged subassembly collapses to one root', roots.length === 2, roots.map(r => r.item.number));
  check('every row is still present in the tree (nothing dropped)',
    roots.reduce((n, r) => n + 1 + countDescendants(r), 0) === rows.length,
    roots.map(r => r.item.number + ':' + countDescendants(r)));
  check('nesting follows the hierarchy, not just the direct parent',
    countDescendants(roots[0]) === 3 && roots[1].item.number === 'LONE-1', roots);

  // No Row Order column -> nothing to group by; must degrade, not crash.
  const flat = groupImOnly(rows.map(r => ({ number: r.number, title: r.title, path: null })), new Map());
  check('degrades to one root per part when the export has no Row Order', flat.length === rows.length, flat.length);

  // A parent that is NOT itself flagged must not swallow its children.
  const presentParent = { number: 'PRESENT-PARENT', title: 'in CAD', path: ['9'] };
  const orphanRows = [{ number: 'CHILD-X', title: 'X', path: ['9', '1'], sourceRow: 7 }];
  check('a child whose parent is not flagged stays top-level',
    groupImOnly(orphanRows, new Map([[orphanRows[0], presentParent]])).length === 1);

  // Ordering must not matter: a child listed BEFORE its flagged ancestor used
  // to break the ancestor walk and silently emit the child as its own root.
  const late = [
    { number: 'CHILD-LATE', title: 'Child', path: ['3', '1'], sourceRow: 20 },
    { number: 'ASSY-LATE', title: 'Assembly', path: ['3'], sourceRow: 21 },
  ];
  const lateRoots = groupImOnly(late, new Map([[late[0], late[1]]]));
  check('child appearing before its flagged ancestor still groups under it',
    lateRoots.length === 1 && lateRoots[0].item.number === 'ASSY-LATE' && countDescendants(lateRoots[0]) === 1,
    lateRoots.map(r => r.item.number + ':' + countDescendants(r)));

  // Virtual-part anchors: a parent that IS in the CAD BOM (so not itself
  // Item-Master-only) can still absorb its orphaned children when passed as an
  // anchor — this is what stops a virtual subassembly's children scattering.
  const vChildren = [
    { number: 'V-CHILD-1', title: 'c1', path: ['4', '1'], sourceRow: 30 },
    { number: 'V-CHILD-2', title: 'c2', path: ['4', '2'], sourceRow: 31 },
  ];
  const vParent = { number: 'V-ASSY', title: 'Virtual assembly', path: ['4'], sourceRow: 29 };
  const vRoots = groupImOnly(vChildren, new Map([[vChildren[0], vParent], [vChildren[1], vParent]]),
    new Map([['V-ASSY', vParent]]));
  check('virtual-part anchor absorbs its orphaned children into one root',
    vRoots.length === 1 && vRoots[0].item.number === 'V-ASSY' && countDescendants(vRoots[0]) === 2,
    vRoots.map(r => r.item.number + ':' + countDescendants(r)));
  check('an anchor that absorbs nothing never appears as an empty finding',
    groupImOnly([], new Map(), new Map([['V-ASSY', vParent]])).length === 0);
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

console.log('\n== synthetic: LLDBO candidate detection (detectLldboCandidates) ==');
{
  const candAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity'],
    ['MACH-01', '-', 'Test Machine', 'root row, should never be a candidate even though it says Motor', '1 Each'],
    ['7-909-00001', '1', 'END OF LINE', 'sentinel row, should never be a candidate even though it says Seal', ''],
    ['7-999-AAA', '2', 'AC Motor 3PH', 'drive motor', '2 Each'],
    ['7-999-AAA', '3', 'AC Motor 3PH', 'drive motor', '2 Each'],          // same PN again at a different position -> dedup
    ['7-999-BBB', '4', 'Seal Kit', 'shaft seal', '5 Each'],
    ['7-234-CCC', '5', 'Motor mounting bracket', 'holds the motor in place', '1 Each'], // review tier: keyword, not 999-numbered
    ['7-234-DDD', '6', 'Pneum.Dichtung DN200', '', '3 Each'],              // review tier: German seal keyword, not 999-numbered
    ['7-100-ORD', '7', 'Ordinary bracket', 'no keyword here', '1 Each'],   // not a candidate at all
  ];
  const candIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, { utils: { sheet_to_json: () => candAoa } });

  const preview = lldboCompare.detectLldboCandidates(candIm, null, {});
  check('preview (no LLDBO loaded): not cross-checked', preview.crossChecked === false);
  check('preview: confident tier = AAA + BBB (999-numbered, deduped, root/EOL excluded)',
    preview.confident.length === 2 && preview.confident.some(c => c.number === '7-999-AAA') && preview.confident.some(c => c.number === '7-999-BBB'),
    preview.confident.map(c => c.number));
  check('preview: review tier = CCC + DDD (keyword match, not 999-numbered)',
    preview.review.length === 2 && preview.review.some(c => c.number === '7-234-CCC') && preview.review.some(c => c.number === '7-234-DDD'),
    preview.review.map(c => c.number));
  check('preview: ordinary part with no keyword never classified',
    !preview.confident.some(c => c.number === '7-100-ORD') && !preview.review.some(c => c.number === '7-100-ORD'));
  check('preview: root row never classified even though its description says "Motor"',
    !preview.confident.some(c => c.number === 'MACH-01') && !preview.review.some(c => c.number === 'MACH-01'));
  check('preview: END OF LINE sentinel never classified even though its description says "Seal"',
    !preview.confident.some(c => c.number === '7-909-00001') && !preview.review.some(c => c.number === '7-909-00001'));
  check('preview: 7-999-AAA at two BOM positions counted once (dedup)',
    preview.confident.filter(c => c.number === '7-999-AAA').length === 1);
  check('preview: matched keyword recorded on each entry',
    preview.confident.find(c => c.number === '7-999-AAA').keyword === 'motor' &&
    preview.review.find(c => c.number === '7-234-DDD').keyword === 'dichtung',
    preview.confident.concat(preview.review).map(c => [c.number, c.keyword]));
  check('preview: candidate entries carry Row # and parent location',
    !!preview.confident[0].sourceRow && preview.confident.every(c => 'parentNumber' in c), preview.confident);
  check('preview: candidate entries carry the Item Master Quantity',
    preview.confident.find(c => c.number === '7-999-AAA').quantity === '2 Each' &&
    preview.confident.find(c => c.number === '7-999-BBB').quantity === '5 Each',
    preview.confident.map(c => [c.number, c.quantity]));

  // Cross-checking: an LLDBO file tracking one confident-tier PN (AAA) and
  // one review-tier PN (CCC) — both should drop out once loaded, proving the
  // cross-check applies uniformly to both tiers.
  const candLldboAoa = [
    ['SR. No', 'PART NO', 'Item Description', 'Specifications', 'Make', 'Qty.', 'Remarks'],
    ['1', '7-999-AAA', 'tracked motor', 'spec', 'MAKE', '1 Nos.', ''],
    ['2', '7-234-CCC', 'tracked bracket (review tier, but already tracked anyway)', 'spec', 'MAKE', '1 Nos.', ''],
  ];
  const candLldbo = lldboParser.parse({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } }, { utils: { sheet_to_json: () => candLldboAoa } });

  const cross = lldboCompare.detectLldboCandidates(candIm, candLldbo, {});
  check('cross-checked: crossChecked true once an LLDBO file is supplied', cross.crossChecked === true);
  check('cross-checked: trackedCount = 2 (one confident + one review PN already tracked)', cross.trackedCount === 2, cross.trackedCount);
  check('cross-checked: confident tier now only BBB (AAA is already tracked)',
    cross.confident.length === 1 && cross.confident[0].number === '7-999-BBB', cross.confident);
  check('cross-checked: review tier now only DDD (CCC is already tracked)',
    cross.review.length === 1 && cross.review[0].number === '7-234-DDD', cross.review);

  // Ignore List support: opts.isIgnored(pn, 'lldboCandidate') suppresses a
  // candidate and records it in ignoredFindings, mirroring compareRevision()
  // — reuses js/compare.js's filterIgnoredFlat() directly, same as it does.
  const candIgnored = lldboCompare.detectLldboCandidates(candIm, null, { isIgnored: (pn, key) => key === 'lldboCandidate' && pn === '7-999-BBB' });
  check('opts.isIgnored suppresses the matching confident-tier candidate',
    candIgnored.confident.length === 1 && candIgnored.confident[0].number === '7-999-AAA', candIgnored.confident);
  check('opts.isIgnored records the suppression in ignoredFindings',
    candIgnored.ignoredFindings.length === 1 && candIgnored.ignoredFindings[0].checkKey === 'lldboCandidate' &&
    candIgnored.ignoredFindings[0].number === '7-999-BBB', candIgnored.ignoredFindings);
  check('opts.isIgnored is only checked against the "lldboCandidate" key',
    lldboCompare.detectLldboCandidates(candIm, null, { isIgnored: (pn, key) => key === 'missing' }).confident.length === 2);
  check('without opts.isIgnored, behaves exactly as before (backward compatible)',
    preview.confident.length === 2 && preview.ignoredFindings.length === 0, preview.ignoredFindings);

  // js/findings.js registry gating: only the confident tier, and only once
  // crossChecked — the review tier and the uncross-checked preview must
  // never feed "Parts needing attention".
  const regCross = findings.buildRegistry({ lldboCandidatesResult: cross });
  check('findings registry: crossChecked confident candidate becomes the primary finding',
    regCross.byPn.get('7-999-BBB') && regCross.byPn.get('7-999-BBB').primary.key === 'lldboCandidate',
    regCross.byPn.get('7-999-BBB'));
  check('findings registry: review tier never recorded, even when crossChecked',
    !regCross.byPn.has('7-234-DDD'));
  const regPreview = findings.buildRegistry({ lldboCandidatesResult: preview });
  check('findings registry: uncross-checked preview never feeds the registry, even with candidates present',
    regPreview.parts.length === 0, regPreview.parts);
}

console.log('\n== synthetic: LLDBO candidate detection skips assemblies (child parts present) ==');
{
  // 7-500-MOTASSY is an assembly (it has a child, MOTCHILD, at Row Order
  // "1.1") whose own title matches the "motor" keyword — it must be
  // skipped, since it is a container, not itself a purchasable part. The
  // real candidate is its child, which must still be classified normally.
  // 7-999-LEAFONLY has no children and must be unaffected by the new rule.
  const assyAoa = [
    ['Number', 'Row Order', 'Title (Item,CO)', 'Description (Item,CO)', 'Quantity'],
    ['MACH-02', '-', 'Test Machine 2', 'root', '1 Each'],
    ['7-500-MOTASSY', '1', 'Motor Cover Assembly', 'wraps the drive motor', '1 Each'],
    ['7-999-MOTCHILD', '1.1', 'AC Motor', 'the actual long-lead motor inside the assembly', '1 Each'],
    ['7-999-LEAFONLY', '2', 'Seal Kit', 'shaft seal, no children', '1 Each'],
  ];
  const assyIm = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, { utils: { sheet_to_json: () => assyAoa } });
  const res = lldboCompare.detectLldboCandidates(assyIm, null, {});
  check('an assembly row (has a child) is never classified, even though it matches a keyword',
    !res.confident.some(c => c.number === '7-500-MOTASSY') && !res.review.some(c => c.number === '7-500-MOTASSY'),
    res.confident.concat(res.review).map(c => c.number));
  check('the actual child part is classified normally (confident tier: 999-numbered)',
    res.confident.some(c => c.number === '7-999-MOTCHILD'), res.confident.map(c => c.number));
  check('a leaf part with no children is unaffected by the new rule',
    res.confident.some(c => c.number === '7-999-LEAFONLY'), res.confident.map(c => c.number));
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
    ['7-000-MACH-01', '-', 'Machine', 'desc', ''],
    ['7-200-ASSY-1', '1', 'An assembly', 'desc', ''],     // assembly (has children below) -> material not expected
    ['7-100-PART-A', '1.1', 'Matches (naming variant)', 'desc', '1.4301'],
    ['7-100-PART-B', '1.2', 'Genuine mismatch', 'desc', 'AISI 304'],
    ['7-999-00001', '1.3', 'Purchased, missing material', 'desc', ''],
    ['7-999-00002', '1.4', 'Purchased, mismatch vs CAD', 'desc', 'AISI 304'],
    ['2-100-PROCURED', '1.5', 'Procured elsewhere, mismatch', 'desc', 'AISI 304'],
    ['7-909-00001', '1.6', 'END OF LINE', '', '.'],       // sentinel, placeholder material
  ];
  const im = itemMasterParser.parse({ SheetNames: ['Sheet'], Sheets: { Sheet: {} } }, {
    utils: { sheet_to_json: () => imAoa },
  });

  const cadSource = {
    kind: 'cad', source: 'flat-xlsx', hasQty: false, hasMaterial: true, items: [
      { number: '7-100-PART-A', title: 'Part A', material: 'AISI 304', isAssembly: false },
      { number: '7-100-PART-B', title: 'Part B', material: 'AISI 304L', isAssembly: false }, // genuine grade difference
      { number: '7-999-00002', title: 'Purchased', material: 'AISI 316', isAssembly: false },
      { number: '2-100-PROCURED', title: 'Procured', material: 'AISI 316', isAssembly: false },
      { number: '7-909-00001', title: 'END OF LINE', material: 'AISI 316 L', isAssembly: false },
    ],
  };

  const noCadRes = materialCompare.compareMaterial([], im);
  check('not applicable with no CAD source carrying material', noCadRes.applicable === false, noCadRes.reason);
  check('bought-out list still populated when not applicable', noCadRes.boughtOut.length === 2, noCadRes.boughtOut.length);

  const res = materialCompare.compareMaterial([cadSource], im);
  check('applicable with a flat-xlsx CAD source', res.applicable === true);
  check('PART-A naming variant not flagged, only PART-B (genuine mismatch)',
    res.mismatches.length === 1 && res.mismatches[0].number === '7-100-PART-B', res.mismatches.map(m => m.number));
  check('non-7 procured part excluded from material mismatches',
    !res.mismatches.some(m => m.number === '2-100-PROCURED'), res.mismatches.map(m => m.number));
  check('END OF LINE sentinel excluded from material mismatches',
    !res.mismatches.some(m => m.number === '7-909-00001'), res.mismatches.map(m => m.number));
  check('purchased parts excluded from the mismatches list', !res.mismatches.some(m => /^\d-999-/.test(m.number)));

  check('bought-out panel lists both purchased parts', res.boughtOut.length === 2, res.boughtOut.map(b => b.number));
  const bo1 = res.boughtOut.find(b => b.number === '7-999-00001');
  check('bought-out: missing IM material flagged, no CAD data', bo1 && bo1.missingMaterial === true && bo1.cadMaterial === '', bo1);
  const bo2 = res.boughtOut.find(b => b.number === '7-999-00002');
  check('bought-out: CAD/IM mismatch flagged for purchased part', bo2 && bo2.mismatch === true && bo2.cadMaterial === 'AISI 316', bo2);
}

/* ---------------- real-sample baseline tests ---------------- */

const [cadPath, imPath, pdf723Path, pdf732Path, inv732Path, pdf733Path, im733Path, lldboPath, invBomPath, imBomMatPath, pdf726Path, invBom22819Path, imBom22819Path, imDiffOldPath, imDiffNewPath, imCandPath, lldboCandPath, imRev498Path, invBom498Path] = process.argv.slice(2);
let pdfjsLib = null;
try { pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js'); } catch (e) { /* npm install to enable PDF tests */ }

async function parsePdf(file) {
  const { pdfExtract } = require(path.join(rootDir, 'js/parsers/pdf-extract.js'));
  const buf = new Uint8Array(fs.readFileSync(file)).buffer;
  const grid = await pdfExtract.extractGrid(buf, { pdfjsLib });
  const parsed = cadLeveledParser.parse(grid.rows, { indents: grid.indents, pageOf: grid.pageOf, source: 'pdf' });
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
  // 49 -> 36: the 13 non-"7-" rows (six 2-*, seven 5-*) are procured from
  // other company locations and can't be fixed here, so they're excluded.
  check('HSG c5 title/desc: 36 flagged, all "7-" parts, all description-missing',
    hsgQc.c5.fail.length === 36 && hsgQc.c5.fail.every(f => f.kind === 'description-missing') &&
    hsgQc.c5.fail.every(f => /^7-/.test(String(f.number).trim())), hsgQc.c5.fail.length);
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
  // 7 -> 6: the END OF LINE marker row (7-909-00001, material ".") is no
  // longer compared — it is an ERP completeness marker, not a part.
  check('6 genuine (deduplicated, normalized) material mismatches on manufactured parts',
    matRes.mismatches.length === 6 && new Set(matRes.mismatches.map(m => m.number)).size === 6, matRes.mismatches.map(m => m.number));
  check('END OF LINE marker row is not a material mismatch',
    !matRes.mismatches.some(m => m.number === '7-909-00001'), matRes.mismatches.map(m => m.number));
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
  // 3 -> 2: the third overlap was a 5-* part flagged by Check 5, which now
  // skips non-"7-" procured parts.
  check('exactly 2 real parts are flagged by more than one check',
    overlaps.length === 2, overlaps.map(p => p.number + ':' + p.issues.map(i => i.key).join('+')));
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
    // 1033 -> 1032 flat (the Item Master's own root row is no longer a
    // finding) and 11 -> 10 roots for the same reason.
    check('vs 732020066: 1032 flat "In Item Master only" rows collapse to 10 roots',
      res732.imOnly.length === 1032 && res732.imOnlyRoots.length === 10,
      { flat: res732.imOnly.length, roots: res732.imOnlyRoots.length });
    check('vs 732020066: the Item Master root row is not reported as a finding',
      !res732.imOnly.some(r => Array.isArray(r.path) && r.path.length === 0));
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
    checkCleanPns('723 PDF', p723.items);

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
    checkCleanPns('733 PDF', p733.items);
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
  // 6 -> 5: the END OF LINE marker row is excluded (see above).
  check('726020768: 5 genuine material mismatches',
    matRes726.mismatches.length === 5 && new Set(matRes726.mismatches.map(m => m.number)).size === 5,
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
  // 380 -> 329. This export reuses a Row Order position for adjacent
  // siblings (40 such positions), which the old path-string ancestor lookup
  // resolved to the wrong branch, inflating those parts' rolled-up Item
  // Master quantities. Verified: all 51 rows that dropped out now have an
  // Item Master total exactly equal to CAD, and no part moved the other way.
  check('329 flat quantity mismatches after the ancestor-resolution fix',
    res22819.qtyMismatches.length === 329, res22819.qtyMismatches.length);
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
  console.log('\n== real samples: virtual parts + description check across all pairs ==');
  {
    const pairs = [
      ['726020768', invBomPath, imBomMatPath, ['2-303-64410', '7-305-20233']],
    ];
    for (const [label, cadP, imP, expected] of pairs) {
      if (!cadP || !imP) continue;
      const c = detect.parseCadFromWorkbook(XLSX.read(fs.readFileSync(cadP), { type: 'buffer' }), XLSX).ok;
      const m = detect.parseItemMasterFromWorkbook(XLSX.read(fs.readFileSync(imP), { type: 'buffer' }), XLSX);
      check(label + ': Inventor export carries the Thumbnail column', c.hasThumbnail === true);
      const v = virtualParts.detectVirtualParts([c], m, indexItemMaster);
      check(label + ': ' + expected.length + ' confirmed virtual parts (' + expected.join(', ') + ')',
        v.confirmed.length === expected.length &&
        expected.every(e => v.confirmed.some(x => x.number === e)),
        v.confirmed.map(x => x.number + ':' + x.childCount));
      check(label + ': nothing is merely "suspected" when Thumbnail is available', v.suspected.length === 0);
      // The virtual parts must actually absorb their orphans in the rollup.
      const grouped = compareAll([c], m, { virtualAnchorRows: v.anchorRows });
      const anchored = grouped.imOnlyRoots.filter(r => r.isAnchor);
      check(label + ': virtual parts act as rollup anchors for their orphaned children',
        anchored.length === expected.length && anchored.every(a => countDescendants(a) > 0),
        anchored.map(a => a.item.number + '+' + countDescendants(a)));
      const ungrouped = compareAll([c], m);
      check(label + ': without the anchors those children would scatter as separate roots',
        ungrouped.imOnlyRoots.length > grouped.imOnlyRoots.length,
        { withAnchors: grouped.imOnlyRoots.length, without: ungrouped.imOnlyRoots.length });
    }
  }

  if (invBom22819Path && imBom22819Path) {
    const c = detect.parseCadFromWorkbook(XLSX.read(fs.readFileSync(invBom22819Path), { type: 'buffer' }), XLSX).ok;
    const m = detect.parseItemMasterFromWorkbook(XLSX.read(fs.readFileSync(imBom22819Path), { type: 'buffer' }), XLSX);
    check('PN22819: this export has NO Thumbnail column', c.hasThumbnail === false);
    const v = virtualParts.detectVirtualParts([c], m, indexItemMaster);
    check('PN22819: the >=3-orphan-child fallback reports nothing here (its 2-child hose/cable rows are not virtual)',
      v.applicable === true && v.confirmed.length === 0 && v.suspected.length === 0,
      { confirmed: v.confirmed.length, suspected: v.suspected.length });
    const td = titleDescCompare.compareTitleDescription([c], m);
    check('PN22819: 31 description mismatches on this site\'s own manufactured parts',
      td.applicable === true && td.mismatches.length === 31, td.mismatches.length);
    check('PN22819: a real dimension error is caught (7-856-20360, "X 3 THK." vs "X 4 THK.")',
      td.mismatches.some(x => x.number === '7-856-20360'), td.mismatches.map(x => x.number).slice(0, 8));
    check('PN22819: spacing-only differences are not reported (7-056-23267 TANK - 300LTRS)',
      !td.mismatches.some(x => x.number === '7-056-23267'));
    check('PN22819: no purchased or non-7 part appears in the description mismatches',
      td.mismatches.every(x => /^7-/.test(x.number) && !/^\d-999-/.test(x.number)),
      td.mismatches.map(x => x.number));
  }

} else if (imDiffOldPath || imDiffNewPath) {
  console.log('\n(the Item Master diff regression needs both the older and newer Item Master paths)');
}

if (imCandPath && lldboCandPath) {
  console.log('\n== real samples: LLDBO candidate detection (SPN016326/PN21902 HSG PRO 200) ==');
  const candImWb = XLSX.read(fs.readFileSync(imCandPath), { type: 'buffer' });
  const candIm = detect.parseItemMasterFromWorkbook(candImWb, XLSX);
  const candLldboWb = XLSX.read(fs.readFileSync(lldboCandPath), { type: 'buffer' });
  const candLldbo = detect.parseLldboFromWorkbook(candLldboWb, XLSX);
  check('candidate-sample IM projectKey = SPN016326 / PN21902',
    candIm.projectKey && candIm.projectKey.spn === 'SPN016326' && candIm.projectKey.pn === 'PN21902', candIm.projectKey);

  const preview = lldboCompare.detectLldboCandidates(candIm, null, {});
  check('IM-only preview: not cross-checked', preview.crossChecked === false);
  check('IM-only preview: 70 unique candidates (29 confident + 41 review)',
    preview.confident.length === 29 && preview.review.length === 41,
    { confident: preview.confident.length, review: preview.review.length });

  const cross = lldboCompare.detectLldboCandidates(candIm, candLldbo, {});
  check('cross-checked: 3 already tracked (2 motors + 1 nozzle), 26 confident remain, review tier unaffected',
    cross.trackedCount === 3 && cross.confident.length === 26 && cross.review.length === 41, cross);
  check('cross-checked: the two tracked motors and the one tracked nozzle no longer flagged',
    !cross.confident.some(c => c.number === '7-999-11415') &&
    !cross.confident.some(c => c.number === '7-999-06562') &&
    !cross.confident.some(c => c.number === '7-999-07016'),
    cross.confident.map(c => c.number));
}

if (imRev498Path && invBom498Path) {
  console.log('\n== real samples: revision ordering (723020498) — no false positives ==');
  // This assembly's real revision mismatches were fully catalogued before
  // building the ordering check: 3 comparable (integer-vs-integer) — all 3
  // go the OPPOSITE direction (Item Master ahead of CAD) — plus several
  // non-comparable ones where the CAD side is Inventor Content-Center
  // placeholder text ("ANY"/"NONE"/"-"), never a real revision. None of the
  // 498 assembly's real mismatches should ever be marked imBehindCad.
  const im498 = detect.parseItemMasterFromWorkbook(XLSX.read(fs.readFileSync(imRev498Path), { type: 'buffer' }), XLSX);
  const invWb498 = XLSX.read(fs.readFileSync(invBom498Path), { type: 'buffer' });
  const invRes498 = detect.parseCadFromWorkbook(invWb498, XLSX);
  const invCad498 = invRes498 && invRes498.ok;
  check('723020498 Inventor BOM parsed', !!invCad498, invRes498);

  const rev498 = revisionCompare.compareRevision([invCad498], im498);
  check('723020498 revision check applicable', rev498.applicable === true, rev498);
  check('723020498: known integer-vs-integer mismatch 7-237-20065 (IM 2 > CAD 1) is NOT marked stale',
    !rev498.mismatches.some(m => m.number === '7-237-20065' && m.imBehindCad === true),
    rev498.mismatches.find(m => m.number === '7-237-20065'));
  check('723020498: no mismatch anywhere in this real sample is marked imBehindCad (documented: none exist in this data)',
    !rev498.mismatches.some(m => m.imBehindCad === true),
    rev498.mismatches.filter(m => m.imBehindCad).map(m => m.number));
  check('723020498: placeholder-text mismatches (e.g. CAD "ANY"/"NONE") stay non-stale, not crash',
    rev498.mismatches.filter(m => !/^\d+$/.test(m.cadRevision)).every(m => m.imBehindCad === false),
    rev498.mismatches.filter(m => !/^\d+$/.test(m.cadRevision)));
}

console.log(failures ? '\n' + failures + ' FAILURE(S)' : '\nall tests passed');
process.exit(failures ? 1 : 0);
