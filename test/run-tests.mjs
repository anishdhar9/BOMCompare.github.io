/*
 * Node smoke tests for the parsers + comparison logic.
 *
 * Usage:
 *   node test/run-tests.mjs [path/to/CAD_Bom.xlsx path/to/Item_Master_BOM.xls]
 *
 * The two real sample exports are NOT committed (BOM data may be sensitive).
 * Without arguments only the synthetic tests run; with the two sample files
 * the full baseline assertions run as well.
 */
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';

const require = createRequire(import.meta.url);
const rootDir = path.dirname(path.dirname(fileURLToPath(import.meta.url)));

const XLSX = require(path.join(rootDir, 'vendor/xlsx.full.min.js'));
const { compare, countDescendants, indexItemMaster } = require(path.join(rootDir, 'js/compare.js'));
const { itemMasterParser } = require(path.join(rootDir, 'js/parsers/itemmaster.js'));
const { cadFlatParser } = require(path.join(rootDir, 'js/parsers/cad-flat-xlsx.js'));
const { cadLeveledParser } = require(path.join(rootDir, 'js/parsers/cad-leveled.js'));
const { detect } = require(path.join(rootDir, 'js/parsers/detect.js'));

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

/* ---------------- real-sample baseline tests ---------------- */

const [cadPath, imPath] = process.argv.slice(2);
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

  console.log('\n== real samples: comparison baseline ==');
  const res = compare(cad, im);
  check('CAD unique count', res.cadUniqueCount === 1231, res.cadUniqueCount);
  check('IM unique count', res.imUniqueCount === 1076, res.imUniqueCount);
  check('missing unique PNs = 183', res.missingTotal === 183, res.missingTotal);
  check('qty comparison disabled', res.qtyMismatches === null);
  check('IM-only = 28', res.imOnly.length === 28, res.imOnly.length);

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

console.log(failures ? '\n' + failures + ' FAILURE(S)' : '\nall tests passed');
process.exit(failures ? 1 : 0);
