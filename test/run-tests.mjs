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
const { compare, compareAll, countDescendants, indexItemMaster, normNumber } = require(path.join(rootDir, 'js/compare.js'));
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

/* ---------------- real-sample baseline tests ---------------- */

const [cadPath, imPath, pdf723Path, pdf732Path, inv732Path, pdf733Path, im733Path] = process.argv.slice(2);
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

if (pdf723Path && imPath && cadPath) {
  if (!pdfjsLib) {
    console.log('\n(pdfjs-dist not installed — run `npm install` to enable the PDF baseline tests)');
  } else {
    console.log('\n== real samples: Vault PDF 7-230-20509 (64 pages) vs Item Master ==');
    const { parsed: p723 } = await parsePdf(pdf723Path);
    check('723 PDF parsed 1820 records', p723.items.length === 1820, p723.items.length);
    check('723 PDF has levels, no qty', p723.hasLevels === true && p723.hasQty === false);
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

console.log(failures ? '\n' + failures + ' FAILURE(S)' : '\nall tests passed');
process.exit(failures ? 1 : 0);
