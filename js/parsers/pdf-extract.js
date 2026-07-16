/*
 * pdf-extract.js — reconstructs the row/column grid from a text-based Vault
 * multi-level BOM PDF (100+ pages) using pdf.js, so cad-leveled.js can parse
 * it like any other table.
 *
 * Method: extract positioned text runs per page -> cluster runs into lines by
 * y-coordinate -> find the header line by known column keywords -> derive
 * column x-boundaries from the header cells -> assign runs of each data line
 * to columns. Repeated page headers are skipped; wrapped description lines
 * (no part number, no qty) are appended to the previous row. The x-position
 * of each row's first run is kept so indentation-based levels can be derived
 * when the PDF encodes hierarchy by indenting.
 *
 * NOTE: written generically against Vault's report layout; tune against a
 * real sample PDF (none was available when this was built).
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // matchField comes from cad-leveled.js (shared keyword table)
  function getMatchField() {
    const bc = typeof self !== 'undefined' ? self.BOMCompare : (typeof global !== 'undefined' ? global.BOMCompare : {});
    if (bc && bc.cadLeveledParser) return bc.cadLeveledParser.matchField;
    if (typeof module !== 'undefined') return require('./cad-leveled.js').cadLeveledParser.matchField;
    throw new Error('cad-leveled.js must be loaded before pdf-extract.js');
  }

  function clusterLines(runs) {
    // runs: {str, x, y, w, h}
    const sorted = runs.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    const lines = [];
    for (const run of sorted) {
      const tol = Math.max(2, (run.h || 8) * 0.5);
      const line = lines.length ? lines[lines.length - 1] : null;
      if (line && Math.abs(line.y - run.y) <= tol) {
        line.runs.push(run);
        line.y = (line.y * (line.runs.length - 1) + run.y) / line.runs.length;
      } else {
        lines.push({ y: run.y, runs: [run] });
      }
    }
    for (const line of lines) line.runs.sort(function (a, b) { return a.x - b.x; });
    return lines;
  }

  // Merge a line's runs into visual cells: a new cell starts when the gap to
  // the previous run exceeds `gap` points.
  function lineCells(line, gap) {
    const cells = [];
    for (const run of line.runs) {
      const s = run.str;
      if (!s || !s.trim()) continue;
      const prev = cells.length ? cells[cells.length - 1] : null;
      if (prev && run.x - prev.end <= gap) {
        prev.text += (run.x - prev.end > 0.5 ? ' ' : '') + s;
        prev.end = run.x + run.w;
      } else {
        cells.push({ text: s, x: run.x, end: run.x + run.w });
      }
    }
    for (const c of cells) c.text = c.text.replace(/\s+/g, ' ').trim();
    return cells.filter(function (c) { return c.text !== ''; });
  }

  function fieldForHeaderText(text, matchField) {
    const normalized = String(text || '').replace(/\s+/g, ' ').trim();
    let f = matchField(normalized);
    // Vault web-client BOM PDFs label the file column as "Name". In generic
    // spreadsheets "Name" can mean title, but in this layout the data under
    // it is the Inventor filename (.iam/.ipt), which is valuable for assembly
    // detection and indentation-based hierarchy.
    if (!f && /^name$/i.test(normalized)) f = 'file';
    if (f === 'title' && /^name$/i.test(normalized)) f = 'file';
    // The rightmost header can arrive as "Part", "Number", "Part Number", or
    // as one text item containing a newline between the words.
    if (!f && /\bpart\b/i.test(normalized) && /\bnumber\b/i.test(normalized)) f = 'number';
    if (!f && /^part$/i.test(normalized)) f = 'part-fragment';
    if (!f && /^number$/i.test(normalized)) f = 'number-fragment';
    return f;
  }

  function findHeaderCells(cells, matchField) {
    let matches = 0;
    const fields = [];
    for (const c of cells) {
      let f = matchField(c.text);
      // Vault web-client BOM PDFs label the file column as "Name". In generic
      // spreadsheets "Name" can mean title, but in this layout the data under
      // it is the Inventor filename (.iam/.ipt), which is valuable for assembly
      // detection and indentation-based hierarchy.
      if (!f && /^name$/i.test(c.text)) f = 'file';
      if (f === 'title' && /^name$/i.test(c.text)) f = 'file';
      // The rightmost header commonly wraps as two visual lines: "Part" on the
      // header line and "Number" below it. Treat the "Part" cell as the part
      // number column so the table is recognized from the first header line.
      if (!f && /^part$/i.test(c.text)) f = 'number';
      fields.push(f);
      if (f) matches++;
    }
    return matches >= 2 ? fields : null;
  }

  function completePartNumber(s) {
    s = (s || '').replace(/\s+/g, '').trim();
    return /^[A-Z0-9]+(?:-[A-Z0-9]+)+$/i.test(s) && !/-$/.test(s);
  }

  function joinCellContinuation(prev, next, isNumber) {
    if (!prev) return next;
    if (!next) return prev;
    if (isNumber) return /-$/.test(prev) ? prev + next : prev + ' ' + next;
    return prev + ' ' + next;
  }

  /**
   * @param data ArrayBuffer of the PDF
   * @param opts { pdfjsLib, onProgress(page, total) }
   * @returns { rows, indents, pages, warnings } — rows[0] is the header row
   */
  async function extractGrid(data, opts) {
    opts = opts || {};
    const pdfjsLib = opts.pdfjsLib || (typeof self !== 'undefined' ? self.pdfjsLib : null);
    if (!pdfjsLib) throw new Error('pdf.js is not loaded');
    const matchField = getMatchField();

    const doc = await pdfjsLib.getDocument({ data: data }).promise;
    const warnings = [];
    const rows = [];
    const indents = [null]; // aligned with rows; header row has no indent
    let header = null;      // { labels:[], centers:[], boundaries:[], fields:[], numberCol, qtyCol, textCols:Set }

    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const content = await page.getTextContent();
      const runs = content.items.map(function (it) {
        return {
          str: it.str,
          x: it.transform[4],
          y: it.transform[5],
          w: it.width || 0,
          h: Math.abs(it.transform[3]) || 8,
        };
      });
      const lines = clusterLines(runs);

      for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
        const line = lines[lineIndex];
        let cells = lineCells(line, 7);
        if (!cells.length) continue;

        if (!header) {
          const nextCells = lineIndex + 1 < lines.length ? lineCells(lines[lineIndex + 1], 7) : [];
          const headerCells = mergedHeaderCells(cells, nextCells, matchField);
          const fields = buildHeaderFromCells(headerCells, matchField);
          if (fields && fields.indexOf('number') !== -1) {
            cells = headerCells;
            const centers = cells.map(function (c) { return (c.x + c.end) / 2; });
            const boundaries = [];
            for (let i = 1; i < centers.length; i++) boundaries.push((centers[i - 1] + centers[i]) / 2);
            const textCols = new Set();
            fields.forEach(function (f, i) {
              if (f === 'title' || f === 'description' || f === 'file' || f === 'number') textCols.add(i);
            });
            header = {
              labels: cells.map(function (c, i) {
                if (fields[i] === 'number' && /^part$/i.test(c.text)) return 'Part Number';
                if (fields[i] === 'file' && /^name$/i.test(c.text)) return 'File';
                return c.text;
              }),
              centers: centers,
              boundaries: boundaries,
              fields: fields,
              numberCol: fields.indexOf('number'),
              qtyCol: fields.indexOf('qty'),
              textCols: textCols,
            };
            rows.push(header.labels.slice());
          }
          continue; // ignore everything above/without a header
        }

        // skip repeated page headers and second-line wrapped header labels such as
        // the standalone "Number" underneath "Part" in Vault's "Part Number".
        const hf = findHeaderCells(cells, matchField);
        if (hf && hf.indexOf('number') !== -1) continue;
        if (cells.length === 1 && /^number$/i.test(cells[0].text)) continue;

        // assign cells to columns by center x
        const rowArr = new Array(header.labels.length).fill('');
        const colsHit = [];
        for (const c of cells) {
          const cx = (c.x + c.end) / 2;
          let col = 0;
          while (col < header.boundaries.length && cx > header.boundaries[col]) col++;
          rowArr[col] = rowArr[col] ? rowArr[col] + ' ' + c.text : c.text;
          if (colsHit.indexOf(col) === -1) colsHit.push(col);
        }

        const hasNumber = rowArr[header.numberCol] !== '';
        const hasQty = header.qtyCol >= 0 && rowArr[header.qtyCol] !== '';

        const onlyContinuationCols = colsHit.every(function (c) { return header.textCols.has(c); });
        const numberLooksComplete = !hasNumber || completePartNumber(rowArr[header.numberCol]);
        if (rows.length > 1 && onlyContinuationCols && (!hasNumber || !numberLooksComplete)) {
          const prev = rows[rows.length - 1];
          for (const c of colsHit) {
            prev[c] = joinCellContinuation(prev[c], rowArr[c], c === header.numberCol);
          }
          continue;
        }

        if (!hasNumber && !hasQty) {
          // wrapped continuation of the previous row — but only if all its
          // content sits in text columns; otherwise it's a footer/noise line
          const allText = colsHit.every(function (c) { return header.textCols.has(c); });
          if (allText && rows.length > 1) {
            const prev = rows[rows.length - 1];
            for (const c of colsHit) prev[c] = prev[c] ? prev[c] + ' ' + rowArr[c] : rowArr[c];
          }
          continue;
        }
        if (!hasNumber) continue; // qty without number: noise

        rows.push(rowArr);
        indents.push(cells[0].x);
      }

      if (opts.onProgress) opts.onProgress(p, doc.numPages);
    }

    if (!header) {
      warnings.push('No table header with a recognizable part-number column was found in the PDF.');
      return { rows: [], indents: [], pages: doc.numPages, warnings: warnings };
    }
    return { rows: rows, indents: indents, pages: doc.numPages, warnings: warnings };
  }

  return { pdfExtract: { extractGrid: extractGrid, clusterLines: clusterLines, lineCells: lineCells } };
});
