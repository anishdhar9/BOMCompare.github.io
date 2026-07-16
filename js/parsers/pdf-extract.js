/*
 * pdf-extract.js — reconstructs the row/column grid from a Vault web-client
 * multi-level BOM PDF (the "Uses" report) using pdf.js, so cad-leveled.js can
 * parse it like any other table.
 *
 * Layout facts (validated against real 23- and 64-page Vault reports):
 *  - The column header (Name / Revision / State / Title / Description /
 *    Part Number) appears once, on page 1 only; column x positions vary per
 *    document, so bands are derived from that header line.
 *  - Page 1 has a preamble above the table that repeats words like "State",
 *    so the header must be found as a LINE where several keywords co-occur,
 *    never by first keyword match.
 *  - "Part Number" wraps: "Part" sits above the header line, "Number" below.
 *  - One logical record spans 2–3 visual lines: the filename line is the
 *    anchor; part numbers wrap as "7-320-" / "20066" (join at trailing "-"),
 *    titles/descriptions wrap as ordinary text.
 *  - Hierarchy is encoded by the filename's indentation (~10–11pt per level).
 *  - "Attachments" blocks (label + .stp file rows, possibly wrapped) must be
 *    swallowed. Records never straddle a page break.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var HEADER_KEYS = ['Name', 'Revision', 'State', 'Title', 'Description'];
  var FILE_RE = /\.(iam|ipt)$/i;
  var ATTACH_RE = /\.stp$/i;

  // Find the header line on page 1: group items into y-lines and take the
  // first (topmost) line where >=3 of the known header keywords co-occur.
  function findHeader(items) {
    var lines = new Map();
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var key = Math.round(it.y / 4) * 4;
      if (!lines.has(key)) lines.set(key, []);
      lines.get(key).push(it);
    }
    var keys = Array.from(lines.keys()).sort(function (a, b) { return b - a; });
    for (var k = 0; k < keys.length; k++) {
      var lineItems = lines.get(keys[k]);
      var hits = HEADER_KEYS.filter(function (name) {
        return lineItems.some(function (it) { return it.str === name; });
      });
      if (hits.length < 3) continue;

      var cols = {};
      var lineY = -Infinity;
      for (var j = 0; j < lineItems.length; j++) {
        var li = lineItems[j];
        if (HEADER_KEYS.indexOf(li.str) !== -1 && cols[li.str] === undefined) {
          cols[li.str] = li.x;
          if (li.y > lineY) lineY = li.y;
        }
      }
      // "Part Number" wraps just above/below the header line
      for (var m = 0; m < items.length; m++) {
        var pi = items[m];
        if (Math.abs(pi.y - lineY) <= 12 && /^(Part|Number|Part Number)$/.test(pi.str)) {
          cols['Part Number'] = Math.min(cols['Part Number'] !== undefined ? cols['Part Number'] : Infinity, pi.x);
        }
      }
      if (cols['Name'] === undefined) continue;
      return { cols: cols, y: lineY };
    }
    return null;
  }

  /**
   * @param data ArrayBuffer of the PDF
   * @param opts { pdfjsLib, onProgress(page, total) }
   * @returns { rows, indents, pages, warnings } — rows[0] is the header row;
   *          indents[i] is the filename x-offset of rows[i] (levels derive
   *          from it in cad-leveled.js), null for the header row.
   */
  async function extractGrid(data, opts) {
    opts = opts || {};
    var pdfjsLib = opts.pdfjsLib || (typeof self !== 'undefined' ? self.pdfjsLib : null);
    if (!pdfjsLib) throw new Error('pdf.js is not loaded');

    var doc = await pdfjsLib.getDocument({ data: data }).promise;
    var warnings = [];

    var pages = [];
    for (var p = 1; p <= doc.numPages; p++) {
      var page = await doc.getPage(p);
      var content = await page.getTextContent();
      var items = [];
      for (var i = 0; i < content.items.length; i++) {
        var t = content.items[i];
        var s = (t.str || '').trim();
        if (s) items.push({ str: s, x: t.transform[4], y: t.transform[5] });
      }
      pages.push(items);
      if (opts.onProgress) opts.onProgress(p, doc.numPages);
    }

    var header = findHeader(pages[0]);
    if (!header) {
      warnings.push('No Vault "Uses" table header (Name / Revision / State / Title / Description) was found on page 1 of the PDF.');
      return { rows: [], indents: [], pages: doc.numPages, warnings: warnings };
    }
    if (header.cols['Part Number'] === undefined) {
      warnings.push('No "Part Number" column found in the PDF — part numbers will be derived from file names, which is unreliable for renamed items.');
    }

    // column bands, sorted by x; label "Name" as "File" (its data is the
    // Inventor filename) so cad-leveled.js maps it to the file field.
    var cols = Object.keys(header.cols).map(function (name) {
      return { name: name === 'Name' ? 'File' : name, x: header.cols[name] };
    }).sort(function (a, b) { return a.x - b.x; });
    var colIndexOf = function (x) {
      var idx = -1;
      for (var c = 0; c < cols.length; c++) {
        if (x >= cols[c].x - 6) idx = c; else break;
      }
      return idx;
    };
    var fileCol = -1, numberCol = -1;
    cols.forEach(function (c, idx) {
      if (c.name === 'File') fileCol = idx;
      if (c.name === 'Part Number') numberCol = idx;
    });

    var rows = [cols.map(function (c) { return c.name; })];
    var indents = [null];

    for (var pn = 0; pn < pages.length; pn++) {
      var anchors = [];
      var rest = [];
      for (var n = 0; n < pages[pn].length; n++) {
        var it = pages[pn][n];
        // skip the page-1 preamble, the header line and its wrapped fragments
        if (pn === 0 && it.y >= header.y - 12) continue;
        var ci = colIndexOf(it.x);
        if (ci === fileCol && FILE_RE.test(it.str)) {
          anchors.push({ x: it.x, y: it.y, file: it.str, sink: false, cells: [] });
        } else if (ci === fileCol && (ATTACH_RE.test(it.str) || it.str === 'Attachments')) {
          // .stp attachment rows and the "Attachments" label swallow the
          // fragments around them
          anchors.push({ x: it.x, y: it.y, file: it.str, sink: true, cells: [] });
        } else {
          rest.push({ str: it.str, x: it.x, y: it.y, col: ci });
        }
      }
      // every non-anchor item belongs to the vertically nearest anchor
      for (var r = 0; r < rest.length; r++) {
        var frag = rest[r];
        var best = null;
        for (var a = 0; a < anchors.length; a++) {
          if (best === null || Math.abs(anchors[a].y - frag.y) < Math.abs(anchors[best].y - frag.y)) best = a;
        }
        if (best !== null) anchors[best].cells.push(frag);
      }
      for (var a2 = 0; a2 < anchors.length; a2++) {
        var anc = anchors[a2];
        if (anc.sink) continue;
        var rowArr = new Array(cols.length).fill('');
        rowArr[fileCol] = anc.file;
        // group fragments per column, top-to-bottom then left-to-right
        anc.cells.sort(function (f, g) { return g.y - f.y || f.x - g.x; });
        for (var f2 = 0; f2 < anc.cells.length; f2++) {
          var cell = anc.cells[f2];
          if (cell.col < 0) continue;
          var prev = rowArr[cell.col];
          if (cell.col === numberCol) {
            rowArr[cell.col] = !prev || /-$/.test(prev) ? prev + cell.str : prev + ' ' + cell.str;
          } else {
            rowArr[cell.col] = prev ? prev + ' ' + cell.str : cell.str;
          }
        }
        if (numberCol >= 0 && !rowArr[numberCol]) {
          // no part-number fragments (hidden column / detached row): derive
          // from the filename — "7-999-00044I00.ipt" -> "7-999-00044"
          rowArr[numberCol] = anc.file.replace(/\.(iam|ipt)$/i, '').replace(/I\d+$/i, '');
        }
        rows.push(rowArr);
        indents.push(anc.x);
      }
    }

    if (rows.length === 1) {
      warnings.push('A Vault table header was found but no component rows below it.');
    }
    return { rows: rows, indents: indents, pages: doc.numPages, warnings: warnings };
  }

  return { pdfExtract: { extractGrid: extractGrid, findHeader: findHeader } };
});
