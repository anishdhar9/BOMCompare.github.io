/*
 * ecr-fill.js — builds this organization's ECR (Engineering Change Request)
 * sheet rows from a two-version Item Master diff, and writes them into the
 * company's fixed ECR Excel template.
 *
 * Reverse-engineered from a real filled sample (ECR_02.xlsx) and its blank
 * template (ECR_02__Copy.xlsx, vendored as vendor/ECR_template.xlsx):
 *
 *   Sheet1 table (header row 11, data from row 12): Line No. | Sub Assly
 *   Number With Rev | Item No. With Rev | Description | Old Quantity |
 *   New Quantity | Action | Reason Code | Remarks.
 *
 *   "Item No. With Rev" fuses the part number and a zero-padded 2-digit
 *   revision suffix ("7-330-20014-02" = part 7-330-20014 at revision 02).
 *   The Item Master carries Number and Revision as separate columns, so
 *   this composes the string.
 *
 *   A revision bump is an old-row/new-row PAIR, not one changed row: the
 *   old number+rev gets Old Qty=N, New Qty=0, Action="Drg. Obsolete"; the
 *   new revision's number+rev gets Old Qty=0, New Qty=N, Action="Drg.
 *   Revised" — both under the same parent.
 *
 *   Action is mechanically derivable from the diff (confirmed against
 *   Sheet2's exact dropdown list: Part Added / Part Deleted / Qty Changed /
 *   Drg. Obsolete / Drg. Revised). Reason Code (a 16-value business-
 *   justification list) is NOT derivable — it needs human judgment about
 *   WHY the BOM changed, so it is left blank for the person filing the ECR
 *   to fill in.
 *
 *   A part whose ONLY changes are Material, Title, Description, or State
 *   (no Quantity or Revision change) has no matching Action code in the
 *   template's vocabulary, so it is not turned into an ECR row — it is
 *   reported separately as "other changes" so it isn't silently dropped.
 *
 * Pure logic (no DOM) except fillEcrTemplate, which needs the XLSX
 * namespace to write into a loaded template workbook (injected, matching
 * this codebase's dependency-injection style).
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function normNumber(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim().toUpperCase();
  }

  // "2" -> "02", "10" -> "10", "" -> "00", "A" -> "A" (letter revisions pass
  // through unpadded — this org's samples only show numeric revisions).
  function padRevision(rev) {
    var s = rev === null || rev === undefined ? '' : String(rev).trim();
    if (!s) return '00';
    if (/^\d+$/.test(s)) return s.length >= 2 ? s : ('0' + s).slice(-2);
    return s;
  }

  function itemNoWithRev(row) {
    return row.number + '-' + padRevision(row.revision);
  }

  function qtyText(row) {
    if (row.quantityText) return row.quantityText;
    if (row.qty === null || row.qty === undefined) return '0';
    return String(row.qty);
  }

  // The immediate parent's "number-rev" composite, via the same Row Order
  // ancestor lookup indexItemMaster already builds (byPath). '' for a
  // top-level row or when the export has no Row Order column.
  function parentComposite(row, imIndex) {
    if (!Array.isArray(row.path) || row.path.length === 0) return '';
    const parent = imIndex.byPath.get(row.path.slice(0, -1).join('.'));
    if (!parent) return '';
    return itemNoWithRev(parent);
  }

  // imOld/imNew: parsed Item Master results. indexItemMaster: js/compare.js's
  // function (injected). Returns { rows: [...ECR table rows], otherChanges }.
  function diffForEcr(imOld, imNew, indexItemMaster) {
    const oldIndex = indexItemMaster(imOld);
    const newIndex = indexItemMaster(imNew);
    const oldPNs = new Set(oldIndex.byNumber.keys());
    const newPNs = new Set(newIndex.byNumber.keys());
    const allPNs = new Set([...oldPNs, ...newPNs]);

    const rows = [];
    const otherChanges = [];

    for (const pn of allPNs) {
      const oldRow = oldPNs.has(pn) ? oldIndex.byNumber.get(pn)[0] : null;
      const newRow = newPNs.has(pn) ? newIndex.byNumber.get(pn)[0] : null;

      if (oldRow && !newRow) {
        rows.push({
          subAssyNumberWithRev: parentComposite(oldRow, oldIndex),
          itemNoWithRev: itemNoWithRev(oldRow),
          description: oldRow.title || oldRow.description || '',
          oldQty: qtyText(oldRow), newQty: '0', action: 'Part Deleted',
        });
        continue;
      }
      if (!oldRow && newRow) {
        rows.push({
          subAssyNumberWithRev: parentComposite(newRow, newIndex),
          itemNoWithRev: itemNoWithRev(newRow),
          description: newRow.title || newRow.description || '',
          oldQty: '0', newQty: qtyText(newRow), action: 'Part Added',
        });
        continue;
      }
      // present in both
      const revChanged = padRevision(oldRow.revision) !== padRevision(newRow.revision);
      const qtyChanged = qtyText(oldRow) !== qtyText(newRow);
      if (revChanged) {
        rows.push({
          subAssyNumberWithRev: parentComposite(oldRow, oldIndex),
          itemNoWithRev: itemNoWithRev(oldRow),
          description: oldRow.title || oldRow.description || '',
          oldQty: qtyText(oldRow), newQty: '0', action: 'Drg. Obsolete',
        });
        rows.push({
          subAssyNumberWithRev: parentComposite(newRow, newIndex),
          itemNoWithRev: itemNoWithRev(newRow),
          description: newRow.title || newRow.description || '',
          oldQty: '0', newQty: qtyText(newRow), action: 'Drg. Revised',
        });
        continue;
      }
      if (qtyChanged) {
        rows.push({
          subAssyNumberWithRev: parentComposite(newRow, newIndex),
          itemNoWithRev: itemNoWithRev(newRow),
          description: newRow.title || newRow.description || '',
          oldQty: qtyText(oldRow), newQty: qtyText(newRow), action: 'Qty Changed',
        });
        continue;
      }
      // No Quantity or Revision change: check for a non-mapped field change
      // (Material / Title / Description / State) to report separately.
      const fields = [];
      const oldMat = (oldRow.material || '').trim(), newMat = (newRow.material || '').trim();
      if (oldMat !== newMat) fields.push('Material');
      const oldTitle = (oldRow.title || '').trim(), newTitle = (newRow.title || '').trim();
      if (oldTitle !== newTitle) fields.push('Title');
      const oldDesc = (oldRow.description || '').trim(), newDesc = (newRow.description || '').trim();
      if (oldDesc !== newDesc) fields.push('Description');
      const oldState = (oldRow.state || '').trim(), newState = (newRow.state || '').trim();
      if (oldState !== newState) fields.push('State');
      if (fields.length) {
        otherChanges.push({ number: newRow.number, title: newRow.title || '', fields: fields });
      }
    }

    const byComposite = function (a, b) { return normNumber(a.itemNoWithRev) < normNumber(b.itemNoWithRev) ? -1 : 1; };
    rows.sort(byComposite);
    otherChanges.sort(function (a, b) { return normNumber(a.number) < normNumber(b.number) ? -1 : 1; });

    return { rows: rows, otherChanges: otherChanges };
  }

  // templateWb: an XLSX workbook already read from vendor/ECR_template.xlsx
  // (XLSX.read/readFile — injected as the `XLSX` param). Mutates and returns
  // the same workbook with the table (from row 12) filled in. Header block
  // (PN, Project Status, Details Of Change, Engg. Comment, the document/
  // department checkboxes, the 5-Why block) is left blank — those are
  // narrative/judgment fields a two-file diff cannot fill in.
  function fillEcrTemplate(templateWb, ecrRows, XLSX) {
    const sheet = templateWb.Sheets['Sheet1'];
    const aoa = ecrRows.map(function (r, i) {
      return [i + 1, r.subAssyNumberWithRev, r.itemNoWithRev, r.description, r.oldQty, r.newQty, r.action, '', ''];
    });
    if (aoa.length) XLSX.utils.sheet_add_aoa(sheet, aoa, { origin: 'A12' });
    return templateWb;
  }

  return { ecrFill: { diffForEcr: diffForEcr, fillEcrTemplate: fillEcrTemplate, padRevision: padRevision, itemNoWithRev: itemNoWithRev } };
});
