/*
 * folder.js — pure helpers for the "load from folder" feature.
 *
 * A browser cannot read a typed filesystem/NAS path — there is no such API,
 * in any browser, for security reasons. The only mechanism that can read
 * *and write back into* a folder the user names is the File System Access
 * API (`window.showDirectoryPicker()`): one click opens a native OS picker,
 * the user browses to the folder (a mapped NAS drive or UNC share both
 * work, since it's the OS's own picker), and the returned
 * FileSystemDirectoryHandle can then be traversed/written to
 * programmatically with no further dialogs. Chrome/Edge only (not
 * Firefox/Safari) — the app's actual picker call lives in app.js since it
 * needs the real `window`; this module holds the pure, Node-testable parts.
 *
 * Expected folder contents (this organization's convention):
 *   "Autodesk Vault- <assembly>.pdf"  (or Vault's own default naming,
 *      "Autodesk_Vault__<assembly>.iam.pdf")     -> CAD BOM (Vault PDF)
 *   "EBOM_<assembly>.xlsx"                        -> Item Master BOM
 *   "INVENTOR_BOM_<assembly>.xlsx"                -> Inventor BOM export (optional second
 *                                                     CAD source — carries quantities and,
 *                                                     when the column was included, material)
 *   "PN<number>_LLDBO*.xlsx"                      -> long-lead parts list (optional)
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.BOMCompare = Object.assign(root.BOMCompare || {}, factory());
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function classifyFolderFile(name) {
    if (!name) return null;
    if (/^PN\d+_LLDBO/i.test(name) && /\.xlsx?$/i.test(name)) return 'lldbo'; // checked before item-master: distinct prefix, no overlap risk
    if (/autodesk[ _-]*vault/i.test(name) && /\.pdf$/i.test(name)) return 'cad-pdf';
    if (/^inventor[ _-]*bom/i.test(name) && /\.xlsx?$/i.test(name)) return 'inventor-bom';
    if (/^ebom/i.test(name) && /\.xlsx?$/i.test(name)) return 'item-master';
    return null;
  }

  // Iterates anything shaped like a FileSystemDirectoryHandle (a real one,
  // or a test mock exposing an async `values()` yielding
  // {kind, name, getFile()} entries) and buckets file entries by
  // classifyFolderFile. Kept separate from any DOM/picker call so it's
  // testable with a plain mock, no browser needed.
  async function scanFolder(directoryHandle) {
    const found = { 'cad-pdf': [], 'item-master': [], 'lldbo': [], 'inventor-bom': [] };
    for await (const entry of directoryHandle.values()) {
      if (entry.kind !== 'file') continue;
      const kind = classifyFolderFile(entry.name);
      if (kind) found[kind].push(entry);
    }
    return found;
  }

  return { folder: { classifyFolderFile: classifyFolderFile, scanFolder: scanFolder } };
});
