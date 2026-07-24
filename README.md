# BOM Compare

A static web app that compares a **CAD BOM** (from Autodesk Inventor / Vault) against the
**Item Master BOM** (Vault items → Navision ERP) and highlights parts that exist in CAD but
never made it into the Item Master — the classic symptom of a component accidentally left as
*BOM Reference* while modelling, which means it silently never gets ordered.

**Privacy:** everything runs client-side in the browser. No BOM data ever leaves your
computer — there is no server and no upload.

## Usage

1. Open the deployed site (or just open `index.html` locally).
2. Drop the **CAD BOM** file(s) on the left box, the **Item Master BOM** export on the right
   box — or, in Chrome/Edge, click **📁 Load from folder** and pick a PNxxxx project folder
   once to do both automatically (see below).
3. Click **Compare BOMs** (skipped automatically in folder mode).
4. Review the result tabs, hide/show columns as needed, filter, and download the
   result workbook via **Download .xlsx**.

Every result table — on screen and in every exported sheet — carries a **Row #** (the row's
position in the source file) and, wherever the row sits inside a BOM hierarchy, a **Parent
Number**/**Parent Title** (its immediate parent assembly), so a flagged row can be found and
placed in context without cross-referencing the source file by hand.

The moment the Item Master loads, an **Item Master data quality** panel also appears —
independent of the CAD side — with its own review list and downloadable report. An optional
**Long-Lead Parts (LLDBO)** panel sits below it for checking early-released long-lead items.

### Load from folder (Chrome/Edge)

A browser cannot read a typed filesystem/NAS path — there's no API for that, in any browser,
for security reasons. The closest (and only) equivalent is the **File System Access API**:
click **📁 Load from folder**, pick the PNxxxx project folder once in the native OS picker,
and everything else is automatic:

- finds the CAD BOM (`Autodesk Vault- <assembly>.pdf`, or Vault's own default naming
  `Autodesk_Vault__<assembly>.iam.pdf`), the Item Master (`EBOM_<assembly>.xlsx`), the
  Inventor BOM export (`INVENTOR_BOM_<assembly>.xlsx`, optional second CAD source), and — if
  present — the long-lead parts list (`PNxxxx_LLDBO.xlsx`) inside it
- loads whatever it finds, runs the comparisons, and writes `BOM-compare-results_<SPN>_<PN>.xlsx`
  straight back into that same folder — no download dialog. The Inventor BOM export and the
  LLDBO file are both optional; their absence isn't treated as a problem
- if a required file is missing or there's more than one candidate, that side is left for you
  to drop in manually via the normal upload boxes, which always work regardless

Firefox and Safari don't support this API — the button is hidden there and a note explains
why; the manual dropzones are the fallback and work in every browser.

**Best results:** drop *two* CAD files together — the Vault multi-level BOM **PDF** and the
Inventor BOM **export (.xlsx)**. They are complementary:

| | Vault "Uses" PDF | Inventor BOM export |
|---|---|---|
| Reference components | **included** | excluded |
| Virtual components (no CAD file) | missing | **included** |
| Quantities | no | **yes** |
| Material | no | **when the column is included** (Vault columns are user-configurable, so this varies by export) |
| Hierarchy | indentation | dotted Item numbers |

With both, the app additionally shows the **Reference items** tab: the exact list of
components currently flagged *BOM Reference* in the model (PDF minus Inventor export),
each marked whether it made it into the Item Master — the direct review list for
"was this meant to be reference?". And when the Inventor export includes a Material
column, it also activates the **Material: CAD vs Item Master** check (see below).

## Supported input formats

### CAD BOM (left box — accepts one or two files)

| Format | Hierarchy | Quantities | Reference components |
|---|---|---|---|
| Multi-level BOM **PDF** from the Vault web client ("Uses" report) | exact (indentation) | no | included |
| **Inventor BOM export** (.xlsx: Item / Part Number / QTY / BOM Structure) | exact (dotted Item) | yes | excluded |
| Any leveled Excel/CSV (Level or dotted Item column + Number) | exact | if a Qty column exists | depends on source |
| Flat Vault **Excel** paste (headerless, depth-first listing) | inferred | no | included |

Vault lets users choose which columns are visible in exports, so columns are auto-detected
by header name and content. If detection fails, the app shows a column-mapping step where
you assign the Part Number / Qty / Level / Title columns manually.

The PDF extractor is tuned against real Vault web-client reports (23- and 64-page samples):
header line found by keyword co-occurrence on page 1, records reassembled from wrapped
lines (part numbers split like `7-320-` / `20066`), `Attachments`/`.stp` blocks skipped,
hierarchy from filename indentation. Scanned (image-only) PDFs are not supported.

### Item Master BOM (right box)

The Vault/ERP item BOM grid export (`.xls`/`.xlsx`) with a `Number` header column.
`Row Order` (dotted position paths like `2.8.1`) enables hierarchy-aware grouping and
quantity roll-up; `Item Qty`/`Quantity` enables quantity comparison.

Columns are located by header keyword, not position — different exports (different plants,
different PLM configurations) don't all spell or order headers the same way, so common
synonyms are recognized (e.g. "Part Number"/"Item Number" for `Number`, "Qty" for `Item Qty`,
"Level"/"Position" for `Row Order`). One exception, deliberately: "PN" is **not** treated as a
`Number` synonym — in this organization's convention "PN" means Producer Number (half of the
project's SPN/PN key), never a part number.

## What the comparison does

- **Match key:** part number, case-insensitive.
- **Missing from Item Master** (red): CAD part numbers absent from the Item Master.
  When a whole assembly is missing (it was set as *Reference*, so its entire BOM is expected
  to be absent), only the assembly is flagged as actionable; its child parts are grouped and
  collapsed beneath it and excluded from the "findings needing action" count.
  - With a leveled CAD source the grouping is exact.
  - With the flat Vault paste it is inferred: the export is a depth-first listing, so a
    missing assembly's subtree is bounded by the next item that the Item Master hierarchy
    knows as a child of an enclosing present assembly.
- **Reference items** (needs both CAD sources): components in the full structure source but
  not in the Inventor BOM export, grouped by subtree, each annotated with Item Master
  presence.
- **Quantity mismatches** (amber): rolled-up total quantity per part number
  (row qty × parent assembly quantities, summed over all occurrences) compared between the
  two BOMs, with a per-parent breakdown. Requires the Inventor BOM export (or any CAD
  source with a Qty column).
- **Revision mismatches** (amber): CAD revision compared directly against the Item Master's
  Revision for every shared part — see "Revision: CAD vs Item Master" below. Requires a
  Revision column on both sides.
- **In Item Master only:** items whose number never appears in the CAD BOM — stale or
  manually added entries worth reviewing.

## Item Master data quality

Runs on the Item Master alone — no CAD BOM needed — the moment it loads. Catches manual
edits made directly in Vault/ERP that don't agree with other fields on the same row, a
different failure mode than CAD-vs-BOM drift:

1. **Producer ↔ Description match** (top-level row only): its Producer/Producer Number
   should appear in its Description.
2. **End of Line integrity**: the "END OF LINE" row should carry the organization's fixed
   part number and a whole-number Row Order.
3. **Quantity vs Item Qty**: these two columns should agree on every row — a mismatch means
   one was edited without updating the other.
4. **Entity Icon status**: should read "Normal" everywhere, when that column is present
   (reports "not applicable" rather than false-flagging every row when it's absent).
5. **Title/Description completeness**: every row should have both. Purchased/catalog parts
   (numbered `X-999-…`) are only flagged when *both* are missing, since one alone is normal
   for catalog hardware; every other part is flagged if *either* is missing.
6. **Material completeness**: every non-assembly row should have a Material. Assemblies are
   detected from the Row Order hierarchy (any row with children under it) and excluded — they
   legitimately don't carry one. Purchased/catalog parts (`X-999-…`) are excluded here too —
   see below.
7. **Revision consistency**: the same part number, used at more than one BOM position, should
   carry the same Revision everywhere it appears — a mismatch usually means one assembly's
   occurrence was updated to a newer released revision and the others weren't picked up.
   Values are compared directly (not normalized); revisions are simple codes, not something
   needing a grade-equivalence lookup like material.

A check can *pass* on its own terms while a related cross-source comparison still finds a
real problem — e.g. every part has a Material value (check 6 passes), but one of those values
doesn't match the CAD model (see "Material: CAD vs Item Master" below). Rather than showing
plain green in that case, the check's card turns amber ("OK, but see below") and points at
the section with the actual finding. Checks 6 and 7 both do this.

The downloadable QC report lists every flagged row per check (with its Row # and parent
assembly), grouped by check.

### Material: CAD vs Item Master, and Bought-Out Parts

Purchased/catalog parts (`X-999-…`) are excluded from Check 6 above and from this comparison's
main findings — verified on real data that 105 of 111 Check-6 flags were purchased parts
(bearings, wheels, cylinders…) where a blank material is often not a real gap, drowning out
the genuine manufactured-part gaps underneath. They get their own **Bought-Out Parts**
panel instead: a full, always-collapsed reference listing of every `X-999-…` part with its
Item Master material and CAD material side by side, mismatches/missing material marked —
informational, never counted toward any flagged total.

For manufactured (non-purchased) parts, a genuine **Material: CAD vs Item Master** check
compares material values — only active once a loaded CAD source actually carries material
data. The multi-level PDF never does; the flat Vault Excel paste and the Inventor BOM export
both *can*, since Vault's exported columns are user-configurable (whichever visible columns
were selected when the export was made) — the app detects this per file (`hasMaterial`)
rather than assuming it from the file format. A raw string comparison is unusable: verified that of 518 shared
manufactured part numbers in a real sample, 38 "differ" as plain text and every one of them
is a naming-convention variant of the same material (`1.4301` = `AISI 304`, DIN vs AISI
grade designation; `AISI 316L` vs `AISI 316 L`, spacing only; `SS316L` vs `AISI316L`,
abbreviation; `Silikon` vs `Silikon/weiß/60°Shore`, CAD simply carrying more descriptive
detail; `Silicon`/`Silikon` and `Borosilicate`/`Borosilikat`, English/German spelling) — not
real errors. The check normalizes before comparing (case/spacing, a DIN↔AISI grade lookup,
one value being a more detailed qualifier of the other) but deliberately keeps a grade's
L-suffix significant (`304` vs `304L`, `316` vs `316L` stay flagged as genuinely different),
since that can be a real weldability/corrosion spec choice, not just formatting. On that same
sample this reduces the 38 false positives to 7 real, worth-reviewing differences.

### Revision: CAD vs Item Master

Compares CAD revision against the Item Master's Revision for every shared part — active once
a loaded CAD source actually carries revision data (`hasRevision`) and the Item Master itself
has a Revision column; both are optional/user-configurable in Vault exports, the same
situation as material. Unlike material, there's no naming-convention ambiguity to normalize
away — verified on a real Vault "Uses" PDF export that revision values are plain integers
("0", "1", "2"…); other organizations may use letters ("A"/"B") — so this is a direct value
comparison (trimmed, case-insensitive), not a grade-equivalence lookup. A mismatch is shown
both as its own row in the "Revision mismatches" summary card and in this section's detail
table, with the same Row #/parent-assembly context as every other check.

## Long-Lead Parts (LLDBO)

Optional — drop a `PNxxxx_LLDBO` file (its own dropzone, below the Item Master QC panel).
Long-lead direct-bought-out parts are released to procurement ahead of the normal BOM
release, to cover supplier lead times; this checks that each one actually made it into the
Item Master once both files are loaded (in either order):

- **Missing from Item Master**: a long-lead part number that was released early but never
  showed up in the Item Master — the process failure this check exists to catch, since it
  means the part may quietly never get ordered through the normal channel either.
- **Quantity mismatches**: the long-lead quantity — summed across all its LLDBO rows, since
  the same catalog part legitimately appears more than once (e.g. the same motor used in two
  different assemblies) — should equal the Item Master's rolled-up total for that part.
- **Project key mismatch warning**: the LLDBO document's own header carries the SPN/PN
  project key (same convention as the Item Master's), read from its "DBO Doc No" line. If it
  doesn't match the loaded Item Master's key, a warning appears before any findings, since
  the more likely explanation is the wrong pair of files was loaded, not real inconsistencies.
- Rows without a Part No yet (not-yet-specified placeholders — seen in the real sample
  document) are counted separately and not treated as findings.

Deliberately **not** routed through the generic CAD file auto-detector: its "PART NO"/"Qty."
headers would otherwise false-match the CAD leveled-table keyword list and get misparsed as
a CAD BOM.

## Development

No build step; plain HTML/CSS/JS. Libraries are vendored in `vendor/`
(`xlsx.full.min.js` is [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style), a
drop-in fork of SheetJS 0.18.5; pdf.js for PDFs) so the app works on locked-down networks.
npm packages are used by the Node tests only.

```
js/compare.js             pure comparison + grouping + qty roll-up (no DOM)
js/parsers/itemmaster.js  Item Master Excel parser
js/parsers/cad-flat-xlsx.js  flat Vault paste parser
js/parsers/cad-leveled.js    leveled table parser (PDF grid / Excel / CSV)
js/parsers/pdf-extract.js    pdf.js Vault-report table reconstruction
js/parsers/lldbo.js          LLDBO (long-lead parts) list parser
js/parsers/detect.js         format detection / role validation
js/imqc.js                Item Master data-quality checks (no DOM)
js/material-compare.js    material CAD-vs-IM comparison + bought-out parts (no DOM)
js/revision-compare.js    revision CAD-vs-IM comparison (no DOM)
js/lldbo-compare.js       LLDBO vs Item Master comparison (no DOM)
js/folder.js              folder auto-load classification/scan (no DOM)
js/app.js                 UI wiring
```

`vendor/xlsx.full.min.js` is [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style)
rather than plain SheetJS — originally needed for a styled export sheet with real cell fills
(since removed as redundant with the "Item Master QC" sheet's per-check tables), kept because
it's a smaller, fully drop-in replacement (same global `XLSX`, same API, confirmed it still
reads the legacy `.xls` Item Master format correctly) with no reason to revert.
`vendor/cpexcel.js` is its codepage-table dependency — only used in Node (the Node test suite
`require`s the same vendored file the browser loads); browsers never fetch it.

### Tests

```
npm install     # once, for the PDF tests (pdfjs-dist, pinned to the vendored version)
node test/run-tests.mjs                                # synthetic tests only
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls   # + flat-export baseline
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls Vault_723.pdf Vault_732.pdf Inventor_732.xlsx \
  Vault_733.pdf Item_Master_733.xls PNxxxx_LLDBO.xlsx  # + PDF, reference & LLDBO baselines
```

Real BOM exports are not committed (potentially sensitive data); pass their paths as
arguments to run the full baseline assertions. The folder-auto-load feature (File System
Access API) additionally has synthetic tests for its pure classification/scan logic in the
same suite; its end-to-end browser behavior (native picker mocked via
`page.addInitScript`, backed by a bridge to real files on disk) isn't part of this Node
suite — see the commit history for the Playwright script used to verify it.

## Deployment — standalone hosting, not under a personal domain

All asset URLs are relative, so the files work opened locally or served from any static
host at any base path.

**Important GitHub Pages behaviour:** if a user account has a custom domain configured on
its user site (`<user>.github.io` repository), **every project site of that account is
served under that domain too** (`https://custom.domain/<repo>/`). No repository setting
overrides this. To host this tool standalone — not linked to anyone's personal domain — the
repository must live under its **own GitHub organization**:

1. Create a free GitHub organization, e.g. **`BOMCompare`** (GitHub → **+** → *New
   organization* → Free plan).
2. Transfer this repository to the organization (repo **Settings → General → Danger Zone →
   Transfer ownership**) and name it **`BOMCompare.github.io`** — for an organization named
   `BOMCompare`, that repository name makes it the organization's *user site*, served at
   the root URL.
3. In the transferred repo: **Settings → Pages → Deploy from a branch** → default branch,
   `/ (root)`.
4. The tool is then live at **`https://bomcompare.github.io/`** with no connection to any
   personal account or domain. (Pick any other org name and the URL becomes
   `https://<orgname>.github.io/` — the repo must be renamed `<orgname>.github.io` to match.)

Do **not** configure a custom domain on the organization unless you want one; the plain
`*.github.io` URL is already standalone.
