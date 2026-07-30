# BOM Compare

BOM Compare is a static web app. It compares a **CAD BOM** (from Autodesk Inventor or Vault)
against the **Item Master BOM** (Vault items sent to the Navision ERP system).

The app finds parts that are in the CAD BOM but not in the Item Master. This is the usual sign
of a component left set to *BOM Reference* by mistake during modeling. A part set to
*BOM Reference* is never ordered.

**Privacy:** the app runs fully in your browser. No BOM data leaves your computer. There is no
server and no upload step.

## Usage

1. Open the deployed site. Or, open the file `index.html` directly on your computer.
2. Put the **CAD BOM** file or files in the left box. Put the **Item Master BOM** export in
   the right box. In Chrome or Edge, you can instead click **📁 Load from folder** and select a
   PNxxxx project folder one time, to load both files automatically (see "Load from folder"
   below).
3. Click **Compare BOMs**. Folder mode skips this step automatically.
4. Review the result tabs. Show or hide columns as needed. Use the filter box to find specific
   parts. Click **Download .xlsx** to save the result workbook.

Every result table, on screen and in every exported sheet, shows a **Row #**. This is the
row's position in the source file. When a row sits inside a BOM hierarchy, the table also
shows a **Parent Number** and a **Parent Title** for its immediate parent assembly. You can
find and understand a flagged row this way, without opening the source file by hand.

When the Item Master file loads, an **Item Master data quality** panel appears. This panel does
not need a CAD file. It has its own review list and downloadable report. An optional
**Long-Lead Parts (LLDBO)** panel sits below it, for checking early-released long-lead items.

### Load from folder (Chrome/Edge)

A browser cannot read a typed filesystem or NAS path. No browser has an API for this, for
security reasons. The closest match is the **File System Access API**.

Click **📁 Load from folder**. Select the PNxxxx project folder one time, in the native OS
picker. The app then does the rest automatically:

- It finds the CAD BOM (`Autodesk Vault- <assembly>.pdf`, or Vault's default name
  `Autodesk_Vault__<assembly>.iam.pdf`), the Item Master (`EBOM_<assembly>.xlsx`), the
  Inventor BOM export (`INVENTOR_BOM_<assembly>.xlsx`, an optional second CAD source), and
  the long-lead parts list, if present (`PNxxxx_LLDBO.xlsx`).
- It loads every file it finds, runs the comparisons, and writes
  `BOM-compare-results_<SPN>_<PN>.xlsx` back into the same folder. There is no download
  dialog. The Inventor BOM export and the LLDBO file are both optional. If they are absent,
  the app does not report a problem.
- If a required file is missing, or if there is more than one candidate file, the app leaves
  that side empty. Drop the file in manually with the normal upload boxes. The upload boxes
  always work, in every browser.

Firefox and Safari do not support the File System Access API. The button stays hidden in
these browsers, and a note explains why. The manual dropzones are the fallback, and they
work in every browser.

**Best results:** drop two CAD files together, the Vault multi-level BOM **PDF** and the
Inventor BOM **export (.xlsx)**. These two files add different data:

| | Vault "Uses" PDF | Inventor BOM export |
|---|---|---|
| Reference components | **included** | excluded |
| Virtual components (no CAD file) | missing | **included** |
| Quantities | no | **yes** |
| Material | no | **when the column is included** (Vault columns are user-configurable, so this varies by export) |
| Hierarchy | indentation | dotted Item numbers |

With both files loaded, the app also shows the **Reference items** tab. This tab lists every
component that is currently flagged *BOM Reference* in the model (the PDF list minus the
Inventor export list). Each entry is marked to show if it also reached the Item Master. Use
this tab to check "was this meant to be Reference?" for every entry.

When the Inventor export includes a Material column, the app also runs the
**Material: CAD vs Item Master** check. See the section below for more detail.

## Supported input formats

### CAD BOM (left box, accepts one or two files)

| Format | Hierarchy | Quantities | Reference components |
|---|---|---|---|
| Multi-level BOM **PDF** from the Vault web client (the "Uses" report) | exact (indentation) | no | included |
| **Inventor BOM export** (.xlsx: Item / Part Number / QTY / BOM Structure) | exact (dotted Item) | yes | excluded |
| Any leveled Excel or CSV file (Level or dotted Item column, plus Number) | exact | if a Qty column exists | depends on the source |
| Flat Vault **Excel** paste (no header row, depth-first list) | inferred | no | included |

Vault lets users choose which columns to show in an export. Because of this, the app finds
columns by header name and content, not by position. If detection fails, the app shows a
column-mapping step. In this step, you assign the Part Number, Qty, Level, and Title columns
yourself.

The PDF extractor is tuned against real Vault web-client reports (23-page and 64-page
samples). It finds the header line by keyword match on page 1. It rebuilds records from
wrapped lines (part numbers split like `7-320-` and `20066`). It skips `Attachments` and
`.stp` blocks. It reads hierarchy from filename indentation. The app does not support
scanned, image-only PDF files.

### Item Master BOM (right box)

This is the Vault or ERP item BOM grid export (`.xls` or `.xlsx`), with a `Number` header
column. A `Row Order` column (dotted position paths like `2.8.1`) enables hierarchy-aware
grouping and quantity roll-up. An `Item Qty` or `Quantity` column enables the quantity
comparison. A `State` column enables the item-state check (see Check 9 below). The app
matches `State` exactly, so it never reads the neighboring `File Link State` or
`State (Historical)` columns in its place. Those two columns describe the CAD file link, not
the item itself.

Some exports carry more than one quantity-like column, for example `Item Qty`, `Quantity`,
and `Unit Qty`, and these columns do not always agree. Real exports have shown `Item Qty`
stuck at a stale value of `1`, while `Quantity` correctly shows a multiplier like `4 Each`.
Because of this, the app always uses `Quantity` (the as-released quantity) for the
rolled-up quantity comparison. The app uses `Item Qty` only as a fallback, when a `Quantity`
column is not present at all.

The app finds columns by header keyword, not by position. Different exports, from different
plants or different PLM setups, do not always use the same header names or the same column
order. Because of this, the app recognizes common synonyms, for example "Part Number" or
"Item Number" for `Number`, "Qty" for `Item Qty`, and "Level" or "Position" for `Row Order`.
When more than one keyword could match a header, the app picks the longest matching keyword.
This way, a specific header always wins over a shorter, more general one, no matter the
column order.

Two exceptions apply on purpose. First, "PN" is **not** a synonym for `Number`. In this
organization, "PN" means Producer Number, one half of the project's SPN/PN key. It is never
a part number. Second, the app never reads a **per-unit** quantity column (for example
"QTY per Unit", "Unit Qty", or "Quantity Per Unit") as the `Quantity` or `Item Qty` column.
A per-unit column holds the quantity for one single unit of the parent, usually `1`, even
when the total Quantity is `4`. Reading it as the main quantity column would create many
false Check-3 mismatches, on every row with a quantity above one.

The same BOM can be exported by different people, with different Vault columns turned on.
Because of this, every result labels the exact source file or files it came from. This keeps
two exports of the same BOM traceable and consistent with each other.

## What the comparison does

- **Match key:** the part number, not case-sensitive.
- **Missing from Item Master** (red): CAD part numbers that are not in the Item Master.
  When a whole assembly is missing, it was likely set to *Reference*, so its entire BOM is
  expected to be absent too. In this case, only the assembly itself is flagged as
  actionable. Its child parts are grouped and collapsed beneath it, and are not part of the
  "findings needing action" count.
  - With a leveled CAD source, this grouping is exact.
  - With the flat Vault paste, the app infers the grouping. The export is a depth-first
    list, so the app finds the missing assembly's subtree boundary at the next item that
    the Item Master hierarchy already knows as a child of a present, enclosing assembly.
- **Reference items** (needs both CAD sources): components that are in the full structure
  source but not in the Inventor BOM export. These are grouped by subtree, and each one is
  marked with its Item Master status.
- **Quantity mismatches** (amber): the app compares the rolled-up total quantity per part
  number between the two BOMs (row quantity multiplied by parent assembly quantities,
  summed over every occurrence), and shows a per-parent breakdown. This check needs the
  Inventor BOM export, or any CAD source with a Qty column.
- **Quantity cascades** (red): a whole Item Master subtree whose direct children are ALL
  released at one clean, uniform ratio of their CAD quantity — usually one root-cause data
  error, not many. See "Quantity cascades" below.
- **Revision mismatches** (amber): the app compares CAD revision directly against the Item
  Master's Revision, for every shared part. See "Revision: CAD vs Item Master" below. This
  check needs a Revision column on both sides.
- **In Item Master only:** items whose number never appears in the CAD BOM. These are stale
  entries, or entries added by hand, and are worth a review. The app groups them under
  their parent assembly (see "Allied parts" below), so a whole subassembly missing from CAD
  counts as one finding.

## One finding per part

Each check runs on its own, so the same part can legitimately be flagged by more than one
check at the same time. For example, a part can be in the Item Master but not in CAD, so it
appears under "In Item Master only". If its `Quantity` and `Item Qty` also disagree, Check 3
flags it too. Read side by side, this looks like several different problems, but it is
really one part that needs one review.

Because of this, all findings go into a single registry (`js/findings.js`). The registry
groups findings by part number, and ranks them by severity. Each part then gets exactly
**one primary finding**, the most serious check that flagged it. Every other check that
also flagged the part becomes a muted cross-reference ("also in: ...") that points at the
primary finding. No finding is dropped. The detail sections still list every row. They
simply stop competing for attention.

Severity order, most serious first:

| # | Check |
|---|---|
| 1 | Sketch part (`7-333-`) in Item Master (release-blocking) |
| 2 | Obsolete or invalid state (release-blocking) |
| 3 | Missing from Item Master |
| 4 | Long-lead part missing from Item Master |
| 5 | Reference item |
| 6 | Quantity cascade (whole subtree released at one ratio) |
| 7 | Quantity mismatch (CAD vs Item Master) |
| 8 | Long-lead quantity mismatch |
| 9 | Revision mismatch vs CAD |
| 10 | Material mismatch vs CAD |
| 11 | In Item Master only |
| 12 | Item Master data-quality checks (Quantity vs Item Qty, Revision consistency, Material, Title/Description, Entity Icon, Producer, End of Line) |
| 13 | Not yet certified (`New` state), ranked low so it cannot hide a real finding |

A **Parts needing attention** table sits at the top of the page. It shows one row per part,
most serious issue first, with every issue found on that part. Click a row to jump to the
section that owns it.

### Allied parts

The app rolls up related parts, instead of repeating them. "Missing from Item Master"
already grouped a missing assembly's children beneath it. **"In Item Master only" now does
the same**, over the Item Master's own `Row Order` hierarchy. When a whole subassembly is
absent from the CAD BOM, that counts as one finding (the subassembly), not one finding per
part. On a real sample, this rollup reduced 1033 flat rows to 11 actionable roots, with
1022 children grouped underneath, and it did not drop a single row. Grouped children are
not part of the actionable count. Click "show N grouped" to reveal them.

Note: the per-occurrence rows inside a quantity mismatch's expander (its `cadBreakdown` and
`imBreakdown`) are *not* duplicates. They show "where this part is used", and stay as they
are.

### Quantity cascades

The plain "Quantity mismatches" check compares each part's total ROLLED-UP quantity. It
cannot see WHERE in the tree a discrepancy starts, so one bad assembly node can surface as
hundreds of separately-flagged descendants, all with the correct value at their own row but
the wrong value once their parent's multiplier is applied. Diagnosed on a real machine
(PN22819): every one of the 92 direct children of one Item Master assembly carried exactly
2× its CAD-required quantity, and that single error cascaded into 380 flagged rows.

The **Quantity cascades** tab and dashboard tile catch this directly. When ALL of an
assembly's mismatched direct children share one clean ratio (2×, 4×, 0.5×, and so on), the
app reports ONE finding for that assembly, with the ratio and child count
("92 of 92 children at 2×"), instead of flagging every downstream part on its own. Every
part in the affected subtree is grouped under that one finding — including parts that also
independently appear in the flat "Quantity mismatches" list — the same demotion the findings
registry already applies to other checks (see "One finding per part" above). Fix the one
root-cause row and the rest of the subtree's findings resolve with it.

This needs the same data the plain quantity check needs (a leveled CAD source with
quantities): it is not a separate optional file.

## Overview dashboard

A single **Overview** panel sits at the top of the page, above every detail section. It
shows one clickable tile per flag value: findings needing action, quantity mismatches,
material and revision mismatches vs CAD, and Item Master quality issues. Each tile is
color-coded (red, amber, or green, or gray when a check does not apply). Tiles appear as
data loads: the material, revision, and quality tiles appear the moment the Item Master
(and a CAD source) load, and the compare tiles appear once you run **Compare BOMs**. Click
a tile to jump straight to that section, which flashes to draw your eye. This gives anyone
unfamiliar with the BOM system the full picture at a glance, with no need to scroll to the
bottom to find a result. The dashboard, and each detail section, always shows the exact
source file or files being compared.

## Item Master data quality

These checks run on the Item Master alone, with no CAD BOM needed, the moment the file
loads. They catch manual edits made directly in Vault or the ERP system that do not agree
with other fields on the same row. This is a different failure mode from CAD-vs-BOM drift.

1. **Producer to Description match** (top-level row only): the row's Producer or Producer
   Number should appear in its Description.
2. **End of Line integrity**: the "END OF LINE" row should carry the organization's fixed
   part number and a whole-number Row Order.
3. **Quantity vs Item Qty**: these two columns should agree on every row. A mismatch means
   one column was edited without updating the other. This check only runs when the export
   carries *both* a `Quantity` column and an `Item Qty` column. With only one of the two
   columns present, there is nothing to cross-check, so the app reports "not applicable"
   instead of flagging every row.
4. **Entity Icon status**: this should read "Normal" on every row, when the column is
   present. The app reports "not applicable", instead of flagging every row, when the
   column is absent.
5. **Title and Description completeness**: every row should have both. Purchased or catalog
   parts (numbered `X-999-…`) are only flagged when *both* fields are missing, since a
   missing Title or Description alone is normal for catalog hardware. Every other part is
   flagged if *either* field is missing.
6. **Material completeness**: every non-assembly row should have a Material value. The app
   detects assemblies from the Row Order hierarchy (any row with children beneath it) and
   excludes them, since they legitimately carry no Material. Purchased or catalog parts
   (`X-999-…`) are excluded here too. See below for more on this.
7. **Revision consistency**: the same part number, used at more than one BOM position,
   should carry the same Revision everywhere it appears. A mismatch usually means one
   occurrence was updated to a newer released revision, and the others were not. The app
   compares Revision values directly, without normalizing them, since revisions are simple
   codes, not values that need a grade-equivalence lookup like material does.
8. **Sketch parts in the Item Master** (release-blocking): `7-333-…` numbers are the
   rough-sketch equivalent in this organization's 3D modeling process. They may exist in
   CAD, but **only ever as Reference**. One reaching the released Item Master is a serious
   release error. The app reports every hit with its Row #, Row Order, immediate parent,
   and the **full parent trail** (for example
   `ASSY-A (Top Assembly) › SUB-B (Sub Assembly)`), so you can trace exactly how the part
   got in. This check runs on the Item Master only, on purpose. Finding a sketch part in
   CAD is expected, and is not flagged. The banned prefix lives in one constant
   (`SKETCH_PART_RE` in `js/imqc.js`), so more prefixes can be added later.
9. **Item state** (release-blocking): a released BOM should hold Certified items only.
   `Obsolete`, `Invalid`, and `Phased Out` states are **errors**: the part was released
   against a dead revision. `New` is shown as a **warning**, not a failure, because genuine
   released exports do contain New rows (six in one real sample). A New finding turns its
   card amber, not red, and ranks low, so it cannot bury a genuinely missing part. This
   check reads the Item Master's `State` column, which the app carefully separates from the
   neighboring `File Link State` (Current or Out of Date) and `State (Historical)` columns.
   Those two columns describe the CAD file link, not the item.

Checks 8 and 9 each get their own summary box. Any hit on either check also raises a 🚨
banner above the detail sections.

A check can *pass* on its own terms, while a related cross-source comparison still finds a
real problem. For example, every part might have a Material value (check 6 passes), but
one of those values might not match the CAD model (see "Material: CAD vs Item Master"
below). In this case, the check's card turns amber ("OK, but see below") instead of plain
green, and it points at the section with the actual finding. Checks 6 and 7 both work this
way.

The downloadable QC report lists every flagged row for each check, with its Row # and
parent assembly, grouped by check.

### Material: CAD vs Item Master, and Bought-Out Parts

The app excludes purchased or catalog parts (`X-999-…`) from Check 6 above, and from this
comparison's main findings. On real data, 105 of 111 Check-6 flags were purchased parts
(bearings, wheels, cylinders, and similar hardware), where a blank material is often not a
real gap. These flags drowned out the genuine manufactured-part gaps underneath. Purchased
parts instead get their own **Bought-Out Parts** panel: a full, always-collapsed reference
list of every `X-999-…` part, with its Item Master material and CAD material side by side,
and mismatches or missing material marked. This panel is informational only, and does not
count toward any flagged total.

For manufactured (non-purchased) parts, a genuine **Material: CAD vs Item Master** check
compares material values. This check only runs once a loaded CAD source actually carries
material data. The multi-level PDF never carries material data. The flat Vault Excel paste
and the Inventor BOM export both *can* carry it, since Vault's exported columns are
user-configurable, so this depends on which columns were included at export time. The app
detects this per file (`hasMaterial`), instead of assuming it from the file format.

A raw text comparison does not work for material values. On real data, 38 of 518 shared
manufactured part numbers "differ" as plain text, and every one of the 38 is a
naming-convention variant of the same material, not a real error. Examples:

- `1.4301` and `AISI 304` (a DIN vs AISI grade designation).
- `AISI 316L` and `AISI 316 L` (spacing only).
- `SS316L` and `AISI316L` (an abbreviation).
- `A2`/`A4` and `AISI 304`/`AISI 316` (an ISO 3506 fastener-grade designation).
- `Silikon` and `Silikon/weiß/60°Shore` (CAD simply carries more detail).
- `Silicon`/`Silikon` and `Borosilicate`/`Borosilikat` (English and German spelling).
- `Silikon/transparent` and `Silikon transparent`, or `St-37` and `St 37` (a separator
  character only).
- `PTFE/weiß` and `PTFE weiss` (a German eszett versus its ASCII spelling).

The check normalizes values before it compares them. It normalizes case, spacing, and
separator punctuation. It applies a DIN, AISI, and ISO grade lookup. It transliterates
German characters. It also treats one value being a more detailed version of the other as a
match. The check keeps a grade's L-suffix significant on purpose (`304` versus `304L`, `316`
versus `316L` stay flagged as genuinely different materials), since this can be a real
weldability or corrosion choice, not just a formatting difference. On the same sample, this
normalization reduces the 38 false positives to 7 real, worth-reviewing differences.

### Revision: CAD vs Item Master

This check compares CAD revision against the Item Master's Revision, for every shared part.
It runs once a loaded CAD source carries revision data (`hasRevision`) and the Item Master
itself has a Revision column. Both are optional and user-configurable in Vault exports, the
same as material. Unlike material, revision values have no naming-convention ambiguity to
normalize away. On a real Vault "Uses" PDF export, revision values are plain integers ("0",
"1", "2", and so on). Other organizations may use letters ("A", "B"). Because of this, the
check compares values directly (trimmed, not case-sensitive), instead of using a
grade-equivalence lookup. A mismatch appears both as its own row in the "Revision
mismatches" summary card, and in this section's detail table, with the same Row # and
parent-assembly context as every other check.

## Long-Lead Parts (LLDBO)

This check is optional. Drop a `PNxxxx_LLDBO` file in its own dropzone, below the Item
Master QC panel. Long-lead direct-bought-out parts are released to procurement ahead of the
normal BOM release, to cover supplier lead times. Once both files are loaded, in either
order, this check confirms that each long-lead part actually reached the Item Master:

- **Missing from Item Master**: a long-lead part number that was released early, but never
  reached the Item Master. This is the process failure the check exists to catch, since it
  means the part may quietly never get ordered through the normal channel either.
- **Quantity mismatches**: the long-lead quantity, summed across all its LLDBO rows (the
  same catalog part can legitimately appear more than once, for example the same motor used
  in two different assemblies), should equal the Item Master's rolled-up total for that
  part.
- **Project key mismatch warning**: the LLDBO document's own header carries the SPN/PN
  project key (the same convention as the Item Master's), read from its "DBO Doc No" line.
  If this key does not match the loaded Item Master's key, a warning appears before any
  findings. The likely cause is the wrong pair of files, not a real inconsistency.
- Rows without a Part No yet (placeholders for items not yet specified, seen in the real
  sample document) are counted separately, and are not treated as findings.

The app does not route this file through the generic CAD file auto-detector, on purpose.
Its "PART NO" and "Qty." headers would otherwise match the CAD leveled-table keyword list,
and get misparsed as a CAD BOM.

## Development

There is no build step. The app is plain HTML, CSS, and JS. Libraries are vendored in
`vendor/` (`xlsx.full.min.js` is [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style),
a drop-in fork of SheetJS 0.18.5, and pdf.js handles PDF files), so the app works on
locked-down networks. npm packages are used by the Node tests only.

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
js/findings.js            cross-check findings registry — one primary finding per part (no DOM)
js/lldbo-compare.js       LLDBO vs Item Master comparison (no DOM)
js/folder.js              folder auto-load classification/scan (no DOM)
js/app.js                 UI wiring
```

`vendor/xlsx.full.min.js` is [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style), not
plain SheetJS. This was originally needed for a styled export sheet with real cell fills.
That sheet has since been removed, since it was redundant with the "Item Master QC" sheet's
per-check tables. The project kept xlsx-js-style anyway, since it is a smaller, fully
drop-in replacement (the same global `XLSX`, the same API), confirmed to still read the
legacy `.xls` Item Master format correctly, with no reason to revert.
`vendor/cpexcel.js` is its codepage-table dependency. Only Node uses it (the Node test
suite `require`s the same vendored file the browser loads). Browsers never fetch it.

### Tests

```
npm install     # once, for the PDF tests (pdfjs-dist, pinned to the vendored version)
node test/run-tests.mjs                                # synthetic tests only
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls   # + flat-export baseline
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls Vault_723.pdf Vault_732.pdf Inventor_732.xlsx \
  Vault_733.pdf Item_Master_733.xls PNxxxx_LLDBO.xlsx  # + PDF, reference & LLDBO baselines
```

Real BOM exports are not committed to this repository, since they may hold sensitive data.
Pass their file paths as arguments to run the full baseline assertions. The folder-auto-load
feature (the File System Access API) also has synthetic tests for its pure classification
and scan logic, in the same test suite. Its end-to-end browser behavior (the native picker
mocked through `page.addInitScript`, backed by a bridge to real files on disk) is not part
of this Node suite. See the commit history for the Playwright script used to verify it.

## Deployment: standalone hosting, not under a personal domain

All asset URLs are relative, so the files work when opened locally, or served from any
static host at any base path.

**Important GitHub Pages behavior:** if a user account has a custom domain set on its user
site (the `<user>.github.io` repository), **every project site under that account is served
under that domain too** (`https://custom.domain/<repo>/`). No repository setting overrides
this. To host this tool standalone, with no link to anyone's personal domain, the
repository must live under its **own GitHub organization**:

1. Create a free GitHub organization, for example **`BOMCompare`** (GitHub → **+** → *New
   organization* → Free plan).
2. Transfer this repository to the organization (repo **Settings → General → Danger Zone →
   Transfer ownership**), and name it **`BOMCompare.github.io`**. For an organization named
   `BOMCompare`, this repository name makes it the organization's *user site*, served at the
   root URL.
3. In the transferred repo, go to **Settings → Pages → Deploy from a branch**, and select
   the default branch, `/ (root)`.
4. The tool is then live at **`https://bomcompare.github.io/`**, with no connection to any
   personal account or domain. You can pick any other organization name. The URL then
   becomes `https://<orgname>.github.io/`, and you must rename the repository to
   `<orgname>.github.io` to match.

Do **not** set a custom domain on the organization, unless you want one. The plain
`*.github.io` URL is already standalone.
