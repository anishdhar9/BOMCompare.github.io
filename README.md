# BOM Compare

A static web app that compares a **CAD BOM** (from Autodesk Inventor / Vault) against the
**Item Master BOM** (Vault items → Navision ERP) and highlights parts that exist in CAD but
never made it into the Item Master — the classic symptom of a component accidentally left as
*BOM Reference* while modelling, which means it silently never gets ordered.

**Privacy:** everything runs client-side in the browser. No BOM data ever leaves your
computer — there is no server and no upload.

## Usage

1. Open the deployed site (or just open `index.html` locally).
2. Drop the **CAD BOM** file(s) on the left box, the **Item Master BOM** export on the right box.
3. Click **Compare BOMs**.
4. Review the result tabs, hide/show columns as needed, filter, and download the
   result workbook via **Download .xlsx**.

**Best results:** drop *two* CAD files together — the Vault multi-level BOM **PDF** and the
Inventor BOM **export (.xlsx)**. They are complementary:

| | Vault "Uses" PDF | Inventor BOM export |
|---|---|---|
| Reference components | **included** | excluded |
| Virtual components (no CAD file) | missing | **included** |
| Quantities | no | **yes** |
| Hierarchy | indentation | dotted Item numbers |

With both, the app additionally shows the **Reference items** tab: the exact list of
components currently flagged *BOM Reference* in the model (PDF minus Inventor export),
each marked whether it made it into the Item Master — the direct review list for
"was this meant to be reference?".

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
- **In Item Master only:** items whose number never appears in the CAD BOM — stale or
  manually added entries worth reviewing.

## Development

No build step; plain HTML/CSS/JS. Libraries are vendored in `vendor/`
(SheetJS for Excel, pdf.js for PDFs) so the app works on locked-down networks.
npm packages are used by the Node tests only.

```
js/compare.js            pure comparison + grouping + qty roll-up (no DOM)
js/parsers/itemmaster.js Item Master Excel parser
js/parsers/cad-flat-xlsx.js  flat Vault paste parser
js/parsers/cad-leveled.js    leveled table parser (PDF grid / Excel / CSV)
js/parsers/pdf-extract.js    pdf.js Vault-report table reconstruction
js/parsers/detect.js         format detection / role validation
js/app.js                UI wiring
```

### Tests

```
npm install     # once, for the PDF tests (pdfjs-dist, pinned to the vendored version)
node test/run-tests.mjs                                # synthetic tests only
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls   # + flat-export baseline
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls Vault_723.pdf Vault_732.pdf Inventor_732.xlsx
                                                       # + PDF & reference baselines
```

Real BOM exports are not committed (potentially sensitive data); pass their paths as
arguments to run the full baseline assertions.

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
