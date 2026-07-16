# BOM Compare

A static web app that compares a **CAD BOM** (from Autodesk Inventor / Vault) against the
**Item Master BOM** (Vault items → Navision ERP) and highlights parts that exist in CAD but
never made it into the Item Master — the classic symptom of a component accidentally left as
*BOM Reference* while modelling, which means it silently never gets ordered.

**Privacy:** everything runs client-side in the browser. No BOM data ever leaves your
computer — there is no server and no upload.

## Usage

1. Open the page (see *Deployment* below, or just open `index.html` locally).
2. Drop the **CAD BOM** on the left box, the **Item Master BOM** export on the right box.
3. Click **Compare BOMs**.
4. Review the three result tabs, hide/show columns as needed, filter, and download the
   result workbook via **Download .xlsx**.

## Supported input formats

### CAD BOM (left box)

| Format | Hierarchy | Quantities |
|---|---|---|
| Multi-level BOM **PDF** from the Vault web client | exact (from levels/indentation) | yes, if a QTY column is visible |
| **Leveled Excel/CSV** export (Level or dotted Item column + Number + Qty) | exact | yes |
| Flat Vault **Excel** paste (headerless, depth-first listing) | inferred | **no** — this export carries no quantity column |

Vault lets users choose which columns are visible in exports, so columns are auto-detected
by header name and content. If detection fails, the app shows a column-mapping step where
you assign the Part Number / Qty / Level / Title columns manually.

> The PDF extractor was written generically against Vault's report layout and needs tuning
> against a real sample — if a PDF fails to parse, please open an issue with the file (or a
> redacted page of it).

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
- **Quantity mismatches** (amber): rolled-up total quantity per part number
  (row qty × parent assembly quantities, summed over all occurrences) compared between the
  two BOMs, with a per-parent breakdown. Requires a CAD source that carries quantities.
- **In Item Master only:** items whose number never appears in the CAD BOM — stale or
  manually added entries worth reviewing.

## Development

No build step; plain HTML/CSS/JS. Libraries are vendored in `vendor/`
(SheetJS for Excel, pdf.js for PDFs) so the app works on locked-down networks.

```
js/compare.js            pure comparison + grouping + qty roll-up (no DOM)
js/parsers/itemmaster.js Item Master Excel parser
js/parsers/cad-flat-xlsx.js  flat Vault paste parser
js/parsers/cad-leveled.js    leveled table parser (PDF grid / Excel / CSV)
js/parsers/pdf-extract.js    pdf.js table reconstruction
js/parsers/detect.js         format detection / role validation
js/app.js                UI wiring
```

### Tests

```
node test/run-tests.mjs                          # synthetic tests only
node test/run-tests.mjs CAD_Bom.xlsx Item_Master.xls   # + real-sample baseline
```

Real BOM exports are not committed (potentially sensitive data); pass their paths as
arguments to run the full baseline assertions.

## Deployment

The site is plain static files with relative paths, so it can be served from any static
host or GitHub Pages configuration (repo Pages, user Pages, or a custom domain):
**Settings → Pages → Deploy from a branch**, select the default branch and `/ (root)`.
