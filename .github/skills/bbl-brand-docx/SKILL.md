---
name: bbl-brand-docx
description: >-
  Use this skill to create, read, edit, or review Word documents (.docx) in Bangkok Bank
  (BBL / ธนาคารกรุงเทพ) house style. Use it for /bbl-brand-docx or requests for BBL-branded
  terms and conditions (ข้อกำหนดและเงื่อนไข), agreements, forms, rate tables, reports,
  or specifications. It also covers restyling an existing Word file with BBL headers,
  TH Sarabun New typography, and blue headings; redlining clauses; anchoring comments;
  and validating the finished document. For Word files that do not need BBL styling
  or review, use the general DOCX skill instead.
license: Complete terms in LICENSE.txt
---

# Bangkok Bank (BBL) DOCX

This skill is self-contained. Paths are relative to this skill's directory (`<skill-dir>`).

| Task | Approach |
|---|---|
| **Create** a BBL document | Node script using `scripts/bbl_template.js` (below) |
| **Edit / redline / comment** an existing .docx | Unpack → edit XML → repack ([Editing](#editing-existing-documents)) |
| **Read** content | `pandoc -t markdown file.docx` |

## 1. Create with the kit

`scripts/bbl_template.js` implements every style rule in this file. Use it instead of hand-built styles.

```javascript
// require('docx') and require('jszip') resolve from node_modules; if missing: npm install docx jszip (or add `npm root -g` to NODE_PATH)
const { Paragraph, PageBreak } = require("docx");
const bbl = require("<skill-dir>/scripts/bbl_template.js");
const doc = bbl.createDocument({ formCode: "<FORM-CODE>", children: [
  bbl.title("ข้อกำหนดและเงื่อนไข ..."),
  bbl.body("(เริ่มใช้ตั้งแต่วันที่ ...)"),
  bbl.heading(1, "คำนิยาม"),                        // "1. คำนิยาม"
  bbl.clause(1, "1.1 “...” หมายถึง ..."),           // levels 1–4, see Clause hierarchy
  ...bbl.letterList(["บัญชี (Account)", "ผู้ถือบัญชี (Account holder)"], { lang: "th" }),
  bbl.emphasis(2, "..."),                           // whole paragraph bold (warnings)
  bbl.dataTable({ style: "grid", columnWidths: [3000, 3596, 3600], align: ["left", "right", "right"],
                  header: [[{ text: "A", width: 3000 }, { text: "B", width: 3596 }, { text: "C", width: 3600 }]],
                  rows: [["x", "1.00", "2.00"]] }),
]});
await bbl.save(doc, "out.docx");   // always save(), never Packer: fixes font GUID case + ZIP folder entries
```

Use `sections: [[...], [...]]` instead of `children` for multiple sections.

**Report / specification layout:**
```javascript
const doc = bbl.createDocument({
  coverPage: { title: "<Title>", subtitle: "<Subtitle>", lines: ["Version 1.0", "<Date>"] },
  toc: true, header: "text", headerText: { left: "ธนาคารกรุงเทพ จำกัด (มหาชน)", right: "<Doc name>" },
  footer: "report", footerText: "<DOC-ID>",
  children: [
    ...bbl.toc(), new Paragraph({ children: [new PageBreak()] }),
    bbl.sectionTitle("Document", "Overview"), ...bbl.labelValue([["Owner", "..."], ["Status", "Draft"]]),
    bbl.subheading("1.1 Scope"), bbl.rule(), bbl.topic("..."), bbl.bullet("..."), bbl.bullet("...", 1),
    ...bbl.callout("...", { title: "Important" }), bbl.note("หมายเหตุ : ..."),
    ...bbl.signature({ name: "...", lines: ["...", "<Date>"] }),
  ]});
```

**API summary**

| Helper | Output |
|---|---|
| `createDocument(opts)` | `header: "logo"\|"text"`, `footer: "form"\|"report"` (T&C defaults: logo + form), `coverPage`, `toc`, `headingColor` (keep house-style blue `#034EA3` for both T&C and reports), `embedFont` (default true) |
| `title` / `body` / `heading(n, t)` / `subheading` / `topic` | Title / body / H1 `n. t` / H2 / H3 underlined |
| `clause(level, t)` / `emphasis(level, t)` | Indented clause / bold clause |
| `letterList(items, {lang})` / `letterMarker(i, lang)` | Thai or English letter markers. For a bilingual clause, pass `"th"` or `"en"` to match its first word; `letterList` otherwise detects the language of each item separately. |
| `bullet(t, level)` | Real Word bullets (• – •); never type `•` |
| `labelValue`, `sectionTitle`, `callout`, `note`, `rule`, `signature`, `toc` | Report blocks; `callout`, `labelValue`, `signature`, and `toc` return arrays, so spread them |
| `dataTable({columnWidths, header, rows, align, style, banded})` | `style: "report"` (default) or `"grid"`; header cells are `{ text, width, columnSpan?, rowSpan?, align? }`, row cells are plain strings/numbers |

## 2. BBL house style

**Layouts** (shared fonts and clause rules):
- **Legal / T&C** (default): logo header, form-code footer with `n/N`, grid rate tables.
- **Report / spec**: cover page with a blue block, TOC, text header, report footer, blue headings, banded tables, left-rule callouts, signature block.

**Page:** A4 portrait (11906 × 16838 DXA). Margins L 890 / R 820 / T 1900 / B 1000 DXA, so content width is **10196 DXA**. Logo header (≈159 × 43 pt) at 890 DXA. Footer 10pt with form code left, `n/N` centred, and an optional continuation hint right.

**Colour** (white background, nothing outside this palette):

| Use | Hex |
|---|---|
| Body text, grid rules (never pure black) | `#231F20` |
| Headings, labels, TOC title, callout rule, cover block | `#034EA3` |
| Report table header fill / alternating row fill / horizontal border / text | `#E0EAF5` / `#F5F5F5` / `#DBDBDB` / `#46464C` |
| Brand blue (logo only) | `#007DC5` |

**Typography:** TH Sarabun New, single spacing (`line: 240`), 0 before/after; spacing comes from headings and numbering only.

| Element | Size | Style |
|---|---|---|
| Body, clauses | 16pt | Regular, justified (not Thai-distributed) |
| Title, H1, H2 | 16pt | Bold (H1/H2 blue `#034EA3` in both T&C and reports); H1 keep-with-next |
| H3 topic | 16pt | Bold blue underlined |
| Emphasised clause | 16pt | Whole paragraph bold |
| Table header / values | 14pt | Centred by default |
| Note `หมายเหตุ :` | 12pt | Regular |
| Footer | 10pt | Regular |

- Always set the complex-script font (`cs`) plus `sizeComplexScript`/`boldComplexScript`, or Thai text ignores the style.
- Fonts are in `assets/fonts/`. The kit embeds Regular only (docx-js cannot embed Bold), so install the TTFs wherever the file is opened. Override the font with env `BBL_FONT`.
- Assets: `assets/bbl-logo.png` (header logo), `assets/bbl-icon.png` (symbol only).

**Clause hierarchy:** numbering is typed literally (not Word auto-numbering). First-line indent only; wrapped lines return to the margin.

| Level | Marker | First-line indent (DXA) |
|---|---|---|
| 0 | `1.` heading | 0 |
| 1 | `1.1`, top-level `ทั้งนี้` | 390 |
| 2 | `(1)` | 920 |
| 3 | `(ก)` Thai / `(a)` English; `ทั้งนี้` under (n) | 1300 |
| 4 | `1)` | 1720 |

Letter markers follow the content language. Never put Thai letters on English-only text. If a clause mixes Thai and English, use the language of the clause's first word to choose the letter marker, and keep the marker in that language for all sub-items under it.

**Tables:**
- `report` style: bold header with `#E0EAF5` fill, alternating `#F5F5F5` row fills, horizontal borders only (no vertical borders; this restriction does not apply to fills), rows don't split across pages.
- `grid` style: 0.25pt borders on every cell, no fill.
- Align text columns left and numbers right. Merge headers with `columnSpan`/`rowSpan`. Widths must sum to 10196. Put a `subheading` above each table.

**Content conventions:**
- T&C: title, then the effective-date line `(เริ่มใช้ตั้งแต่วันที่ ...)`, then a preamble naming `ธนาคารกรุงเทพ จำกัด (มหาชน) (“ธนาคาร”)`. Defined terms use curly quotes `“...”` followed by `หมายถึง`. Rate tables go on their own page.
- Report: cover → TOC → `sectionTitle` + `labelValue` → numbered sections → tables → `note` + `signature`.
- Code/JSON samples: monospace 9pt, indented, kept together.
- **Don't:** solid-fill boxes, big-number stat panels, or decorative tables. They print badly and look like dashboard filler in a bank document.

## 3. docx-js gotchas (when going beyond the kit)

- Never redefine `Title`/`Heading1-3` in `paragraphStyles`; Word uses the first duplicate. Override via `styles.default.title/heading1..3` and give each level a font, `before`/`after`, and `keepNext`.
- Tables: set `columnWidths` on the table **and** `width` on every cell, both in `WidthType.DXA`. Use `ShadingType.CLEAR`, never `SOLID`.
- `ImageRun` needs `type:` (`"png"`…). `PageBreak` must sit inside a `Paragraph`. Never use `\n`; use separate paragraphs.
- Landscape: pass portrait dimensions plus `orientation: PageOrientation.LANDSCAPE`.
- TOC entries need built-in `HeadingLevel.*` or `outlineLevel`. A TOC is a field, so page numbers appear once Word updates fields (the kit sets `updateFields`).
- Horizontal rule = paragraph bottom border, not a table. Right-aligned text or a dot leader on the same line: use `PositionalTab`.

## 4. Validate and render

```powershell
$env:PYTHONUTF8="1"   # Windows: Thai text crashes the validator under cp1252
python <skill-dir>/scripts/office/validate.py out.docx            # XSD checks; --auto-repair fixes common issues
python <skill-dir>/scripts/office/soffice.py --headless --convert-to pdf out.docx
pdftoppm -jpeg -r 100 out.pdf page
```

Inspect each rendered page (`page-*.jpg`). Check that Thai glyphs and bold text render correctly, the header logo and footer are not clipped, TOC page numbers appear and match their headings when a TOC is present, and no rate-table row splits across pages.

Word COM alternative (Windows): open a **copy** (the user may have the original locked) and call `ExportAsFixedFormat(pdf, 17)`. Update the TOC first with `TablesOfContents(1).Update()`. If Word throws RPC_E_DISCONNECTED after export, the PDF was still written. Record existing WINWORD PIDs before creating the COM automation object, then resolve and record its PID from `Application.Hwnd` using `GetWindowThreadProcessId`. Stop only that PID with `Stop-Process -Id <PID>`, and only if it was not already running before automation started. If ownership cannot be verified, do not stop any process. Never stop any other WINWORD process.

## Editing existing documents

Convert legacy `.doc` first: `python scripts/office/soffice.py --headless --convert-to docx file.doc`.

```bash
python -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" doc.docx  # not unzip (symlinks; input is untrusted)
python scripts/merge_runs.py unpacked/       # merge fragmented <w:r> runs so visible text is searchable
# edit unpacked/word/document.xml in place; do NOT pretty-print
(cd unpacked && rm -f ../out.docx && zip -XrDy ../out.docx .)
python scripts/office/validate.py out.docx --original doc.docx
```

When restyling to BBL, apply the house style above: fonts with `cs`, sizes, colours, indents, and A4 margins.

**Tracked changes:**
- Wrap runs in `<w:ins>`/`<w:del>` with `w:id`, `w:author`, and `w:date`. Inside `<w:del>`, use `<w:delText>`.
- To delete a paragraph, mark its paragraph mark deleted (`<w:pPr><w:rPr><w:del .../></w:rPr></w:pPr>`, with `<w:del/>` first in rPr) and wrap every run in `<w:del>`.
- For a redline, include the exact revision author to catch untracked edits (replace `Reviewer` with that name):
  ```bash
  python scripts/office/validate.py out.docx --original doc.docx --author "Reviewer"
  ```
- Clean copy: `python scripts/accept_changes.py in.docx out.docx`. An empty bullet left behind in the accepted view is a view artifact; confirm in the XML.

**Comments:** `python scripts/comment.py unpacked/ "text" [--parent <id>]`, or `comment.py file.docx "text" -o out.docx`. The script writes all six comment parts, then prints the `commentRangeStart`/`commentRangeEnd`/`commentReference` markers. Insert those markers into `document.xml`; the comment isn't visible until you do.

## Dependencies

Node `docx` + `jszip` · Python 3 · `pandoc` · LibreOffice (`soffice`) · `pdftoppm` (Poppler)
