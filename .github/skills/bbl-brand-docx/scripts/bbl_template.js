// Bangkok Bank (BBL) document style kit for docx-js.
// Styling extracted from BBL "Terms & Conditions" PDF (NFKL 097-2) and the
// BBL 56-1 One Report 2025 attachments (blue headings, cover, TOC, running header, report tables).
// Usage: const bbl = require("<skill>/scripts/bbl_template.js");
//        const doc = bbl.createDocument({ formCode: "NFKL 097-2 (07/69)", children: [...] });
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer, ImageRun, AlignmentType,
  PageNumber, TabStopType, Table, TableRow, TableCell, WidthType, BorderStyle,
  VerticalAlign, HeightRule, CharacterSet, ShadingType, TableOfContents, LevelFormat,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TableLayoutType,
} = require("docx");

const ASSETS = path.join(__dirname, "..", "assets");
const BBL = {
  // TH Sarabun New replaces the proprietary BBL Sans Looped. TTFs are bundled in assets/fonts.
  font: "TH Sarabun New",
  fontFile: path.join(ASSETS, "fonts", "THSarabunNew.ttf"),
  colors: {
    text: "231F20", blue: "007DC5", rule: "231F20",
    heading: "034EA3",      // One Report heading / label / cover-block blue
    // report table (FX-rate table look)
    tableHeader: "E0EAF5", tableBand: "F5F5F5", tableRule: "DBDBDB", tableText: "46464C",
  },
  // Sarabun renders ~20% smaller than BBL Sans at equal pt, so sizes are scaled up from the source
  // (13pt → 16pt body, 12pt → 14pt table, 8pt → 10pt footer). Half-points.
  size: { body: 32, table: 28, tableHeader: 28, footer: 20, note: 24, coverTitle: 56, coverSub: 36 },
  line: 240,                        // single spacing (Sarabun's own line height ≈ source density)
  indent: { l1: 390, l2: 920, l3: 1300, l4: 1720 }, // first-line indents in DXA
  page: {
    size: { width: 11906, height: 16838 },          // A4
    margin: { top: 1900, right: 820, bottom: 1000, left: 890, header: 890, footer: 260 },
  },
  contentWidth: 11906 - 890 - 820,                  // 10196 DXA
  logo: path.join(ASSETS, "bbl-logo.png"),
  icon: path.join(ASSETS, "bbl-icon.png"),
};

function font() {
  const name = process.env.BBL_FONT || BBL.font;
  return { ascii: name, hAnsi: name, cs: name, eastAsia: name };
}

// T&C and reports share the default BBL heading blue (#034EA3).
const styles = ({ headingColor = BBL.colors.heading } = {}) => { const HEAD = { bold: true, boldComplexScript: true,
  size: BBL.size.body, sizeComplexScript: BBL.size.body, color: headingColor }; return {
  default: {
    document: {
      run: { font: font(), size: BBL.size.body, sizeComplexScript: BBL.size.body, color: BBL.colors.text },
      paragraph: { spacing: { line: BBL.line, lineRule: "auto", before: 0, after: 0 } },
    },
    // Built-in title/headings MUST be overridden here, not in paragraphStyles: docx-js always emits its own
    // Title/Heading1..6 (blue #2E74B5, not bold) and Word uses the first definition of a duplicated styleId.
    title: { run: { ...HEAD, color: BBL.colors.text }, paragraph: { spacing: { before: 0, after: 0 } } },
    // One Report hierarchy: H1 section title, H2 subheading, H3 underlined topic — all bold #034EA3.
    heading1: { run: HEAD, paragraph: { keepNext: true, keepLines: true, outlineLevel: 0, spacing: { before: 240, after: 60 } } },
    heading2: { run: HEAD, paragraph: { keepNext: true, keepLines: true, outlineLevel: 1, spacing: { before: 120, after: 0 } } },
    heading3: { run: { ...HEAD, underline: {} }, paragraph: { keepNext: true, keepLines: true, outlineLevel: 2, spacing: { before: 120, after: 0 } } },
  },
  paragraphStyles: [
    { id: "TOCHeading", name: "TOC Heading", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { bold: true, boldComplexScript: true, color: BBL.colors.heading },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 120 } } },
    // TOC entries: plain black, page number right-aligned, no dot leader (One Report TOC)
    ...[1, 2].map(n => ({ id: `TOC${n}`, name: `toc ${n}`, basedOn: "Normal", next: "Normal",
      paragraph: { indent: { left: (n - 1) * BBL.indent.l1 },
        tabStops: [{ type: TabStopType.RIGHT, position: BBL.contentWidth, leader: "none" }] } })),
    { id: "Note", name: "BBL Note", basedOn: "Normal", quickFormat: true,
      run: { size: BBL.size.note, sizeComplexScript: BBL.size.note } },
    { id: "Clause1", name: "BBL Clause 1", basedOn: "Normal", quickFormat: true,
      paragraph: { indent: { firstLine: BBL.indent.l1 }, alignment: AlignmentType.JUSTIFIED } },
    { id: "Clause2", name: "BBL Clause 2", basedOn: "Normal", quickFormat: true,
      paragraph: { indent: { firstLine: BBL.indent.l2 }, alignment: AlignmentType.JUSTIFIED } },
    { id: "Clause3", name: "BBL Clause 3", basedOn: "Normal", quickFormat: true,
      paragraph: { indent: { firstLine: BBL.indent.l3 }, alignment: AlignmentType.JUSTIFIED } },
    { id: "Clause4", name: "BBL Clause 4", basedOn: "Normal", quickFormat: true,
      paragraph: { indent: { firstLine: BBL.indent.l4 }, alignment: AlignmentType.JUSTIFIED } },
    { id: "BodyText", name: "Body Text", basedOn: "Normal", quickFormat: true,
      paragraph: { alignment: AlignmentType.JUSTIFIED } },
    { id: "TableTitle", name: "BBL Table Title", basedOn: "Normal", quickFormat: true,
      run: { bold: true, boldComplexScript: true },
      paragraph: { alignment: AlignmentType.CENTER, keepNext: true, spacing: { after: 240 } } },
  ],
}; };

// Text helpers. Pass a string or an array of TextRun/strings.
const runs = (c, o = {}) => (Array.isArray(c) ? c : [c]).map(x => typeof x === "string"
  ? new TextRun({ text: x, ...o, ...(o.bold ? { boldComplexScript: true } : {}) }) : x);
const title = (t) => new Paragraph({ style: "Title", children: runs(t) });
const body = (t, o) => new Paragraph({ style: "BodyText", children: runs(t, o) });
const heading = (num, t) => new Paragraph({ style: "Heading1", children: runs(num === null || num === undefined || num === "" ? t : `${num}. ${t}`) });
// One Report section title "เอกสารแนบ 1<tab>title" with hanging indent so a wrapped title aligns after the label.
const sectionTitle = (label, t, hang = 1700) => new Paragraph({ style: "Heading1",
  indent: { left: hang, hanging: hang }, tabStops: [{ type: TabStopType.LEFT, position: hang }],
  children: runs(`${label}\t${t}`) });
const subheading = (t) => new Paragraph({ style: "Heading2", children: runs(t) });   // blue bold
const topic = (t) => new Paragraph({ style: "Heading3", children: runs(t) });        // blue bold underlined
const clause = (level, t, o) => new Paragraph({ style: `Clause${level}`, children: runs(t, o) });
// Emphasised legal paragraph (source uses bold for consumer-warning clauses)
const emphasis = (level, t) => clause(level, t, { bold: true });

// Level-3 letter marker: Thai (ก) ข ค ... when the item text contains Thai, else (a) (b) (c).
const THAI_LETTERS = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ";
const isThai = (t) => /[\u0E00-\u0E7F]/.test(t);
function letterMarker(i, lang = "en") {
  if (lang === "th") return `(${THAI_LETTERS[i]})`;
  let s = ""; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(97 + ((n - 1) % 26)) + s;
  return `(${s})`;
}
// Lettered list at level 3; `lang` "auto" (default) picks per item from its text.
const letterList = (items, { level = 3, lang = "auto" } = {}) =>
  items.map((t, i) => clause(level, `${letterMarker(i, lang === "auto" ? (isThai(t) ? "th" : "en") : lang)} ${t}`));

const B = (o = {}) => ({ bold: true, boldComplexScript: true, ...o });
const sz = (s) => ({ size: s, sizeComplexScript: s });

// Bullet list (real Word numbering, "•" as in One Report). Levels 0–2.
const bullet = (t, level = 0, o) => new Paragraph({ numbering: { reference: "bbl-bullets", level },
  alignment: AlignmentType.JUSTIFIED, children: runs(t, o) });
const NUMBERING = { config: [{ reference: "bbl-bullets", levels: [0, 1, 2].map(l => ({
  level: l, format: LevelFormat.BULLET, text: l === 1 ? "–" : "•", alignment: AlignmentType.LEFT,
  style: { paragraph: { indent: { left: BBL.indent.l1 + 360 + l * 400, hanging: 300 } } },
})) }] };

// Label/value lines (One Report profile layout): label bold blue, value black, wrapped value aligns to the tab.
const labelValue = (pairs, { labelWidth = 3400, indent = 0 } = {}) => pairs.map(([k, v]) => new Paragraph({
  indent: { left: indent + labelWidth, hanging: labelWidth },
  tabStops: [{ type: TabStopType.LEFT, position: indent + labelWidth }],
  children: [new TextRun({ text: k, ...B({ color: BBL.colors.heading }) }), new TextRun("\t"), ...runs(v)],
}));

// Thin black rule under a category heading (paragraph border — never a table).
const rule = (color = BBL.colors.rule) => new Paragraph({ spacing: { after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color, space: 1 } }, children: [] });

// Small "หมายเหตุ :" / "Note:" paragraph (12pt).
const note = (t) => new Paragraph({ style: "Note", alignment: AlignmentType.JUSTIFIED, children: runs(t) });

// Print-friendly callout: thin blue left rule, no fill, black body text, optional blue bold `title`
// ("Important", "หมายเหตุ"). `content`: string | string[] (one paragraph each). Returns Paragraph[].
function callout(content, { title, color = BBL.colors.heading } = {}) {
  const items = Array.isArray(content) ? content : [content];
  const left = { style: BorderStyle.SINGLE, size: 12, color, space: 10 };
  return items.map((t, i) => new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { left: 220 }, border: { left },
    keepLines: true, keepNext: i < items.length - 1,
    spacing: { before: i ? 0 : 160, after: i === items.length - 1 ? 160 : 0 },
    children: [
      ...(title && i === 0 ? [new TextRun({ text: `${title}: `, ...B({ color }) })] : []),
      ...runs(t),
    ] }));
}

// Right-aligned signature block: "(name)" blue bold, then title / date lines in black.
const signature = ({ name, lines = [] }) => [
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 360 }, keepNext: true,
    children: [new TextRun({ text: `(${name})`, ...B({ color: BBL.colors.heading }) })] }),
  ...lines.map((l, i) => new Paragraph({ alignment: AlignmentType.RIGHT, keepNext: i < lines.length - 1, children: runs(l) })),
];

// Table of contents: centred blue "สารบัญ" title, blue "หน้า" over the page column, Word TOC field (H1–H2).
// createDocument sets updateFields so Word fills the page numbers on open.
const toc = ({ title: t = "สารบัญ", pageLabel = "หน้า", levels = "1-2" } = {}) => [
  new Paragraph({ style: "TOCHeading", children: runs(t) }),
  new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: pageLabel, ...B({ color: BBL.colors.heading }) })] }),
  new TableOfContents(t, { hyperlink: true, headingStyleRange: levels }),
];

function logoHeader() {
  return new Header({ children: [new Paragraph({ children: [new ImageRun({
    type: "png", data: fs.readFileSync(BBL.logo),
    transformation: { width: 212, height: 57 }, // ≈159pt × 43pt, as in source
    altText: { title: "Bangkok Bank", description: "Bangkok Bank logo", name: "bbl-logo" },
  })] })] });
}

// One Report running header: bold 10pt, bank name left · document name right, no rule, no logo.
function textHeader(left = "ธนาคารกรุงเทพ จำกัด (มหาชน)", right = "") {
  return new Header({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: BBL.contentWidth }],
    children: [new TextRun({ text: left, ...B(sz(BBL.size.footer)) }), new TextRun({ text: right ? `\t${right}` : "", ...B(sz(BBL.size.footer)) })],
  })] });
}

// Footer: form code (left) · page x/y (center) · continuation hint (right), 8pt
function formFooter(formCode = "", nextHint = "") {
  const s = sz(BBL.size.footer);
  return new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.CENTER, position: Math.round(BBL.contentWidth / 2) },
               { type: TabStopType.RIGHT, position: BBL.contentWidth }],
    children: [
      new TextRun({ text: formCode, ...s }),
      new TextRun({ children: ["\t", PageNumber.CURRENT, "/", PageNumber.TOTAL_PAGES], ...s }),
      new TextRun({ text: nextHint ? `\t${nextHint}` : "", ...s }),
    ],
  })] });
}

// One Report footer: section name bold left · page number bold right.
function reportFooter(left = "") {
  const s = B(sz(BBL.size.footer));
  return new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: BBL.contentWidth }],
    children: [new TextRun({ text: left, ...s }), new TextRun({ children: ["\t", PageNumber.CURRENT], ...s })],
  })] });
}

// Data table.
//  style "report" (default) — FX-rate look: #E0EAF5 header with bold text, white/#F5F5F5 banded rows,
//    thin #DBDBDB horizontal rules only (no vertical lines), #46464C text, generous padding.
//  style "grid" — T&C rate schedule: 0.25pt #231F20 borders on every cell, no fill, centred.
// align: per-column "left" | "center" | "right" (body and, unless a header cell sets its own `align`, header).
//  Default: report → first column left, others left; grid → all centred. Put numbers in "right" columns.
function dataTable({ columnWidths, header, rows, align, style = "report", banded = true }) {
  const AL = { left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT };
  const report = style !== "grid";
  const line = (c, s) => ({ style: BorderStyle.SINGLE, size: s, color: c });
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const borders = report
    ? { top: line(BBL.colors.tableRule, 4), bottom: line(BBL.colors.tableRule, 4), left: none, right: none }
    : { top: line(BBL.colors.rule, 2), bottom: line(BBL.colors.rule, 2), left: line(BBL.colors.rule, 2), right: line(BBL.colors.rule, 2) };
  const pad = report ? { top: 90, bottom: 90, left: 160, right: 160 } : { top: 20, bottom: 20, left: 80, right: 80 };
  const colAlign = (i) => (align && align[i]) || (report ? "left" : "center");
  const cell = (t, w, o = {}) => new TableCell({
    borders, width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: pad,
    ...(o.fill ? { shading: { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } } : {}), ...o.cell,
    children: [new Paragraph({ alignment: AL[o.align] || AlignmentType.CENTER,
      children: runs(typeof t === "number" ? String(t) : t, { ...sz(o.size || BBL.size.table),
        ...(report ? { color: BBL.colors.tableText } : {}), ...(o.bold ? { bold: true } : {}) }) })],
  });
  const total = columnWidths.reduce((a, c) => a + c, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths, layout: TableLayoutType.FIXED,
    rows: [
      ...(header || []).map(r => { let col = 0; return new TableRow({ tableHeader: true, cantSplit: true,
        children: r.map(h => { const i = col; col += h.columnSpan || 1;
          return cell(h.text, h.width, { size: h.size || BBL.size.tableHeader, bold: report,
            fill: report ? BBL.colors.tableHeader : undefined,
            align: h.align || (report && !h.columnSpan ? colAlign(i) : "center"),
            cell: { ...(h.columnSpan ? { columnSpan: h.columnSpan } : {}), ...(h.rowSpan ? { rowSpan: h.rowSpan } : {}) } }); }) }); }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, height: { value: report ? 400 : 320, rule: HeightRule.ATLEAST },
        children: r.map((v, i) => cell(v, columnWidths[i], { align: colAlign(i),
          fill: report && banded && ri % 2 === 1 ? BBL.colors.tableBand : undefined })) })),
    ],
  });
}

// 1×1 solid-colour PNG (stretched for the cover block).
function solidPng(hex) {
  const crcT = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (b) => { let c = 0xFFFFFFFF; for (const x of b) c = crcT[(c ^ x) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.from([0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
  const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(Buffer.from([0, r, g, b]))), chunk("IEND", Buffer.alloc(0))]);
}

// One Report cover: large logo top-left, black bold title lines low on the page,
// full-bleed #034EA3 block across the bottom ~25% of the page. Returns section children.
function cover({ title: t, subtitle, lines = [], blockRatio = 0.25 } = {}) {
  const pageW = 793.7, pageH = 1122.5;            // A4 in px @96dpi
  const blockH = Math.round(pageH * blockRatio);
  return [
    new Paragraph({ children: [new ImageRun({ type: "png", data: fs.readFileSync(BBL.logo),
      transformation: { width: 318, height: 86 },
      altText: { title: "Bangkok Bank", description: "Bangkok Bank logo", name: "bbl-cover-logo" } })] }),
    new Paragraph({ spacing: { before: 6600 }, keepNext: true,
      children: runs(t, B(sz(BBL.size.coverTitle))) }),
    ...(subtitle ? [new Paragraph({ keepNext: true, children: runs(subtitle, B(sz(BBL.size.coverSub))) })] : []),
    ...lines.map(l => new Paragraph({ keepNext: true, children: runs(l) })),
    new Paragraph({ children: [new ImageRun({ type: "png", data: solidPng(BBL.colors.heading),
      transformation: { width: Math.ceil(pageW) + 2, height: blockH },
      floating: {
        horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
        verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: Math.round((pageH - blockH) * 9525) },
        behindDocument: true, allowOverlap: true, lockAnchor: true,
      },
      altText: { title: "Cover block", description: "Bangkok Bank blue cover band", name: "bbl-cover-block" } })] }),
  ];
}

// embedFont: embed THSarabunNew.ttf (regular) so the document renders where the font isn't installed.
// docx-js can only embed the regular face; bold falls back to the installed bold or synthetic bold.
//  header: "logo" (T&C, default) | "text" (One Report running header; headerText: { left, right })
//  footer: "form" (T&C form code + n/N, default) | "report" (footerText bold left + page no. right)
//  coverPage: { title, subtitle, lines } → separate first section without header/footer
//  headingColor: defaults to BBL heading blue (#034EA3) for both layouts
//  updateFields: true when the document contains bbl.toc() (auto-set when coverPage or toc option given)
function createDocument({ formCode = "", children = [], sections, title: docTitle = "Bangkok Bank",
  logo = true, embedFont = true, header = "logo", headerText = {}, footer = "form", footerText = "",
  coverPage, toc: hasToc = false, updateFields, headingColor } = {}) {
  const margin = BBL.page.margin;
  const props = (extra = {}) => ({ page: { size: BBL.page.size, margin, ...extra } });
  const hdr = header === "text" ? textHeader(headerText.left, headerText.right) : (logo ? logoHeader() : undefined);
  const ftr = footer === "report" ? reportFooter(footerText) : formFooter(formCode);
  const sec = (c, i) => ({ properties: props(coverPage && i === 0 ? { pageNumbers: { start: 1 } } : {}),
    headers: hdr ? { default: hdr } : undefined, footers: { default: ftr }, children: c });
  const fonts = embedFont && !process.env.BBL_FONT && fs.existsSync(BBL.fontFile)
    ? [{ name: BBL.font, data: fs.readFileSync(BBL.fontFile), characterSet: CharacterSet.THAI }] : undefined;
  const body = sections ? sections.map(sec) : [sec(children, 0)];
  const emptyH = new Header({ children: [new Paragraph({ children: [] })] });
  const emptyF = new Footer({ children: [new Paragraph({ children: [] })] });
  const coverSec = coverPage ? [{ properties: props(), headers: { default: emptyH }, footers: { default: emptyF },
    children: cover(coverPage) }] : [];
  return new Document({
    title: docTitle, creator: "Bangkok Bank", styles: styles({ headingColor }), numbering: NUMBERING, ...(fonts ? { fonts } : {}),
    ...((updateFields ?? hasToc) ? { features: { updateFields: true } } : {}),
    sections: [...coverSec, ...body],
  });
}

// Pack + write, fixing two docx-js quirks so the file passes OOXML schema validation:
// lowercase embedded-font GUIDs (schema requires uppercase) and folder entries in the ZIP.
async function save(doc, outPath) {
  const JSZip = require("jszip");
  const zip = await JSZip.loadAsync(await Packer.toBuffer(doc));
  const out = new JSZip();
  for (const [name, f] of Object.entries(zip.files)) {
    if (f.dir) continue;
    if (name === "word/fontTable.xml") {
      const xml = (await f.async("string")).replace(/w:fontKey="(\{[^"]+\})"/g, (m, g) => `w:fontKey="${g.toUpperCase()}"`);
      out.file(name, xml, { createFolders: false });
    } else out.file(name, await f.async("nodebuffer"), { createFolders: false });
  }
  const buf = await out.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(outPath, buf);
  return buf;
}

module.exports = { BBL, styles, title, body, heading, sectionTitle, subheading, topic, clause, emphasis,
  bullet, labelValue, rule, note, callout, signature, toc, cover, dataTable, letterMarker, letterList,
  logoHeader, textHeader, formFooter, reportFooter, createDocument, save, runs };
