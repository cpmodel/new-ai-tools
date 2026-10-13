// Render a simple block list (JSON) to .docx with docx-js: headings, paragraphs, bullets, tables.
// Usage: node render_docx_v1_0.js blocks.json out.docx
// Block types: {t:'title'|'h1'|'h2'|'p'|'note', text} | {t:'bullets', items:[...]} |
//              {t:'table', head:[...], rows:[[...]], widths:[twips...]}. Inline **bold** is supported in text.
const fs = require('fs');
const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
        LevelFormat, AlignmentType, Footer, PageNumber, TableOfContents, BorderStyle } = docx;

const [, , src, out] = process.argv;
const doc = JSON.parse(fs.readFileSync(src, 'utf8'));
const FONT = 'Arial';
const PAGE_W = 11906, MARGIN = 1134;          // A4, 2 cm margins
const BODY_W = PAGE_W - 2 * MARGIN;

function runs(text, opts = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*)/);
  return parts.filter(p => p.length).map(p => p.startsWith('**') && p.endsWith('**')
    ? new TextRun({ text: p.slice(2, -2), bold: true, font: FONT, size: opts.size || 21 })
    : new TextRun({ text: p, font: FONT, size: opts.size || 21, italics: !!opts.italics, color: opts.color }));
}

function cell(text, width, head) {
  const kids = head ? [new TextRun({ text: String(text), bold: true, font: FONT, size: 18 })] : runs(text, { size: 18 });
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE6F1' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: kids, spacing: { after: 0 } })],
  });
}

const children = [];
for (const b of doc.blocks) {
  if (b.t === 'title') {
    children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: b.text, font: FONT, size: 40, bold: true })] }));
  } else if (b.t === 'subtitle') {
    children.push(new Paragraph({ children: runs(b.text, { size: 22, color: '595959' }), spacing: { after: 240 } }));
  } else if (b.t === 'toc') {
    children.push(new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }));
  } else if (b.t === 'h1') {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: b.text, font: FONT, size: 30, bold: true, color: '1F3864' })], spacing: { before: 360, after: 120 } }));
  } else if (b.t === 'h2') {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: b.text, font: FONT, size: 24, bold: true, color: '2F5496' })], spacing: { before: 240, after: 80 } }));
  } else if (b.t === 'p') {
    children.push(new Paragraph({ children: runs(b.text), spacing: { after: 120 } }));
  } else if (b.t === 'note') {
    children.push(new Paragraph({ children: runs(b.text, { size: 18, italics: true, color: '595959' }), spacing: { after: 120 } }));
  } else if (b.t === 'bullets') {
    for (const it of b.items) {
      children.push(new Paragraph({ numbering: { reference: 'bullets', level: 0 }, children: runs(it), spacing: { after: 60 } }));
    }
  } else if (b.t === 'table') {
    const w = b.widths || b.head.map(() => Math.floor(BODY_W / b.head.length));
    const scale = BODY_W / w.reduce((a, c) => a + c, 0);
    const ws = w.map(x => Math.floor(x * scale));
    ws[ws.length - 1] += BODY_W - ws.reduce((a, c) => a + c, 0);
    const rows = [new TableRow({ tableHeader: true, children: b.head.map((h, i) => cell(h, ws[i], true)) })]
      .concat(b.rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, ws[i], false)) })));
    children.push(new Table({ width: { size: BODY_W, type: WidthType.DXA }, columnWidths: ws, rows }));
    children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
  }
}

const document = new Document({
  creator: 'CPAT-AI', title: doc.title,
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [doc.footer + '  |  page ', PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children,
  }],
});
Packer.toBuffer(document).then(buf => { fs.writeFileSync(out, buf); console.log('wrote', out); });
