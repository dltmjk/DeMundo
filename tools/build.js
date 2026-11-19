// Usage: node tools/build.js translations/FM1.json
// Builds translations/<ID>.docx and translations/<ID>.md from a unit JSON file.
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  HeadingLevel, ShadingType, LevelFormat, PageOrientation, AlignmentType } = require('docx');

const src = process.argv[2];
const d = JSON.parse(fs.readFileSync(src, 'utf8'));
const base = src.replace(/\.json$/, '');

const W = 15840 - 2 * 1080; // landscape letter, 0.75" margins
const COL = [W / 2, W / 2];
const border = { style: 'single', size: 4, color: '999999' };
const borders = { top: border, bottom: border, left: border, right: border };

// Render text, italicising [bracketed] insertions lightly is avoided; keep plain.
const p = (t, opts = {}) => new Paragraph({ children: [new TextRun({ text: t, ...opts })], spacing: { after: 80 } });
const cell = (t, shade, bold) => new TableCell({
  borders, width: { size: W / 2, type: WidthType.DXA },
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  shading: shade ? { fill: shade, type: ShadingType.CLEAR, color: 'auto' } : undefined,
  children: [p(t, bold ? { bold: true } : {})],
});

const kids = [];
kids.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(`De Mundo — ${d.unit}`)] }));
kids.push(p(`Latin title: ${d.latin_title}`, { italics: true }));
kids.push(p(`English title: ${d.english_title}`));
kids.push(p(`Pages: ${d.pages}`));
kids.push(p(`Source: Thomas White, De mundo dialogi tres (Paris: D. Moreau, 1642), Google Books scan.`, { size: 18 }));

for (const s of d.sections) {
  kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(s.heading)] }));
  const rows = [new TableRow({ tableHeader: true, children: [cell('Latin', 'DDDDDD', true), cell('English', 'DDDDDD', true)] })];
  for (const [la, en] of s.rows) rows.push(new TableRow({ cantSplit: false, children: [cell(la), cell(en)] }));
  kids.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COL, rows }));
}

kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Words and phrases with alternative or special meanings')] }));
for (const [la, en, alt, why] of d.terms) {
  kids.push(new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 60 }, children: [
    new TextRun({ text: la, bold: true, italics: true }),
    new TextRun(` — rendered “${en}”. ${alt}${why ? ' ' + why : ''}`)] }));
}
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Translator's notes")] }));
for (const n of d.notes) kids.push(new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 60 }, children: [new TextRun(n)] }));

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Garamond', size: 22 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: 'Garamond', size: 36, bold: true }, paragraph: { spacing: { before: 0, after: 160 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', run: { font: 'Garamond', size: 28, bold: true }, paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 0 } },
    ],
  },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } }, children: kids }],
});
Packer.toBuffer(doc).then(b => fs.writeFileSync(base + '.docx', b));

// Markdown copy
const esc = t => t.replace(/\|/g, '\\|');
let md = `# De Mundo — ${d.unit}\n\n*Latin title:* ${d.latin_title}  \n*English title:* ${d.english_title}  \n*Pages:* ${d.pages}\n\n`;
for (const s of d.sections) {
  md += `## ${s.heading}\n\n| Latin | English |\n|---|---|\n`;
  for (const [la, en] of s.rows) md += `| ${esc(la)} | ${esc(en)} |\n`;
  md += '\n';
}
md += `## Words and phrases with alternative or special meanings\n\n`;
for (const [la, en, alt, why] of d.terms) md += `- ***${la}*** — rendered “${en}”. ${alt}${why ? ' ' + why : ''}\n`;
md += `\n## Translator's notes\n\n`;
for (const n of d.notes) md += `- ${n}\n`;
fs.writeFileSync(base + '.md', md);
