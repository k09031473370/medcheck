// html2json.py が作った JSON から Word (.docx) を作る。
// 使い方: node build_docx.js manual.json 出力.docx
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, LevelFormat, HeadingLevel, Footer, PageNumber, VerticalAlign,
} = require('docx');

const [, , SRC, DST] = process.argv;
const blocks = JSON.parse(fs.readFileSync(SRC, 'utf8'));

const FONT = { ascii: 'Yu Gothic', hAnsi: 'Yu Gothic', eastAsia: 'Yu Gothic', cs: 'Yu Gothic' };
const MONO = { ascii: 'MS Gothic', hAnsi: 'MS Gothic', eastAsia: 'MS Gothic', cs: 'MS Gothic' };
const BODY = 22;            // 11pt
const GREEN = '0B6B4F';
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134;   // A4, 余白 2cm
const CONTENT_W = PAGE_W - MARGIN * 2;

function runs(rs, base) {
  base = base || {};
  const out = [];
  let brNext = false;
  for (const r of rs) {
    if (r.br) { brNext = true; continue; }
    const o = { text: r.t, font: FONT, size: base.size || BODY };
    if (base.bold || r.b) o.bold = true;
    if (base.color) o.color = base.color;
    if (r.code) { o.font = MONO; o.shading = { type: ShadingType.CLEAR, fill: 'EEF1F5', color: 'auto' }; }
    if (r.btn) { o.bold = true; o.border = { style: BorderStyle.SINGLE, size: 4, color: '8A94A3', space: 1 }; o.shading = { type: ShadingType.CLEAR, fill: 'F1F3F6', color: 'auto' }; }
    if (r.small) { o.size = (base.size || BODY) - 3; o.color = '555F6D'; }
    if (brNext) { o.break = 1; brNext = false; }
    out.push(new TextRun(o));
  }
  return out;
}
const plain = rs => rs.filter(r => r.t).map(r => r.t).join('');

const children = [];
let numInst = 0;

function para(rs, opt) {
  opt = opt || {};
  return new Paragraph(Object.assign({ spacing: { before: 60, after: 100, line: 320 }, children: runs(rs, opt.base) }, opt.p || {}));
}

function callout(kind, rs, extra) {
  const fill = kind === 'warn' ? 'FFF4E0' : 'F1F6FA';
  const bar = kind === 'warn' ? 'D97706' : GREEN;
  return new Paragraph(Object.assign({
    spacing: { before: 0, after: 0, line: 320 },
    indent: { left: 200, right: 120 },
    shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: bar, space: 6 } },
    children: runs(rs),
  }, extra || {}));
}

function table(rows) {
  const n = Math.max(...rows.map(r => r.cells.length));
  // 列の幅: 文字数の多い列を広く (最小 16%)
  const w = [];
  for (let c = 0; c < n; c++) {
    let m = 4;
    for (const r of rows) if (r.cells[c]) m = Math.max(m, Math.min(plain(r.cells[c]).length, 46));
    w.push(Math.sqrt(m));
  }
  let sum = w.reduce((a, b) => a + b, 0);
  let widths = w.map(x => Math.max(0.16, x / sum));
  sum = widths.reduce((a, b) => a + b, 0);
  widths = widths.map(x => Math.floor(CONTENT_W * x / sum));
  widths[n - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const line = { style: BorderStyle.SINGLE, size: 4, color: 'B8C0CC' };
  const borders = { top: line, bottom: line, left: line, right: line };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(r => new TableRow({
      tableHeader: r.head, cantSplit: true,
      children: r.cells.map((c, i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        borders, verticalAlign: VerticalAlign.TOP,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        shading: r.head ? { type: ShadingType.CLEAR, fill: 'E8EEF3', color: 'auto' } : undefined,
        children: [new Paragraph({ spacing: { before: 0, after: 0, line: 290 }, children: runs(c, { size: 20, bold: r.head }) })],
      })),
    })),
  });
}
const gap = (n) => new Paragraph({ spacing: { before: 0, after: n || 120 }, children: [] });

for (const b of blocks) {
  switch (b.type) {
    case 'h1':
      children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: plain(b.runs), font: FONT })] }));
      break;
    case 'lead':
      children.push(para(b.runs, { base: { color: '4A5563' }, p: { spacing: { before: 0, after: 200 } } }));
      break;
    case 'flow':
      children.push(new Paragraph({
        spacing: { before: 60, after: 200, line: 340 },
        children: b.steps.flatMap((s, i) => [
          new TextRun({ text: s, font: FONT, size: BODY, bold: true }),
          ...(i < b.steps.length - 1 ? [new TextRun({ text: '  →  ', font: FONT, size: BODY, color: '6B7480' })] : []),
        ]),
      }));
      break;
    case 'toc':
      children.push(callout('note', [{ t: 'もくじ', b: true }]));
      children.push(callout('note', [{ t: b.items.join(' ／ ') }]));
      children.push(gap());
      break;
    case 'h2':
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: plain(b.runs), font: FONT })] }));
      break;
    case 'h3':
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: plain(b.runs), font: FONT })] }));
      break;
    case 'p':
      children.push(para(b.runs));
      break;
    case 'small':
      children.push(para(b.runs, { base: { size: 19, color: '555F6D' } }));
      break;
    case 'ul':
    case 'ol': {
      const inst = ++numInst;
      const ref = b.type === 'ol' ? 'steps' : 'bullets';
      for (const it of b.items) {
        children.push(new Paragraph({
          numbering: { reference: ref, level: 0, instance: inst },
          spacing: { before: 30, after: 60, line: 320 }, children: runs(it.runs),
        }));
        for (const s of it.sub) {
          children.push(new Paragraph({
            numbering: { reference: 'bullets', level: 1, instance: 1000 + inst },
            spacing: { before: 20, after: 40, line: 320 }, children: runs(s),
          }));
        }
      }
      children.push(gap(60));
      break;
    }
    case 'table':
      children.push(table(b.rows));
      children.push(gap());
      break;
    case 'callout':
      children.push(gap(40));
      if (b.items) {
        b.items.forEach(it => children.push(callout(b.kind, [{ t: '・' }].concat(it), { indent: { left: 440, hanging: 240, right: 120 } })));
      } else {
        children.push(callout(b.kind, b.runs));
      }
      children.push(gap());
      break;
    default:
      throw new Error('知らないブロック: ' + b.type);
  }
}

const doc = new Document({
  title: '受診票の取込 説明書',
  styles: {
    default: {
      document: { run: { font: FONT, size: BODY } },
      title: { run: { font: FONT, size: 40, bold: true, color: '1B2430' }, paragraph: { spacing: { before: 0, after: 80 }, keepNext: true } },
      heading1: {
        run: { font: FONT, size: 28, bold: true, color: '1B2430' },
        paragraph: {
          spacing: { before: 360, after: 140 }, keepNext: true,
          shading: { type: ShadingType.CLEAR, fill: 'EEF3F1', color: 'auto' },
          border: { left: { style: BorderStyle.SINGLE, size: 36, color: GREEN, space: 6 } },
          indent: { left: 160 },
        },
      },
      heading2: { run: { font: FONT, size: 24, bold: true, color: GREEN }, paragraph: { spacing: { before: 220, after: 80 }, keepNext: true } },
    },
  },
  numbering: {
    config: [
      { reference: 'steps', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 380 } } } }] },
      { reference: 'bullets', levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 280 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 900, hanging: 280 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: '受診票の取込 説明書　', font: FONT, size: 16, color: '6B7480' }), new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: '6B7480' })],
      })] }),
    },
    children,
  }],
});

// 日本語と英数字の間に Word が自動で入れるすき間を止める (ファイル名に空白があるように見えるため)。
// ZIP のフォルダ項目も入れない。
const JSZip = require(require.resolve('jszip', { paths: [require.resolve('docx')] }));
Packer.toBuffer(doc).then(async buf => {
  const zin = await JSZip.loadAsync(buf);
  const zout = new JSZip();
  for (const name of Object.keys(zin.files)) {
    const f = zin.files[name];
    if (f.dir) continue;
    let data = await f.async('nodebuffer');
    if (name === 'word/styles.xml') {
      let x = data.toString('utf8');
      const fix = '<w:pPrDefault><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/></w:pPr></w:pPrDefault>';
      if (!x.includes('<w:pPrDefault/>')) throw new Error('pPrDefault が見つからない');
      x = x.replace('<w:pPrDefault/>', fix);
      data = Buffer.from(x, 'utf8');
    }
    if (name === 'word/document.xml') {
      // 既定値を見ないソフト向けに、段落にも同じ指定を入れる (w:spacing の直前が正しい位置)
      let x = data.toString('utf8');
      x = x.split('<w:spacing ').join('<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:spacing ');
      data = Buffer.from(x, 'utf8');
    }
    zout.file(name, data, { createFolders: false });
  }
  const out = await zout.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(DST, out);
  console.log('written', DST, out.length);
});
