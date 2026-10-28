// usage: node build_manuals.js <outdir>
const fs = require('fs'), path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  HeadingLevel, AlignmentType, BorderStyle, ShadingType, LevelFormat, Header, Footer, PageNumber,
} = require('docx');

const FONT = 'Yu Gothic';
const TEAL = '0A7B6C', NAVY = '0F2847', PALE = 'E6F5F3', WARN = 'FDF1E6', WARN_LINE = 'D9822B', GRAY = 'F2F5F7', BORDER = 'B8CBD4';
const PAGE_W = 11906, MARGIN = 1134, CW = PAGE_W - MARGIN * 2;

function makeDoc(title, subtitle, base, build) {
  const SZ = base;                       // body size in half-points
  const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: SZ, ...o });
  const kids = (t) => typeof t === 'string' ? [run(t)] : t;
  const p = (t, o = {}) => new Paragraph({ spacing: { after: 120, line: 320 }, ...o, children: kids(t) });
  const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 140 },
    shading: { type: ShadingType.CLEAR, fill: TEAL, color: 'auto' }, children: [run(' ' + t, { bold: true, size: SZ + 8, color: 'FFFFFF' })] });
  const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 2 } }, children: [run(t, { bold: true, size: SZ + 3, color: TEAL })] });
  let listId = 0; const refs = [];
  const steps = (items) => { const ref = 'n' + (listId++); refs.push(ref);
    return items.map(t => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 90, line: 320 }, children: kids(t) })); };
  const bullets = (items) => items.map(t => new Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 70, line: 320 }, children: kids(t) }));
  const box = (lines, warn = false) => new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    rows: [new TableRow({ children: [new TableCell({ width: { size: CW, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: warn ? WARN : PALE, color: 'auto' },
      borders: { top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL }, left: { style: BorderStyle.SINGLE, size: 24, color: warn ? WARN_LINE : TEAL } },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: lines.map((l, i) => new Paragraph({ spacing: { after: i === lines.length - 1 ? 0 : 60, line: 300 }, children: kids(l) })) })] })] });
  const table = (headers, rows, parts) => {
    const s = parts.reduce((a, b) => a + b, 0); const w = parts.map(x => Math.floor(CW * x / s)); w[w.length - 1] += CW - w.reduce((a, b) => a + b, 0);
    const cell = (t, i, hdr, ri) => new TableCell({ width: { size: w[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: hdr ? NAVY : (ri % 2 ? GRAY : 'FFFFFF'), color: 'auto' },
      borders: ['top', 'bottom', 'left', 'right'].reduce((a, k) => (a[k] = { style: BorderStyle.SINGLE, size: 4, color: BORDER }, a), {}),
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ spacing: { after: 0, line: 280 }, children: [run(t, { size: SZ - 2, bold: hdr || i === 0, color: hdr ? 'FFFFFF' : '000000' })] })] });
    return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: w,
      rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true, 0)) }),
             ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, ri)) }))] });
  };
  const sp = () => new Paragraph({ spacing: { after: 40 }, children: [] });
  const B = (t) => run(t, { bold: true });

  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [run('インフルエンザ予防接種 受付', { size: SZ + 6, color: NAVY, bold: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [run(title, { size: SZ + 22, color: TEAL, bold: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BORDER, space: 6 } }, children: [run(subtitle, { size: SZ - 1, color: '5A6B7A' })] }),
    ...build({ run, p, h1, h2, steps, bullets, box, table, sp, B }),
  ];
  return new Document({
    creator: '事務局', title: 'インフルエンザ予防接種 ' + title,
    styles: { default: { document: { run: { font: FONT, size: SZ } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, bold: true }, paragraph: { outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, bold: true }, paragraph: { outlineLevel: 1 } }] },
    numbering: { config: [
      { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
      ...refs.map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 600, hanging: 380 } }, run: { font: FONT, bold: true, color: TEAL } } }] })) ] },
    sections: [{ properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1134, bottom: 1134, left: MARGIN, right: MARGIN } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 17, color: '8299AB' }), new TextRun({ text: ' / ', font: FONT, size: 17, color: '8299AB' }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 17, color: '8299AB' })] })] }) },
      children }] });
}

// ───────────────────────── 事務員用 ─────────────────────────
const clerk = makeDoc('事務員用マニュアル', '名簿を受け取ってから、現場に渡すファイルを作るまで', 21, ({ run, p, h1, h2, steps, bullets, box, table, sp, B }) => [
  box([[B('ルール：'), run('1事業所・1日につき1ファイル。日付シートは当日に現場で作るので、事務員は作らない。')]]),
  h1('1. ファイルを作る'),
  ...steps([
    '「インフル予防接種受付_テンプレート.xlsm」をコピーする（テンプレート本体は触らない）。',
    [run('名前を変える。例：'), B('20261105_KENWOOD.xlsm'), run('（実施日_事業所名。.xlsm のまま）')],
    '開いて、上の黄色い帯の「コンテンツの有効化」を押す。',
  ]),
  box(['赤い帯（セキュリティ リスク）が出たら：Excel を閉じる → ファイルを右クリック →「プロパティ」→ 一番下の「許可する」にチェック → OK → 開き直す。'], true),
  h1('2. 会社マスタに事業所名を入れる'),
  ...steps(['「会社マスタ」シートの B4 に事業所名を入れる。', '名簿の会社名と1文字も違わないようにする（全角・半角、スペースも同じに）。']),
  h1('3. 受診者マスタに名簿を貼る'),
  p('申込名簿から1列ずつコピーし、受診者マスタに「値」で貼る（右クリック → 貼り付けのオプション → 123 のアイコン）。'),
  table(['貼る列', '項目', ''], [
    ['D', '氏名', '必ず'], ['E', 'カナ氏名', '必ず（プルダウンの並び順になる）'], ['I', '会社名', '必ず（会社マスタと同じ文字）'],
    ['F・G', '性別・生年月日', 'あれば（生年月日を入れると年齢は自動）'], ['O', 'フリー3', 'あれば（受診者番号など。番号で検索できる）'],
  ], [1.2, 2.5, 6]),
  sp(),
  box([[B('貼らない列：'), run('B（受付済み）・C（検索キー）・K（正式コース名）・L（年齢）。数式なので消さない。')]], true),
  p('会社名は I4 に1件入れて、セル右下の■をダブルクリックすると最後まで埋まる。'),
  h1('4. 確認して保存する'),
  ...steps([
    '「入力シート」を開き、会社名の▼に事業所名が出るか見る。',
    '会社名を選び、氏名の▼に名簿の人がカナ順で出るか見る。',
    '氏名を選び、カナ氏名・生年月日・年齢が表示されるか見る。',
    '確認できたら「クリア」を押す。',
    'Ctrl + S で保存して、現場の PC に入れる。',
  ]),
  box([[B('「受付登録」は押さない。'), run('登録まで試したいときは、ファイルをもう1つコピーし、コピーの方で表紙の「日付シート作成」→ OK → 受付登録を試す。終わったらコピーは捨てる。')]], true),
  h1('5. 間違えたとき'),
  table(['こんなとき', 'こうする'], [
    ['貼る列を間違えた', 'すぐなら Ctrl + Z で戻す。後で気づいたら、その列を選んで Delete で消し、正しく貼り直す。'],
    ['B・C・K・L 列を消してしまった', 'Ctrl + Z で戻す。戻せなければテンプレートのコピーから作り直す。'],
    ['氏名の▼に誰も出ない', '会社マスタの事業所名と、受診者マスタ I 列の会社名が同じ文字か確認する。'],
    ['年齢が出ない', '生年月日が空、または文字で貼られている。日付として入れ直す。'],
    ['名簿に人を追加・削除したい', '追加は、名簿の一番下のすぐ下の空いている行に入力する。削除は、その人の氏名・カナ氏名・会社名などのセルを選んで Delete で消す。'],
    ['確認で「受付登録」を押してしまった', '①一番右の日付シートを削除 ②受診者マスタ B 列の「済」を消す ③入力シートの受付No.（C3）を 1 に戻す。'],
  ], [3.2, 6.8]),
]);

// ───────────────────────── 現場用 ─────────────────────────
const field = makeDoc('現場用マニュアル', '当日の受付のやり方', 24, ({ run, p, h1, h2, steps, bullets, box, table, sp, B }) => [
  h1('1. 朝の準備'),
  ...steps([
    'パソコンを起動し、事務員から渡されたファイル（例：20261105_KENWOOD.xlsm）を開く。',
    '上の黄色い帯の「コンテンツの有効化」を押す。',
    [run('「表紙」の'), B('「日付シート作成」'), run('を押す → 今日の日付が入っているので'), B('そのまま OK'), run('。')],
    '「作成しました」と出たら、下の「入力シート」を開く。',
  ]),
  box(['日付はパソコンの今日の日付が自動で入ります。入力シートの日付も触らないでください。'], true),
  h1('2. 受付（1人ずつくり返す）'),
  ...steps([
    [B('会社名'), run(' の▼から事業所を選ぶ。')],
    [B('氏名'), run(' の▼から来た人を選ぶ（カナ順）。番号で探すときは「フリー3で検索」。')],
    'カナ氏名・生年月日を本人に確認する。',
    [B('「受付登録」'), run('を押す →「はい」→「OK」。')],
    '入力欄が空になり、受付No.が次の番号になる。次の人へ。',
  ]),
  p([run('受付した人は、一番右の日付シートにたまります。人数は日付シート上部の「受付人数」で見られます（見終わったら「入力シート」に戻る）。')]),
  h1('3. 終わったら'),
  p('Ctrl + S で保存して、ファイルを閉じる。'),
  h1('4. 間違えたとき'),
  h2('登録する前に気づいた'),
  p('「クリア」を押してやり直す。確認画面で「いいえ」を押しても取り消せる。'),
  h2('違う人を登録してしまった'),
  ...steps([
    '一番右の日付シートで、間違えた人の行番号（左端の数字）を右クリック →「削除」。',
    '「受診者マスタ」で、その人の B 列の「済」を消す（選んで Delete）。',
    '「入力シート」に戻る（受付No.は自動で直る）。正しい人を登録し直す。',
  ]),
  box(['その日の最初の1人を消したときだけ、受付No.（C3）に 1 を手で入れる。'], true),
  h2('その他'),
  table(['こんなとき', 'こうする'], [
    ['「シートが存在しません」と出た', '朝の準備の3（日付シート作成）をしていない。表紙でボタンを押して OK。'],
    ['「既に受付済みです」と出た', 'もう受付が済んでいる人。「いいえ」を押して、本人に確認する。'],
    ['名簿に名前が無い人が来た', '「受診者マスタ」で、名簿の一番下のすぐ下の空いている行に、氏名（D列）・カナ氏名（E列）・会社名（I列）を入れる → 入力シートに戻ると選べる。'],
    ['ボタンを押しても動かない', 'ファイルを開き直して「コンテンツの有効化」を押す。'],
  ], [3.4, 6.6]),
]);

const out = process.argv[2] || '.';
Promise.all([
  Packer.toBuffer(clerk).then(b => fs.writeFileSync(path.join(out, '事務員用マニュアル.docx'), b)),
  Packer.toBuffer(field).then(b => fs.writeFileSync(path.join(out, '現場用マニュアル.docx'), b)),
]).then(() => console.log('ok'));
