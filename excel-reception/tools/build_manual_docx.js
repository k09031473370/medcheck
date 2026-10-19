const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  HeadingLevel, AlignmentType, BorderStyle, ShadingType, LevelFormat, PageBreak,
  Header, Footer, PageNumber, TabStopType,
} = require('docx');

const FONT = 'Yu Gothic';
const TEAL = '0A7B6C', NAVY = '0F2847', PALE = 'E6F5F3', GRAY = 'F2F5F7', BORDER = 'B8CBD4';
const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - MARGIN * 2; // A4

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const p = (t, o = {}) => new Paragraph({ spacing: { after: 100, line: 300 }, ...o, children: typeof t === 'string' ? [run(t, o.run || {})] : t });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, children: [run(t, { bold: true, size: 30, color: 'FFFFFF' })], shading: { type: ShadingType.CLEAR, fill: TEAL, color: 'auto' }, indent: { left: 0 } });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 }, children: [run(t, { bold: true, size: 24, color: TEAL })], border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 2 } } });
const num = (t, ref) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 300 }, children: typeof t === 'string' ? [run(t)] : t });
const bul = (t) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60, line: 300 }, children: typeof t === 'string' ? [run(t)] : t });
const note = (lines) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
  rows: [new TableRow({ children: [new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: PALE, color: 'auto' },
    borders: { top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL }, left: { style: BorderStyle.SINGLE, size: 24, color: TEAL } },
    margins: { top: 100, bottom: 100, left: 160, right: 160 },
    children: lines.map((l, i) => new Paragraph({ spacing: { after: i === lines.length - 1 ? 0 : 60, line: 280 }, children: [run((i === 0 ? '※ ' : '　 ') + l, { size: 19 })] })),
  })] })],
});
const spacer = () => new Paragraph({ spacing: { after: 60 }, children: [] });

function table(headers, rows, widths, boldCol = -1) {
  const cell = (t, w, hdr, i, ri) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: hdr ? NAVY : (ri % 2 ? GRAY : 'FFFFFF'), color: 'auto' },
    borders: ['top', 'bottom', 'left', 'right'].reduce((a, k) => (a[k] = { style: BorderStyle.SINGLE, size: 4, color: BORDER }, a), {}),
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { after: 0, line: 260 }, children: [run(t, { size: 19, bold: hdr || i === boldCol, color: hdr ? 'FFFFFF' : '000000' })] })],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true, i, 0)) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false, i, ri)) })),
    ],
  });
}
const scale = (parts) => { const s = parts.reduce((a, b) => a + b, 0); const w = parts.map(x => Math.floor(CONTENT_W * x / s)); w[w.length - 1] += CONTENT_W - w.reduce((a, b) => a + b, 0); return w; };

const children = [];
// Title block
children.push(new Paragraph({ spacing: { before: 1200, after: 200 }, alignment: AlignmentType.CENTER, children: [run('インフルエンザ予防接種', { size: 40, bold: true, color: NAVY })] }));
children.push(new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.CENTER, children: [run('受付ファイル作成手順書', { size: 52, bold: true, color: TEAL })] }));
children.push(new Paragraph({ spacing: { after: 600 }, alignment: AlignmentType.CENTER, children: [run('事務員向け', { size: 28, color: NAVY })] }));
children.push(new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER, children: [run('対象ファイル：インフル予防接種受付_テンプレート.xlsm', { size: 21 })] }));
children.push(new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER, children: [run('必要環境：Excel 2016 以降（Microsoft 365 も可）・マクロ有効', { size: 21 })] }));
children.push(new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER, children: [run('作成日：2026年10月', { size: 21 })] }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// Overview
children.push(h1('1. この手順書について'));
children.push(p('事業所から申込名簿（Excel）を受け取ったら、この手順で「その事業所・その実施日用の受付ファイル」を1つ作ります。現場のスタッフはそのファイルを開いて、当日の受付を行います。'));
children.push(bul([run('ルール：', { bold: true }), run('1事業所 × 1実施日 ＝ 1ファイル（同じ事業所で複数日ある場合は第10章参照）')]));
children.push(bul([run('所要時間：', { bold: true }), run('慣れれば5分程度')]));
children.push(bul([run('必要なもの：', { bold: true }), run('Excel、テンプレート「インフル予防接種受付_テンプレート.xlsm」、申込名簿')]));
children.push(spacer());
children.push(h2('ファイルの構成（参考）'));
children.push(table(['シート名', '役割'], [
  ['表紙', '説明と「日付シート作成」「雑費請求書作成」ボタン'],
  ['入力シート', '当日の受付画面。会社名→氏名の連動プルダウン（氏名はカナ順）、受付No.自動採番、「受付登録」ボタン。健診用の便潜血・オプション・備考欄は非表示にしてある'],
  ['受診者マスタ', '申込名簿を貼り付ける場所。受付が済むと「済」が付く。年齢は生年月日から自動計算'],
  ['会社マスタ', '事業所名の一覧。入力シートの会社名プルダウンの元'],
  ['検診コースマスタ', '健診用の設定。インフルでは使わない'],
  ['【原本】受付名簿', '日付シートの雛形。触らない'],
  ['【原本】雑費請求書', '健診用の請求書雛形。インフルでは使わない'],
  ['YYYYMMDD（例 20261105）', '実施日ごとの受付名簿。「日付シート作成」ボタンで作る'],
], scale([3, 7]), 0));

// Step 1
children.push(h1('2. 手順1　テンプレートをコピーして名前を付ける'));
children.push(num('「インフル予防接種受付_テンプレート.xlsm」を右クリック → コピー → 貼り付け でコピーします。テンプレート本体は直接編集しないでください。', 'steps1'));
children.push(num([run('コピーしたファイルの名前を変えます。例：'), run('20261105_KENWOOD.xlsm', { bold: true }), run('（実施日8桁_事業所名）')], 'steps1'));
children.push(num('拡張子は「.xlsm」のままにします。「.xlsx」にするとボタンが動かなくなります。', 'steps1'));

// Step 2
children.push(h1('3. 手順2　ファイルを開いてマクロを有効にする'));
children.push(p('このファイルはボタン（マクロ）で動くため、開いたときに表示される帯の種類によって対応が変わります。'));
children.push(table(['開いたときの表示', '対応'], [
  ['黄色い帯「セキュリティの警告　マクロが無効にされました」', '「コンテンツの有効化」を押す。これで使えます'],
  ['赤い帯「セキュリティ リスク　このファイルのソースが信頼できないため、Microsoft によりマクロの実行がブロックされました」', '下の「赤い帯が出たとき」の手順でブロックを解除する（ファイルごとに1回だけ）'],
  ['帯が出ない', 'そのまま使えます'],
], scale([4.5, 5.5]), -1));
children.push(spacer());
children.push(h2('赤い帯が出たとき（ブロックの解除）'));
children.push(p('メールやダウンロードで受け取ったファイルは、Windows が「他のコンピューターから取得したファイル」として扱い、マクロを止めます。次の手順で解除します。'));
children.push(num('Excel を閉じます。', 'steps2'));
children.push(num('エクスプローラーでそのファイルを右クリック → 「プロパティ」。', 'steps2'));
children.push(num('「全般」タブの一番下、「セキュリティ：このファイルは他のコンピューターから取得したものです…」の右にある「許可する」にチェックを入れて OK。', 'steps2'));
children.push(num('もう一度ファイルを開きます。黄色い帯に変わるので「コンテンツの有効化」を押します。', 'steps2'));
children.push(note(['テンプレートで一度解除しておけば、そこからコピーした各日ファイルには引き継がれます。ただし、新しくダウンロードしたファイルはそのたびに解除が必要です。']));
children.push(h2('おすすめ：受付ファイル用フォルダを「信頼できる場所」に登録'));
children.push(p('一度設定すると、そのフォルダに置いたファイルは赤い帯が出なくなり、毎回の解除が不要になります。事務員の PC と現場の PC の両方で設定しておくと安心です。'));
children.push(num('受付ファイルを置くフォルダを作る（例：ドキュメントの中に「インフル受付」）。', 'steps2b'));
children.push(num('Excel の「ファイル → オプション → トラストセンター → トラストセンターの設定」を開く。', 'steps2b'));
children.push(num('「信頼できる場所」→「新しい場所の追加」→「参照」で、作ったフォルダを選ぶ。', 'steps2b'));
children.push(num('「この場所のサブフォルダーも信頼する」にチェック → OK。', 'steps2b'));
children.push(num('以後、テンプレートや受付ファイルはこのフォルダに入れてから開く。', 'steps2b'));

// Step 3
children.push(h1('4. 手順3　会社マスタに事業所名を入れる'));
children.push(num('「会社マスタ」シートを開きます。', 'steps3'));
children.push(num('B4 セルに事業所名を入力します（例：KENWOOD）。', 'steps3'));
children.push(num('名簿に複数の事業所が混ざっている場合は B5、B6 … と続けて入力します。', 'steps3'));
children.push(note(['ここに入れた名前が、入力シートの「会社名」プルダウンの選択肢になります。', '手順4で貼る名簿の会社名と、文字が完全に一致している必要があります（全角/半角、スペースも含む）。']));

// Step 4
children.push(h1('5. 手順4　申込名簿を「受診者マスタ」に貼り付ける'));
children.push(p('「受診者マスタ」シートの表（4行目以降）に、申込名簿の内容を列ごとに貼り付けます。申込名簿の列の並びはテンプレートと違うので、まとめて貼らず、1列ずつ「値のみ貼り付け」します。'));
children.push(table(['列', '項目', '貼り付け', '補足'], [
  ['B', '受付済み', '貼らない', '当日の受付で自動的に「済」が付く'],
  ['C', '検索キー', '貼らない（数式）', '消すと入力シートが動かなくなる'],
  ['D', '氏名', '必ず貼る', '姓と名の間は全角スペース推奨'],
  ['E', 'カナ氏名', '必ず貼る', '氏名プルダウンの並び順（カナ順）に使う'],
  ['F', '性別', 'あれば貼る', '「男性」「女性」'],
  ['G', '生年月日', 'あれば貼る', '日付として貼る。入れると L列の年齢が自動計算される'],
  ['H', '健保名', 'あれば貼る', ''],
  ['I', '会社名', '必ず貼る', '手順3で入れた名前と同じ文字にする'],
  ['J', '検診コース', '空欄のまま', '健診用。インフルでは使わない'],
  ['K, L', '正式コース名・年齢', '貼らない（数式）', '自動列。年齢は生年月日から、その年度末（3/31）時点で計算'],
  ['M', 'フリー1', '任意', '去年は事業所コードを入れていた'],
  ['N', 'フリー2', '任意', '去年は社員番号を入れていた'],
  ['O', 'フリー3', '任意', '去年は受診者番号。入力シートの「フリー3で検索」ボタンで使える'],
], scale([1, 2.6, 2.4, 5.5]), 2));
children.push(spacer());
children.push(h2('貼り付けのやり方'));
children.push(num('申込名簿で氏名の列を選択してコピー。', 'steps4a'));
children.push(num('受診者マスタの D4 セルを選んで、右クリック → 貼り付けのオプション → 「値」（123 のアイコン）で貼り付け。', 'steps4a'));
children.push(num('カナ氏名（E列）、性別（F列）、生年月日（G列）、会社名（I列）、フリー1〜3（M〜O列）も同じように貼る。', 'steps4a'));
children.push(num('会社名が全員同じなら、I4 に1件入れて、セル右下の■をダブルクリックすると最終行まで一気に埋まります。', 'steps4a'));
children.push(h2('貼り付け後の確認'));
children.push(bul('C列（検索キー）に「KENWOOD_山田　太郎」のように「会社名_氏名」が全員分表示されている。'));
children.push(bul('生年月日を貼った人は、L列に年齢が表示されている。'));
children.push(bul('表の最終行より下に貼った場合、表（青い枠）が自動で広がっている。広がっていなければ、表の右下の小さな「┘」マークをドラッグして広げる。'));

// Step 5
children.push(h1('6. 手順5　実施日のシートを作る'));
children.push(num('「表紙」シートに戻ります。', 'steps5'));
children.push(num('右側の「✨ 日付シート作成」ボタンを押します。', 'steps5'));
children.push(num([run('実施日を8桁の数字で入力します。例：2026年11月5日 → '), run('20261105', { bold: true })], 'steps5'));
children.push(num('「作成しました」と出たら OK。シートの一番右に「20261105」という名前のシートができます。', 'steps5'));
children.push(note(['このシートが無いと、当日の「受付登録」でエラーになります。必ず実施日の分を作ってください。', '同じ事業所で実施日が2日ある場合は、このボタンを2回押して2枚作ります。1ファイルで2日分を管理でき、受付No.は2日通しで連番になります。']));

// Step 6
children.push(h1('7. 手順6　動作確認'));
children.push(num('「入力シート」を開きます。', 'steps6'));
children.push(num('右上の受付No.が「1」になっている（初回はこの番号から採番されます。別の番号から始めたい場合は C3 セルを書き換える）。', 'steps6'));
children.push(num('「会社名」のプルダウンに手順3で入れた事業所名が出る。', 'steps6'));
children.push(num('会社名を選ぶと「氏名」のプルダウンに、その会社の人がカナ順で出る。', 'steps6'));
children.push(num('氏名を選ぶと、カナ氏名・性別・生年月日・年齢・フリー1〜3 が自動表示される。', 'steps6'));
children.push(num('確認が終わったら「🧹 クリア」ボタンで入力を消す。', 'steps6'));
children.push(note(['確認で「受付登録」は押さないでください。押してしまった場合は、日付シートのその行を削除し、受診者マスタの「済」を消し、入力シートの受付No.を 1 に戻してください。']));

// Step 7
children.push(h1('8. 手順7　保存して現場へ渡す'));
children.push(num('Ctrl + S で保存（形式は .xlsm のまま）。', 'steps7'));
children.push(num('現場で使う PC・共有フォルダにコピーします。', 'steps7'));
children.push(num('OneDrive 上で直接使う場合、同期の遅れでエラーが出ることがあるため、できればローカルにコピーして使ってください。', 'steps7'));

// Day-of flow
children.push(h1('9. 当日の受付の流れ（現場スタッフ向けメモ）'));
children.push(num('ファイルを開き「コンテンツの有効化」。', 'steps9'));
children.push(num('「入力シート」で会社名 → 氏名をプルダウンから選ぶ（または「🔍 フリー3で検索」で番号入力）。カナ氏名・生年月日などで本人確認する。', 'steps9'));
children.push(num('「➕ 受付登録」を押す → 確認ダイアログで「はい」。', 'steps9'));
children.push(num('日付シートに1行追加され、入力欄が自動で消え、受付No.が次の番号になる。', 'steps9'));
children.push(num('同じ人を二度登録しようとすると「既に受付済みです」と警告が出る（通常は「いいえ」）。', 'steps9'));
children.push(num('日付シート上部の「受付人数」で当日の人数が確認できます。', 'steps9'));

// Troubleshooting
children.push(h1('10. よくあるトラブルと対処'));
children.push(table(['症状', '原因と対処'], [
  ['ボタンを押しても何も起きない', 'マクロが無効。手順2を確認。拡張子が .xlsx になっていないか確認'],
  ['「○月○日のシートが存在しません」', '実施日のシートが未作成。手順5を実施。入力シートの日付（F2）は今日の日付が自動で入るため、前日にテストする場合は F2 を実施日に手で書き換える'],
  ['氏名のプルダウンに誰も出ない', '受診者マスタの会社名（I列）と会社マスタの名前が一致していない。または C列の検索キー数式が消えている'],
  ['カナ氏名などが「#NAME?」になる', '古いテンプレート（去年のファイル）を使っている。このテンプレートで作り直す'],
  ['年齢が空欄', '受診者マスタの生年月日（G列）が空、または文字として貼られている。日付として入れ直す'],
  ['「400」などのエラーが出る', 'Excel が古い PC で去年のファイルを開いている。このテンプレートで作り直す'],
  ['受付No.が 1 から始まらない', '動作確認で登録してしまった、または日付シートにデータが残っている。日付シートの行を削除し、C3 を 1 に戻す'],
  ['「既に受付済み」が誤って出る', '同姓同名が名簿にいる。確認のうえ「はい」で登録を続ける'],
  ['便潜血・オプションの欄を使いたい', '入力シートの19〜33行目が非表示になっている。行番号18と34の間を右クリック →「再表示」で出せる。行は削除しないこと（ボタンが動かなくなる）'],
], scale([3.5, 6.5]), 0));
children.push(spacer());
children.push(h2('複数事業所・複数日の整理'));
children.push(bul([run('事業所が違う → ', { bold: true }), run('ファイルを分ける（手順1からやり直す）')]));
children.push(bul([run('同じ事業所で日が違う → ', { bold: true }), run('同じファイルに日付シートを追加（手順5を日数分）')]));
children.push(bul('ファイル名は「実施日_事業所名.xlsm」で統一すると、あとで探しやすくなります。'));

const numbered = (ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } }, run: { font: FONT, bold: true, color: TEAL } } }] });

const doc = new Document({
  creator: '事務局',
  title: 'インフルエンザ予防接種 受付ファイル作成手順書',
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 30, bold: true, color: 'FFFFFF' }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 24, bold: true, color: TEAL }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
    ...['steps1','steps2','steps2b','steps3','steps4a','steps5','steps6','steps7','steps9'].map(numbered),
  ] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1300, bottom: 1300, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run('インフルエンザ予防接種 受付ファイル作成手順書', { size: 17, color: '8299AB' })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('', { size: 17 }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 17, color: '8299AB' }), run(' / ', { size: 17, color: '8299AB' }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 17, color: '8299AB' })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('ok', b.length); });
