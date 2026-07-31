const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, LevelFormat,
} = require('docx');
const fs = require('fs');

const FONT = { ascii: 'Yu Gothic', eastAsia: 'Yu Gothic', hAnsi: 'Yu Gothic' };
const ACCENT = '1F4E79';
const LIGHT = 'DCE6F1';
const WARN = 'FDE9D9';
const OKBG = 'E2EFDA';

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.size || 21, bold: !!opts.bold, color: opts.color, ...opts });
}
function p(text, opts = {}) {
  return new Paragraph({
    children: Array.isArray(text) ? text : [run(text, opts)],
    spacing: { after: opts.after != null ? opts.after : 120, line: 300 },
  });
}
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 340, after: 160 },
    children: [run(text, { size: 30, bold: true, color: ACCENT })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 260, after: 120 },
    children: [run(text, { size: 24, bold: true, color: ACCENT })],
  });
}
function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: 'bul', level: 0 }, spacing: { after: 80, line: 300 },
    children: Array.isArray(text) ? text : [run(text, opts)],
  });
}
function num(ref, text) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 }, spacing: { after: 90, line: 300 },
    children: Array.isArray(text) ? text : [run(text)],
  });
}
function mono(text) {
  return new Paragraph({
    spacing: { after: 120 }, shading: { type: ShadingType.CLEAR, fill: 'F2F2F2' },
    children: [new TextRun({ text, font: { ascii: 'Consolas', eastAsia: 'MS Gothic', hAnsi: 'Consolas' }, size: 18 })],
  });
}
const TW = 9800;
function cell(text, w, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { after: 0, line: 280 }, children: [run(text, { bold: !!opts.bold, size: 19 })] })],
  });
}
function table(widths, rows) {
  return new Table({
    columnWidths: widths, width: { size: TW, type: WidthType.DXA },
    rows: rows.map((r, i) => new TableRow({
      children: r.map((c, j) => cell(c, widths[j], i === 0 ? { fill: LIGHT, bold: true } : {})),
    })),
  });
}
function box(lines, fill) {
  return new Table({
    columnWidths: [TW], width: { size: TW, type: WidthType.DXA },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: TW, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: fill },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: lines.map((l, i) => new Paragraph({
        spacing: { after: i === lines.length - 1 ? 0 : 60 },
        children: [run(l.text, { size: 19, bold: !!l.bold })],
      })),
    })] })],
  });
}

const numbering = {
  config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 240 } } } }] },
    ...['sA','sB','sC','sD','sE','sF','sG','sH','sI'].map(ref => ({
      reference: ref,
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 360 } } } }],
    })),
  ],
};

const children = [
  new Paragraph({ spacing: { before: 200, after: 60 }, alignment: AlignmentType.CENTER, children: [run('健診結果 取込マニュアル', { size: 44, bold: true, color: ACCENT })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [run('芝浦健診クリニック / KENSHIN navi 11', { size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [run('第3版 2026/07/31', { size: 19, color: '666666' })] }),

  // ---------------------------------------------------------------- 1
  h1('1. はじめに'),
  p('健診の結果をExcelから健診ナビへ自動で登録するツールです。手入力の代わりに使います。'),
  box([
    { text: '覚えておくこと', bold: true },
    { text: '・「プレビュー」で内容を確認してから「書込実行」。プレビューだけならデータは変わりません。' },
    { text: '・書込む前のデータは自動でバックアップされます(backupフォルダ)。' },
    { text: '・書込んだあとは、健診ナビで必ず「自動判定」を実行します。' },
  ], OKBG),
  p(''),
  p([run('ファイルを選べば、どちらのレイアウトかはツールが自動で判別します。', { bold: true }), run('(うまく判別されないときだけ「レイアウト」で手動選択)')]),
  p([run('ツールの場所: ', { bold: true }), run('C:\\Users\\User\\Documents\\excel_tools  (USBに入れて持ち歩くこともできます)')]),
  p([run('起動のしかた: ', { bold: true }), run('そのフォルダの start.bat をダブルクリック')]),

  // ---------------------------------------------------------------- 2
  h1('2. データの種類と作業の早見表'),
  p('扱うデータは3種類あります。どれをやるのかを最初に確認してください。'),
  table([2200, 3000, 4600], [
    ['種類', 'ファイルの見た目', 'やること'],
    ['リアンパターン', '氏名・受診日・結果が入った横長のExcel(106列)', '第3章。予約登録から結果取込まで一式'],
    ['芝浦巡回AIデータ', '受付NOと結果だけのExcel(44列)。氏名なし', '第4章。結果の取込だけ'],
    ['心電図の結果', '心電計から出るCSV(ecgresult〜)', '第5章'],
  ]),
  p(''),
  p([run('血液検査', { bold: true }), run(': リアンはSRLから届くExcelをこのツールで取り込みます(3-5)。それ以外は従来どおり江東微研から取り込みます。')]),

  // ---------------------------------------------------------------- 3
  h1('3. リアンパターン の手順'),
  p('氏名・受診日が入っているので、予約の登録から結果の登録まで通してできます。'),

  h2('3-1. 予約取込ファイルを作る 〈ツール〉'),
  p('※ 予約取込フォーマットに手入力済みのファイルがある場合は、この手順は飛ばして 3-2 へ。'),
  num('sA', 'start.bat をダブルクリックしてツールを開く。'),
  num('sA', '「参照...」で名簿のExcelを選ぶ。'),
  num('sA', '「予約取込ファイルを作成」を押す。'),
  num('sA', 'デスクトップに「予約取込_日付.xlsx」ができるので、開いて内容を確認する。'),

  h2('3-2. 受診者を登録する 〈健診ナビ〉'),
  num('sB', '健診ナビの「予約データ取込」を開く。'),
  num('sB', '3-1 で作ったExcel(または手入力済みのファイル)を読み込む。'),
  num('sB', '受診者と検査項目の枠が作られる。'),

  h2('3-3. 受付番号を入れる 〈ツール〉'),
  p('当日発行した受付番号を、氏名で照合して健診ナビへ一括で入れます。手入力は不要です。'),
  num('sC', 'ツールの「参照...」で当日の結果Excelを選ぶ(レイアウトは自動で判別されます)。'),
  num('sC', '「受付番号を設定」を押す → 一覧が出るので確認 → 「はい」で設定。'),
  p([run('※ すでに受付番号が入っている人は「設定済み(変更なし)」と出てそのままです。', { size: 19 })]),

  h2('3-4. 結果を取り込む 〈ツール〉'),
  num('sD', '同じファイル・同じレイアウトのまま、受付番号を入れる(空欄なら全員)。'),
  num('sD', [run('「受診日」は空欄のままにする。', { bold: true }), run(' リアンのExcelには日付の列があるので、そこから読みます。')]),
  num('sD', '「3. プレビュー」を押す。'),
  num('sD', '「状態」の列がすべて OK になっていることを確認する。'),
  num('sD', '「4. 書込実行」→ 確認の画面で「はい」。'),

  h2('3-5. 血液を取り込む 〈ツール〉'),
  p('SRLから届くExcel(パスワード付き)を使います。他社の受診者も混ざっているので、名簿でしぼり込みます。'),
  num('sH', '「参照...」でSRLのExcelを選ぶ(パスワードは自動で解除されます)。'),
  num('sH', [run('「名簿を選ぶ」でリアンのExcel(結果が入っているもの)を指定する。', { bold: true }), run(' ← これをしないと他社の人まで対象になります。')]),
  num('sH', '「受診日」は空欄でOK(SRLのファイルに受診日が入っています)。'),
  num('sH', '「3. プレビュー」→ 状態がすべて OK か確認する。'),
  num('sH', '「4. 書込実行」。'),
  p('名簿に無い人は「名簿に無い ○ 人は取り込みませんでした」と出ます。これは正常です。'),
  p([run('入る項目: ', { bold: true }), run('クレアチニン / 尿酸 / 総コレステロール / LDL / HDL / 中性脂肪 / AST / ALT / ALP / γ-GTP / 血糖 / HbA1c / 白血球 / 赤血球 / 血色素 / ヘマトクリット / 血小板 / eGFR / MCV / MCH / MCHC / 便潜血1回目・2回目', { size: 19 })]),

  h2('3-5b. 心電図を取り込む 〈ツール〉'),
  p([run('リアンのExcelは、心電図・所見(101列)に書いてある人だけが結果取込(3-4)で一緒に入ります。')]),
  p([run('所見が空の人', { bold: true }), run('は、「撮って異常なし」なのか「撮っていない」のか、Excelからは区別が付きません。そのまま異常なしを入れると、撮っていない人に正常の結果が付いてしまいます。そこで次のどちらかにしてください。')]),
  box([
    { text: '方法A(おすすめ・Excelに1を入れる)', bold: true },
    { text: '撮って異常が無かった人の 心電図・判定(105列) に 1 を入れてから取り込む。' },
    { text: '→ その人だけ「正常範囲」が自動で入ります。空欄の人には何も入りません。' },
    { text: '(所見(101列)の方に直接 1 と書いても同じ結果になります)' },
  ], OKBG),
  p(''),
  box([
    { text: '方法B(心電計のCSVから入れる)', bold: true },
    { text: '心電計から出るCSVを 心電図取込.bat に入れる(第5章)。撮った人しかCSVに出てこないので、こちらが一番確実です。' },
  ]),
  p('※ 全員が必ず撮ると確認できれば、判定欄に1を入れなくても自動で入るように変更できます。check7.bat で「枠がある人数」と「所見が入っている人数」を過去の分で比べられるので、いつも同じなら担当者に連絡してください。'),

  h2('3-6. 自動判定 〈健診ナビ〉'),
  p([run('結果も血液も入れ終わってから実行します。', { bold: true }), run('先に実行すると血液抜きの判定になります。')]),
  num('sE', '健診ナビで対象者を開き、値が入っていることを目で確認する。'),
  num('sE', '「自動判定」を実行する(BMI・標準体重・判定はここで計算されます)。'),
  num('sE', '東振協などの提出データを出力する。'),

  p(''),
  box([
    { text: '「受診日」欄について', bold: true },
    { text: '空欄にしておくと、Excelの日付の列から自動で読みます(リアンはこちら)。' },
    { text: '入力すると、Excelの日付より入力した日付が優先されます。その場合は確認の画面が出るので、内容を見て「はい」か「いいえ」を選んでください。' },
    { text: '日付の列が無いファイル(芝浦巡回AIデータ)では、受診日の入力が必須です(確認の画面は出ません)。' },
    { text: '書き方は 2026/07/12 です。260712 や 20260712 でも通りますが、迷ったら 2026/07/12 の形で入れてください。' },
  ]),

  // ---------------------------------------------------------------- 4
  h1('4. 芝浦巡回AIデータ の手順'),
  p('このExcelには氏名も受診日も入っていません。予約登録と受付番号の入力が済んでいることが前提です。'),
  h2('4-1. 結果を取り込む'),
  num('sF', 'start.bat でツールを開く。'),
  num('sF', '「参照...」でそのExcelを選ぶ(レイアウトは自動で判別されます)。'),
  num('sF', [run('「受診日」に実施日を入力する(例 2026/05/21)。', { bold: true }), run(' ← 日付の列が無いので必須です。')]),
  num('sF', '受付番号は空欄のまま(全員が対象)。'),
  num('sF', '「3. プレビュー」→ 状態がすべて OK なら「4. 書込実行」。'),
  num('sF', '健診ナビで「自動判定」を実行する。'),
  p(''),
  box([
    { text: '「受診者が見つかりません」と出たら', bold: true },
    { text: 'その受付番号が健診ナビにまだ入っていません。下の 4-2 で入れてから、もう一度実行してください。' },
  ], WARN),

  h2('4-2. 受付番号が入っていないとき'),
  p('芝浦巡回AIデータには氏名が無いので、そのファイルだけでは誰が何番か分かりません。氏名と受付番号の一覧を作って、ツールに入れさせます。'),
  num('sG', 'ツールのフォルダにある 受付番号リスト_ひな形.xlsx をデスクトップにコピーする。'),
  num('sG', '開いて、2行目から 受付番号・氏名(フリガナは任意)を入力して保存する。'),
  num('sG', 'ツールの「参照...」でそのファイルを選ぶ(レイアウトは自動で判別されます)。'),
  num('sG', [run('「受診日」に実施日を入力する(例 2026/05/21)。', { bold: true }), run(' ← このファイルには日付の列が無いので必須です。')]),
  num('sG', '「受付番号を設定」を押す → 一覧を確認 → 「はい」。'),
  p('そのあと 4-1 に戻って、AIデータの取込をやり直してください。'),
  p([run('※ 数人だけなら、健診ナビの受診者の画面で受付番号を直接入力しても構いません。', { size: 19 })]),
  p(''),
  box([
    { text: '「該当者なし」と出たら', bold: true },
    { text: 'その氏名がその受診日で健診ナビに登録されていません。受診日と、氏名の表記(スペース・旧字体)を確認してください。' },
  ], WARN),

  // ---------------------------------------------------------------- 5
  h1('5. 心電図の結果を取り込む'),
  p('心電計から出るCSVを使います。ID列がその日の受付番号、検査日列が受診日として扱われます。'),
  p([run('心電図取込.bat にCSVをドラッグ＆ドロップ', { bold: true }), run('してください。プレビューが出て、確認してから y を押すと書き込まれます。')]),
  p('(ダブルクリックしてからパスを貼り付けても同じです)'),
  bullet('所見が無く判定Aの人には「正常範囲」が自動で入ります。'),
  bullet('判定(A・B・C12など)も一緒に登録されます。'),
  bullet('1人だけにしたいときは -Only 13 のようにIDを付けます。'),

  // ---------------------------------------------------------------- 6
  h1('6. 困ったときは'),
  table([2000, 3600, 4200], [
    ['画面に出る言葉', '意味', 'どうするか'],
    ['枠なし', 'その人のコースにその検査が無い', '対象外の検査なら、そのままで問題ありません'],
    ['受診者が見つかりません', '受付番号か受診日が健診ナビと合っていない', '受診日と受付番号を確認する'],
    ['名簿でしぼり込んだ結果、対象が0人', '名簿と取込ファイルの受診日が違う', '画面に出る日付を見て、受診日欄に名簿と同じ日付を入れる'],
    ['氏名が一致しません', 'Excelの氏名と健診ナビの氏名が違う。受付番号が別人に付いている可能性', '受付番号を確認する。別人に書き込まないよう、その人は飛ばされます'],
    ['結果入力画面で編集中', 'その人の画面が健診ナビで開いている', '健診ナビでその人の画面を閉じて、もう一度実行'],
    ['変換表に無いコード', 'Excelの値がまだ登録されていない', '担当者に連絡(対応表の追加が必要)'],
    ['所見CD未登録', '所見のコードが健診ナビに無い', '担当者に連絡'],
    ['接続エラー', 'サーバーに繋がっていない', 'VPN・ネットワークを確認。それでも駄目なら担当者へ'],
  ]),
  p(''),
  p('エラーが出た人は書き込まれません。他の人の処理は続きます。'),

  // ---------------------------------------------------------------- 7
  h1('7. やってはいけないこと'),
  box([
    { text: '・プレビューを見ないで「書込実行」を押さない', bold: true },
    { text: '・取込む人の結果入力画面を健診ナビで開いたままにしない(書いた内容が消えます)' },
    { text: '・「エラーを飛ばす」「氏名の違いを無視する」のチェックは、理由が分かっているとき以外は使わない' },
    { text: '・form フォルダの中のCSVをExcelで編集しない(先頭の 0 が消えて壊れます)' },
    { text: '・江東微研のコード変換マスタ、検査センター.csv は変更しない' },
  ], WARN),

  // ---------------------------------------------------------------- 8
  h1('8. 元に戻したいとき'),
  p('書込むたびに、その人の書込前のデータが backup フォルダにCSVで保存されています。'),
  p('1〜2項目だけなら、健診ナビの画面から手で直すのが一番早いです。'),
  p([run('まとめて戻したいときは 元に戻す.bat をダブルクリック', { bold: true }), run('してください。')]),
  num('sI', 'バックアップの一覧が出ます。'),
  num('sI', '戻したいファイル名をコピーして貼り付ける。'),
  num('sI', '何が戻るかが表示されます(まだ戻りません)。'),
  num('sI', 'y を押すと戻ります。'),
  p('戻す前の状態も保存されるので、やり直せます。'),
  p([run('戻したあとは健診ナビで自動判定をやり直してください。', { bold: true })]),
  p('判断に迷うときは担当者に相談してください。'),

  h1('付録: ボタンの意味'),
  table([2600, 7200], [
    ['ボタン', '何をするか'],
    ['参照...', '取り込むExcelやCSVを選ぶ'],
    ['1. 列確認', 'Excelの何列目が何の項目かを表示する(確認用・データは変わらない)'],
    ['2. 枠一覧(DB)', 'その人の検査項目の一覧を健診ナビから読んで表示する(確認用)'],
    ['受付番号を設定', '氏名で照合して、受付番号を健診ナビに入れる'],
    ['予約取込ファイルを作成', '名簿から、健診ナビの予約データ取込で読ませるExcelを作る'],
    ['3. プレビュー', '何がどう書き込まれるかを表示する(データは変わらない)'],
    ['4. 書込実行', '実際に健診ナビへ書き込む'],
    ['名簿でしぼり込む', '名簿(リアン等)を指定すると、その名簿に載っている人だけを取り込む。空欄ならファイル内の全員'],
    ['氏名の違いを無視する', '通常はチェックしない。Excelと健診ナビで氏名が違う人を、あえて取り込むときだけ使う'],
  ]),

  h1('付録: 同梱ファイル'),
  table([3400, 6400], [
    ['ファイル', '何に使うか'],
    ['受付番号リスト_ひな形.xlsx', '氏名と受付番号だけの一覧。芝浦巡回など、結果ファイルに氏名が無いときに受付番号を入れるために使う'],
    ['予約取込フォーマット.xlsx', '「予約取込ファイルを作成」で使うテンプレート。健診ナビからもらった原本をこの名前で置く'],
    ['心電図取込.bat', '心電図CSVの取込。CSVをドラッグ＆ドロップ(第5章)'],
    ['元に戻す.bat', '書込前のバックアップから元に戻す(第8章)'],
    ['USBで使うときは.txt', 'USBに入れて使うときの注意'],
  ]),
];

const doc = new Document({
  numbering,
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{
    properties: { page: { margin: { top: 1100, bottom: 1100, left: 1100, right: 1100 } } },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/user/medcheck/kenshin-navi-tool/取込マニュアル.docx', buf);
  console.log('written', buf.length, 'bytes');
});
