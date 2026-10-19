# excel-reception

インフルエンザ予防接種（元は健康診断）の受付用 Excel マクロアプリ一式。

| ファイル | 用途 |
|---|---|
| `インフル予防接種受付_テンプレート.xlsm` | 名簿・受付データを空にしたテンプレート（Microsoft 365 / Excel 2021 以降用）。これをコピーして各事業所・各日のファイルを作る |
| `インフル予防接種受付_テンプレート_Excel2019以前用.xlsm` | 同上の Excel 2016 / 2019 互換版。XLOOKUP / FILTER / SORTBY を INDEX+MATCH と INDIRECT に置き換えてある。氏名プルダウンは名簿全員が名簿順で出る。フリー1集計は無し |
| `事務員向け_受付ファイル作成手順書.docx` | 事務員が名簿を受け取ってから現場用ファイルを作るまでの手順書（Word） |
| `tools/build_manual_docx.js` | 手順書 Word を生成するスクリプト（`node build_manual_docx.js 出力.docx`） |
| `tools/make_compat.py` | 新しい方のテンプレートから Excel2019 以前用を生成するスクリプト |
| `tools/make_template.py` | 使用済みファイル（去年の 1128 ファイル）からテンプレートを再生成したスクリプト。VBA を壊さないよう OOXML を直接編集している |

標準テンプレートは Microsoft 365 / Excel 2021 以降が前提（FILTER, SORTBY, XLOOKUP を使用）。Excel 2016 / 2019 では「Excel2019以前用」を使う。
