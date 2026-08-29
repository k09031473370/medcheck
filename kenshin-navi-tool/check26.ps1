<#
  受付番号の振られ方を調べる (check26.ps1)
  check26.bat をダブルクリックすると実行され、結果 r_check26.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check25 で分かったこと:
    ・08/21(城西学園)は98人全員 SEQ1 が空。まだ受付処理をしていない。
      これが自動判定エラー・文字色の違い・報告書0件 すべての原因。ツールは無実。
    ・SEQ1 = 受診日8桁 + 予約No 4桁 (受付番号ではなかった)
    ・08/20(院内)の受付番号は 1595〜1614 というクリニック全体の通し番号

  ここで確かめたい心配ごと:
    ツールは秋葉さんに受付番号 6 を入れた。
    もし健診ナビの受付処理が受付番号を自前で採番し直すなら、
    取り込んだ受付番号が上書きされ、血液データの紐付けが崩れる。

    巡回健診では日ごとに 1 から振るのか、院内と同じ通し番号なのか。
    城西学園の過去の受診データを見れば分かる。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check26.txt'

# 城西学園
$DANTAI = '0000000370'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 受付番号の振られ方 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★本命: 城西学園の過去の受診 (去年はどう振られたか) ---' @"
SELECT TOP 60 s.D_KENSIN AS 受診日, COUNT(*) AS 人数,
       MIN(s.UKE_NO_KENSA) AS 受付番号の最小,
       MAX(s.UKE_NO_KENSA) AS 受付番号の最大,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL THEN 1 ELSE 0 END) AS 受付番号なし,
       MIN(s.SEQ1) AS SEQ1の最小, MAX(s.SEQ1) AS SEQ1の最大
FROM T_KENSIN s
WHERE s.DANTAI_CD1 = '$DANTAI' AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN
ORDER BY s.D_KENSIN DESC
"@ 80

Q '--- 2. 城西学園の去年の受診 明細 (先頭20人。1から振られているか通し番号か) ---' @"
SELECT TOP 20 s.PK_SEQ, s.D_KENSIN AS 受診日, s.SEQ1, s.SEQ2,
       s.UKE_NO_KENSA AS 受付番号
FROM T_KENSIN s
WHERE s.DANTAI_CD1 = '$DANTAI' AND s.F_TORIKESI = 0
  AND s.D_KENSIN < '2026/08/21'
ORDER BY s.D_KENSIN DESC, s.SEQ2
"@ 30

Q '--- 3. 受付番号は日ごとに1から振り直すのか、ずっと通し番号か (日ごとの範囲) ---' @"
SELECT TOP 40 s.D_KENSIN AS 受診日, COUNT(*) AS 人数,
       MIN(s.UKE_NO_KENSA) AS 最小, MAX(s.UKE_NO_KENSA) AS 最大
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/06/01' AND s.D_KENSIN <= '2026/08/28'
  AND s.UKE_NO_KENSA IS NOT NULL
GROUP BY s.D_KENSIN
ORDER BY s.D_KENSIN
"@ 60

Q '--- 4. 同じ受付番号が別の日にも使われているか (使い回しなら日ごと、ユニークなら通し番号) ---' @"
SELECT TOP 20 s.UKE_NO_KENSA AS 受付番号, COUNT(DISTINCT s.D_KENSIN) AS 使われた日数,
       COUNT(*) AS 件数
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.UKE_NO_KENSA IS NOT NULL
GROUP BY s.UKE_NO_KENSA
ORDER BY COUNT(DISTINCT s.D_KENSIN) DESC
"@ 30

Q '--- 5. 巡回らしい日 (人数が多い日) の受付番号の振られ方 ---' @"
SELECT TOP 30 s.D_KENSIN AS 受診日,
       LTRIM(RTRIM(s.DANTAI_CD1)) AS 団体CD, d.MEISYO1 AS 団体名,
       COUNT(*) AS 人数,
       MIN(s.UKE_NO_KENSA) AS 受付最小, MAX(s.UKE_NO_KENSA) AS 受付最大
FROM T_KENSIN s
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.UKE_NO_KENSA IS NOT NULL
  AND s.D_KENSIN >= '2026/04/01'
GROUP BY s.D_KENSIN, LTRIM(RTRIM(s.DANTAI_CD1)), d.MEISYO1
HAVING COUNT(*) >= 15
ORDER BY COUNT(*) DESC
"@ 40

Q '--- 6. 受付番号が入っていない過去の受診はどれくらいあるか ---' @"
SELECT CASE WHEN UKE_NO_KENSA IS NULL THEN N'受付番号なし' ELSE N'受付番号あり' END AS 区分,
       COUNT(*) AS 件数, MIN(D_KENSIN) AS 最古, MAX(D_KENSIN) AS 最新
FROM T_KENSIN
WHERE F_TORIKESI = 0
GROUP BY CASE WHEN UKE_NO_KENSA IS NULL THEN N'受付番号なし' ELSE N'受付番号あり' END
"@ 10

W ''
W '=== 読み方 ==='
W '  1・2 が本命です。城西学園の去年の受診で受付番号が'
W '     1, 2, 3 … と 1 から振られていた → 巡回は日ごとに1から。'
W '        ツールが入れた 6 は正しい形。受付処理をしても上書きされない可能性が高い。'
W '     1600 台のような通し番号だった → 受付処理で採番し直される恐れがある。'
W '        その場合は「受付処理を先、取込を後」の順にすれば安全です。'
W ''
W '  4 で同じ受付番号が何日も使われている → 日ごとに1から振り直す方式。'
W '     1日しか使われていない番号ばかり → クリニック全体の通し番号。'
W ''
W '  結論がどちらでも、本番の手順は決まります:'
W '     日ごとに1から  → 今のまま「取込 → 受付処理 → 自動判定」でOK'
W '     通し番号      → 「受付処理 → 取込 → 自動判定」の順に変更'

notepad $out
