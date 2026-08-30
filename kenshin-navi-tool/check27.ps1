<#
  片付けが済んだか確認する (check27.ps1)
  check27.bat をダブルクリックすると実行され、結果 r_check27.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  確かめること:
    テストで使った秋葉さん(PK_SEQ 2005094)が、
    同じ日の他の97人と同じ状態に戻っているか。

    「他の人と見分けがつかない」なら片付け完了です。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check27.txt'

$PK = 2005094

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 片付け確認 (PK_SEQ $PK) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★受付まわり: 秋葉さん vs 同じ日の他の人 ---' @"
SELECT N'★秋葉' AS 区分, s.PK_SEQ,
       '[' + ISNULL(s.SEQ1,'NULL') + ']' AS SEQ1,
       '[' + ISNULL(s.SEQ2,'NULL') + ']' AS SEQ2,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),'NULL') + ']' AS 受付番号,
       s.F_UKETUKE AS 受付フラグ
FROM T_KENSIN s WHERE s.PK_SEQ = $PK
UNION ALL
SELECT TOP 5 N'同じ日の他の人', s.PK_SEQ,
       '[' + ISNULL(s.SEQ1,'NULL') + ']',
       '[' + ISNULL(s.SEQ2,'NULL') + ']',
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),'NULL') + ']',
       s.F_UKETUKE
FROM T_KENSIN s
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0 AND s.PK_SEQ <> $PK
ORDER BY 区分, PK_SEQ
"@ 20

Q '--- 2. 08/21 で受付番号が入っている人が他にいないか (いなければ秋葉さんだけ浮いている) ---' @"
SELECT CASE WHEN UKE_NO_KENSA IS NULL THEN N'受付番号なし' ELSE N'受付番号あり' END AS 区分,
       COUNT(*) AS 人数
FROM T_KENSIN
WHERE D_KENSIN = '2026/08/21' AND F_TORIKESI = 0
GROUP BY CASE WHEN UKE_NO_KENSA IS NULL THEN N'受付番号なし' ELSE N'受付番号あり' END
"@ 10

Q '--- 3. ★結果の値: 秋葉さんに残っている値の数 vs 他の人 ---' @"
SELECT N'★秋葉' AS 区分, k.PK_SEQ,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値あり,
       COUNT(*) AS 枠の数
FROM T_KENSA k WHERE k.PK_SEQ = $PK
GROUP BY k.PK_SEQ
UNION ALL
SELECT TOP 5 N'同じ日の他の人', k.PK_SEQ,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END),
       COUNT(*)
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0 AND k.PK_SEQ <> $PK
GROUP BY k.PK_SEQ
ORDER BY 区分, PK_SEQ
"@ 20

Q '--- 4. 秋葉さんにまだ値が残っている項目の一覧 (他の人にも同じものが入っていれば正常) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, k.KEKKA AS 値,
       (SELECT COUNT(*) FROM T_KENSA k2
         JOIN T_KENSIN s2 ON s2.PK_SEQ = k2.PK_SEQ
         WHERE s2.D_KENSIN = '2026/08/21' AND s2.F_TORIKESI = 0 AND k2.PK_SEQ <> $PK
           AND LTRIM(RTRIM(k2.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
           AND LTRIM(RTRIM(ISNULL(k2.KEKKA,''))) <> '') AS 他の人にも入っている数
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.PK_SEQ = $PK AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
ORDER BY 4 ASC, 1
"@ 200

W ''
W '=== 読み方 ==='
W '  1 で秋葉さんの行が他の人と同じ形(SEQ1が[NULL]、受付番号が[NULL]、受付フラグ0)'
W '     → 片付け完了です。'
W '     受付番号に [6] が残っている → まだ戻っていません。元に戻す.bat で'
W '        UKE_NO_20260829_160802.csv を指定してください。'
W '     SEQ1 に値が入っている → 健診ナビで受付処理をしてしまっています。'
W '        健診ナビの受付画面で受付取消をしてください。'
W ''
W '  3 で秋葉さんの「値あり」の数が他の人とほぼ同じ → 計測データは消えています。'
W '     他の人より多い → まだ残っています。'
W ''
W '  4 で「他の人にも入っている数」が 0 の項目 → それがテストの消し残しです。'
W '     全部の行に他の人の数が並んでいれば、残っているのは元からの項目だけなので正常です。'

notepad $out
