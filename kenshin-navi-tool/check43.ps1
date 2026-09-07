<#
  C以上の所見が総合判定に載っているかを調べる (check43.ps1)
  check43.bat をダブルクリックすると実行され、結果 r_check43.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  心電図・胸部X線・胃部X線で判定が C 以上(C/C3/C6/C12/D/E/F…)の人について、
    ・各科の判定 (600130 心電図 / 600090 胸部X線 / 600160 消化器) に反映されているか
    ・総合判定のコメントにその検査のことが書かれているか
  を確かめます。

  「A・B は正常〜軽度」なので、C以上だけを対象にします。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check43.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

# 所見の枠 → 各科の判定項目
#   心電図 067112A → 600130 / 胸部X線 077010A → 600090 / 胃部X線 077300A → 600160
"=== C以上の所見が総合判定に載っているか ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★C以上の人の一覧 (所見・所見の判定・各科の判定) ---' @"
SELECT g.KANJI_SIMEI AS 氏名,
       CASE LTRIM(RTRIM(k.KOMOKU_CD))
            WHEN '067112A' THEN N'心電図'
            WHEN '077010A' THEN N'胸部X線'
            WHEN '077300A' THEN N'胃部X線' END AS 検査,
       LEFT(k.KEKKA, 28) AS 所見,
       '[' + LTRIM(RTRIM(k.HANTEI_KIGO)) + ']' AS 所見の判定,
       '[' + ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(j.HANTEI_KIGO,'')))
                       FROM T_KENSA j WHERE j.PK_SEQ = s.PK_SEQ
                        AND LTRIM(RTRIM(j.KOMOKU_CD)) =
                            CASE LTRIM(RTRIM(k.KOMOKU_CD))
                                 WHEN '067112A' THEN '600130'
                                 WHEN '077010A' THEN '600090'
                                 WHEN '077300A' THEN '600160' END), '') + ']' AS 各科の判定,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('067112A','077010A','077300A')
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('', 'A', 'B')
ORDER BY g.KANA_SIMEI, k.KOMOKU_CD
"@ 60

Q '--- 2. ★★その人たちの総合判定コメント (全文) ---' @"
SELECT g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       k.KEKKA AS コメント,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1100%')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND s.PK_SEQ IN (
      SELECT s2.PK_SEQ FROM T_KENSA k2 JOIN T_KENSIN s2 ON s2.PK_SEQ = k2.PK_SEQ
      WHERE s2.D_KENSIN = '$Ymd' AND s2.F_TORIKESI = 0
        AND LTRIM(RTRIM(k2.KOMOKU_CD)) IN ('067112A','077010A','077300A')
        AND LTRIM(RTRIM(ISNULL(k2.HANTEI_KIGO,''))) NOT IN ('', 'A', 'B'))
ORDER BY g.KANA_SIMEI, k.KOMOKU_CD
"@ 200

Q '--- 3. ★★まとめ: 人ごとに「載っているか」を判定 ---' @"
SELECT x.氏名,
       x.心電図, x.胸部, x.胃部,
       CASE WHEN x.心電図 = '' THEN '' WHEN x.文中心電図 > 0 THEN 'あり' ELSE N'★無し' END AS 総合に心電図,
       CASE WHEN x.胸部   = '' THEN '' WHEN x.文中胸部   > 0 THEN 'あり' ELSE N'★無し' END AS 総合に胸部,
       CASE WHEN x.胃部   = '' THEN '' WHEN x.文中胃部   > 0 THEN 'あり' ELSE N'★無し' END AS 総合に胃部,
       x.PK_SEQ
FROM (
  SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ, s.PK_SEQ,
         ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                  WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '067112A'
                    AND LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) NOT IN ('','A','B')), '') AS 心電図,
         ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                  WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '077010A'
                    AND LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) NOT IN ('','A','B')), '') AS 胸部,
         ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                  WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '077300A'
                    AND LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) NOT IN ('','A','B')), '') AS 胃部,
         (SELECT COUNT(*) FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ
           AND (LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1100%')
           AND c.KEKKA LIKE N'%心電図%')                                    AS 文中心電図,
         (SELECT COUNT(*) FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ
           AND (LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1100%')
           AND (c.KEKKA LIKE N'%胸部%' OR c.KEKKA LIKE N'%肺%'))            AS 文中胸部,
         (SELECT COUNT(*) FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ
           AND (LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1100%')
           AND (c.KEKKA LIKE N'%胃%' OR c.KEKKA LIKE N'%消化%'))            AS 文中胃部
  FROM T_KENSIN s
  LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
WHERE x.心電図 <> '' OR x.胸部 <> '' OR x.胃部 <> ''
ORDER BY x.カナ
"@ 60

Q '--- 4. くらべる: 各科の判定 (600xxx) の分布 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定, COUNT(*) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('600090','600130','600160')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), k.HANTEI_KIGO
ORDER BY 1, 4 DESC
"@ 30

W ''
W '=== 読み方 ==='
W '  1 が、心電図・胸部・胃部で C以上だった人の一覧です。'
W '     「各科の判定」に同じ記号が入っていれば、自動判定はそこまでは拾えています。'
W '     ここが空なら、各科の判定に反映されていません。'
W ''
W '  3 が答えです。「★無し」と出た人は、'
W '     C以上の所見があるのに総合判定の文章にその検査のことが書かれていません。'
W '     全部「あり」なら、総合判定にちゃんと載っています。'
W ''
W '  2 で実際の文章を確かめてください。'
W '     3 は文字を探しているだけなので、言い回しによっては取りこぼします。'
W '     最終的な判断は 2 の文章を読んでからにしてください。'

notepad $out
