<#
  自動判定エラーの原因さがし その2 (check21.ps1)
  check21.bat をダブルクリックすると実行され、結果 r_check21.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check20 の結果:
    値の長さ・形は実データと一致(調査1・2が0件)。それでも自動判定が落ちる。

  check20 の調査1には盲点があった:
    「他の受診者にも値が入っている項目」としか比べていないので、
    他の誰も値を入れたことがない項目は、そもそも表に出てこない。

  ここで確かめること:
    秋葉さんに値を入れた項目のうち、実データでは誰も(ほとんど)値を入れていない項目。
    そういう項目は判定マスタが未整備でも今までエラーにならなかったはずで、
    今回全枠に値を入れたことで初めて踏んだ、というのが本命の筋書き。
    名前からして「血糖(未判別)」「中性脂肪(未判別)」が怪しい。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check21.txt'

# テスト受診者 秋葉達也 2026/08/21 受付6
$PK = 2005094

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 自動判定エラーの原因さがし その2 (PK_SEQ $PK) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 本命: テストで値を入れた項目ごとに、他の受診者で値が入っている件数 (少ない順) ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD,
       m.MEISYO1                 AS 項目名,
       a.KEKKA                   AS テストの値,
       (SELECT COUNT(*) FROM T_KENSA k
         WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
           AND k.PK_SEQ <> a.PK_SEQ
           AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS 他で値入り件数
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
ORDER BY 4 ASC, 1
"@ 200

Q '--- 2. 同じコース(YSA1)の実データで、その項目に値が入る率 ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD,
       m.MEISYO1                 AS 項目名,
       (SELECT COUNT(*) FROM T_KENSA k
         JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
         WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
           AND k.PK_SEQ <> a.PK_SEQ
           AND s.COURSE_CD = 'YSA1'
           AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS 同コースで値入り,
       (SELECT COUNT(*) FROM T_KENSA k
         JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
         WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
           AND k.PK_SEQ <> a.PK_SEQ
           AND s.COURSE_CD = 'YSA1') AS 同コースで枠あり
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
ORDER BY 3 ASC, 1
"@ 200

Q '--- 3. 判定の基準を持っていそうなテーブル (次の手がかり用) ---' @"
SELECT TABLE_NAME AS テーブル
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
  AND (TABLE_NAME LIKE '%HANTEI%' OR TABLE_NAME LIKE '%KIJUN%'
       OR TABLE_NAME LIKE '%JUDGE%' OR TABLE_NAME LIKE '%RANGE%')
ORDER BY TABLE_NAME
"@ 40

Q '--- 4. 「未判別」の項目は実運用でどの項目に振り分けられているか (血糖・中性脂肪の親戚) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, COUNT(*) AS 値入り件数
FROM T_KENSA k
JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE (m.MEISYO1 LIKE N'%血糖%' OR m.MEISYO1 LIKE N'%中性脂肪%' OR m.MEISYO1 LIKE N'%ﾄﾘｸﾞ%' OR m.MEISYO1 LIKE N'%TG%')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY COUNT(*) DESC
"@ 60

W ''
W '=== 読み方 ==='
W '  1 で「他で値入り件数」が 0 (かゼロに近い) 項目が犯人の最有力です。'
W '     誰も値を入れたことがない項目は判定マスタが未整備でも表面化しません。'
W '     今回そこに値を入れたので、自動判定が初めてそのマスタを読んで落ちた、という筋です。'
W '  4 で「血糖(未判別)」ではなく「空腹時血糖」などに件数が集まっていたら、'
W '     実運用では判別後の項目に入れるのが正しく、テストデータの入れ先が違ったということです。'
W ''
W '  対処 (テストを続ける場合):'
W '    結果入力画面で、1 に出た件数0の項目の値を消して空欄に戻し、もう一度「自動判定」。'
W '    それで通れば原因確定。取込側の練習データ作成も、その項目を避けるように直します。'

notepad $out
