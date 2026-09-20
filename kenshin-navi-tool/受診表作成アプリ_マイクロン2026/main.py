import os
import sys
import csv
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# ==========================================
# 見出し(1行目)の名前で列をさがす
#   予約取込の様式が 33列でも 35列でも、列がずれても正しく読めるようにする。
#   見出しが読めないときは、これまでどおりの列位置を使う。
# ==========================================
FIELD_DEFS = [
    ("氏名",            ["氏名"],                    1),
    ("氏名カナ",        ["氏名カナ", "氏名ｶﾅ"],      2),
    ("性別",            ["性別"],                    3),
    ("生年月日",        ["生年月日"],                4),
    ("保険証記号",      ["保険証記号"],             12),
    ("保険証番号",      ["保険証番号"],             13),
    ("事業所名",        ["事業所名"],               18),
    ("コース名",        ["コース名"],               20),
    ("予約日",          ["予約日"],                 21),
]
for _i in range(1, 11):
    FIELD_DEFS.append(("オプション検査%d" % _i, ["オプション検査%d" % _i], 22 + _i))


def build_index(header):
    ix = {}
    norm = []
    if header:
        for h in header:
            norm.append((h or "").strip().replace(" ", "").replace("\u3000", ""))
    for key, names, fallback in FIELD_DEFS:
        found = None
        for n in names:                       # 完全一致を優先
            if n in norm:
                found = norm.index(n); break
        if found is None:                     # 次に前方一致
            for n in names:
                for i, h in enumerate(norm):
                    if h and h.startswith(n):
                        found = i; break
                if found is not None: break
        ix[key] = found if found is not None else fallback
    return ix


def main():
    # ==========================================
    # フォントの準備
    # ==========================================
    try:
        pdfmetrics.registerFont(TTFont('Mincho', 'C:/Windows/Fonts/msmincho.ttc'))
        pdfmetrics.registerFont(TTFont('Gothic', 'C:/Windows/Fonts/msgothic.ttc'))
        fnt_m = 'Mincho'
        fnt_g = 'Gothic'
    except:
        fnt_m = fnt_g = 'Helvetica'

    # exe・py どちらで起動しても、このファイルと同じフォルダを見る
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    csv_file = os.path.join(base_dir, "data.csv")
    if not os.path.exists(csv_file):
        print("エラー: data.csv が見つかりません。")
        print("　　　  main.exe と同じフォルダに data.csv を置いてください。")
        input("Enterキーを押して閉じてください...")
        return

    try:
        with open(csv_file, 'r', encoding='utf-8-sig') as f:
            f.read()
        enc = 'utf-8-sig'
    except:
        enc = 'shift_jis'

    pdf_file = os.path.join(base_dir, "Complete_Health_Check.pdf")
    pw, ph = landscape(A3)
    c = canvas.Canvas(pdf_file, pagesize=(pw, ph))
    
    aw = 210 * mm
    margin_right = 15 * mm
    xb = pw - aw + margin_right  
    w = aw - 30 * mm             

    def chk(x, y, txt):
        c.setDash(); c.rect(x, y, 12, 12)
        c.setFont(fnt_g, 11); c.drawString(x + 17, y + 2, txt)
    
    def num(x, y, d1, d2):
        c.setDash(2, 2); bw, bh = 18, 24
        c.line(x, y, x + d1*bw, y); c.line(x, y + bh, x + d1*bw, y + bh)
        for i in range(d1 + 1): c.line(x + i*bw, y, x + i*bw, y + bh)
        c.setDash()
        if d2 > 0:
            c.setFont(fnt_g, 14); c.drawString(x + d1*bw + 4, y + 4, ".")
            c.setDash(2, 2); sx = x + d1*bw + 12
            c.line(sx, y, sx + d2*bw, y); c.line(sx, y + bh, sx + d2*bw, y + bh)
            for i in range(d2 + 1): c.line(sx + i*bw, y, sx + i*bw, y + bh)
        c.setDash()

    def bp(x, y):
        c.setDash(2, 2); bw, bh = 18, 24
        c.line(x, y, x + 3*bw, y); c.line(x, y + bh, x + 3*bw, y + bh)
        for i in range(4): c.line(x + i*bw, y, x + i*bw, y + bh)
        c.setDash(); c.setFont(fnt_g, 14); c.drawString(x + 58, y + 4, "/")
        c.setDash(2, 2); sx = x + 72
        c.line(sx, y, sx + 3*bw, y); c.line(sx, y + bh, sx + 3*bw, y + bh)
        for i in range(4): c.line(sx + i*bw, y, sx + i*bw, y + bh)
        c.setDash()

    top_y = ph - 75; red_top = top_y - 30; center_y = ph / 2.0; gap = 20
    doc_y = center_y + (gap / 2.0); top_b = center_y - (gap / 2.0)
    opt_h, gap_h = 80, 10; rem_h = red_top - opt_h - gap_h - doc_y
    test_h = rem_h / 4.0; red_y = doc_y + (test_h * 3); red_h = red_top - red_y
    row_h = red_h / 8.0; doc_h = red_y - doc_y - 22; lbl_w = 80

    page_count = 0
    tally = {}
    n_ibu_strike = 0
    with open(csv_file, 'r', encoding=enc, errors='replace') as f:
        all_rows = list(csv.reader(f))
        # 見出しの行をさがす。名簿シートの1行目に集計行が入っていても大丈夫にする。
        hdr_i = 0
        for i, r in enumerate(all_rows[:10]):
            if any((cc or "").strip().replace(" ", "").replace("\u3000", "") == "氏名" for cc in r):
                hdr_i = i; break
        header = all_rows[hdr_i] if all_rows else None
        ix = build_index(header)
        for row in all_rows[hdr_i + 1:]:
            if not row or not any(row): continue
            row = row + [""] * 40

            def g(key, _row=row):
                i = ix.get(key)
                if i is None or i >= len(_row): return ""
                return (_row[i] or "").strip()

            emp_name = g("氏名"); furigana = g("氏名カナ")
            if not emp_name and not furigana: continue
            page_count += 1
            gender = g("性別")
            birth_dt = g("生年月日"); kigou = g("保険証記号"); bango = g("保険証番号")
            company = g("事業所名"); course = g("コース名")
            
            # 【新ルール】コース名の自動置換（入力の手間を減らす）
            c_map = {
                "SA": "定期健診A",
                "SB": "定期健診B",
                "SB2": "定期健診B＋血液",
                "TA2": "東振協A2",
                "TB": "東振協B",
                "KY101": "協会 一般健診",
                "KY301": "協会 若年健診",
                "定健A": "定期健診A",
                "定健B": "定期健診B",
                "FA": "福生A",
                "FB": "福生B",
                "TKA2": "東振協化粧品A2",
                "TKB": "東振協化粧品B"
            }
            course_display = c_map.get(course, course)
            
            reserve_date = g("予約日")
            options = [g("オプション検査%d" % i) for i in range(1, 11)]
            options = [o for o in options if o]

            # ガイド位置
            birth_x_off = 0 if birth_dt else 0
            gender_x_off = 0 if gender else 0
            age_x_off = 0 if birth_dt else -5

            age_str = ""
            if birth_dt:
                try:
                    b_date = datetime.strptime(birth_dt.replace("-", "/"), "%Y/%m/%d")
                    target = datetime(2027, 3, 31)
                    age_str = str(target.year - b_date.year - ((target.month, target.day) < (b_date.month, b_date.day)))
                except: age_str = ""
            else: age_str = "   歳"

            # --- 斜線の判定 (マイクロンメモリジャパン 2026年10月 専用) ---
            #   人間ドックD１コース       … 胃部・採血・心電図すべてあり → 斜線なし
            #   一般健診A2コース          … 胃部X線は希望者のみ
            #   生活習慣病予防健診Bコース … 胃部X線は希望者のみ
            #   → オプションに「胃部X線検査」がある人だけ胃部あり。無ければ胃部に斜線。
            #   採血・心電図は3コースとも全員あるので斜線なし。
            has_ibu = any(("胃部" in o) or ("胃カメラ" in o) or ("胃ｶﾒﾗ" in o) for o in options)
            strike_ecg = False
            strike_blood = False
            if not course:
                strike_ibu = False            # コースが空なら全検査あり
            elif "ドック" in course_display:
                strike_ibu = False            # ドックは胃部込み
            else:
                strike_ibu = not has_ibu      # A2・B は申込があった人だけ
            tally[course_display] = tally.get(course_display, 0) + 1
            if strike_ibu: n_ibu_strike += 1

            c.setDash(); c.setFont(fnt_g, 15); c.drawString(xb, top_y + 35, "健康診断受診票")
            ox, ow = xb + w * 0.65 + 5, w - (w * 0.65) - 5
            c.drawString(ox + 15, top_y + 35, f"受診日:{reserve_date}"); c.drawString(ox + 15, top_y + 15, "受付No:")
            c.setFont(fnt_g, 12); c.rect(xb, top_y + 3, w * 0.65, 24); c.line(xb + lbl_w, top_y + 3, xb + lbl_w, top_y + 27)
            c.drawString(xb + 10, top_y + 11, "検査項目"); c.drawString(xb + lbl_w + 10, top_y + 11, course_display)
            
            c.setFont(fnt_g, 9); c.setFillColorRGB(1, 0, 0)
            c.drawString(xb + 5, top_y - 8, "※太枠内の空欄部分を記入して下さい。")
            c.drawString(xb + 5, top_y - 22, "※食事時間は、受診時に食後何時間経過しているか〇で囲んでください。")
            c.setFillColorRGB(0, 0, 0)
            c.setStrokeColorRGB(0.8, 0, 0); c.setLineWidth(3.0); c.rect(xb, red_y, w * 0.65, red_h)
            c.setStrokeColorRGB(0, 0, 0); c.setLineWidth(0.5); c.setDash()
            for i in range(1, 8):
                sx = xb + lbl_w if i == 1 else xb
                c.line(sx, red_y + i * row_h, xb + w * 0.65, red_y + i * row_h)
            c.line(xb + lbl_w, red_y, xb + lbl_w, red_y + red_h)
            c.line(xb + 180, red_y + 7 * row_h, xb + 180, red_y + red_h)
            for lx in [160, 220, 250, 280]: c.line(xb + lx, red_y + 3 * row_h, xb + lx, red_y + 4 * row_h)
            
            def row_txt(r, x, t, s=9, y_off=0, font_name=fnt_g):
                c.setFont(font_name, s); c.drawString(xb + x, red_y + r * row_h + (row_h - s)/2 + 1 + y_off, t)
            
            display_name = f"{emp_name}  様" if emp_name else ""
            row_txt(7, 5, "保険証"); row_txt(7, lbl_w + 5, f"記号  {kigou}"); row_txt(7, 185, f"番号  {bango}")
            row_txt(6, 5, "事業所"); row_txt(6, lbl_w + 5, company); row_txt(5, 5, "フリガナ"); row_txt(5, lbl_w + 5, furigana)
            row_txt(4, 5, "氏名"); row_txt(4, lbl_w + 5, display_name, 14, font_name=fnt_m)
            
            # ガイド印字
            row_txt(3, 5, "生年月日"); row_txt(3, lbl_w + 5 + birth_x_off, "    /   /  " if not birth_dt else birth_dt)
            row_txt(3, 165, "年度末年齢"); row_txt(3, 228 + age_x_off, age_str)
            row_txt(3, 255, "性別"); row_txt(3, 285 + gender_x_off, "男 ・ 女" if not gender else gender)
            
            row_txt(2, 5, "食事時間"); row_txt(2, lbl_w + 5, "①食後10h以上  ②食後3.5〜10h未満  ③食後3.5h未満")
            row_txt(1, 5, "※女性のみ", 9, -3); row_txt(0, 5, "回答ください", 9, 3)
            row_txt(1, lbl_w + 5, "現在妊娠中または可能性あり( はい ・ いいえ )", 8); row_txt(0, lbl_w + 5, "現在生理中である( はい ・ いいえ )", 8)
            
            c.setDash(); c.rect(xb, doc_y, w * 0.65, doc_h); cw = 95; c.line(xb + cw, doc_y, xb + cw, doc_y + doc_h)
            c.setFont(fnt_g, 11); c.drawString(xb+5, doc_y+doc_h+4, "医師記入欄／他覚所見")
            c.line(xb, doc_y+60, xb+w*0.65, doc_y+60); c.line(xb, doc_y+26, xb+w*0.65, doc_y+26)
            chk(xb + 15, doc_y+60+(doc_h-60)/2.0-5, "所見なし")
            sy, ds, sx, sx2 = doc_y+doc_h-22, 18, xb+cw+10, xb+cw+115
            chk(sx, sy, "呼吸音異常"); chk(sx2, sy, "心雑音"); chk(sx, sy-ds, "不整脈"); chk(sx2, sy-ds, "頻脈")
            chk(sx, sy-ds*2, "貧血"); chk(sx2, sy-ds*2, "甲状腺腫脹"); chk(sx, sy-ds*3, "その他所見"); c.drawString(sx+85, sy-ds*3+2, "〔                〕")
            c.drawString(xb+15, doc_y+40, "診察判定"); chk(sx, doc_y+44, "要経過観察"); c.setFont(fnt_g, 11); c.drawString(sx + 72, doc_y+46, "〔3ヶ月・6ヶ月・12ヶ月〕")
            chk(sx, doc_y+28, "要精密検査"); chk(sx+95, doc_y+28, "要医療"); chk(sx+165, doc_y+28, "治療中"); c.drawString(xb+10, doc_y+9, "診察医師の氏名")
            
            c.setDash(); c.rect(ox, red_top-opt_h, ow, opt_h); c.setFont(fnt_g, 10); c.drawString(ox+5, red_top-15, "オプション等")
            
            # 【新ルール】便潜血の印字条件（東振協B or 協会一般）
            if any(x in course_display for x in ["東振協B", "協会 一般健診"]):
                c.setFillColorRGB(0, 0, 1); c.setFont(fnt_g, 14); c.drawRightString(ox + ow - 5, red_top - 15, "便潜血"); c.setFillColorRGB(0, 0, 0)
            
            # オプション描画：個数に応じて1列／2列／3列を切り替え
            if options:
                c.setFont(fnt_g, 9)
                start_y = red_top - 28
                line_h = 12
                if len(options) <= 5:
                    # 5個以下：左寄せ1列
                    for i, opt in enumerate(options):
                        c.drawString(ox + 5, start_y - i * line_h, "●" + opt)
                elif len(options) <= 10:
                    # 6〜10個：2列レイアウト
                    for i in range(5):
                        c.drawString(ox + 5, start_y - i * line_h, "●" + options[i])
                    col2_x = ox + ow / 2
                    for i in range(5, len(options)):
                        c.drawString(col2_x, start_y - (i - 5) * line_h, "●" + options[i])
                else:
                    # 11個以上：3列、フォント小さめ
                    c.setFont(fnt_g, 8)
                    start_y_3 = red_top - 24
                    line_h_3 = 11
                    per_col = (len(options) + 2) // 3  # 切り上げ
                    col_w = ow / 3
                    for i, opt in enumerate(options):
                        col = i // per_col
                        row = i % per_col
                        y = start_y_3 - row * line_h_3
                        if y < red_top - opt_h + 4:
                            break
                        c.drawString(ox + 5 + col * col_w, y, "●" + opt)
            
            # テスト項目ごとの斜線描画
            test_map = {"心電図": strike_ecg, "採血": strike_blood, "胃部": strike_ibu}
            for i, t in enumerate(["心電図", "採血", "胃部", "胸部"]):
                by = doc_y+(i*test_h); c.setDash(); c.rect(ox, by, ow, test_h)
                c.setFont(fnt_g, 12); c.drawString(ox+5, by+test_h/2-4, t); c.setFont(fnt_g, 10); c.drawString(ox+55, by+test_h/2-4, "NO.")
                if test_map.get(t, False):
                    c.line(ox, by, ox+ow, by+test_h)
            
            by_a, bh_a = 15, top_b-15; c.setDash(); c.rect(xb, by_a, w, bh_a)
            y_s = top_b-25; c.setFont(fnt_g, 14); c.drawString(xb+10, y_s, "【身体計測】")
            r1, r2 = y_s-30, y_s-62
            c.setFont(fnt_g, 13); c.drawString(xb+25, r1+6, "身長"); num(xb+65, r1, 3, 1); c.setFont(fnt_g, 12); c.drawString(xb+160, r1+6, "cm")
            c.setFont(fnt_g, 13); c.drawString(xb+200, r1+6, "腹囲"); num(xb+245, r1, 3, 1); c.setFont(fnt_g, 12); c.drawString(xb+345, r1+6, "cm")
            c.setFont(fnt_g, 13); c.drawString(xb+25, r2+6, "体重"); num(xb+65, r2, 3, 1); c.setFont(fnt_g, 12); c.drawString(xb+160, r2+6, "kg")
            y_v = r2-32; c.setFont(fnt_g, 14); c.drawString(xb+10, y_v, "【視力】"); c.drawString(xb+260, y_v, "【血圧】") 
            c.setFont(fnt_g, 12); c.drawString(xb+70, y_v, "裸眼"); c.drawString(xb+108, y_v, "〔矯正/ｺﾝﾀｸﾄ･眼鏡〕")
            v1, v2 = y_v-28, y_v-60
            c.setFont(fnt_g, 13); c.drawString(xb+25, v1+6, "右"); num(xb+55, v1, 1, 1); c.setFont(fnt_g, 12); c.drawString(xb+115, v1+6, "〔"); num(xb+135, v1, 1, 1); c.drawString(xb+205, v1+6, "〕")
            c.setFont(fnt_g, 13); c.drawString(xb+274, v1+6, "1回目"); bp(xb+324, v1); c.setFont(fnt_g, 12); c.drawString(xb+470, v1+6, "mmHg")
            c.setFont(fnt_g, 13); c.drawString(xb+25, v2+6, "左"); num(xb+55, v2, 1, 1); c.setFont(fnt_g, 12); c.drawString(xb+115, v2+6, "〔"); num(xb+135, v2, 1, 1); c.drawString(xb+205, v2+6, "〕")
            c.setFont(fnt_g, 13); c.drawString(xb+274, v2+6, "2回目"); bp(xb+324, v2); c.setFont(fnt_g, 12); c.drawString(xb+470, v2+6, "mmHg")
            y_a = v2-32; c.setFont(fnt_g, 14); c.drawString(xb+10, y_a, "【聴力】"); ay = y_a-75; c.setDash(); c.rect(xb+30, ay, 460, 70)
            c.line(xb+30, ay+35, xb+490, ay+35); c.line(xb+70, ay, xb+70, ay+70); c.line(xb+140, ay, xb+140, ay+70)
            c.setFont(fnt_g, 13); c.drawString(xb+40, ay+48, "右"); c.setFont(fnt_g, 12); c.drawString(xb+80, ay+54, "1000Hz"); c.drawString(xb+80, ay+38, "4000Hz")
            c.drawString(xb+150, ay+54, "1. 所見なし  2. 所見あり"); c.drawString(xb+150, ay+38, "1. 所見なし  2. 所見あり"); c.drawString(xb+330, ay+46, "※補聴器使用")
            c.setFont(fnt_g, 13); c.drawString(xb+40, ay+13, "左"); c.setFont(fnt_g, 12); c.drawString(xb+80, ay+19, "1000Hz"); c.drawString(xb+80, ay+3, "4000Hz")
            c.drawString(xb+150, ay+19, "1. 所見なし  2. 所見あり"); c.drawString(xb+150, ay+3, "1. 所見なし  2. 所見あり"); c.drawString(xb+330, ay+11, "※補聴器使用")
            y_u = ay-25; c.setFont(fnt_g, 14); c.drawString(xb+10, y_u, "【尿検査】")
            xos = [90, 190, 250, 310, 370]
            for i, itm in enumerate(["蛋白", "糖", "潜血"]):
                ry = y_u-22-(i*19); c.setFont(fnt_g, 12); c.drawString(xb+35, ry, itm)
                for j, o in enumerate(["所見なし", "±", "+", "++", "+++"]): chk(xb+xos[j], ry-2, o)
            c.showPage()

    if page_count == 0:
        print("data.csv に受診者データがありません。")
        print("　　　  1行目の見出しの下に、1人1行で入力して保存してください。")
        input("Enterキーを押して閉じてください...")
        return

    c.save()
    print(f"PDF作成完了：{page_count}人分（Complete_Health_Check.pdf）")
    print("")
    print("--- コース別の人数 ---")
    for k in sorted(tally, key=lambda x: -tally[x]):
        print("  %-28s %4d 人" % (k if k else "(コース名なし)", tally[k]))
    print("")
    print("  胃部に斜線を入れた人 : %d 人" % n_ibu_strike)
    print("  胃部の検査がある人   : %d 人" % (page_count - n_ibu_strike))
    print("  ※ オプションに「胃部X線検査」がある人だけ胃部ありにしています。")
    print("")
    input("Enterキーを押して閉じてください...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("エラーが発生しました：", e)
        print("※PDFを開いたままだと上書きできません。閉じてからもう一度実行してください。")
        input("Enterキーを押して閉じてください...")
