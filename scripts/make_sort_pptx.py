#!/usr/bin/env python3
"""ソートアルゴリズム教材スライド (sort.pptx) を生成する。
html/sort.html のビジュアライザと対になる講義用スライド。
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---- カラーパレット (html/sort.html と統一) ----
BG     = RGBColor(0xF4, 0xF6, 0xFB)
PANEL  = RGBColor(0xFF, 0xFF, 0xFF)
PANEL2 = RGBColor(0xEE, 0xF1, 0xF7)
LINE   = RGBColor(0xD6, 0xDC, 0xE8)
TEXT   = RGBColor(0x1A, 0x20, 0x30)
MUTED  = RGBColor(0x5D, 0x6B, 0x82)
ACCENT = RGBColor(0x2F, 0x7F, 0xD9)
INK    = RGBColor(0x06, 0x12, 0x2A)   # 明色面（バッジ/見出し帯/明るいセル）に乗せる濃文字
BAR    = RGBColor(0x4A, 0x90, 0xD9)
COMPARE= RGBColor(0xFF, 0xD2, 0x4A)
SWAP   = RGBColor(0xFF, 0x6B, 0x6B)
SORTED = RGBColor(0x4A, 0xDE, 0x80)
PIVOT  = RGBColor(0xC0, 0x84, 0xFC)

FONT = "Meiryo"
MONO = "Consolas"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


def slide_bg(slide, color=BG):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def add_text(slide, x, y, w, h, text, size=18, color=TEXT, bold=False,
             align=PP_ALIGN.LEFT, font=FONT, anchor=MSO_ANCHOR.TOP, line_spacing=1.1):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        # インラインで **強調** を解釈
        segs = ln.split("**")
        for j, seg in enumerate(segs):
            if seg == "":
                continue
            r = p.add_run()
            r.text = seg
            r.font.size = Pt(size)
            r.font.name = font
            r.font.color.rgb = ACCENT if (j % 2 == 1) else color
            r.font.bold = bold or (j % 2 == 1)
    return tb


def add_rect(slide, x, y, w, h, fill, line=None, radius=False):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shp_type, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def card(slide, x, y, w, h):
    return add_rect(slide, x, y, w, h, PANEL, line=LINE, radius=True)


def section_header(slide, title, subtitle=None, num=None):
    """各内容スライド上部の共通ヘッダ。"""
    add_rect(slide, 0, 0, EMU_W, Inches(1.15), PANEL)
    add_rect(slide, 0, Inches(1.15), EMU_W, Pt(3), ACCENT)
    if num is not None:
        badge = add_rect(slide, Inches(0.45), Inches(0.3), Inches(0.55), Inches(0.55), ACCENT, radius=True)
        tf = badge.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(num)
        r.font.size = Pt(22); r.font.bold = True; r.font.name = FONT
        r.font.color.rgb = INK
    tx = Inches(1.2) if num is not None else Inches(0.5)
    add_text(slide, tx, Inches(0.18), Inches(11), Inches(0.6), title, size=28, bold=True)
    if subtitle:
        add_text(slide, tx, Inches(0.72), Inches(11.5), Inches(0.4), subtitle, size=14, color=MUTED)


def draw_bars(slide, x, y, w, h, values, colors=None, max_val=None, show_val=False, gap=Inches(0.04)):
    """縦棒グラフ（配列の可視化）を描く。"""
    n = len(values)
    if max_val is None:
        max_val = max(values)
    bar_w = (w - gap * (n - 1)) / n
    for i, v in enumerate(values):
        bh = Emu(int(h * (v / max_val)))
        bx = Emu(int(x + i * (bar_w + gap)))
        by = Emu(int(y + h - bh))
        c = colors[i] if colors else BAR
        shp = add_rect(slide, bx, by, Emu(int(bar_w)), bh, c)
        shp.adjustments  # noop
        if show_val:
            add_text(slide, bx, Emu(int(by - Inches(0.32))), Emu(int(bar_w)), Inches(0.3),
                     str(v), size=11, color=MUTED, align=PP_ALIGN.CENTER)


def legend(slide, x, y):
    items = [("未ソート", BAR), ("比較中", COMPARE), ("交換/書込", SWAP), ("基準(pivot)", PIVOT), ("確定", SORTED)]
    cx = x
    for label, col in items:
        add_rect(slide, cx, y + Inches(0.03), Inches(0.18), Inches(0.18), col, radius=True)
        tb = add_text(slide, cx + Inches(0.24), y - Inches(0.02), Inches(1.5), Inches(0.3),
                      label, size=11, color=MUTED)
        cx += Inches(0.24) + Inches(len(label) * 0.13 + 0.35)


# ====================================================================
# 1. タイトルスライド
# ====================================================================
s = prs.slides.add_slide(BLANK)
slide_bg(s)
# 装飾の棒（ランダム風）
demo = [30, 55, 20, 80, 45, 95, 65, 40, 75, 50, 85, 35]
demo_colors = [SORTED if i < 4 else (COMPARE if i in (5, 6) else BAR) for i in range(len(demo))]
draw_bars(s, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.4), demo, demo_colors, max_val=100)
add_text(s, Inches(0.8), Inches(1.4), Inches(11.7), Inches(1.0),
         "ソートアルゴリズム入門", size=54, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(2.6), Inches(11.7), Inches(0.6),
         "縦棒で見て、手を動かして学ぶ", size=24, color=ACCENT, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(3.4), Inches(11.7), Inches(0.5),
         "比較 ・ 交換 ・ 計算量 を体感する教材", size=16, color=MUTED, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(6.9), Inches(11.7), Inches(0.4),
         "対応ビジュアライザ:  html/sort.html", size=13, color=MUTED, align=PP_ALIGN.CENTER, font=MONO)

# ====================================================================
# 2. ソートとは / 学習の流れ
# ====================================================================
s = prs.slides.add_slide(BLANK)
slide_bg(s)
section_header(s, "ソートとは？", "データを一定の順序に並べ替えること", num=1)
card(s, Inches(0.5), Inches(1.5), Inches(6.0), Inches(2.4))
add_text(s, Inches(0.8), Inches(1.7), Inches(5.5), Inches(0.5), "身近な例", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.8), Inches(2.3), Inches(5.5), Inches(1.5),
         "・ テストの点数を高い順に並べる\n・ 名前を五十音順に並べる\n・ ファイルを更新日時順に表示\n・ 検索の前処理（並んでいると速い）",
         size=15, line_spacing=1.3)
card(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(2.4))
add_text(s, Inches(7.1), Inches(1.7), Inches(5.5), Inches(0.5), "なぜ複数の方法を学ぶ？", size=18, bold=True, color=ACCENT)
add_text(s, Inches(7.1), Inches(2.3), Inches(5.5), Inches(1.5),
         "同じ「並べ替え」でも手順は何通りもある。\nそれぞれ **速さ** ・ **メモリ** ・ **わかりやすさ** が違う。\n違いを知ると「どれを使うか」を選べるようになる。",
         size=15, line_spacing=1.3)
# 並べ替えのイメージ
add_text(s, Inches(0.5), Inches(4.15), Inches(6), Inches(0.4), "並べ替え前（ランダム）", size=14, color=MUTED)
draw_bars(s, Inches(0.5), Inches(4.6), Inches(5.6), Inches(2.3), [40, 90, 20, 70, 30, 100, 50, 80, 10, 60], max_val=100, show_val=True)
add_text(s, Inches(6.9), Inches(4.15), Inches(6), Inches(0.4), "並べ替え後（昇順）", size=14, color=SORTED)
draw_bars(s, Inches(6.9), Inches(4.6), Inches(5.6), Inches(2.3), [10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
          colors=[SORTED]*10, max_val=100, show_val=True)

# ====================================================================
# 3. 共通ルール: 比較と交換 / 計算量とは
# ====================================================================
s = prs.slides.add_slide(BLANK)
slide_bg(s)
section_header(s, "2つの基本操作と「計算量」", "アルゴリズムの速さは操作の回数で測る", num=2)
card(s, Inches(0.5), Inches(1.5), Inches(3.9), Inches(2.5))
add_text(s, Inches(0.7), Inches(1.65), Inches(3.5), Inches(0.4), "① 比較 (compare)", size=17, bold=True, color=COMPARE)
add_text(s, Inches(0.7), Inches(2.15), Inches(3.5), Inches(1.7),
         "2つの要素の大小を\n調べる。\n\n例: a[i] と a[j]、\nどちらが大きい？", size=14, line_spacing=1.25)
card(s, Inches(4.6), Inches(1.5), Inches(3.9), Inches(2.5))
add_text(s, Inches(4.8), Inches(1.65), Inches(3.5), Inches(0.4), "② 交換 (swap)", size=17, bold=True, color=SWAP)
add_text(s, Inches(4.8), Inches(2.15), Inches(3.5), Inches(1.7),
         "2つの要素の位置を\n入れ替える。\n\n例: a[i] と a[j] を\nひっくり返す", size=14, line_spacing=1.25)
card(s, Inches(8.7), Inches(1.5), Inches(4.1), Inches(2.5))
add_text(s, Inches(8.9), Inches(1.65), Inches(3.7), Inches(0.4), "計算量 (Big-O)", size=17, bold=True, color=ACCENT)
add_text(s, Inches(8.9), Inches(2.15), Inches(3.7), Inches(1.7),
         "要素数 n が増えると\n操作回数がどれくらい\n増えるかの目安。\n\n**O(n²)** は n が10倍で\n操作が約100倍に！", size=14, line_spacing=1.2)
# Big-O の比較イメージ
card(s, Inches(0.5), Inches(4.2), Inches(12.3), Inches(2.9))
add_text(s, Inches(0.7), Inches(4.35), Inches(11), Inches(0.4),
         "n が増えたときの操作回数の伸び（小さいほど速い）", size=15, bold=True)
rows = [
    ("O(1)",       "一定",       "★★★★★", SORTED),
    ("O(log n)",   "とても緩やか", "★★★★★", SORTED),
    ("O(n)",       "比例",       "★★★★☆", SORTED),
    ("O(n log n)", "やや増える",  "★★★☆☆", ACCENT),
    ("O(n²)",      "急増",       "★★☆☆☆", COMPARE),
    ("O(n³)〜",    "爆発的",     "★☆☆☆☆", SWAP),
]
ry = Inches(4.9)
for label, desc, stars, col in rows:
    add_text(s, Inches(0.8), ry, Inches(2.2), Inches(0.35), label, size=14, bold=True, color=col, font=MONO)
    add_text(s, Inches(3.1), ry, Inches(3.0), Inches(0.35), desc, size=13, color=TEXT)
    add_text(s, Inches(6.2), ry, Inches(2.5), Inches(0.35), stars, size=13, color=col)
    # 伸びを示すバー
    widths = {"O(1)":0.5,"O(log n)":1.0,"O(n)":2.0,"O(n log n)":3.2,"O(n²)":5.0,"O(n³)〜":6.0}
    add_rect(s, Inches(8.8), ry + Inches(0.03), Inches(widths[label]*0.66), Inches(0.22), col, radius=True)
    ry += Inches(0.36)

# ====================================================================
# アルゴリズム解説スライド共通関数
# ====================================================================
def algo_slide(num, name_jp, name_en, idea, steps, before, snapshot_label, snapshot_vals, snapshot_colors,
               best, avg, worst, space, stable, point):
    s = prs.slides.add_slide(BLANK)
    slide_bg(s)
    section_header(s, f"{name_jp}", name_en, num=num)
    # 左: 考え方とステップ
    card(s, Inches(0.5), Inches(1.5), Inches(6.4), Inches(5.4))
    add_text(s, Inches(0.75), Inches(1.65), Inches(6.0), Inches(0.4), "考え方", size=17, bold=True, color=ACCENT)
    add_text(s, Inches(0.75), Inches(2.1), Inches(6.0), Inches(0.9), idea, size=14, line_spacing=1.25)
    add_text(s, Inches(0.75), Inches(3.05), Inches(6.0), Inches(0.4), "手順", size=17, bold=True, color=ACCENT)
    step_txt = "\n".join(f"{i+1}. {t}" for i, t in enumerate(steps))
    add_text(s, Inches(0.75), Inches(3.5), Inches(6.0), Inches(2.0), step_txt, size=14, line_spacing=1.3)
    add_text(s, Inches(0.75), Inches(6.05), Inches(6.0), Inches(0.8),
             "ポイント: " + point, size=13, color=COMPARE, line_spacing=1.2)
    # 右上: スナップショット（棒グラフ）
    card(s, Inches(7.2), Inches(1.5), Inches(5.6), Inches(3.0))
    add_text(s, Inches(7.45), Inches(1.6), Inches(5.2), Inches(0.35), snapshot_label, size=13, color=MUTED)
    draw_bars(s, Inches(7.5), Inches(2.15), Inches(5.0), Inches(1.7), snapshot_vals,
              snapshot_colors, max_val=max(snapshot_vals), show_val=True)
    legend(s, Inches(7.45), Inches(4.05))
    # 右下: 計算量カード
    card(s, Inches(7.2), Inches(4.7), Inches(5.6), Inches(2.2))
    add_text(s, Inches(7.45), Inches(4.82), Inches(5), Inches(0.35), "計算量", size=15, bold=True, color=ACCENT)
    comp = [("最良", best), ("平均", avg), ("最悪", worst), ("メモリ", space), ("安定性", stable)]
    cy = Inches(5.25)
    for k, v in comp:
        add_text(s, Inches(7.5), cy, Inches(2.2), Inches(0.3), k, size=13, color=MUTED)
        add_text(s, Inches(9.6), cy, Inches(3.0), Inches(0.3), v, size=13, color=TEXT, font=MONO)
        cy += Inches(0.32)
    return s


# ====================================================================
# 4-10. 各アルゴリズム
# ====================================================================
B = BAR; C = COMPARE; W = SWAP; S = SORTED; P = PIVOT

algo_slide(
    3, "バブルソート", "Bubble Sort",
    "隣り合う2つを比べて、大きい方を後ろへ送る。これを繰り返すと\n大きい値が泡のように右端へ浮かび上がっていく。",
    ["先頭から隣同士を比較", "左 > 右 なら交換", "右端まで行くと最大値が確定", "残りで繰り返す"],
    None, "比較中: a[3] と a[4]（右端は確定済み）",
    [40, 60, 30, 90, 50, 20, 70, 80, 100],
    [B, B, B, C, C, B, B, S, S],
    "O(n)", "O(n²)", "O(n²)", "O(1)", "安定 ○",
    "シンプルで理解しやすいが遅い。教材の定番。")

algo_slide(
    4, "選択ソート", "Selection Sort",
    "未ソート部分から最小値を探し、その先頭と交換する。\n交換回数が少ないのが特徴。",
    ["未ソート部分の最小値を探す", "先頭の要素と交換", "確定範囲を1つ広げる", "全部確定するまで繰り返す"],
    None, "最小値を探索中（紫=現在地, 緑=確定）",
    [10, 20, 30, 90, 50, 25, 70, 80, 60],
    [S, S, S, P, C, B, B, B, B],
    "O(n²)", "O(n²)", "O(n²)", "O(1)", "不安定 ×",
    "比較は多いが交換は最小限（最大 n-1 回）。")

algo_slide(
    5, "挿入ソート", "Insertion Sort",
    "整列済みの列に、新しい要素を正しい位置へ「差し込む」。\nトランプを手札に並べる動きと同じ。",
    ["左端を整列済みとみなす", "次の要素を取り出す(key)", "整列済みの中で入る位置を探す", "ずらして空けた所へ挿入"],
    None, "key=35 を左の整列済み部分へ挿入中",
    [10, 20, 50, 70, 90, 35, 60, 80, 100],
    [S, S, S, S, S, P, B, B, B],
    "O(n)", "O(n²)", "O(n²)", "O(1)", "安定 ○",
    "ほぼ整列済みのデータに非常に強い（O(n)に近い）。")

algo_slide(
    6, "クイックソート", "Quick Sort",
    "基準値(pivot)を決め、それより小さい組・大きい組に分ける。\n各組を同じ方法で再帰的に並べ替える「分割統治」。",
    ["pivot を1つ選ぶ", "pivot 未満を左、以上を右へ分割", "pivot を境界に置いて確定", "左右をそれぞれ再帰的に処理"],
    None, "pivot=70 で分割中（紫=pivot）",
    [30, 50, 20, 70, 90, 60, 100, 40, 80],
    [B, B, B, P, C, B, B, B, B],
    "O(n log n)", "O(n log n)", "O(n²)", "O(log n)", "不安定 ×",
    "平均的に最速級。実用で広く使われる。")

algo_slide(
    7, "マージソート", "Merge Sort",
    "配列を半分ずつに分割し、小さくしてから\n整列しつつ「併合(merge)」して戻す。安定で最悪も速い。",
    ["配列を半分に分ける", "分割できなくなるまで再帰", "2つの整列列を比較しながら併合", "全体が1つに戻れば完成"],
    None, "2つの整列済み列を併合中（黄=比較）",
    [20, 50, 70, 30, 60, 90, 40, 80, 100],
    [C, B, B, C, B, B, B, B, B],
    "O(n log n)", "O(n log n)", "O(n log n)", "O(n)", "安定 ○",
    "最悪でも O(n log n) で安定。ただし追加メモリが必要。")

algo_slide(
    8, "ヒープソート", "Heap Sort",
    "配列を「最大ヒープ（親≧子の木）」に組み、根の最大値を\n末尾へ運んで確定…を繰り返す。",
    ["配列を最大ヒープに構築", "根(最大値)を末尾と交換し確定", "残りを再びヒープに整える", "全部確定するまで繰り返す"],
    None, "根の最大値を末尾へ移動して確定",
    [80, 60, 70, 30, 50, 20, 40, 90, 100],
    [P, B, B, B, B, B, B, S, S],
    "O(n log n)", "O(n log n)", "O(n log n)", "O(1)", "不安定 ×",
    "追加メモリ O(1) で最悪も O(n log n)。木構造の応用。")

# ====================================================================
# 11. 比較まとめ表
# ====================================================================
s = prs.slides.add_slide(BLANK)
slide_bg(s)
section_header(s, "アルゴリズム比較まとめ", "用途に合わせて選ぶための早見表", num=9)
table_rows = [
    ("アルゴリズム", "最良", "平均", "最悪", "メモリ", "安定", "ひとこと"),
    ("バブル",     "O(n)",      "O(n²)",      "O(n²)",      "O(1)",     "○", "学習用・最も素直"),
    ("選択",       "O(n²)",     "O(n²)",      "O(n²)",      "O(1)",     "×", "交換回数が最少"),
    ("挿入",       "O(n)",      "O(n²)",      "O(n²)",      "O(1)",     "○", "ほぼ整列済みに強い"),
    ("クイック",   "O(n log n)","O(n log n)", "O(n²)",      "O(log n)", "×", "平均最速・実用的"),
    ("マージ",     "O(n log n)","O(n log n)", "O(n log n)", "O(n)",     "○", "安定＆最悪も速い"),
    ("ヒープ",     "O(n log n)","O(n log n)", "O(n log n)", "O(1)",     "×", "省メモリで安定的速度"),
]
nrows = len(table_rows); ncols = len(table_rows[0])
tbl_x, tbl_y = Inches(0.5), Inches(1.5)
tbl_w, tbl_h = Inches(12.3), Inches(4.6)
gtbl = s.shapes.add_table(nrows, ncols, tbl_x, tbl_y, tbl_w, tbl_h).table
col_w = [2.1, 1.7, 1.7, 1.7, 1.4, 0.9, 2.8]
for i, w in enumerate(col_w):
    gtbl.columns[i].width = Inches(w)
for r in range(nrows):
    gtbl.rows[r].height = Inches(0.66)
    for c in range(ncols):
        cell = gtbl.cell(r, c)
        cell.margin_left = Inches(0.08); cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if r == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = ACCENT
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb = PANEL if r % 2 else PANEL2
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
        run = p.add_run(); run.text = table_rows[r][c]
        run.font.size = Pt(14 if r == 0 else 13)
        run.font.name = FONT
        run.font.bold = (r == 0)
        if r == 0:
            run.font.color.rgb = INK
        elif c in (1, 2, 3, 4):
            run.font.color.rgb = TEXT; run.font.name = MONO
        elif c == 5:
            run.font.color.rgb = SORTED if table_rows[r][c] == "○" else MUTED
        else:
            run.font.color.rgb = TEXT
add_text(s, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.8),
         "選び方の目安:  まず学ぶなら **挿入/選択** ・ 実用で迷ったら **クイック** ・ 安定性が必要なら **マージ** ・ メモリが厳しいなら **ヒープ**",
         size=14, color=TEXT, line_spacing=1.2)

# ====================================================================
# 12. まとめ & 演習
# ====================================================================
s = prs.slides.add_slide(BLANK)
slide_bg(s)
section_header(s, "まとめと演習", "手を動かして確かめよう", num=10)
card(s, Inches(0.5), Inches(1.5), Inches(6.0), Inches(3.0))
add_text(s, Inches(0.75), Inches(1.65), Inches(5.5), Inches(0.4), "今日のポイント", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.75), Inches(2.2), Inches(5.5), Inches(2.2),
         "・ ソートは **比較** と **交換** の繰り返し\n"
         "・ 速さは **計算量 (Big-O)** で測る\n"
         "・ O(n²) 組と O(n log n) 組で大きく差が出る\n"
         "・ 「速さ・メモリ・安定性」のバランスで選ぶ\n"
         "・ 万能な1つはない。場面で使い分ける",
         size=15, line_spacing=1.35)
card(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(3.0))
add_text(s, Inches(7.05), Inches(1.65), Inches(5.5), Inches(0.4), "演習（ビジュアライザで）", size=18, bold=True, color=ACCENT)
add_text(s, Inches(7.05), Inches(2.2), Inches(5.5), Inches(2.2),
         "1. 同じ配列でバブルとクイックの\n   「比較回数」を比べよう\n"
         "2. 要素数を10→60に増やすと\n   回数はどれだけ増える？\n"
         "3. ほぼ整列済みの列で挿入ソートが\n   速くなるのを確かめよう",
         size=15, line_spacing=1.3)
add_rect(s, Inches(0.5), Inches(4.9), Inches(12.3), Inches(1.5), PANEL2, line=ACCENT, radius=True)
add_text(s, Inches(0.8), Inches(5.05), Inches(11.7), Inches(0.5),
         "▶ 付属ビジュアライザで体験", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.8), Inches(5.55), Inches(11.7), Inches(0.8),
         "html/sort.html をブラウザで開く  →  🎲ランダム生成 →  「次へ」で1ステップずつ /  ◀戻る で逆再生 /  計算量カウントをリアルタイム表示",
         size=14, color=TEXT, font=MONO, line_spacing=1.2)
add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
         "おしまい — 自分の手で動かすのがいちばんの近道！", size=14, color=MUTED, align=PP_ALIGN.CENTER)

# ---- 保存 ----
out = "/Users/kimio/AI/Algo/sort.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
