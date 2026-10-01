#!/usr/bin/env python3
"""迷路探索の教材スライド (maze.pptx) を生成する。
docs/maze.html と対になる講義用スライド。
実際に迷路を生成して DFS / BFS / A* / ランダムウォークを走らせ、
その探索結果（訪問マス・経路）を図として描く。
"""
import random
from collections import deque
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---- カラーパレット (docs/maze.html と統一) ----
BG     = RGBColor(0xF4, 0xF6, 0xFB)
PANEL  = RGBColor(0xFF, 0xFF, 0xFF)
PANEL2 = RGBColor(0xEE, 0xF1, 0xF7)
LINE   = RGBColor(0xD6, 0xDC, 0xE8)
TEXT   = RGBColor(0x1A, 0x20, 0x30)
MUTED  = RGBColor(0x5D, 0x6B, 0x82)
ACCENT = RGBColor(0x2F, 0x7F, 0xD9)
INK    = RGBColor(0x06, 0x12, 0x2A)   # 明色面に乗せる濃文字
WALL   = RGBColor(0x2B, 0x33, 0x44)
OPEN   = RGBColor(0xE9, 0xED, 0xF5)
VISITED= RGBColor(0x2F, 0x5D, 0x8A)
FRONT  = RGBColor(0xFF, 0xD2, 0x4A)
CURRENT= RGBColor(0xFF, 0x6B, 0x6B)
PATH   = RGBColor(0x4A, 0xDE, 0x80)
START  = RGBColor(0x4A, 0xDE, 0x80)
GOAL   = RGBColor(0xC0, 0x84, 0xFC)

FONT = "Meiryo"
MONO = "Consolas"
EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

# ====================================================================
# 迷路生成と探索（描画用データを作る）
# ====================================================================
H = W = 15                       # 奇数。通路は奇数座標。
def gen_maze(seed, braid=0.0):
    rnd = random.Random(seed)
    grid = [[0]*W for _ in range(H)]     # 0=壁, 1=通路
    grid[1][1] = 1
    stack = [(1, 1)]
    dirs = [(-2, 0), (2, 0), (0, -2), (0, 2)]
    while stack:
        r, c = stack[-1]
        cand = [(r+dr, c+dc, dr, dc) for dr, dc in dirs
                if 0 < r+dr < H-1 and 0 < c+dc < W-1 and grid[r+dr][c+dc] == 0]
        if not cand:
            stack.pop(); continue
        nr, nc, dr, dc = rnd.choice(cand)
        grid[r+dr//2][c+dc//2] = 1
        grid[nr][nc] = 1
        stack.append((nr, nc))
    # braiding: 行き止まりを確率で壊してループ（複数経路）を作る
    for r in range(1, H-1, 2):
        for c in range(1, W-1, 2):
            opens = sum(1 for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)] if grid[r+dr][c+dc] == 1)
            if opens == 1 and rnd.random() < braid:
                closed = [(dr, dc) for dr, dc in dirs
                          if 0 < r+dr < H-1 and 0 < c+dc < W-1 and grid[r+dr//2][c+dc//2] == 0]
                if closed:
                    dr, dc = rnd.choice(closed)
                    grid[r+dr//2][c+dc//2] = 1
    return grid

START_RC = (1, 1)
GOAL_RC = (H-2, W-2)

def nbrs(grid, r, c):
    out = []
    for dr, dc in ((-1, 0), (0, 1), (1, 0), (0, -1)):  # 上右下左
        nr, nc = r+dr, c+dc
        if 0 <= nr < H and 0 <= nc < W and grid[nr][nc] == 1:
            out.append((nr, nc))
    return out

def reconstruct(parent, goal):
    path = []
    p = goal
    while p is not None:
        path.append(p); p = parent.get(p)
    return path[::-1]

def run_dfs(grid):
    stack = [START_RC]; parent = {START_RC: None}; visited = []; seen = set()
    while stack:
        node = stack.pop()
        if node in seen: continue
        seen.add(node); visited.append(node)
        if node == GOAL_RC: break
        for nb in nbrs(grid, *node):
            if nb not in seen and nb not in parent:
                parent[nb] = node; stack.append(nb)
    return set(visited), reconstruct(parent, GOAL_RC)

def run_bfs(grid):
    q = deque([START_RC]); parent = {START_RC: None}; visited = []; seen = {START_RC}
    while q:
        node = q.popleft(); visited.append(node)
        if node == GOAL_RC: break
        for nb in nbrs(grid, *node):
            if nb not in seen:
                seen.add(nb); parent[nb] = node; q.append(nb)
    return set(visited), reconstruct(parent, GOAL_RC)

def run_astar(grid):
    gr, gc = GOAL_RC
    def h(n): return abs(n[0]-gr) + abs(n[1]-gc)
    openset = {START_RC}; parent = {START_RC: None}
    g = {START_RC: 0}; closed = set(); visited = []
    while openset:
        node = min(openset, key=lambda n: g[n] + h(n))
        openset.remove(node)
        if node in closed: continue
        closed.add(node); visited.append(node)
        if node == GOAL_RC: break
        for nb in nbrs(grid, *node):
            if nb in closed: continue
            t = g[node] + 1
            if t < g.get(nb, 1e9):
                g[nb] = t; parent[nb] = node; openset.add(nb)
    return set(visited), reconstruct(parent, GOAL_RC)

def run_random(grid, seed):
    rnd = random.Random(seed)
    node = START_RC; prev = None; visited = {node}; steps = 0
    cap = H*W*8
    while node != GOAL_RC and steps < cap:
        ns = nbrs(grid, *node)
        choices = [n for n in ns if n != prev] or ns
        nxt = rnd.choice(choices)
        prev = node; node = nxt; visited.add(node); steps += 1
    return visited, steps, node == GOAL_RC

def run_wall(grid, right_hand):
    """壁伝い法（右手法/左手法）。地図も座標も持たず、向きと『どちらの手を
    壁につけるか』だけで進む。docs/maze.html の wallFollower と同じ規則。"""
    DIR = [(-1, 0), (0, 1), (1, 0), (0, -1)]     # 0:北 1:東 2:南 3:西
    node = START_RC; d = 1                        # 最初は東を向く
    visited = {node}; steps = 0; cap = H*W*8
    while node != GOAL_RC and steps < cap:
        # 右手法は「右→直進→左→後ろ」、左手法はその左右逆
        order = ([(d+1) % 4, d, (d+3) % 4, (d+2) % 4] if right_hand
                 else [(d+3) % 4, d, (d+1) % 4, (d+2) % 4])
        moved = False
        for nd in order:
            nr, nc = node[0]+DIR[nd][0], node[1]+DIR[nd][1]
            if 0 <= nr < H and 0 <= nc < W and grid[nr][nc] == 1:
                d = nd; node = (nr, nc); visited.add(node); steps += 1
                moved = True; break
        if not moved:
            break
    return visited, steps, node == GOAL_RC

# 迷路を1つ固定（再現性のためシード固定）。ループ入りで各アルゴリズムの個性が出る。
MAZE = gen_maze(18, 0.25)
dfs_v, dfs_p   = run_dfs(MAZE)
bfs_v, bfs_p   = run_bfs(MAZE)
astar_v, astar_p = run_astar(MAZE)
rand_v, rand_steps, rand_ok = run_random(MAZE, 129)
wallL_v, wallL_steps, wallL_ok = run_wall(MAZE, False)   # 左手法
wallR_v, wallR_steps, wallR_ok = run_wall(MAZE, True)    # 右手法

# ====================================================================
# 描画ヘルパ
# ====================================================================
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
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = line_spacing
        for j, seg in enumerate(ln.split("**")):
            if seg == "": continue
            r = p.add_run(); r.text = seg
            r.font.size = Pt(size); r.font.name = font
            r.font.color.rgb = ACCENT if (j % 2 == 1) else color
            r.font.bold = bold or (j % 2 == 1)
    return tb

def add_rect(slide, x, y, w, h, fill, line=None, radius=False):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(t, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp

def add_oval(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    shp.line.fill.background(); shp.shadow.inherit = False
    return shp

def card(slide, x, y, w, h):
    return add_rect(slide, x, y, w, h, PANEL, line=LINE, radius=True)

def section_header(slide, title, subtitle=None, num=None):
    add_rect(slide, 0, 0, EMU_W, Inches(1.15), PANEL)
    add_rect(slide, 0, Inches(1.15), EMU_W, Pt(3), ACCENT)
    if num is not None:
        b = add_rect(slide, Inches(0.45), Inches(0.3), Inches(0.55), Inches(0.55), ACCENT, radius=True)
        p = b.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(num)
        r.font.size = Pt(22); r.font.bold = True; r.font.name = FONT; r.font.color.rgb = INK
    tx = Inches(1.2) if num is not None else Inches(0.5)
    add_text(slide, tx, Inches(0.18), Inches(11), Inches(0.6), title, size=28, bold=True)
    if subtitle:
        add_text(slide, tx, Inches(0.72), Inches(11.5), Inches(0.4), subtitle, size=14, color=MUTED)

def draw_maze(slide, ox, oy, total, visited=None, path=None, frontier=None):
    """迷路を描く。visited/path/frontier は (r,c) の集合/リスト。"""
    visited = visited or set(); frontier = frontier or set()
    pathset = set(path or [])
    cell = int(total / max(W, H))
    for r in range(H):
        for c in range(W):
            if MAZE[r][c] == 0:        col = WALL
            elif (r, c) in pathset:    col = PATH
            elif (r, c) in frontier:   col = FRONT
            elif (r, c) in visited:    col = VISITED
            else:                      col = OPEN
            add_rect(slide, Emu(int(ox + c*cell)), Emu(int(oy + r*cell)),
                     Emu(cell), Emu(cell), col)
    # S / G マーカー
    for (r, c), col, ch in ((START_RC, START, "S"), (GOAL_RC, GOAL, "G")):
        m = Emu(int(cell*0.16))
        ov = add_oval(slide, Emu(int(ox + c*cell + cell*0.16)), Emu(int(oy + r*cell + cell*0.16)),
                      Emu(int(cell*0.68)), Emu(int(cell*0.68)), col)
        tf = ov.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        rn = p.add_run(); rn.text = ch
        rn.font.size = Pt(9); rn.font.bold = True; rn.font.name = FONT; rn.font.color.rgb = INK
    return cell * W

def legend(slide, x, y, items):
    cx = x
    for label, col in items:
        add_rect(slide, cx, y + Inches(0.03), Inches(0.18), Inches(0.18), col, radius=True)
        add_text(slide, cx + Inches(0.24), y - Inches(0.02), Inches(2.0), Inches(0.3),
                 label, size=11, color=MUTED)
        cx += Inches(0.24) + Inches(len(label) * 0.16 + 0.35)

# ====================================================================
# 1. タイトル
# ====================================================================
s = prs.slides.add_slide(BLANK); slide_bg(s)
draw_maze(s, Inches(8.7), Inches(1.55), Inches(4.3), visited=bfs_v, path=bfs_p)
add_text(s, Inches(0.7), Inches(2.0), Inches(8), Inches(1.0),
         "迷路で学ぶ探索アルゴリズム", size=46, bold=True)
add_text(s, Inches(0.7), Inches(3.1), Inches(8), Inches(0.6),
         "DFS ・ BFS ・ A* ・ 壁伝い法 ・ ランダムウォーク", size=22, color=ACCENT)
add_text(s, Inches(0.7), Inches(3.9), Inches(8), Inches(0.5),
         "「どう探すか」で、速さも経路も変わる", size=16, color=MUTED)
add_text(s, Inches(0.7), Inches(6.7), Inches(8), Inches(0.4),
         "対応ビジュアライザ:  docs/maze.html", size=13, color=MUTED, font=MONO)

# ====================================================================
# 2. 迷路＝グラフ探索
# ====================================================================
s = prs.slides.add_slide(BLANK); slide_bg(s)
section_header(s, "迷路は「グラフ探索」の問題", "通路をたどってスタート S からゴール G へ", num=1)
card(s, Inches(0.5), Inches(1.5), Inches(6.6), Inches(5.2))
add_text(s, Inches(0.75), Inches(1.65), Inches(6.1), Inches(0.4), "考え方", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.75), Inches(2.15), Inches(6.1), Inches(2.2),
         "・ 迷路の各マス = **ノード（点）**\n"
         "・ 隣り合う通路のつながり = **エッジ（辺）**\n"
         "・ S から G へ、つながりをたどって進む\n"
         "・ 一度行ったマスを覚えれば、無駄な往復を防げる",
         size=15, line_spacing=1.35)
add_text(s, Inches(0.75), Inches(4.35), Inches(6.1), Inches(0.4), "探索で出てくる言葉", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.75), Inches(4.85), Inches(6.1), Inches(1.6),
         "・ **フロンティア**: 次に調べる候補のマス（黄）\n"
         "・ **訪問済み**: もう調べたマス（青）\n"
         "・ **経路**: 見つかった道すじ（緑）\n"
         "違いは「フロンティアからどれを次に選ぶか」だけ！",
         size=15, line_spacing=1.35)
card(s, Inches(7.4), Inches(1.5), Inches(5.4), Inches(5.2))
add_text(s, Inches(7.65), Inches(1.62), Inches(5), Inches(0.35),
         "BFS で探索した例（青=訪問, 緑=最短経路）", size=12, color=MUTED)
draw_maze(s, Inches(7.7), Inches(2.1), Inches(4.7), visited=bfs_v, path=bfs_p)
legend(s, Inches(7.65), Inches(6.25),
       [("壁", WALL), ("通路", OPEN), ("訪問", VISITED), ("経路", PATH)])

# ====================================================================
# 3. フロンティアの選び方で性格が決まる
# ====================================================================
s = prs.slides.add_slide(BLANK); slide_bg(s)
section_header(s, "「次にどれを選ぶか」が決め手", "同じ枠組みでも、選び方で結果が大きく変わる", num=2)
cols = [
    ("DFS", "スタック", "最後に入れた\n候補を選ぶ", "奥へ奥へ進む", CURRENT),
    ("BFS", "キュー", "最初に入れた\n候補を選ぶ", "近い順に広げる", ACCENT),
    ("A*", "優先度キュー", "f = g + h が\n最小を選ぶ", "ゴール方向へ絞る", PATH),
    ("ランダム", "なし", "その場で\nランダムに選ぶ", "あてもなく歩く", FRONT),
]
cw = Inches(3.0); gap = Inches(0.18); x0 = Inches(0.5)
for i, (name, ds, how, char, col) in enumerate(cols):
    x = x0 + i*(cw+gap)
    card(s, x, Inches(1.6), cw, Inches(4.0))
    add_rect(s, x, Inches(1.6), cw, Inches(0.6), PANEL2, radius=True)
    add_text(s, x, Inches(1.66), cw, Inches(0.5), name, size=20, bold=True, color=col, align=PP_ALIGN.CENTER)
    add_text(s, x+Inches(0.2), Inches(2.4), cw-Inches(0.4), Inches(0.4), "データ構造", size=11, color=MUTED)
    add_text(s, x+Inches(0.2), Inches(2.75), cw-Inches(0.4), Inches(0.4), ds, size=15, bold=True)
    add_text(s, x+Inches(0.2), Inches(3.35), cw-Inches(0.4), Inches(0.4), "選び方", size=11, color=MUTED)
    add_text(s, x+Inches(0.2), Inches(3.7), cw-Inches(0.4), Inches(0.8), how, size=14, line_spacing=1.15)
    add_text(s, x+Inches(0.2), Inches(4.6), cw-Inches(0.4), Inches(0.4), "性格", size=11, color=MUTED)
    add_text(s, x+Inches(0.2), Inches(4.95), cw-Inches(0.4), Inches(0.5), char, size=15, bold=True, color=col)
add_text(s, Inches(0.5), Inches(5.9), Inches(12.3), Inches(0.8),
         "ポイント: フロンティア（次の候補）から **どれを取り出すか** を変えるだけで、"
         "探索の順番・速さ・得られる経路が変わる。", size=15, color=TEXT, line_spacing=1.2)

# ====================================================================
# アルゴリズム解説スライド共通
# ====================================================================
def algo_slide(num, name_jp, name_en, idea, steps, visited, path, note,
               struct, tcomp, scomp, shortest, badge_col, extra_stat=None):
    s = prs.slides.add_slide(BLANK); slide_bg(s)
    section_header(s, name_jp, name_en, num=num)
    # 左: 解説
    card(s, Inches(0.5), Inches(1.5), Inches(6.4), Inches(5.4))
    add_text(s, Inches(0.75), Inches(1.65), Inches(6.0), Inches(0.4), "考え方", size=17, bold=True, color=ACCENT)
    add_text(s, Inches(0.75), Inches(2.1), Inches(6.0), Inches(0.9), idea, size=14, line_spacing=1.25)
    add_text(s, Inches(0.75), Inches(3.05), Inches(6.0), Inches(0.4), "手順", size=17, bold=True, color=ACCENT)
    add_text(s, Inches(0.75), Inches(3.5), Inches(6.0), Inches(1.9),
             "\n".join(f"{i+1}. {t}" for i, t in enumerate(steps)), size=14, line_spacing=1.3)
    add_text(s, Inches(0.75), Inches(6.0), Inches(6.0), Inches(0.8),
             "ポイント: " + note, size=13, color=FRONT, line_spacing=1.2)
    # 右上: 迷路図（この迷路での実際の探索結果）
    card(s, Inches(7.2), Inches(1.5), Inches(5.6), Inches(3.55))
    vcount = len(visited)
    plen = f" / 経路 {len(path)} マス" if path else ""
    sub = f"この迷路での結果 — 訪問 {vcount} マス{plen}" + (f" / {extra_stat}" if extra_stat else "")
    add_text(s, Inches(7.45), Inches(1.6), Inches(5.2), Inches(0.35), sub, size=12, color=MUTED)
    draw_maze(s, Inches(8.55), Inches(2.05), Inches(2.9), visited=visited, path=path)
    legend(s, Inches(7.45), Inches(4.65),
           [("訪問", VISITED), ("経路", PATH), ("S", START), ("G", GOAL)])
    # 右下: 性質
    card(s, Inches(7.2), Inches(5.25), Inches(5.6), Inches(1.65))
    info = [("データ構造", struct), ("時間", tcomp), ("空間", scomp), ("最短経路", shortest)]
    cy = Inches(5.4)
    for k, v in info:
        add_text(s, Inches(7.45), cy, Inches(2.0), Inches(0.3), k, size=12, color=MUTED)
        add_text(s, Inches(9.3), cy, Inches(3.3), Inches(0.3), v, size=12, color=TEXT, font=MONO)
        cy += Inches(0.30)

algo_slide(
    3, "DFS（深さ優先探索）", "Depth-First Search",
    "行けるところまで奥へ進み、行き止まりになったら\n分岐点まで戻って別の道を試す。",
    ["スタックに S を入れる", "一番上を取り出して訪問", "未訪問の隣を上に積む", "ゴールまで繰り返す"],
    dfs_v, dfs_p,
    "実装が簡単で記憶も少ない。ただし遠回りの経路になりやすく、最短は保証されない。",
    "スタック", "O(V + E)", "O(V)", "× 保証なし", CURRENT)

algo_slide(
    4, "BFS（幅優先探索）", "Breadth-First Search",
    "スタートから近いマスを順に、波紋のように\n全方向へ等しく広げていく。",
    ["キューに S を入れる", "先頭を取り出して訪問", "未訪問の隣を後ろに並べる", "ゴールまで繰り返す"],
    bfs_v, bfs_p,
    "重みのない迷路では最短経路を必ず見つけられる。ただし広く探すので訪問マスは多い。",
    "キュー", "O(V + E)", "O(V)", "○ 最短を保証", ACCENT)

algo_slide(
    5, "A*（エースター）", "A-star Search",
    "実コスト g とゴールまでの推定 h を足した\nf = g + h が小さいマスを優先して選ぶ。",
    ["g=距離, h=ゴールまでの推定", "f=g+h が最小のマスを選ぶ", "隣の g を更新して候補に追加", "ゴールまで繰り返す"],
    astar_v, astar_p,
    "h（推定）が妥当なら最短を保証しつつ、ゴール方向に絞って探すので BFS より無駄が少ない。",
    "優先度キュー", "O(E)", "O(V)", "○ 最短を保証", PATH)

algo_slide(
    6, "ランダムウォーク", "Random Walk",
    "今いるマスから、隣の通路をランダムに選んで\n進むだけ。記憶も計画もしない。",
    ["S に立つ", "隣の通路を1つ無作為に選ぶ", "そこへ移動（同じ道も通る）", "運よくゴールに着くまで歩く"],
    rand_v, None,
    "同じ所を何度も通り、非常に効率が悪い。『戦略がないとどうなるか』の比較用。",
    "なし", "期待値は大（不定）", "O(1)", "× 保証なし", FRONT,
    extra_stat=(f"{rand_steps} 歩で到達" if rand_ok else f"{rand_steps} 歩でも未到達"))

# ====================================================================
# 7. 壁伝い法（左手法 / 右手法）— 全体を見ない局所ルール
# ====================================================================
s = prs.slides.add_slide(BLANK); slide_bg(s)
section_header(s, "壁伝い法（左手法・右手法）", "Wall Follower — 地図を持たず、片手を壁につけて歩く", num=7)
# 左: 解説
card(s, Inches(0.5), Inches(1.5), Inches(6.4), Inches(5.4))
add_text(s, Inches(0.75), Inches(1.65), Inches(6.0), Inches(0.4), "考え方", size=17, bold=True, color=ACCENT)
add_text(s, Inches(0.75), Inches(2.1), Inches(6.0), Inches(1.0),
         "片方の手をずっと壁につけたまま歩く。分かれ道でも"
         "**つけている側の壁**に沿って曲がるだけ。\n"
         "迷路全体の地図も、座標も、訪問記録もいらない（記憶 O(1)）。",
         size=14, line_spacing=1.25)
add_text(s, Inches(0.75), Inches(3.35), Inches(6.0), Inches(0.4), "手順（右手法の例）", size=17, bold=True, color=ACCENT)
add_text(s, Inches(0.75), Inches(3.8), Inches(6.0), Inches(1.6),
         "1. 進む向きを1つ決める\n"
         "2. 右手側が空いていれば右へ曲がる\n"
         "3. ふさがっていれば 直進 → 左 → 後戻り の順で試す\n"
         "4. これを繰り返すだけ（左手法は左右が逆）",
         size=14, line_spacing=1.3)
add_text(s, Inches(0.75), Inches(5.95), Inches(6.0), Inches(0.9),
         "ポイント: **ループの無い迷路なら必ずゴールに着く**（ただし遠回り）。"
         "ループのある迷路では同じ所を巡り続けて戻れないことがある。",
         size=13, color=FRONT, line_spacing=1.2)
# 右: 同じ迷路で左手 / 右手の歩いた跡を並べる
card(s, Inches(7.2), Inches(1.5), Inches(5.6), Inches(5.4))
add_text(s, Inches(7.45), Inches(1.6), Inches(5.1), Inches(0.35),
         "同じ迷路・同じルール。つける手が違うだけで道が変わる", size=12, color=MUTED)
add_text(s, Inches(7.6), Inches(2.0), Inches(2.4), Inches(0.35), "左手法", size=14, bold=True, color=ACCENT)
draw_maze(s, Inches(7.6), Inches(2.45), Inches(2.4), visited=wallL_v)
add_text(s, Inches(7.6), Inches(4.95), Inches(2.4), Inches(0.4),
         f"訪問 {len(wallL_v)} / {wallL_steps} 歩", size=12, color=TEXT)
add_text(s, Inches(10.2), Inches(2.0), Inches(2.4), Inches(0.35), "右手法", size=14, bold=True, color=PATH)
draw_maze(s, Inches(10.2), Inches(2.45), Inches(2.4), visited=wallR_v)
add_text(s, Inches(10.2), Inches(4.95), Inches(2.4), Inches(0.4),
         f"訪問 {len(wallR_v)} / {wallR_steps} 歩", size=12, color=TEXT)
legend(s, Inches(7.45), Inches(5.5),
       [("歩いた跡", VISITED), ("S", START), ("G", GOAL)])
add_text(s, Inches(7.45), Inches(6.0), Inches(5.1), Inches(0.85),
         "データ構造: なし（足元の規則だけ）  空間: O(1)  最短: ×\n"
         "DFS/BFS/A* が『全体を見て次を選ぶ』のに対し、壁伝いは『今の足元だけ』で決める。",
         size=12, color=MUTED, line_spacing=1.2)

# ====================================================================
# 8. 比較まとめ（同じ迷路での訪問マス数）
# ====================================================================
s = prs.slides.add_slide(BLANK); slide_bg(s)
section_header(s, "4つの探索を同じ迷路で比較", "訪問したマスが少ないほど効率的（経路長は最短なら同じ）", num=8)
# 表
rows = [
    ("アルゴリズム", "データ構造", "最短", "経路長", "訪問マス", "ひとこと"),
    ("DFS",     "スタック",     "×",  str(len(dfs_p)),   str(len(dfs_v)),   "遠回りしがち"),
    ("BFS",     "キュー",       "○",  str(len(bfs_p)),   str(len(bfs_v)),   "波紋状・必ず最短"),
    ("A*",      "優先度キュー", "○",  str(len(astar_p)), str(len(astar_v)), "ゴール方向・無駄少"),
    ("ランダム","なし",         "×",  f"{rand_steps}歩", str(len(rand_v)),  "あてもなく低効率"),
]
nrows, ncols = len(rows), len(rows[0])
tbl = s.shapes.add_table(nrows, ncols, Inches(0.5), Inches(1.55), Inches(7.0), Inches(3.4)).table
widths = [1.4, 1.7, 0.8, 1.0, 1.1]
for i, w in enumerate(widths):
    tbl.columns[i].width = Inches(w)
tbl.columns[5].width = Inches(7.0) - Inches(sum(widths))
for r in range(nrows):
    tbl.rows[r].height = Inches(0.68)
    for c in range(ncols):
        cell = tbl.cell(r, c)
        cell.margin_left = Inches(0.07); cell.margin_right = Inches(0.03)
        cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = ACCENT if r == 0 else (PANEL if r % 2 else PANEL2)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if c in (2,3,4) else PP_ALIGN.LEFT
        rn = p.add_run(); rn.text = rows[r][c]
        rn.font.size = Pt(13 if r == 0 else 12); rn.font.name = FONT; rn.font.bold = (r == 0)
        if r == 0: rn.font.color.rgb = INK
        elif c == 2: rn.font.color.rgb = PATH if rows[r][c] == "○" else MUTED
        elif c in (3, 4): rn.font.color.rgb = TEXT; rn.font.name = MONO
        else: rn.font.color.rgb = TEXT
# 訪問マス数バー
card(s, Inches(7.8), Inches(1.55), Inches(5.0), Inches(3.4))
add_text(s, Inches(8.0), Inches(1.7), Inches(4.6), Inches(0.4), "訪問マス数（短いほど効率的）", size=13, bold=True)
bars = [("DFS", len(dfs_v), CURRENT), ("BFS", len(bfs_v), ACCENT),
        ("A*", len(astar_v), PATH), ("ランダム", len(rand_v), FRONT)]
mx = max(v for _, v, _ in bars)
by = Inches(2.3)
for name, v, col in bars:
    add_text(s, Inches(8.0), by, Inches(1.1), Inches(0.35), name, size=12, color=TEXT)
    w = Inches(3.0) * (v / mx)
    add_rect(s, Inches(9.15), by + Inches(0.03), w, Inches(0.26), col, radius=True)
    add_text(s, Inches(9.15)+w+Inches(0.05), by, Inches(0.8), Inches(0.35), str(v), size=11, color=MUTED)
    by += Inches(0.62)
# 迷路サムネ2枚
card(s, Inches(0.5), Inches(5.15), Inches(12.3), Inches(1.95))
add_text(s, Inches(0.7), Inches(5.25), Inches(5), Inches(0.35), "BFS（広く探す）", size=12, color=ACCENT)
draw_maze(s, Inches(0.7), Inches(5.6), Inches(1.4), visited=bfs_v, path=bfs_p)
add_text(s, Inches(2.6), Inches(5.25), Inches(5), Inches(0.35), "A*（ゴール方向に絞る）", size=12, color=PATH)
draw_maze(s, Inches(2.6), Inches(5.6), Inches(1.4), visited=astar_v, path=astar_p)
add_text(s, Inches(4.6), Inches(5.45), Inches(8), Inches(1.5),
         f"BFS と A* は同じ最短 {len(bfs_p)} マス。でも A* は **訪問が少ない**"
         f"（{len(astar_v)} 対 {len(bfs_v)}）。\n"
         f"DFS は経路 {len(dfs_p)} マスと大遠回り（最短の {len(dfs_p)//len(bfs_p)} 倍超）。\n"
         f"ランダムは {rand_steps} 歩も歩いてやっと到達＝桁違いに非効率。\n"
         "※ h を 0 にすると A* は BFS と一致する。",
         size=14, line_spacing=1.25)

# ====================================================================
# 9. まとめ & 演習
# ====================================================================
s = prs.slides.add_slide(BLANK); slide_bg(s)
section_header(s, "まとめと演習", "手を動かして確かめよう", num=9)
card(s, Inches(0.5), Inches(1.5), Inches(6.0), Inches(3.1))
add_text(s, Inches(0.75), Inches(1.65), Inches(5.5), Inches(0.4), "今日のポイント", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.75), Inches(2.2), Inches(5.5), Inches(2.3),
         "・ 迷路探索は **フロンティアから次を選ぶ** 繰り返し\n"
         "・ 選び方（スタック/キュー/優先度）で性格が決まる\n"
         "・ **BFS・A*** は最短を保証、**DFS** はしない\n"
         "・ **A*** はゴール方向に絞れて訪問マスが少ない\n"
         "・ **壁伝い法** は地図なしの局所ルール、**ランダム** は戦略なし",
         size=15, line_spacing=1.35)
card(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(3.1))
add_text(s, Inches(7.05), Inches(1.65), Inches(5.5), Inches(0.4), "演習（ビジュアライザで）", size=18, bold=True, color=ACCENT)
add_text(s, Inches(7.05), Inches(2.2), Inches(5.5), Inches(2.3),
         "1. 同じ迷路で BFS と A* の\n   「訪問マス数」を比べよう\n"
         "2. 壁伝い法で左手法と右手法を切替\n   ゴールへの道はどう変わる？\n"
         "3. 「ループあり」の迷路だと\n   壁伝い法は戻れる？ 各法はどう変わる？",
         size=15, line_spacing=1.3)
add_rect(s, Inches(0.5), Inches(4.9), Inches(12.3), Inches(1.5), PANEL2, line=ACCENT, radius=True)
add_text(s, Inches(0.8), Inches(5.05), Inches(11.7), Inches(0.5),
         "▶ 付属ビジュアライザで体験", size=18, bold=True, color=ACCENT)
add_text(s, Inches(0.8), Inches(5.55), Inches(11.7), Inches(0.8),
         "docs/maze.html をブラウザで開く → 🌀新しい迷路 → アルゴリズムを選び「次へ」で1歩ずつ / "
         "◀戻るで逆再生 / ↻やり直しで同じ迷路を別アルゴリズムと比較",
         size=14, color=TEXT, font=MONO, line_spacing=1.2)
add_text(s, Inches(0.5), Inches(6.95), Inches(12.3), Inches(0.4),
         "おしまい — 自分の手で動かすのがいちばんの近道！", size=14, color=MUTED, align=PP_ALIGN.CENTER)

out = "/Users/kimio/AI/Algo/maze.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
print(f"訪問マス数  DFS={len(dfs_v)}  BFS={len(bfs_v)}  A*={len(astar_v)}  Random={len(rand_v)}({rand_steps}歩, {'到達' if rand_ok else '未到達'})")
print(f"経路長      DFS={len(dfs_p)}  BFS={len(bfs_p)}  A*={len(astar_p)}")
print(f"壁伝い法    左手={len(wallL_v)}マス({wallL_steps}歩,{'到達' if wallL_ok else '未到達'})  "
      f"右手={len(wallR_v)}マス({wallR_steps}歩,{'到達' if wallR_ok else '未到達'})")
