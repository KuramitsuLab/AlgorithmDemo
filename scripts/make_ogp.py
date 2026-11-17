#!/usr/bin/env python3
"""docs/*.html に OGP タグを入れ、サムネイル docs/img/X.png（1200x630）を Chrome ヘッドレスで撮影する。

撮影方法: 一時コピーの </body> 直前に教材ごとの JS（ADVANCE）を差し込んで「少し進めた状態」にし、
ページ全体を縦長ウィンドウで撮ってから、.layout（なければ .stage）の上端から 1200x630 を切り出す。
→ ヘッダ・設定バー・再生ボタンが写らない。

使い方: python3 scripts/make_ogp.py [X ...] [--force]
  --force : og:image が既にあるページも画像を撮り直す（タグは重複して入れない）
"""
import html, re, subprocess, sys, tempfile
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
DOCS, IMG = ROOT / "docs", ROOT / "docs" / "img"
BASE = "https://kuramitsulab.github.io/AlgorithmDemo"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
SKIP = {"gcd"}  # 未公開・作業中
W, H = 1200, 630

def fr(f):  # frames[] の f 割の位置へ
    return f"stop();goTo(Math.floor(frames.length*{f}));"

# 教材ごとの「少し進めた状態」にする JS。anchor を 0 にすると切り出さず先頭から撮る。
ADVANCE = {
    "sort": fr(0.45), "maze": fr(0.5), "bpe": fr(0.1), "sieve": fr(0.35),
    "beam": fr(0.5), "grad": fr(0.4), "jug": fr(0.4), "markov": fr(0.3),
    "mnist": fr(0.5), "nw": fr(0.5), "nw2": fr(0.5), "recommend": fr(0.5),
    "euclid": "document.querySelector('button[data-v=rect]').click();stop();goTo(2);",
    "bucket": "[2,6,2,6,3,6].forEach(apply);",
    "monte": "stop();var s=document.getElementById('nmax');s.value='10000';"
             "s.dispatchEvent(new Event('change'));stop();goTo(2000);",
    "ttt": "onCell(4);stop();goTo(Math.floor(frames.length*0.5));",
    "othello": "human=0;for(var k=0;k<4;k++){place(legal(pos,turn)[0]);}human=B;stop();"
               "onCell(legal(pos,turn)[0]);stop();goTo(Math.floor(frames.length*0.5));",
    "hanoi": "goTab('viz');goTo(Math.floor(frames.length*0.4));",
    "logic": "select('xor');Object.keys(state).slice(0,1).forEach(toggle);",
    "adder": "select('full');Object.keys(state).slice(0,2).forEach(toggle);",
    "tm": "['inc','inc','inc','write','right','inc','inc','write','right'].forEach(function(o){"
          "document.querySelector('[data-op='+o+']').click();});",
    "gpa": "",
}
ANCHOR_SEL = {"hanoi": "#modeViz .layout"}   # 既定は .layout / .stage 以外を基準にしたい教材
MARGIN = {"hanoi": -4}   # 切り出し位置の余白（既定 12px）
ANCHOR_TOP = {"logic", "adder", "tm"}   # 再生ボタンが無い対話型：先頭から撮る

def chrome(*args):
    return subprocess.run([CHROME, "--headless", "--disable-gpu", "--hide-scrollbars",
                           "--virtual-time-budget=3000", *args], capture_output=True, text=True)

def shoot(name):
    IMG.mkdir(exist_ok=True)
    src = (DOCS / f"{name}.html").read_text(encoding="utf-8")
    js = ADVANCE.get(name, "")
    sel = ANCHOR_SEL.get(name)
    q = f"document.querySelector('{sel}')" if sel else "(document.querySelector('.layout')||document.querySelector('.stage'))"
    probe = (f"var e={q};"
             "document.title='Y='+Math.round(e?e.getBoundingClientRect().top+scrollY:0);")
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d) / f"{name}.html"
        tmp.write_text(src.replace("</body>", f"<script>{js}{probe}</script></body>"), encoding="utf-8")
        y = 0
        if name not in ANCHOR_TOP:
            out = chrome(f"--window-size={W},{H}", "--dump-dom", f"file://{tmp}").stdout
            m = re.search(r"<title>Y=(\d+)", out)
            y = max(0, int(m.group(1)) - MARGIN.get(name, 12)) if m else 0
        full = Path(d) / "full.png"
        chrome(f"--window-size={W},{y + H}", f"--screenshot={full}", f"file://{tmp}")
        Image.open(full).crop((0, y, W, y + H)).save(IMG / f"{name}.png")

def tags(name, title):
    parts = re.split(r"\s*—\s*", title, maxsplit=1)
    t, d = parts[0], (parts[1] if len(parts) > 1 else parts[0])
    e = lambda s: html.escape(s, quote=True)
    return (f'<meta name="description" content="{e(title)}">\n'
            f'<meta property="og:type" content="website">\n'
            f'<meta property="og:title" content="{e(t)}">\n'
            f'<meta property="og:description" content="{e(d)}">\n'
            f'<meta property="og:url" content="{BASE}/{name}.html">\n'
            f'<meta property="og:image" content="{BASE}/img/{name}.png">\n'
            f'<meta property="og:image:width" content="{W}">\n'
            f'<meta property="og:image:height" content="{H}">\n'
            f'<meta name="twitter:card" content="summary_large_image">')

def main():
    force = "--force" in sys.argv
    names = [a for a in sys.argv[1:] if not a.startswith("--")] or \
            sorted(p.stem for p in DOCS.glob("*.html") if p.stem not in SKIP)
    for n in names:
        p = DOCS / f"{n}.html"
        s = p.read_text(encoding="utf-8")
        if "og:image" in s and not force:
            print(f"skip {n}"); continue
        if "og:image" not in s:
            m = re.search(r"<title>(.*?)</title>", s, re.S)
            p.write_text(s[:m.end()] + "\n" + tags(n, html.unescape(m.group(1).strip())) + s[m.end():],
                         encoding="utf-8")
        shoot(n); print(f"ok   {n}")

main()
