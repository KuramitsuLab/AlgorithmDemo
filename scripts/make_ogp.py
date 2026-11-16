#!/usr/bin/env python3
"""docs/*.html に OGP タグを挿入し、docs/img/X.png（1200x630）を Chrome ヘッドレスで撮影する。
使い方: python3 scripts/make_ogp.py [X ...]   （省略時は og:image 未設定の全教材）
既に og:image があるページは HTML をスキップ（--force で画像のみ撮り直し可）。"""
import html, re, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS, IMG = ROOT / "docs", ROOT / "docs" / "img"
BASE = "https://kuramitsulab.github.io/AlgorithmDemo"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
SKIP = {"gcd"}  # 未公開・作業中

def shoot(name):
    IMG.mkdir(exist_ok=True)
    subprocess.run([CHROME, "--headless", "--disable-gpu", "--hide-scrollbars",
                    "--window-size=1200,630", "--virtual-time-budget=3000",
                    f"--screenshot={IMG / (name + '.png')}", f"file://{DOCS / (name + '.html')}"],
                   check=True, capture_output=True)

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
            f'<meta property="og:image:width" content="1200">\n'
            f'<meta property="og:image:height" content="630">\n'
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
            s = s[:m.end()] + "\n" + tags(n, html.unescape(m.group(1).strip())) + s[m.end():]
            p.write_text(s, encoding="utf-8")
        shoot(n); print(f"ok   {n}")

main()
