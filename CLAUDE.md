# CLAUDE.md

アルゴリズムのビジュアル教材集。
各テーマごとに **インタラクティブな Web ビジュアライザ**（`docs/*.html`）と、それと対になる **講義用スライド**（`slides/*.pptx`、`scripts/make_*_pptx.py` で生成）をセットで用意する。
UI・解説はすべて日本語。テーマは「探索 → 状態空間 → 学習 → 生成」と AI 時代へ向かう流れで構成している。

## 教材セット一覧

| テーマ | ビジュアライザ | スライド生成 | 主な内容 |
|---|---|---|---|
| 整列 | `docs/sort.html` | `make_sort_pptx.py` → `sort.pptx` | バブル/選択/挿入/シェル/クイック/マージ/ヒープ |
| 迷路探索 | `docs/maze.html` | `make_maze_pptx.py` → `maze.pptx` | DFS / BFS / A* / ランダムウォーク、迷路生成(枝分かれ可変) |
| バケツ問題 | `docs/jug.html` | `make_jug_pptx.py` → `jug.pptx` | 状態空間グラフを DFS / BFS で探索 |
| 最適化 | `docs/grad.html` | `make_gradient_pptx.py` → `gradient.pptx` | 学習率/収束/発散/局所最適、オプティマイザ選択（SGD/Momentum/AdaGrad/RMSProp/Adam） |
| ビームサーチ | `docs/beam.html` | `make_beam_pptx.py` → `beam.pptx` | LLM の文章生成、貪欲(k=1) vs ビーム |
| レコメンド | `docs/rec.html` | （未作成） | 協調フィルタリング：ユーザー/アイテムベース、コサイン/ピアソン類似度 |
| ニューラルネット | `docs/nw.html` | （未作成） | 2次元分類を小さなMLPが学習：順伝播/逆伝播/勾配降下、活性化関数、決定境界 |
| マルコフ連鎖 | `docs/markov.html` | （未作成） | 文字単位 N-gram（1〜3）で『吾輩は猫である』(`neko.txt`)を学習し1文字ずつ生成：任意の書き出し(プロンプト)の続き・確率分布・くじ引き・温度・top-k/top-p・貪欲の暴走。解説 `docs/markov.md` |
| ニューラルネット2 | `docs/nw2.html` | `make_nw2_pptx.py` → `nw2.pptx` | nw.html を多層化：隠れ層 1〜4 層を選べる（深さ＝表現力 vs 勾配消失）。講義ノート `nw2.md`／TeX清書 `nw2.tex`（lualatex）→ `nw2.pdf` あり |

**命名規約**: 1テーマ = 短い名前 `X` で統一 → `docs/X.html` / `scripts/make_X_pptx.py` / `slides/X.pptx`。

## ビルド・実行

```bash
python3 scripts/make_sort_pptx.py        # → sort.pptx を生成（他テーマも同様）
open docs/sort.html              # ブラウザでビジュアライザを開く
open slides/sort.pptx                   # スライドを開く
```

- ビジュアライザは**単一ファイル完結・外部ライブラリ非依存**。サーバ不要、`open` でそのまま動く。
- スライド生成には **python-pptx**（環境に 1.0.2）。`python3` は `/Users/kimio/python-3.11.7`。
- JS 構文チェック: `node -e "...new Function(script)..."`（DOM は実行しないので構文確認のみ）。

## 共通カラーパレット（HTML / Python で一致させる）

**白基調（ライトテーマ）**。背景・パネル・文字などのニュートラル系は明るく、アクセント・状態色（黄/赤/青/緑/紫）はそのまま使う。

| 意味 | HEX | HTML 変数 | 役割の例 |
|---|---|---|---|
| 背景 | `#f4f6fb` | `--bg` | |
| パネル | `#ffffff` / `#eef1f7` | `--panel` / `--panel2` | カード背景 |
| 罫線 | `#d6dce8` | `--line` | |
| 文字 | `#1a2030` / `#5d6b82` | `--text` / `--muted` | 本文 / 補足 |
| アクセント | `#2f7fd9` | `--accent` | 見出し・強調・BFS |
| 黄 | `#ffd24a` | compare/frontier/cand | 比較中・候補 |
| 赤 | `#ff6b6b` | swap/current/ball | 交換・現在地 |
| 青 | `#2f5d8a` | visited | 訪問済み |
| 緑 | `#4ade80` | sorted/path/best/step | 確定・経路・最良 |
| 紫 | `#c084fc` | pivot/goal/minima/end | 基準・目標・最小値 |

補助色（テーマ固有）:
- **canvas 背景**: `#f7f9fc`（jug/gradient/beam/rec のキャンバス）。
- **maze の壁/通路**: 壁 `#2b3344` / 通路 `#e9edf5`（白地で「濃い壁・明るい通路」）。
- **明色面に乗せる濃文字**: `#06122a`（Python では `INK`）。バッジ番号・表ヘッダ（アクセント面）・明るいノードのラベル等。primary ボタン文字は白 `#ffffff`。
- **濃色面に乗せる淡文字**: `#e6ecf5`（Python では `ONDARK`）。訪問/未確定の濃いノードや枝刈りノードのラベル。

Python 側は同じ値を `RGBColor(0x..,0x..,0x..)` で定義。フォントは **Meiryo**（等幅は **Consolas**）。
（旧ダーク基調 `#0f1420` 系から白基調へ移行済み。HTML・pptx ともこのパレットで統一。）

## ビジュアライザの設計パターン（docs/*.html）

すべて同じ骨格・操作感で実装する。新規テーマもこれを踏襲する。

- **画面構成**: `header` → `controls`（アルゴリズム/パラメータ選択 + 速度スライダ）→
  `layout`（左: `stage` に canvas または棒、右: `side` に統計/性質/凡例カード）→ `playbar`。
- **フレーム録画方式（最重要）**: アルゴリズムは1ステップずつ状態を `yield`/push する
  **ジェネレータ or builder** で書き、各ステップのスナップショットを配列 `frames[]` に**全部記録**する。
  描画は `frames[cur]` を引くだけ。これにより**前進・逆再生・シーク**がすべて成立する。
  （例: `sort.html` の `record()`, `maze.html` の `dfsBfs()`, `jug.html` の `buildSearch()`,
  `grad.html` の `build()`, `beam.html` の `build()`）
- **共通の制御関数**: `rebuild()`/`regenerate()`（パラメータ変更で作り直し）, `renderFrame(i)`,
  `goTo(i)`, `play()`, `tick()`, `stop()`。`$ = id => document.getElementById(id)`。
- **playbar**: `⏮ / ◀ 戻る / ▶ 再生 / 次へ ▶ / ⏭` + `#progress` シークバー + `#stepLabel`（`i / 最大`）。
- **キーボード**: `←`/`→` で1ステップ、`Space` で再生切替（テーマにより `N`/`R` で再生成）。
- **速度**: スライダ値 → `setTimeout` の遅延に変換（おおむね `基準 - speed*係数` ms）。
- 状態は**コード（整数）で持ち**、`frames` には軽量スナップショット（`Uint8Array.slice()` 等）を保存。
- canvas は毎フレーム全再描画（要素数が中規模なので十分軽い）。グラフ系のエッジは弧で重なりを回避。

## スライド生成の設計パターン（scripts/make_*_pptx.py）

- **python-pptx**、16:9（`Inches(13.333) × Inches(7.5)`）、`slide_layouts[6]`（空白）。
- **実データで作図する**: スライド用の図は飾りでなく、**Python 側で実際にアルゴリズムを走らせた結果**
  （探索の訪問集合・経路、降下の軌跡、ビームの木など）を描く。再現性のため乱数は**シード固定**。
- **共通ヘルパ関数**（各スクリプトに同名で実装。新規もこれを流用）:
  `slide_bg`, `add_text`, `add_rect`, `add_oval`, `add_line`, `card`,
  `section_header`（= 一部スクリプトでは `header`）, `legend`。
  - `add_text` は本文中の `**…**` を**アクセント色＋太字**にインライン変換する（自前パーサ）。
  - グラフ/曲線は connector（`add_line`）と freeform（`build_freeform`）で描画。
- **スライド構成の定番アーク**: 表紙 → 概念 → 仕組み → 各アルゴリズム（統一レイアウト）→
  比較表（実数値）→ まとめ & 演習（対応する `docs/X.html` を案内）。
- **単位の注意**: 座標は EMU。`build_freeform` / connector の座標は **int(EMU)**。
  フォントサイズ `Pt(...)` は **100〜400000 の範囲**（EMU をそのまま渡さない。`rad/12700` で pt 換算）。

## 編集の指針

- パラメータや題材を変えるときは `scripts/make_X_pptx.py` 冒頭の定数（例 `A,B,T=5,3,4`、`gen_maze(seed, braid)`）を編集して再生成。
- 配色・操作・スライド構成は**テーマ間で一貫**させる。逸脱する場合は理由を明確に。
- ビジュアライザに新アルゴリズムを足すときは「ジェネレータを1本書いて、選択肢・性質テーブルに登録」で済む設計を保つ。
