#!/usr/bin/env python3
"""勾配降下法の教材スライド (gradient.pptx) を生成する。
html/gradient.html と対になる講義用スライド。
実際に反復を回した軌跡（ボールの動き）と損失推移を作図する。
"""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

BG=RGBColor(0xF4,0xF6,0xFB); PANEL=RGBColor(0xFF,0xFF,0xFF); PANEL2=RGBColor(0xEE,0xF1,0xF7)
LINE=RGBColor(0xD6,0xDC,0xE8); TEXT=RGBColor(0x1A,0x20,0x30); MUTED=RGBColor(0x5D,0x6B,0x82)
ACCENT=RGBColor(0x2F,0x7F,0xD9); CURVE=RGBColor(0x2F,0x7F,0xD9); BALL=RGBColor(0xFF,0x6B,0x6B)
INK=RGBColor(0x06,0x12,0x2A)   # 明色面に乗せる濃文字
GRAD=RGBColor(0xFF,0xD2,0x4A); STEPC=RGBColor(0x4A,0xDE,0x80); MINIMA=RGBColor(0xC0,0x84,0xFC)
FONT="Meiryo"; MONO="Consolas"
EMU_W=Inches(13.333); EMU_H=Inches(7.5)

# ---- 関数と反復 ----
def bowl(x):   return 0.4*x*x
def bowl_d(x): return 0.8*x
def dbl(x):    return 0.06*x**4 - 0.5*x*x + 0.2*x
def dbl_d(x):  return 0.24*x**3 - x + 0.2

def descend(f, df, x, eta, steps=40, bound=12):
    xs=[x]
    for _ in range(steps):
        g=df(x); xn=x-eta*g
        if abs(xn)>bound: xs.append(max(-bound,min(bound,xn))); return xs,"発散"
        x=xn; xs.append(x)
        if abs(df(x))<1e-3: return xs,"収束"
    return xs,"途中"

prs=Presentation(); prs.slide_width=EMU_W; prs.slide_height=EMU_H
BLANK=prs.slide_layouts[6]

def slide_bg(s,c=BG): s.background.fill.solid(); s.background.fill.fore_color.rgb=c
def add_text(s,x,y,w,h,t,size=18,color=TEXT,bold=False,align=PP_ALIGN.LEFT,font=FONT,
             anchor=MSO_ANCHOR.TOP,ls=1.1):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,ln in enumerate(t.split("\n")):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align; p.line_spacing=ls
        for j,seg in enumerate(ln.split("**")):
            if seg=="":continue
            r=p.add_run(); r.text=seg; r.font.size=Pt(size); r.font.name=font
            r.font.color.rgb=ACCENT if j%2 else color; r.font.bold=bold or bool(j%2)
    return tb
def add_rect(s,x,y,w,h,fill,line=None,radius=False):
    t=MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh=s.shapes.add_shape(t,x,y,w,h); sh.fill.solid(); sh.fill.fore_color.rgb=fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb=line; sh.line.width=Pt(1)
    sh.shadow.inherit=False; return sh
def add_oval(s,x,y,w,h,fill,ring=None,ring_w=2.0):
    sh=s.shapes.add_shape(MSO_SHAPE.OVAL,x,y,w,h); sh.fill.solid(); sh.fill.fore_color.rgb=fill
    if ring is None: sh.line.fill.background()
    else: sh.line.color.rgb=ring; sh.line.width=Pt(ring_w)
    sh.shadow.inherit=False; return sh
def add_line(s,x1,y1,x2,y2,color,width=1.0,dash=None):
    cn=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,int(x1),int(y1),int(x2),int(y2))
    cn.line.color.rgb=color; cn.line.width=Pt(width); cn.shadow.inherit=False
    return cn
def card(s,x,y,w,h): return add_rect(s,x,y,w,h,PANEL,line=LINE,radius=True)
def header(s,title,sub=None,num=None):
    add_rect(s,0,0,EMU_W,Inches(1.15),PANEL); add_rect(s,0,Inches(1.15),EMU_W,Pt(3),ACCENT)
    if num is not None:
        b=add_rect(s,Inches(0.45),Inches(0.3),Inches(0.55),Inches(0.55),ACCENT,radius=True)
        p=b.text_frame.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
        r=p.add_run(); r.text=str(num); r.font.size=Pt(22); r.font.bold=True; r.font.name=FONT; r.font.color.rgb=INK
    tx=Inches(1.2) if num is not None else Inches(0.5)
    add_text(s,tx,Inches(0.18),Inches(11),Inches(0.6),title,size=28,bold=True)
    if sub: add_text(s,tx,Inches(0.72),Inches(11.5),Inches(0.4),sub,size=14,color=MUTED)

def yrange(f,xmin,xmax):
    ys=[f(xmin+(xmax-xmin)*i/200) for i in range(201)]
    lo,hi=min(ys),max(ys); pad=(hi-lo)*0.12+1e-6; return lo-pad*0.4,hi+pad

def draw_func(s, ox,oy,w,h, f,df, xmin,xmax, traj=None, color=CURVE, minima=None,
              dot_color=BALL, show_axis=True):
    ylo,yhi=yrange(f,xmin,xmax)
    ox,oy,w,h=int(ox),int(oy),int(w),int(h)
    def X(x): return int(ox + (x-xmin)/(xmax-xmin)*w)
    def Y(y): return int(oy + h - (y-ylo)/(yhi-ylo)*h)
    if show_axis:
        add_line(s,ox,oy+h,ox+w,oy+h,LINE,1.0)
    # 最小値の縦線
    for m in (minima or []):
        add_line(s,X(m),oy,X(m),oy+h,MINIMA,1.0)
    # 曲線（freeform）
    pts=[(X(xmin+(xmax-xmin)*i/120), Y(f(xmin+(xmax-xmin)*i/120))) for i in range(121)]
    fb=s.shapes.build_freeform(pts[0][0],pts[0][1],scale=1)
    fb.add_line_segments(pts[1:],close=False)
    sh=fb.convert_to_shape(); sh.fill.background(); sh.line.color.rgb=color; sh.line.width=Pt(2.5)
    sh.shadow.inherit=False
    # 軌跡
    if traj:
        # 接続線（薄い緑）
        for i in range(len(traj)-1):
            add_line(s,X(traj[i]),Y(f(traj[i])),X(traj[i+1]),Y(f(traj[i+1])),STEPC,1.25)
        for i,xx in enumerate(traj):
            r=Emu(int(min(w,h)*0.020))
            c = dot_color if i==len(traj)-1 else RGBColor(0xE8,0x88,0x88)
            if i==0: c=GRAD
            add_oval(s,Emu(X(xx))-r,Emu(Y(f(xx)))-r,2*r,2*r,c)
    return X,Y

# ====================================================================
# 1. タイトル
# ====================================================================
s=prs.slides.add_slide(BLANK); slide_bg(s)
tr,_=descend(bowl,bowl_d,-4.0,0.6)
draw_func(s, Inches(7.6),Inches(2.2),Inches(5.0),Inches(3.0), bowl,bowl_d,-5,5, traj=tr, minima=[0])
add_text(s,Inches(0.7),Inches(2.0),Inches(7),Inches(1.0),"勾配降下法",size=50,bold=True)
add_text(s,Inches(0.7),Inches(3.15),Inches(7),Inches(0.6),"AI が「学習する」仕組みの心臓",size=22,color=ACCENT)
add_text(s,Inches(0.7),Inches(3.95),Inches(7),Inches(0.5),"誤差の坂を下って、いちばん低い谷を探す",size=16,color=MUTED)
add_text(s,Inches(0.7),Inches(6.7),Inches(7),Inches(0.4),"対応ビジュアライザ:  html/gradient.html",size=13,color=MUTED,font=MONO)

# ====================================================================
# 2. 学習＝誤差の最小化
# ====================================================================
s=prs.slides.add_slide(BLANK); slide_bg(s)
header(s,"AI の学習＝「誤差を小さくする」","正解とのズレ（損失）が最小になるパラメータを探す",num=1)
card(s,Inches(0.5),Inches(1.5),Inches(6.3),Inches(5.2))
add_text(s,Inches(0.75),Inches(1.65),Inches(5.8),Inches(0.4),"考え方",size=18,bold=True,color=ACCENT)
add_text(s,Inches(0.75),Inches(2.15),Inches(5.8),Inches(2.0),
         "・ モデルには調整つまみ＝**パラメータ**(重み)がある\n"
         "・ 予測と正解のズレを **損失 f(x)** で測る\n"
         "・ 損失を一番小さくするつまみの値を見つけたい\n"
         "・ でも、つまみは何百万個。手では無理",size=15,ls=1.35)
add_text(s,Inches(0.75),Inches(4.2),Inches(5.8),Inches(0.4),"勾配降下法のアイデア",size=18,bold=True,color=ACCENT)
add_text(s,Inches(0.75),Inches(4.7),Inches(5.8),Inches(1.6),
         "目隠しで山にいる人が、足元の **傾き** だけを\n"
         "頼りに、低い方へ一歩ずつ下りていく。\n"
         "やがて谷の底（損失が最小）にたどり着く。",size=15,ls=1.4)
card(s,Inches(7.1),Inches(1.5),Inches(5.7),Inches(5.2))
add_text(s,Inches(7.35),Inches(1.62),Inches(5.2),Inches(0.35),"損失の地形を下るボール",size=13,color=MUTED)
draw_func(s, Inches(7.7),Inches(2.4),Inches(4.5),Inches(2.7), bowl,bowl_d,-5,5,
          traj=descend(bowl,bowl_d,-4.0,0.6)[0], minima=[0])
add_text(s,Inches(7.35),Inches(5.5),Inches(5.2),Inches(1.1),
         "横軸 = パラメータ x、縦軸 = 損失 f(x)。\n"
         "黄=スタート、赤=現在地。低い方へ一歩ずつ。",size=12.5,color=MUTED,ls=1.35)

# ====================================================================
# 3. 1ステップの中身（更新式）
# ====================================================================
s=prs.slides.add_slide(BLANK); slide_bg(s)
header(s,"1ステップの中身","勾配（傾き）の逆向きに、学習率の分だけ動く",num=2)
add_rect(s,Inches(0.5),Inches(1.55),Inches(12.3),Inches(1.1),PANEL2,line=ACCENT,radius=True)
add_text(s,Inches(0.5),Inches(1.72),Inches(12.3),Inches(0.8),
         "x  ←  x  −  η · f′(x)",size=30,bold=True,align=PP_ALIGN.CENTER,font=MONO)
items=[("x","パラメータ（つまみ）","いま調整している値",ACCENT),
       ("f′(x)","勾配（傾き）","坂がどっち向きに急か",GRAD),
       ("−","逆向き","上りではなく下りへ",STEPC),
       ("η","学習率","一歩の大きさ",BALL)]
cw=Inches(2.95); x0=Inches(0.5)
for i,(sym,name,desc,col) in enumerate(items):
    x=x0+i*(cw+Inches(0.13))
    card(s,x,Inches(2.95),cw,Inches(1.9))
    add_text(s,x,Inches(3.1),cw,Inches(0.6),sym,size=26,bold=True,color=col,align=PP_ALIGN.CENTER,font=MONO)
    add_text(s,x+Inches(0.2),Inches(3.8),cw-Inches(0.4),Inches(0.4),name,size=14,bold=True,align=PP_ALIGN.CENTER)
    add_text(s,x+Inches(0.2),Inches(4.2),cw-Inches(0.4),Inches(0.6),desc,size=12,color=MUTED,align=PP_ALIGN.CENTER,ls=1.2)
card(s,Inches(0.5),Inches(5.05),Inches(12.3),Inches(1.85))
add_text(s,Inches(0.75),Inches(5.2),Inches(11.8),Inches(0.4),"なぜ「傾きの逆向き」？",size=16,bold=True,color=ACCENT)
add_text(s,Inches(0.75),Inches(5.65),Inches(11.8),Inches(1.1),
         "傾き f′(x) は「x を増やすと損失がどれだけ増えるか」。プラスなら右が上り坂 → 左（−方向）へ動けば損失が減る。\n"
         "だから **勾配の符号と逆** に動かす。これを何百万個のパラメータすべてに同時に行うのが AI の学習。",
         size=14.5,ls=1.4)

# ====================================================================
# 4. 学習率の効き方
# ====================================================================
s=prs.slides.add_slide(BLANK); slide_bg(s)
header(s,"学習率 η が結果を左右する","小さすぎ→遅い / ちょうど→収束 / 大きすぎ→発散",num=3)
panels=[("η = 0.1（小さすぎ）","なかなか進まない＝学習が遅い",0.1,MUTED),
        ("η = 0.6（ちょうど良い）","すっと谷の底へ収束",0.6,STEPC),
        ("η = 2.6（大きすぎ）","行き過ぎて振動し発散",2.6,BALL)]
pw=Inches(4.0)
for i,(title,sub,eta,col) in enumerate(panels):
    x=Inches(0.5)+i*(pw+Inches(0.15))
    card(s,x,Inches(1.55),pw,Inches(5.3))
    add_text(s,x+Inches(0.2),Inches(1.68),pw-Inches(0.4),Inches(0.4),title,size=16,bold=True,color=col)
    add_text(s,x+Inches(0.2),Inches(2.1),pw-Inches(0.4),Inches(0.4),sub,size=12.5,color=MUTED)
    tr,st=descend(bowl,bowl_d,-4.0,eta)
    draw_func(s, x+Inches(0.25),Inches(2.7),pw-Inches(0.5),Inches(2.6), bowl,bowl_d,-5,5,
              traj=tr, minima=[0])
    nsteps = len(tr)-1
    msg = (f"{nsteps} 歩でもまだ底に届かず" if st=="途中"
           else f"{nsteps} 歩で収束 ✅" if st=="収束" else f"{nsteps} 歩で発散 ⚠️")
    add_text(s,x+Inches(0.2),Inches(5.5),pw-Inches(0.4),Inches(0.5),msg,size=14,bold=True,color=col)
    add_text(s,x+Inches(0.2),Inches(6.0),pw-Inches(0.4),Inches(0.8),
             "黄=開始、赤=最後の位置。点は1ステップごとの位置。",size=11.5,color=MUTED,ls=1.25)

# ====================================================================
# 5. 局所最適のわな
# ====================================================================
s=prs.slides.add_slide(BLANK); slide_bg(s)
header(s,"局所最適のわな","谷が複数あると、スタート位置で行き着く先が変わる",num=4)
card(s,Inches(0.5),Inches(1.5),Inches(7.4),Inches(5.3))
add_text(s,Inches(0.75),Inches(1.62),Inches(6.8),Inches(0.35),
         "二つの谷。左から始めると深い谷（全体最適）、右から始めると浅い谷（局所最適）へ",size=12.5,color=MUTED)
# 左スタート（深い谷へ）
trL,_=descend(dbl,dbl_d,-3.0,0.18)
draw_func(s, Inches(0.9),Inches(2.4),Inches(6.6),Inches(3.3), dbl,dbl_d,-3.6,3.6,
          traj=trL, minima=[-2.0,2.0], dot_color=ACCENT)
# 右スタート（浅い谷へ）— 点だけ別色で重ねる
trR,_=descend(dbl,dbl_d,1.0,0.18)
ylo,yhi=yrange(dbl,-3.6,3.6)
def Xc(x): return int(Inches(0.9)+(x-(-3.6))/7.2*Inches(6.6))
def Yc(y): return int(Inches(2.4)+Inches(3.3)-(y-ylo)/(yhi-ylo)*Inches(3.3))
for i in range(len(trR)-1):
    add_line(s,Xc(trR[i]),Yc(dbl(trR[i])),Xc(trR[i+1]),Yc(dbl(trR[i+1])),RGBColor(0xFF,0xA0,0x4A),1.25)
for i,xx in enumerate(trR):
    r=Emu(int(Inches(6.6)*0.020))
    add_oval(s,Emu(Xc(xx))-r,Emu(Yc(dbl(xx)))-r,2*r,2*r,RGBColor(0xFF,0xA0,0x4A))
card(s,Inches(8.1),Inches(1.5),Inches(4.7),Inches(5.3))
add_text(s,Inches(8.3),Inches(1.65),Inches(4.3),Inches(0.4),"何が起きている？",size=16,bold=True,color=ACCENT)
add_text(s,Inches(8.3),Inches(2.15),Inches(4.3),Inches(2.6),
         "● **青** … x=−3 から開始\n"
         f"  → 深い谷 (x≈{trL[-1]:.1f}) ＝全体最適\n\n"
         "● **橙** … x=+1 から開始\n"
         f"  → 浅い谷 (x≈{trR[-1]:.1f}) ＝局所最適\n\n"
         "勾配降下は「今いる谷」しか下れない。",size=14,ls=1.4)
add_text(s,Inches(8.3),Inches(5.0),Inches(4.3),Inches(1.6),
         "対策の例:\n・スタートを変えて何度も試す\n・モーメンタム/学習率の工夫\n・乱数を加える（確率的勾配降下）",
         size=13,color=MUTED,ls=1.35)

# ====================================================================
# 6. まとめ & 演習
# ====================================================================
s=prs.slides.add_slide(BLANK); slide_bg(s)
header(s,"まとめと演習","手を動かして確かめよう",num=5)
card(s,Inches(0.5),Inches(1.5),Inches(6.0),Inches(3.1))
add_text(s,Inches(0.75),Inches(1.65),Inches(5.5),Inches(0.4),"今日のポイント",size=18,bold=True,color=ACCENT)
add_text(s,Inches(0.75),Inches(2.2),Inches(5.5),Inches(2.3),
         "・ 学習＝**損失を最小化**するパラメータ探し\n"
         "・ **勾配（傾き）の逆向き**へ一歩ずつ下る\n"
         "・ 一歩の大きさ＝**学習率 η** が超重要\n"
         "・ 小さすぎ→遅い、大きすぎ→発散\n"
         "・ 谷が複数だと**局所最適**にはまることも",size=15,ls=1.35)
card(s,Inches(6.8),Inches(1.5),Inches(6.0),Inches(3.1))
add_text(s,Inches(7.05),Inches(1.65),Inches(5.5),Inches(0.4),"演習（ビジュアライザで）",size=18,bold=True,color=ACCENT)
add_text(s,Inches(7.05),Inches(2.2),Inches(5.5),Inches(2.3),
         "1. η を上げていくと何歩で収束する？\n   どこから発散する？\n"
         "2. 「二つの谷」でスタートを左右に\n   変えると行き先はどう変わる？\n"
         "3. 「でこぼこ」だと局所最適に\n   はまる様子を観察しよう",size=15,ls=1.3)
add_rect(s,Inches(0.5),Inches(4.9),Inches(12.3),Inches(1.5),PANEL2,line=ACCENT,radius=True)
add_text(s,Inches(0.8),Inches(5.05),Inches(11.7),Inches(0.5),"▶ 付属ビジュアライザで体験",size=18,bold=True,color=ACCENT)
add_text(s,Inches(0.8),Inches(5.55),Inches(11.7),Inches(0.8),
         "html/gradient.html を開く → 関数・学習率・スタートを設定 → 「次へ」で1ステップずつ降下 / "
         "損失グラフが下がる様子と発散をその場で確認",size=14,color=TEXT,font=MONO,ls=1.2)
add_text(s,Inches(0.5),Inches(6.95),Inches(12.3),Inches(0.4),
         "この一歩の積み重ねが、巨大な AI モデルの学習を支えている。",size=14,color=MUTED,align=PP_ALIGN.CENTER)

out="/Users/kimio/AI/Algo/gradient.pptx"
prs.save(out)
print("saved:",out,"slides:",len(prs.slides._sldIdLst))
