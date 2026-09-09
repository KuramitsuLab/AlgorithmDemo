# -*- coding: utf-8 -*-
"""
nw2.pptx を生成するスクリプト（ニューラルネットワーク・多層版 講義スライド）。
 - 数式は matplotlib(mathtext) で画像化して埋め込む
 - 図は numpy で実際に MLP を学習させて描く（決定境界の学習進行・ネットワーク図・損失曲線）
 - 配色/レイアウトは CLAUDE.md の共通パレット・設計パターンに準拠（対になる教材: html/nw2.html, nw2.md）
実行: python3 make_nw2_pptx.py
"""
import os, struct, tempfile, itertools
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ====================================================================
# 共通パレット（CLAUDE.md）
# ====================================================================
BG    = RGBColor(0xf4,0xf6,0xfb)
PANEL = RGBColor(0xff,0xff,0xff)
PANEL2= RGBColor(0xee,0xf1,0xf7)
LINE  = RGBColor(0xd6,0xdc,0xe8)
TEXT  = RGBColor(0x1a,0x20,0x30)
MUTED = RGBColor(0x5d,0x6b,0x82)
ACCENT= RGBColor(0x2f,0x7f,0xd9)
YELLOW= RGBColor(0xff,0xd2,0x4a)
RED   = RGBColor(0xff,0x6b,0x6b)
BLUE2 = RGBColor(0x2f,0x5d,0x8a)
GREEN = RGBColor(0x4a,0xde,0x80)
PURPLE= RGBColor(0xc0,0x84,0xfc)
INK   = RGBColor(0x06,0x12,0x2a)
ONDARK= RGBColor(0xe6,0xec,0xf5)
FONT  = "Meiryo"
MONO  = "Consolas"
# hex 文字列（matplotlib 用）
HX = lambda c: "#%02x%02x%02x" % (c[0],c[1],c[2])
H_ACCENT, H_INK, H_RED, H_BLUE, H_YEL, H_GRN, H_PUR, H_MUT = \
  HX((0x2f,0x7f,0xd9)),HX((0x06,0x12,0x2a)),HX((0xff,0x6b,0x6b)),HX((0x2f,0x5d,0x8a)),\
  HX((0xff,0xd2,0x4a)),HX((0x2b,0xa5,0x5f)),HX((0xc0,0x84,0xfc)),HX((0x5d,0x6b,0x82))

EMU_IN = 914400
# 図は永続フォルダ figs_nw2/ に出力（nw2.tex の \includegraphics からも参照する）
ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figs_nw2")
os.makedirs(ASSETS, exist_ok=True)
_counter = itertools.count()

# ====================================================================
# PNG サイズ読み取り（PIL 非依存）
# ====================================================================
def png_size(path):
    with open(path,"rb") as f:
        head=f.read(24)
    w,h=struct.unpack(">II", head[16:24])
    return w,h

# ====================================================================
# 数式レンダリング（mathtext → 透過 PNG）
# ====================================================================
DPI = 300
def render_eq(latex, color=H_INK, fs=24):
    fig = plt.figure(figsize=(0.01,0.01))
    fig.text(0,0, "$%s$" % latex, fontsize=fs, color=color)
    path = os.path.join(ASSETS, "eq%d.png" % next(_counter))
    fig.savefig(path, dpi=DPI, transparent=True, bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    return path

def add_eq(s, latex, cx, cy, h_in, color=H_INK, fs=24, center=True):
    """数式を高さ h_in インチで配置。center=True なら (cx,cy) を中心、False なら左上。"""
    p = render_eq(latex, color, fs)
    pw,ph = png_size(p)
    ar = pw/ph
    w_in = h_in*ar
    if center:
        x = cx - w_in/2; y = cy - h_in/2
    else:
        x = cx; y = cy
    s.shapes.add_picture(p, Inches(x), Inches(y), Inches(w_in), Inches(h_in))
    return w_in, h_in

def add_image_fit(s, path, x, y, w, h):
    pw,ph = png_size(path); ar = pw/ph; boxar = w/h
    if ar > boxar: iw=w; ih=w/ar
    else: ih=h; iw=h*ar
    ix=x+(w-iw)/2; iy=y+(h-ih)/2
    s.shapes.add_picture(path, Inches(ix), Inches(iy), Inches(iw), Inches(ih))

# ====================================================================
# pptx 描画ヘルパ
# ====================================================================
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def slide_bg(s, color=BG):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color

def _set_runs(p, text, size, color, bold, font):
    """**...** を ACCENT+bold にインライン変換して runs を作る。"""
    parts = text.split("**")
    for i,seg in enumerate(parts):
        if seg=="" : continue
        r = p.add_run(); r.text = seg
        r.font.size = Pt(size); r.font.name = font
        if i%2==1:
            r.font.bold = True; r.font.color.rgb = ACCENT
        else:
            r.font.bold = bold; r.font.color.rgb = color

def add_text(s, x,y,w,h, text, size=18, color=TEXT, bold=False,
             align=PP_ALIGN.LEFT, font=FONT, anchor=MSO_ANCHOR.TOP, line_spacing=1.05):
    tb = s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    tf.vertical_anchor = anchor
    for li,line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if li==0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = line_spacing
        _set_runs(p, line, size, color, bold, font)
    return tb

def add_rrect(s, x,y,w,h, fill=PANEL, line=LINE, line_w=1.0, radius=0.055):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x),Inches(y),Inches(w),Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(line_w)
    sh.shadow.inherit = False
    try: sh.adjustments[0] = radius
    except Exception: pass
    return sh

def card(s, x,y,w,h, title=None, fill=PANEL, accent=None):
    sh = add_rrect(s, x,y,w,h, fill=fill, line=LINE)
    if accent:  # 左に細いアクセント帯
        bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x),Inches(y),Inches(0.09),Inches(h))
        bar.fill.solid(); bar.fill.fore_color.rgb = accent; bar.line.fill.background(); bar.shadow.inherit=False
        try: bar.adjustments[0]=0.5
        except Exception: pass
    if title:
        add_text(s, x+0.22, y+0.14, w-0.44, 0.4, title, size=15, bold=True, color=TEXT)
    return sh

def header(s, title, sub=None, num=None):
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0),Inches(0),Inches(13.333),Inches(0.95))
    bar.fill.solid(); bar.fill.fore_color.rgb = PANEL; bar.line.fill.background(); bar.shadow.inherit=False
    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0),Inches(0.95),Inches(13.333),Inches(0.035))
    accent.fill.solid(); accent.fill.fore_color.rgb = ACCENT; accent.line.fill.background(); accent.shadow.inherit=False
    add_text(s, 0.55, 0.14, 10.8, 0.6, title, size=27, bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    if sub:
        add_text(s, 0.57, 0.62, 10.8, 0.3, sub, size=13, color=MUTED)
    if num is not None:
        add_text(s, 12.2, 0.14, 0.9, 0.6, num, size=14, color=MUTED,
                 align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, font=MONO)

def legend(s, x, y, items, size=12):
    """items: [(color, label), ...] を横並び。"""
    cx = x
    for col,lab in items:
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx),Inches(y+0.03),Inches(0.16),Inches(0.16))
        dot.fill.solid(); dot.fill.fore_color.rgb = col; dot.line.fill.background(); dot.shadow.inherit=False
        add_text(s, cx+0.22, y-0.03, 2.4, 0.3, lab, size=size, color=MUTED)
        cx += 0.30 + 0.11*len(lab) + 0.35

def pagenum(s, n):
    add_text(s, 12.5, 7.08, 0.7, 0.3, "%d / %d" % (n, NPAGES), size=10, color=MUTED,
             align=PP_ALIGN.RIGHT, font=MONO)

# ====================================================================
# numpy 版 MLP（html/nw2.html と同じ数理）
# ====================================================================
def sigmoid(z): return 1.0/(1.0+np.exp(-np.clip(z,-30,30)))
ACT = {
  "tanh":   (lambda z: np.tanh(z),           lambda z,a: 1-a*a,        False),
  "relu":   (lambda z: np.maximum(0,z),      lambda z,a: (z>0)*1.0,    True ),
  "sigmoid":(lambda z: sigmoid(z),           lambda z,a: a*(1-a),      False),
}

def gen(name, seed=1, N=120):
    rng = np.random.default_rng(seed)
    if name=="circle":
        r = np.sqrt(rng.random(N))*np.where(rng.random(N)<0.5,0.5,1.0)
        t = rng.random(N)*2*np.pi
        x = r*np.cos(t)*0.95; y = r*np.sin(t)*0.95
        lab = ((x*x+y*y)<0.28).astype(float)
    elif name=="xor":
        x=(rng.random(N)*2-1)*0.92; y=(rng.random(N)*2-1)*0.92
        lab=((x*y)>0).astype(float)
    elif name=="sine":
        x=(rng.random(N)*2-1)*0.92; y=(rng.random(N)*2-1)*0.92
        lab=(y>0.5*np.sin(3.2*x)).astype(float)
    elif name=="blobs":
        c=(np.arange(N)>=N//2).astype(float)
        cx=np.where(c>0,0.45,-0.45); cy=np.where(c>0,0.45,-0.45)
        x=np.clip(cx+rng.standard_normal(N)*0.22,-1,1); y=np.clip(cy+rng.standard_normal(N)*0.22,-1,1)
        lab=c
    else:  # spiral
        n=N//2; xs=[]; ys=[]; ls=[]
        for c in range(2):
            i=np.arange(n); r=i/n*0.92
            th=1.75*i/n*2*np.pi + c*np.pi + rng.standard_normal(n)*0.12
            xs.append(r*np.cos(th)); ys.append(r*np.sin(th)); ls.append(np.full(n,c,float))
        x=np.concatenate(xs); y=np.concatenate(ys); lab=np.concatenate(ls)
    X=np.stack([x,y],0)   # (2,N)
    return X, lab

def init_net(sizes, seed, he):
    rng=np.random.default_rng(seed*104729+7)
    W=[]; b=[]
    for l in range(len(sizes)-1):
        fin,fout=sizes[l],sizes[l+1]
        s=(np.sqrt(2) if he else 1.0)/np.sqrt(fin)
        W.append(rng.standard_normal((fout,fin))*s); b.append(np.zeros(fout))
    return W,b

def forward_all(W,b,act,A0):
    f=act[0]; A=[A0]; Z=[]
    M=len(W)
    for l in range(M):
        z=W[l]@A[l]+b[l][:,None]
        a=sigmoid(z) if l==M-1 else f(z)
        Z.append(z); A.append(a)
    return A,Z

def train(name, sizes, actname="tanh", opt="adam", eta=0.3, epochs=800, seed=1, snap_epochs=None):
    f,df,he=ACT[actname]; act=(f,df,he)
    X,y=gen(name,seed); N=X.shape[1]; Y=y[None,:]
    W,b=init_net(sizes,seed,he); M=len(W)
    mom=0.9; b1=0.9; b2=0.999; eps=1e-8
    elr=eta*(0.1 if opt=="adam" else 1.0)
    vW=[np.zeros_like(w) for w in W]; vb=[np.zeros_like(bb) for bb in b]
    sW=[np.zeros_like(w) for w in W]; sb=[np.zeros_like(bb) for bb in b]
    hist_loss=[]; hist_acc=[]; snaps={}
    snap_epochs=set(snap_epochs or [])
    def record(ep):
        A,_=forward_all(W,b,act,X); p=A[-1]
        q=np.clip(p,1e-7,1-1e-7)
        loss=float(np.mean(-(Y*np.log(q)+(1-Y)*np.log(1-q))))
        acc=float(np.mean((p>0.5)==(Y>0.5)))
        hist_loss.append(loss); hist_acc.append(acc)
        if ep in snap_epochs:
            snaps[ep]=([w.copy() for w in W],[bb.copy() for bb in b])
        return loss,acc
    record(0)
    t=0
    for ep in range(1,epochs+1):
        A,Z=forward_all(W,b,act,X)
        delta=(A[-1]-Y)            # (1,N)
        gW=[None]*M; gb=[None]*M
        for l in range(M-1,-1,-1):
            gW[l]=delta@A[l].T / N
            gb[l]=delta.mean(axis=1)
            if l>0:
                da=W[l].T@delta
                delta=da*df(Z[l-1],A[l])
        t+=1
        for l in range(M):
            for (P,gp,v,sm) in ((W[l],gW[l],vW[l],sW[l]),):
                pass
        # 更新（W と b をまとめて）
        for l in range(M):
            if opt=="sgd":
                W[l]-=elr*gW[l]; b[l]-=elr*gb[l]
            elif opt=="mom":
                vW[l]=mom*vW[l]-elr*gW[l]; W[l]+=vW[l]
                vb[l]=mom*vb[l]-elr*gb[l]; b[l]+=vb[l]
            else: # adam
                vW[l]=b1*vW[l]+(1-b1)*gW[l]; sW[l]=b2*sW[l]+(1-b2)*gW[l]**2
                mh=vW[l]/(1-b1**t); sh=sW[l]/(1-b2**t); W[l]-=elr*mh/(np.sqrt(sh)+eps)
                vb[l]=b1*vb[l]+(1-b1)*gb[l]; sb[l]=b2*sb[l]+(1-b2)*gb[l]**2
                mhb=vb[l]/(1-b1**t); shb=sb[l]/(1-b2**t); b[l]-=elr*mhb/(np.sqrt(shb)+eps)
        record(ep)
    return dict(W=W,b=b,act=act,X=X,y=y,loss=hist_loss,acc=hist_acc,snaps=snaps,sizes=sizes)

# ====================================================================
# 図の生成（matplotlib）
# ====================================================================
CMAP = LinearSegmentedColormap.from_list("bwr2",
        [(47/255,127/255,217/255),(0.97,0.98,0.99),(255/255,107/255,107/255)])

def _grid_pred(W,b,act,res=240):
    xs=np.linspace(-1,1,res)
    XX,YY=np.meshgrid(xs,xs)
    P=np.stack([XX.ravel(),YY.ravel()],0)
    A,_=forward_all(W,b,act,P)
    return A[-1].reshape(res,res)

def fig_boundary(path, W,b,act,X,y,title=None):
    fig,ax=plt.subplots(figsize=(2.5,2.5),dpi=200)
    Pg=_grid_pred(W,b,act)
    ax.imshow(Pg,extent=[-1,1,-1,1],origin="lower",cmap=CMAP,vmin=0,vmax=1,alpha=0.9,aspect="auto")
    m0=y<0.5; m1=~m0
    ax.scatter(X[0,m0],X[1,m0],s=12,c=H_ACCENT,edgecolors="white",linewidths=0.5)
    ax.scatter(X[0,m1],X[1,m1],s=12,c=H_RED,edgecolors="white",linewidths=0.5)
    ax.set_xlim(-1,1); ax.set_ylim(-1,1); ax.set_xticks([]); ax.set_yticks([])
    if title: ax.set_title(title,fontsize=11,color=H_INK,pad=4)
    for sp in ax.spines.values(): sp.set_color(HX((0xd6,0xdc,0xe8)))
    fig.tight_layout(pad=0.2); fig.savefig(path,transparent=True); plt.close(fig)

def fig_activations(path):
    z=np.linspace(-3,3,300)
    fig,ax=plt.subplots(figsize=(4.2,2.4),dpi=200)
    ax.axhline(0,color="#c9d2e0",lw=0.8); ax.axvline(0,color="#c9d2e0",lw=0.8)
    ax.plot(z,np.tanh(z),color=H_ACCENT,lw=2.2,label="tanh")
    ax.plot(z,np.maximum(0,z),color=H_RED,lw=2.2,label="ReLU")
    ax.plot(z,sigmoid(z),color=H_PUR,lw=2.2,label="sigmoid")
    ax.set_xlim(-3,3); ax.set_ylim(-1.2,2.2)
    ax.legend(fontsize=10,loc="upper left",framealpha=0.9)
    ax.set_xlabel("z",fontsize=10); ax.set_ylabel("f(z)",fontsize=10)
    ax.tick_params(labelsize=8)
    for sp in ax.spines.values(): sp.set_color("#d6dce8")
    fig.tight_layout(pad=0.3); fig.savefig(path,transparent=True); plt.close(fig)

def fig_network(path, W,b,act, X):
    """学習済みネットの構造図（線=重み・色=符号・太さ=大きさ）。"""
    sizes=[W[0].shape[1]]+[w.shape[0] for w in W]
    C=len(sizes)
    fig,ax=plt.subplots(figsize=(4.6,3.1),dpi=200)
    xs=np.linspace(0.06,0.94,C)
    def ys(n):
        if n==1: return np.array([0.5])
        r=min(0.42,0.10+n*0.05); return np.linspace(0.5+r,0.5-r,n)
    pos=[np.stack([np.full(sizes[c],xs[c]),ys(sizes[c])],1) for c in range(C)]
    wmax=max(np.abs(w).max() for w in W)+1e-9
    for l in range(len(W)):
        for i in range(W[l].shape[0]):
            for j in range(W[l].shape[1]):
                w=W[l][i,j]
                col=H_ACCENT if w>=0 else H_RED
                ax.plot([pos[l][j,0],pos[l+1][i,0]],[pos[l][j,1],pos[l+1][i,1]],
                        color=col,lw=0.3+2.6*abs(w)/wmax,alpha=0.55,zorder=1)
    for c in range(C):
        for k in range(sizes[c]):
            ax.scatter(pos[c][k,0],pos[c][k,1],s=190,c="#eef1f7",edgecolors="#b9c4d8",
                       linewidths=1.2,zorder=2)
    for k,lab in ((0,"x"),(1,"y")):
        ax.text(pos[0][k,0],pos[0][k,1],lab,ha="center",va="center",fontsize=10,color=H_INK,zorder=3)
    ax.text(pos[-1][0,0],pos[-1][0,1],"p",ha="center",va="center",fontsize=10,color=H_INK,zorder=3)
    labs=["input"]+["hidden%d"%c for c in range(1,C-1)]+["output"]
    for c in range(C):
        ax.text(xs[c],0.99,labs[c],ha="center",va="top",fontsize=8,color=H_MUT)
    ax.set_xlim(0,1); ax.set_ylim(0,1.03); ax.axis("off")
    fig.tight_layout(pad=0.1); fig.savefig(path,transparent=True); plt.close(fig)

def fig_losscurves(path, runs, ylabel="loss", key="loss", title=None):
    fig,ax=plt.subplots(figsize=(4.4,2.6),dpi=200)
    for (lab,col,r) in runs:
        ax.plot(r[key],color=col,lw=2.0,label=lab)
    ax.set_xlabel("epoch",fontsize=10); ax.set_ylabel(ylabel,fontsize=10)
    ax.legend(fontsize=9,framealpha=0.9); ax.tick_params(labelsize=8)
    ax.grid(True,color="#eef1f7");
    for sp in ax.spines.values(): sp.set_color("#d6dce8")
    if title: ax.set_title(title,fontsize=10,color=H_INK)
    fig.tight_layout(pad=0.3); fig.savefig(path,transparent=True); plt.close(fig)

def fig_scatter(path, name, sep=None):
    X,y=gen(name,1)
    fig,ax=plt.subplots(figsize=(2.4,2.4),dpi=200)
    m0=y<0.5; m1=~m0
    ax.scatter(X[0,m0],X[1,m0],s=14,c=H_ACCENT,edgecolors="white",linewidths=0.5)
    ax.scatter(X[0,m1],X[1,m1],s=14,c=H_RED,edgecolors="white",linewidths=0.5)
    if sep is not None:
        xs=np.linspace(-1,1,10); ax.plot(xs,sep(xs),color=H_INK,lw=1.6,ls="--")
    ax.set_xlim(-1,1); ax.set_ylim(-1,1); ax.set_xticks([]); ax.set_yticks([])
    for sp in ax.spines.values(): sp.set_color("#d6dce8")
    fig.tight_layout(pad=0.15); fig.savefig(path,transparent=True); plt.close(fig)

# 生成物パス
P = lambda n: os.path.join(ASSETS, n)

print("学習と作図中 ...")
# 学習進行のスナップショット（spiral / 2-8-8-1 / Adam）
run_prog = train("spiral",[2,8,8,1],"tanh","adam",0.3,800,seed=3,snap_epochs=[0,60,800])
for ep,tag in ((0,"init"),(60,"epoch 60"),(800,"final")):
    Ws,bs=run_prog["snaps"][ep]
    fig_boundary(P("prog_%d.png"%ep),Ws,bs,run_prog["act"],run_prog["X"],run_prog["y"],title=tag)
fig_losscurves(P("prog_loss.png"),[("loss",H_YEL,run_prog)],ylabel="loss")

# 活性化関数
fig_activations(P("acts.png"))
# ネットワーク図（circle / 2-6-6-1）
run_net = train("circle",[2,6,6,1],"tanh","adam",0.3,400,seed=2)
fig_network(P("net.png"),run_net["W"],run_net["b"],run_net["act"],run_net["X"])
# 活性化と境界（sine: tanh vs relu）
run_sin_t=train("sine",[2,8,1],"tanh","adam",0.3,600,seed=1)
run_sin_r=train("sine",[2,8,1],"relu","adam",0.3,600,seed=1)
fig_boundary(P("sine_tanh.png"),run_sin_t["W"],run_sin_t["b"],run_sin_t["act"],run_sin_t["X"],run_sin_t["y"],title="tanh")
fig_boundary(P("sine_relu.png"),run_sin_r["W"],run_sin_r["b"],run_sin_r["act"],run_sin_r["X"],run_sin_r["y"],title="ReLU")
# 最適化比較（spiral / 2-8-8-1）
r_sgd=train("spiral",[2,8,8,1],"tanh","sgd",0.8,800,seed=3)
r_mom=train("spiral",[2,8,8,1],"tanh","mom",0.8,800,seed=3)
r_adm=train("spiral",[2,8,8,1],"tanh","adam",0.3,800,seed=3)
fig_losscurves(P("opt_loss.png"),[("SGD",H_MUT,r_sgd),("Momentum",H_ACCENT,r_mom),("Adam",H_RED,r_adm)],ylabel="loss")
fig_losscurves(P("opt_acc.png"),[("SGD",H_MUT,r_sgd),("Momentum",H_ACCENT,r_mom),("Adam",H_RED,r_adm)],ylabel="accuracy",key="acc")
# 学習率（circle / momentum: 小・適・大）
r_lo=train("circle",[2,6,1],"tanh","mom",0.05,300,seed=2)
r_ok=train("circle",[2,6,1],"tanh","mom",0.6,300,seed=2)
r_hi=train("circle",[2,6,1],"tanh","mom",2.6,300,seed=2)
fig_losscurves(P("lr_loss.png"),[("small 0.05",H_ACCENT,r_lo),("good 0.6",H_GRN,r_ok),("too large",H_RED,r_hi)],ylabel="loss")
# 深さの効果（spiral: 1層 vs 3層）
r_d1=train("spiral",[2,8,1],"tanh","adam",0.3,800,seed=3)
r_d3=train("spiral",[2,8,8,8,1],"tanh","adam",0.3,800,seed=3)
fig_boundary(P("depth1.png"),r_d1["W"],r_d1["b"],r_d1["act"],r_d1["X"],r_d1["y"],title="1 hidden layer")
fig_boundary(P("depth3.png"),r_d3["W"],r_d3["b"],r_d3["act"],r_d3["X"],r_d3["y"],title="3 hidden layers")
# 問題設定（blobs 直線 / circle / spiral）
fig_scatter(P("d_blobs.png"),"blobs",sep=lambda x:-x)
fig_scatter(P("d_circle.png"),"circle")
fig_scatter(P("d_spiral.png"),"spiral")
print("作図完了。スライド生成 ...")

# ====================================================================
# スライド構築
# ====================================================================
NPAGES = 15
def newslide():
    s = prs.slides.add_slide(BLANK); slide_bg(s); return s

# ---- 1. 表紙 ----------------------------------------------------------------
s=newslide()
band=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0),Inches(0),Inches(13.333),Inches(7.5))
band.fill.solid(); band.fill.fore_color.rgb=PANEL; band.line.fill.background(); band.shadow.inherit=False
strip=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0),Inches(2.95),Inches(13.333),Inches(0.06))
strip.fill.solid(); strip.fill.fore_color.rgb=ACCENT; strip.line.fill.background(); strip.shadow.inherit=False
add_text(s,0.9,1.5,11.5,1.0,"🧠 ニューラルネットワーク",size=46,bold=True,color=TEXT)
add_text(s,0.9,2.35,11.5,0.6,"多層パーセプトロンの仕組みを数式とデモで理解する",size=22,color=ACCENT)
add_text(s,0.9,3.25,11.5,1.6,
  "順伝播 → 損失 → **逆伝播** → 最適化 という「学習」の一連の流れを、\n"
  "**数式**と体験アプリ **html/nw2.html** の観察でつなぐ講義スライド。\n"
  "対になる教材：詳細ノート **nw2.md** ／ 勾配降下法 **html/gradient.html**",
  size=16,color=MUTED,line_spacing=1.3)
add_text(s,0.9,6.5,11,0.4,"アルゴリズム・ビジュアル教材集",size=13,color=MUTED,font=MONO)
add_eq(s,r"p=\sigma\!\left(W^{(M)}f(\cdots f(W^{(1)}\mathbf{x}+\mathbf{b}^{(1)})\cdots)+\mathbf{b}^{(M)}\right)",
       10.0,5.9,0.42,color=H_MUT,fs=20)
pagenum(s,1)

# ---- 2. 問題設定 ------------------------------------------------------------
s=newslide(); header(s,"1. 問題設定 ― 何を学習するのか","2次元の点を2クラスに分ける関数をデータから作る","01")
add_text(s,0.55,1.15,7.0,2.2,
 "平面上の点 **x=(x₁,x₂)** を受け取り、\n「赤である確率 **p**」を返す関数\n"
 "  **p = F(x ; θ)**\nを作る。**θ** は全ての重みとバイアス。\n"
 "「学習」とは、データに合うよう **θ** を調整すること。",
 size=16,color=TEXT,line_spacing=1.3)
card(s,0.55,3.7,7.0,3.1,title="なぜ直線では足りないか",accent=ACCENT)
add_text(s,0.8,4.2,6.5,2.4,
 "・**2つの塊** は1本の直線で分けられる（左図の破線）。\n"
 "・**まる**・**渦巻き** は直線では絶対に不可能＝**非線形**。\n"
 "・この「曲がった境界」を作る力こそ\n  ニューラルネットワークの本質。",
 size=15,color=TEXT,line_spacing=1.35)
for i,(img,cap) in enumerate([("d_blobs.png","2つの塊：直線でOK"),
                              ("d_circle.png","まる：曲線が要る"),
                              ("d_spiral.png","渦巻き：さらに複雑")]):
    x=8.05; y=1.25+i*1.98
    add_image_fit(s,P(img),x,y,1.85,1.6)
    add_text(s,x+2.0,y+0.5,3.0,0.7,cap,size=13,color=MUTED,anchor=MSO_ANCHOR.MIDDLE)
pagenum(s,2)

# ---- 3. 順伝播①ニューロン --------------------------------------------------
s=newslide(); header(s,"2. 順伝播 ① ニューロン1個","加重和にバイアスを足し、活性化関数に通す","02")
card(s,0.55,1.2,7.0,2.2,accent=ACCENT)
add_eq(s,r"z=\sum_j w_j\,x_j+b,\qquad a=f(z)",4.05,2.05,0.62,color=H_ACCENT,fs=26)
add_text(s,0.8,2.55,6.5,0.8,
 "**w**：各入力の重視度（重み）／ **b**：発火しやすさの下駄／ **f**：非線形の「くせ」",
 size=13.5,color=MUTED)
card(s,0.55,3.7,7.0,3.1,title="活性化関数がなぜ要る？",accent=PURPLE)
add_text(s,0.8,4.2,6.5,2.4,
 "**f** が無い（=恒等）と、何層重ねても\n合成はただの1次式＝**直線の分類器**にしかならない。\n\n"
 "**f** の非線形性が、層を重ねるほど\n複雑な境界を生む源泉になる。",
 size=15,color=TEXT,line_spacing=1.35)
card(s,7.8,1.2,5.0,5.6,title="活性化関数 3種")
add_image_fit(s,P("acts.png"),8.0,1.75,4.6,2.6)
add_text(s,8.05,4.45,4.6,2.2,
 "**tanh**：出力(−1,1)・なめらか\n**ReLU**：軽い・折れ線・深い網に強い\n**sigmoid**：出力(0,1)・勾配消失しやすい",
 size=13.5,color=TEXT,line_spacing=1.4)
pagenum(s,3)

# ---- 4. 順伝播②ネットワーク全体 -------------------------------------------
s=newslide(); header(s,"3. 順伝播 ② ネットワーク全体","層を行列で書き、入力から出力へ伝える","03")
card(s,0.55,1.2,7.1,2.35,title="層の計算（行列表記）",accent=ACCENT)
add_eq(s,r"\mathbf{z}^{(l)}=W^{(l)}\mathbf{a}^{(l-1)}+\mathbf{b}^{(l)},\quad \mathbf{a}^{(l)}=f\!\left(\mathbf{z}^{(l)}\right)",
       4.1,2.35,0.5,color=H_INK,fs=22)
add_text(s,0.8,2.9,6.6,0.6,"入力 **a⁽⁰⁾=x** から順に適用。**出力層だけ σ** で 0〜1 の確率に。",size=13.5,color=MUTED)
card(s,0.55,3.75,7.1,3.05,title="出力＝確率")
add_eq(s,r"p=\sigma\!\left(z^{(M)}\right)=\frac{1}{1+e^{-z^{(M)}}}\in(0,1)",4.1,4.5,0.5,color=H_ACCENT,fs=22)
add_text(s,0.8,5.15,6.6,1.5,
 "**M** は重み行列の数（＝隠れ層の数 **+1**）。\n"
 "図の各線が重み **W** の1成分、丸がニューロン。\n"
 "デモでは右キャンバスにマウスを乗せると、\nこの伝播が光って見える。",
 size=14,color=TEXT,line_spacing=1.35)
card(s,7.9,1.2,4.9,5.6,title="ネットワーク図（2→6→6→1）")
add_image_fit(s,P("net.png"),8.05,1.7,4.6,4.4)
add_text(s,8.05,6.05,4.6,0.6,"線の色＝重みの符号（青＋/赤−）、太さ＝大きさ",size=11.5,color=MUTED)
pagenum(s,4)

# ---- 5. 損失関数 ------------------------------------------------------------
s=newslide(); header(s,"4. 損失関数 ― どれだけ間違えているか","二値クロスエントロピー(BCE)","04")
card(s,0.55,1.2,12.2,1.7,accent=ACCENT)
add_eq(s,r"\mathcal{L}=\frac{1}{N}\sum_{i=1}^{N}-\left[\,y_i\ln p_i+(1-y_i)\ln(1-p_i)\,\right]",
       6.65,2.05,0.66,color=H_INK,fs=24)
card(s,0.55,3.1,6.0,3.7,title="直感")
add_eq(s,r"y=1:\ \ \mathcal{L}=-\ln p",3.55,3.95,0.42,color=H_ACCENT,fs=20)
add_eq(s,r"y=0:\ \ \mathcal{L}=-\ln(1-p)",3.55,4.65,0.42,color=H_ACCENT,fs=20)
add_text(s,0.85,5.15,5.4,1.5,
 "正解が1なら **p を1へ**、0なら **p を0へ**。\n外すほど **−ln** が急増して強く罰する。",
 size=14.5,color=TEXT,line_spacing=1.35)
card(s,6.75,3.1,6.0,3.7,title="なぜ二乗誤差でなく BCE か",accent=PURPLE)
add_text(s,7.05,3.65,5.5,3.0,
 "出力がシグモイドのとき、BCE だと勾配が\nきれいな **p − y** になり（次頁）、\n"
 "外したときに勾配が大きく **速く安定**して学習。\n\n"
 "二乗誤差はシグモイドの平坦部で勾配がつぶれ、\n自信満々で外しても直りが遅い。",
 size=14.5,color=TEXT,line_spacing=1.35)
pagenum(s,5)

# ---- 6. 勾配降下 ------------------------------------------------------------
s=newslide(); header(s,"5. 学習＝最適化 ― 勾配降下法","損失の坂を、勾配の逆向きに少しずつ下る","05")
card(s,0.55,1.2,7.0,1.7,accent=ACCENT)
add_eq(s,r"\theta \leftarrow \theta-\eta\,\nabla_\theta \mathcal{L}",4.05,2.05,0.62,color=H_ACCENT,fs=26)
add_text(s,0.55,3.15,7.0,2.0,
 "**∇θL**：各パラメータを増やすと損失がどれだけ増えるか。\n"
 "その**逆向き**に進めば損失が減る。**η**（学習率）は1歩の大きさ。\n\n"
 "アプリは毎エポック、**全データ(フルバッチ)の平均勾配**で1歩進める。",
 size=15,color=TEXT,line_spacing=1.35)
add_eq(s,r"\nabla_\theta \mathcal{L}=\frac{1}{N}\sum_{i=1}^{N}\nabla_\theta \mathcal{L}_i",
       4.05,5.6,0.5,color=H_INK,fs=20)
card(s,7.8,1.2,5.0,5.6,title="学習率 η の効き方")
add_image_fit(s,P("lr_loss.png"),7.95,1.75,4.7,2.9)
add_text(s,7.95,4.75,4.7,1.9,
 "**小さすぎ**：なかなか進まない。\n**ちょうど良い**：すっと収束。\n**大きすぎ**：行き過ぎて損失が**暴れる/発散**。",
 size=14,color=TEXT,line_spacing=1.4)
pagenum(s,6)

# ---- 7. 逆伝播①出力誤差 ----------------------------------------------------
s=newslide(); header(s,"6. 逆伝播 ① 出力層の誤差","連鎖律で δ を求めると p−y に約分される","06")
card(s,0.55,1.2,12.2,2.5,title="導出（BCE＋シグモイド）",accent=ACCENT)
add_eq(s,r"\frac{\partial \mathcal{L}}{\partial p}=\frac{p-y}{p(1-p)},\qquad \frac{\partial p}{\partial z^{(M)}}=\sigma'(z^{(M)})=p(1-p)",
       6.65,2.15,0.55,color=H_INK,fs=21)
add_eq(s,r"\Rightarrow\ \ \delta^{(M)}=\frac{\partial \mathcal{L}}{\partial z^{(M)}}=\frac{p-y}{p(1-p)}\cdot p(1-p)=\,p-y",
       6.65,3.05,0.55,color=H_ACCENT,fs=22)
card(s,0.55,4.0,12.2,2.8,title="意味")
add_text(s,0.85,4.6,11.6,2.0,
 "**p(1−p) が約分**され、出力層の誤差は **「予測 − 正解」** という最短の形に。\n"
 "・**δ⁽ᴹ⁾ > 0**（p が大きすぎ）→ p を下げる向きに重みを直す。\n"
 "・**δ⁽ᴹ⁾ < 0**（p が小さすぎ）→ p を上げる向きに直す。\n"
 "これが「BCE＋シグモイドは二値分類の定番」である理由。実装の `delta = out − label` はまさにこれ。",
 size=15,color=TEXT,line_spacing=1.4)
pagenum(s,7)

# ---- 8. 逆伝播②隠れ層 ------------------------------------------------------
s=newslide(); header(s,"7. 逆伝播 ② 隠れ層への伝播と勾配","誤差を後ろから前へ配り直す","07")
card(s,0.55,1.2,12.2,1.9,title="1つ後ろの層から手前の層へ（連鎖律）",accent=ACCENT)
add_eq(s,r"\boldsymbol{\delta}^{(l)}=\left(W^{(l+1)\top}\boldsymbol{\delta}^{(l+1)}\right)\odot f'\!\left(\mathbf{z}^{(l)}\right)",
       6.65,2.25,0.6,color=H_INK,fs=23)
card(s,0.55,3.35,6.0,3.45,title="重み・バイアスの勾配")
add_eq(s,r"\frac{\partial \mathcal{L}}{\partial W^{(l)}}=\boldsymbol{\delta}^{(l)}\,\left(\mathbf{a}^{(l-1)}\right)^{\top}",
       3.55,4.25,0.62,color=H_ACCENT,fs=20)
add_eq(s,r"\frac{\partial \mathcal{L}}{\partial \mathbf{b}^{(l)}}=\boldsymbol{\delta}^{(l)}",
       3.55,5.35,0.62,color=H_ACCENT,fs=20)
card(s,6.75,3.35,6.0,3.45,title="手順まとめ",accent=GREEN)
add_text(s,7.05,3.95,5.5,2.7,
 "① **順伝播**で z,a を保存\n"
 "② 出力誤差 **δ⁽ᴹ⁾ = p − y**\n"
 "③ **逆伝播**で δ⁽ˡ⁾ を順に計算\n"
 "④ 各層の勾配を求める\n"
 "⑤ 勾配降下で θ を1歩更新",
 size=15,color=TEXT,line_spacing=1.5)
pagenum(s,8)

# ---- 9. 勾配消失 ------------------------------------------------------------
s=newslide(); header(s,"8. 勾配消失 ― 深い網が学習しにくい理由","f′ が何度も掛かって入力側の勾配が消える","08")
card(s,0.55,1.2,7.0,2.2,accent=RED)
add_eq(s,r"\boldsymbol{\delta}^{(1)}\ \sim\ \left(f'\cdot f'\cdots f'\right)\times(\cdots)",
       4.05,2.05,0.55,color=H_RED,fs=22)
add_text(s,0.8,2.62,6.5,0.3,"（f′ が層の数だけ掛かる）",size=11.5,color=MUTED,align=PP_ALIGN.CENTER)
add_text(s,0.8,2.98,6.5,0.5,"tanh は **f′≤1**、sigmoid は **f′≤0.25**。掛け続けると 0 へ。",size=13.5,color=MUTED)
card(s,0.55,3.7,7.0,3.1,title="対策",accent=GREEN)
add_text(s,0.8,4.25,6.5,2.3,
 "・**ReLU**：正の領域で f′=1（消えにくい）\n"
 "・**良い初期化**（Xavier / He、次頁）\n"
 "・**Adam** など適応的な最適化\n"
 "・（発展）残差接続・正規化 など",
 size=15,color=TEXT,line_spacing=1.45)
card(s,7.8,1.2,5.0,5.6,title="デモで観察")
add_text(s,8.05,1.7,4.6,5.0,
 "**渦巻き** で **隠れ層=4**・**sigmoid** に\nすると学習が停滞しやすい。\n\n"
 "**tanh / ReLU / Adam** に変えると\n改善することがある。\n\n"
 "**R** キーで初期値を変えて\n再挑戦できる（初期化の効果）。",
 size=15,color=TEXT,line_spacing=1.5)
pagenum(s,9)

# ---- 10. 最適化手法 ---------------------------------------------------------
s=newslide(); header(s,"9. 最適化手法 ― SGD / モメンタム / Adam","勾配を得たあと、どう1歩進めるか","09")
cw=4.02
for i,(ti,acc) in enumerate([("SGD",MUTED),("モメンタム",ACCENT),("Adam",RED)]):
    x=0.55+i*(cw+0.13)
    card(s,x,1.2,cw,2.5,title=ti,accent=acc)
# 各カード内に式
add_eq(s,r"\theta\leftarrow\theta-\eta\,g",0.55+cw/2,2.55,0.42,color=H_INK,fs=20)
add_eq(s,r"\mathbf{v}\leftarrow\mu\mathbf{v}-\eta g,\ \ \theta\leftarrow\theta+\mathbf{v}",0.55+(cw+0.13)+cw/2,2.55,0.34,color=H_INK,fs=15)
add_eq(s,r"\theta\leftarrow\theta-\eta\,\dfrac{\hat m}{\sqrt{\hat s}+\varepsilon}",0.55+2*(cw+0.13)+cw/2,2.5,0.5,color=H_INK,fs=16)
add_text(s,0.7,3.15,cw-0.3,0.5,"素直・遅い。谷でジグザグ。",size=12.5,color=MUTED)
add_text(s,0.83+cw+0.13,3.15,cw-0.3,0.5,"慣性で加速(μ=0.9)。振動が減る。",size=12.5,color=MUTED)
add_text(s,0.7+2*(cw+0.13),2.98,cw-0.3,0.7,"重みごとに歩幅を自動調整。\n速く安定・難データに強い。",size=12.5,color=MUTED,line_spacing=1.2)
card(s,0.55,3.95,6.0,2.85,title="Adam の中身",accent=RED)
add_eq(s,r"m\leftarrow\beta_1 m+(1-\beta_1)g,\ \ s\leftarrow\beta_2 s+(1-\beta_2)g^2",3.55,4.6,0.34,color=H_INK,fs=15)
add_eq(s,r"\hat m=\dfrac{m}{1-\beta_1^{\,t}},\quad \hat s=\dfrac{s}{1-\beta_2^{\,t}}",3.55,5.35,0.45,color=H_INK,fs=15)
add_text(s,0.8,5.95,5.5,0.8,"√ŝ で割る＝勾配のスケールに鈍感。**アプリは Adam の実効 η を ×0.1**。",size=12.5,color=MUTED)
card(s,6.75,3.95,6.0,2.85,title="収束比較（渦巻き・2-8-8-1）")
add_image_fit(s,P("opt_loss.png"),6.95,4.45,5.6,2.25)
pagenum(s,10)

# ---- 11. 初期化 -------------------------------------------------------------
s=newslide(); header(s,"10. 重みの初期化","大きさ(分散)の設計が学習の成否を分ける","10")
card(s,0.55,1.3,12.2,1.9,accent=ACCENT)
add_eq(s,r"\text{tanh/sigmoid (Xavier)}:\ \ \mathrm{std}=\frac{1}{\sqrt{n_{\mathrm{in}}}}\qquad\qquad \text{ReLU (He)}:\ \ \mathrm{std}=\sqrt{\frac{2}{n_{\mathrm{in}}}}",
       6.65,2.25,0.6,color=H_INK,fs=20)
card(s,0.55,3.5,12.2,3.3,title="なぜスケールが大事か")
add_text(s,0.9,4.1,11.4,2.5,
 "・**大きすぎ**：活性が飽和（tanh が ±1 に張り付き）→ f′≈0 で勾配消失。\n"
 "・**小さすぎ**：信号がどんどん縮む → やはり勾配が消える。\n"
 "・入力数 **n_in**（ファンイン）に応じて分散を決め、層をまたいで信号の大きさを保つ。\n"
 "・ReLU は負side を捨てて分散が半減するため **√2 倍** で補う（He 初期化）。\n"
 "  アプリは活性化に応じてこの初期化を自動選択。**R** キーで初期値だけ変えて再学習できる。",
 size=15.5,color=TEXT,line_spacing=1.5)
pagenum(s,11)

# ---- 12. デモ：学習の進行 --------------------------------------------------
s=newslide(); header(s,"11. デモ：境界が「学習」で立ち上がる","渦巻き・2隠れ層・Adam のスナップショット","11")
for i,(ep,cap) in enumerate([(0,"init.png"),(60,"epoch 60"),(800,"final")]):
    pass
imgs=[("prog_0.png","初期（ランダム）"),("prog_60.png","学習中（60ep）"),("prog_800.png","学習後（最終）")]
for i,(img,cap) in enumerate(imgs):
    x=0.6+i*3.05
    card(s,x,1.25,2.9,3.5)
    add_image_fit(s,P(img),x+0.15,1.5,2.6,2.6)
    add_text(s,x,4.05,2.9,0.5,cap,size=13,color=TEXT,align=PP_ALIGN.CENTER,bold=True)
card(s,10.0,1.25,2.85,3.5,title="損失の推移")
add_image_fit(s,P("prog_loss.png"),10.1,1.75,2.65,2.7)
card(s,0.6,5.05,12.25,1.75,title="操作",accent=ACCENT)
add_text(s,0.9,5.6,11.7,1.1,
 "**▶ 再生**でエポックを進め、**◀/▶**や**シークバー**で前後に移動。**← →** で1コマ、**Space** 再生、**R** 初期値replay。\n"
 "でたらめな境界が、逆伝播による重み更新の繰り返しで、データの模様に沿って立ち上がる様子を観察できる。",
 size=14.5,color=TEXT,line_spacing=1.4)
pagenum(s,12)

# ---- 13. デモ：活性化と深さ -------------------------------------------------
s=newslide(); header(s,"12. デモ：活性化の個性と『深さ』の効果","同じデータでも設定で境界が変わる","12")
card(s,0.55,1.2,6.0,5.6,title="活性化：サイン波の境界")
add_image_fit(s,P("sine_tanh.png"),0.9,1.75,2.5,2.5)
add_image_fit(s,P("sine_relu.png"),3.45,1.75,2.5,2.5)
add_text(s,0.85,4.4,5.4,2.2,
 "**tanh** はなめらかな曲線、**ReLU** は\n区分線形ゆえ境界が**折れ線**になる。\n"
 "デモの活性化を切り替えて見比べよう。",
 size=15,color=TEXT,line_spacing=1.4)
card(s,6.75,1.2,6.0,5.6,title="深さ：渦巻き（1層 vs 3層）")
add_image_fit(s,P("depth1.png"),7.1,1.75,2.5,2.5)
add_image_fit(s,P("depth3.png"),9.65,1.75,2.5,2.5)
add_text(s,7.05,4.4,5.4,2.2,
 "**1層**では渦を追い切れないが、**層を重ねる**と\n前の層の特徴を組み合わせ、複雑な境界を描ける。\n"
 "ただし深いほど学習は難しくなる（勾配消失）。",
 size=15,color=TEXT,line_spacing=1.4)
pagenum(s,13)

# ---- 14. 比較表 -------------------------------------------------------------
s=newslide(); header(s,"13. まとめ表：データ難易度と収束（実測）","本教材で実際に学習させた結果（tanh, 2-8-8-1）","13")
rows=[
 ("データ","直線で可?","1層で解ける?","深さ/Adamの効果"),
 ("2つの塊","○","○","不要（線形分離）"),
 ("まる / サイン波","×","○（曲線1本）","少ない幅でOK"),
 ("市松(XOR)","×","△","隠れ層が必須"),
 ("チェッカー / 三重丸","×","△","幅・深さで表現力↑"),
 ("渦巻き","×","×","深さ＋Adamで解ける"),
]
tx,ty,tw=0.7,1.3,11.9; rh=0.82; cw=[2.9,2.0,2.6,4.4]
for r,row in enumerate(rows):
    yy=ty+r*rh
    head = (r==0)
    xx=tx
    for c,val in enumerate(row):
        fill = ACCENT if head else (PANEL if r%2 else PANEL2)
        add_rrect(s,xx,yy,cw[c],rh,fill=fill,line=LINE,radius=0.04)
        add_text(s,xx+0.14,yy,cw[c]-0.28,rh,val,size=14.5 if head else 14,
                 bold=head, color=(RGBColor(0xff,0xff,0xff) if head else TEXT),
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT if c==0 or c==3 else PP_ALIGN.CENTER)
        xx+=cw[c]
add_text(s,0.7,6.35,11.9,0.6,
 "○=容易 / △=不安定・工夫要 / ×=不可。**渦巻き**は SGD・モメンタムでは頭打ち、**Adam** でほぼ解ける。",
 size=13.5,color=MUTED)
pagenum(s,14)

# ---- 15. まとめ＆演習 -------------------------------------------------------
s=newslide(); header(s,"14. まとめ & 演習","一枚チートシートと、手を動かす課題","14")
card(s,0.55,1.2,7.2,5.6,title="ワンページ・チートシート",accent=ACCENT)
add_eq(s,r"\mathbf{z}^{(l)}=W^{(l)}\mathbf{a}^{(l-1)}+\mathbf{b}^{(l)},\ \ \mathbf{a}^{(l)}=f(\mathbf{z}^{(l)}),\ \ p=\sigma(z^{(M)})",
       4.15,2.05,0.34,color=H_INK,fs=14)
add_eq(s,r"\mathcal{L}=\frac{1}{N}\sum_i-\left[y_i\ln p_i+(1-y_i)\ln(1-p_i)\right]",
       4.15,2.75,0.34,color=H_INK,fs=14)
add_eq(s,r"\delta^{(M)}=p-y,\ \ \boldsymbol{\delta}^{(l)}=\left(W^{(l+1)\top}\boldsymbol{\delta}^{(l+1)}\right)\odot f'(\mathbf{z}^{(l)})",
       4.15,3.45,0.34,color=H_ACCENT,fs=14)
add_eq(s,r"\frac{\partial\mathcal{L}}{\partial W^{(l)}}=\boldsymbol{\delta}^{(l)}\mathbf{a}^{(l-1)\top},\quad \frac{\partial\mathcal{L}}{\partial\mathbf{b}^{(l)}}=\boldsymbol{\delta}^{(l)}",
       4.15,4.2,0.4,color=H_INK,fs=14)
add_eq(s,r"\theta\leftarrow\theta-\eta\,g\ \ /\ \ \text{Momentum}\ \ /\ \ \text{Adam}",
       4.15,5.0,0.34,color=H_INK,fs=14)
add_text(s,0.8,5.5,6.7,1.2,
 "順伝播→損失→**逆伝播**→更新。この1周を\n何百回も回すのが「学習」。",
 size=14,color=MUTED,line_spacing=1.3)
card(s,7.9,1.2,4.9,5.6,title="演習（nw2.md に略解）",accent=GREEN)
add_text(s,8.15,1.75,4.5,5.0,
 "**A. 数式**\n"
 "・最小ネットで順伝播/逆伝播/1更新を手計算\n"
 "・σ′=σ(1−σ) の導出\n"
 "・XOR が直線分離不可の証明\n\n"
 "**B. デモ実験**\n"
 "・まるを解く最小ニューロン数\n"
 "・深さ vs 幅（渦巻き）\n"
 "・最適化レース（SGD/Mom/Adam）\n"
 "・学習率の崖／勾配消失の観察\n\n"
 "**C. 発展**\n"
 "・BCE と二乗誤差の勾配比較\n"
 "・Adam のスケール不変性",
 size=13,color=TEXT,line_spacing=1.3)
pagenum(s,15)

out=os.path.join(os.path.dirname(os.path.abspath(__file__)),"nw2.pptx")
prs.save(out)
print("保存:", out)
print("スライド枚数:", len(prs.slides._sldIdLst))
