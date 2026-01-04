import schemdraw
import schemdraw.elements as elm
import matplotlib.pyplot as plt
from matplotlib import font_manager

from schemdraw import elements as elm
from schemdraw.segments import Segment, SegmentCircle

# -------------------------
# フォントの読み込み

# font_path_reg = 'Fonts/SourceSans3-Regular.ttf'
# font_path_it  = 'Fonts/SourceSans3-Italic.ttf'

font_path_reg = 'Fonts/Helvetica Neue LT Pro 55 Roman.otf'
font_path_it  = 'Fonts/Helvetica Neue LT Pro 56 Italic.otf'

font_size = 14

# --- 2. フォントマネージャーに追加 ---
try:
    font_manager.fontManager.addfont(font_path_reg)
    font_manager.fontManager.addfont(font_path_it)
    
    # 正確なフォント名を取得
    prop_reg = font_manager.FontProperties(fname=font_path_reg)
    prop_it  = font_manager.FontProperties(fname=font_path_it)
    font_name = prop_reg.get_name()
    
    print(f"フォント読み込み成功: {font_name}")


except Exception as e:
    print(f"フォント読み込み失敗: {e}")
    font_name = 'sans-serif' # フォールバック

# --- 3. 数式フォントをカスタム設定にする ---
# まず 'custom' モードを有効化
plt.rcParams['mathtext.fontset'] = 'custom'
plt.rcParams['mathtext.default'] = 'regular'  # 数式のデフォルトスタイルを設定
plt.rcParams['font.size'] = font_size  # 数式フォントのサイズを明示的に設定

# 数式の各役割に、読み込んだフォントを割り当てる
plt.rcParams['mathtext.rm'] = font_name          # ローマン体（数字、単位など）
plt.rcParams['mathtext.it'] = f'{font_name}:italic' # イタリック体（変数 V, R など）
plt.rcParams['mathtext.bf'] = f'{font_name}:bold'   # ボールド体（ベクトルなど）

# ※通常のテキストフォントも合わせておく
plt.rcParams['font.family'] = font_name

# -------------------------


# =========================
# Vccを作成
class Vcc(elm.Vdd):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.scale(2) # Vdd自身の大きさを2倍する
        self.segments.append(SegmentCircle((0,0), 0.1)) # 中心に半径0.1の円を追加する

# =========================
# 抵抗
resw = 1.0 / 6
resh = .18

class Res(elm.Element2Term):
    def __init__(self, **kwargs):
            super().__init__(**kwargs)
            self.segments.append(Segment(
            [(0, 0), (0.5*resw, resh), (1.5*resw, -resh),
             (2.5*resw, resh), (3.5*resw, -resh),
             (4.5*resw, resh), (5.5*resw, -resh), (6*resw, 0)]))
            
# =========================
# 可変抵抗
class Pot(Res):
    _element_defaults = {
        'arrowwidth': .15,
        'arrowlength': .25,
        'tap_length': 0.72,
        }
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        potheight = self.params['tap_length']
        self.anchors['tap'] = (resw*3, potheight)
        self.params['lblloc'] = 'bot'
        self.segments.append(Segment(
            [(resw*3, potheight), (resw*3, resw*1.5)],
            arrow='->', arrowwidth=self.params['arrowwidth'],
            arrowlength=self.params['arrowlength']))
        
# =========================


with schemdraw.Drawing() as d:
    d.config(fontsize=font_size, font=font_name, lw=1.25) 
    
    # 入力
    IN = elm.Dot(open=True).label('Input', loc='left')
    L1 = elm.Line().right(2)
    elm.Line().down(2).dot()
    Rie1 = Res().down(3).label('$r_{ie1}$',loc='bottom')
    elm.GroundSignal(lead=False)

    elm.Gap().right(3.5).at(L1.end)
    elm.GroundSignal(lead=False)
    elm.SourceI().reverse().label('$h_{fe1} \\cdot i_{b1}$',color='red')
    elm.Line().right(3).dot()
    Lc1 = elm.Line().down(1)
    Rc1 = Res().down(1).label('$R_{C1}$',loc='bottom')
    elm.Line().down(1)
    elm.GroundSignal(lead=False)
    elm.Line().right(3).at(Lc1.start).dot()
    Lie2 = elm.Line().down(1)
    Rie2 = Res().down(1).label('$R_{ie2}$',loc='bottom')
    elm.Line().down(1)
    elm.GroundSignal(lead=False)

    elm.Line().right(5).at(Rie1.start)
    Rf = Res().right(3).label('$R_F$',loc='bottom')
    elm.Line().right(5).dot()
    Re2 = Res().down(3).label('$R_{e2}$',loc='bottom')
    elm.GroundSignal(lead=False)

    elm.Line().up(3).at(Re2.start).toy(Rie2.end)
    elm.SourceI().up(3).reverse().dot().label('$h_{fe2} \\cdot i_{b2}$',color='red')
    elm.Line().right(3).dot()
    Rc2 = Res().down(3).label('$R_{C2}$',loc='bottom')
    elm.GroundSignal(lead=False)

    elm.Line().right(2).at(Rc2.start)
    OUT = elm.Dot(open=True).label('out',loc='top')

    # Current Label
    elm.CurrentLabel(length=1).at(L1).label('$i_1$',ofst=-.15).color('red')
    elm.CurrentLabel(length=1,ofst=-.6).at(Rie1).label('$i_{b1}$',loc='top',ofst=-.1).color('red')
    elm.CurrentLabel(length=1, ofst=.3).at(Lc1).label('$i_{c1}$',ofst=-.1).color('red')
    elm.CurrentLabel(length=1, ofst=.3).at(Lie2).label('$i_{ie2}$',ofst=-.1).color('red')
    elm.CurrentLabel(length=1).at(Rf).label('$i_f$',ofst=-.1).color('red').reverse()
    elm.CurrentLabel(length=1,ofst=-.6).at(Re2).label('$i_{e2}$',loc='top',ofst=-.1).color('red')
    elm.CurrentLabel(length=1,ofst=-.6).at(Rc2).label('$i_{c2}$',loc='top',ofst=-.1).color('red')

    # Voltage Label
    Gin = elm.Gap().at(L1.start).toy(Rie1.end)
    elm.GroundSignal(lead=False).color('blue')
    elm.VoltageLabelArc(length=4.8).at(Gin).label('$v_1$').reverse().color('blue')

    """
    # 電流ラベル
    elm.CurrentLabel(length=1, ofst=0.3).at(LB1).label('$I_{B1}$')
    elm.CurrentLabel(length=1.25, ofst=0.3).at(LE2).label('$I_{B2}$')
    elm.CurrentLabel(length=1.25, ofst=0.3).at(L2e).label('$I_{E2}$')

    elm.CurrentLabel(top=False, length=1, ofst=0.3).at(R1).label('$I_{C1}+I_{B2}$').reverse()

    # 電圧ラベル
    gap_EB1 = elm.Gap().at(Q1.emitter).to(Q1.base)
    elm.VoltageLabelArc(length=1, ofst=.3).at(gap_EB1).label('$V_{BE}$')
    """


    # 保存
    d.save('FuzzFace_eq.svg', dpi=300)
    d.save('FuzzFace_eq.png', dpi=300, transparent=False)
    print("FuzzFace images has been exported.")