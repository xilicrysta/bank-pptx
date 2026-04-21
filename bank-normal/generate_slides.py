import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- Constants & Design System (Normal / Simplified) ---
BG_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x1C, 0x1B, 0x19)
SUB_GREY = RGBColor(0x75, 0x75, 0x75)
LIGHT_GREY = RGBColor(0xF0, 0xF0, 0xF0)
ACCENT_LINE = RGBColor(0xD0, 0xD0, 0xD0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT_SANS = "Hiragino Sans"

class PresentationGenerator:
    def __init__(self, title, subtitle):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.title = title
        self.subtitle = subtitle

    def _add_base_design(self, slide):
        # A simple top decorative line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(12.333), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = ACCENT_LINE
        line.line.visible = False


    def add_cover_slide(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        if os.path.exists("assets/cover_bg.png"):
            slide.shapes.add_picture("assets/cover_bg.png", 0, 0, width=self.prs.slide_width, height=self.prs.slide_height)
        else:
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = BG_WHITE
            bg.line.visible = False
        
        # Center Box
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(2), Inches(10.333), Inches(3.5))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = ACCENT_LINE
        box.line.width = Pt(1.5)
        
        tf = box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = self.title
        p.font.name = FONT_SANS
        p.font.size = Pt(54)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        p.alignment = PP_ALIGN.CENTER
        
        p2 = tf.add_paragraph()
        p2.text = self.subtitle
        p2.font.name = FONT_SANS
        p2.font.size = Pt(24)
        p2.font.color.rgb = SUB_GREY
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(30)

    def add_transition_slide(self, section_num, title):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
        # Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_WHITE
        bg.line.visible = False

        # Left bar accent
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(0.15), Inches(5.9))
        bar.fill.solid()
        bar.fill.fore_color.rgb = TEXT_DARK
        bar.line.visible = False
        
        num_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10), Inches(1))
        p = num_box.text_frame.paragraphs[0]
        p.text = f"CHAPTER {section_num}"
        p.font.name = FONT_SANS
        p.font.size = Pt(28)
        p.font.color.rgb = SUB_GREY
        
        title_box = slide.shapes.add_textbox(Inches(1.5), Inches(3.2), Inches(10), Inches(2))
        p = title_box.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_SANS
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        title_box.text_frame.word_wrap = True

    def add_message_slide(self, title, message, body_items=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        # Force white bg
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = WHITE
        bg.line.visible = False

        self._add_base_design(slide)
        
        # Title
        t_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.6), Inches(12), Inches(1))
        p = t_box.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_SANS
        p.font.size = Pt(24)
        p.font.color.rgb = SUB_GREY
        
        # Message
        m_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.333), Inches(1.2))
        tf = m_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = message
        p.font.name = FONT_SANS
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        
        if body_items:
            b_box = slide.shapes.add_textbox(Inches(0.4), Inches(3.0), Inches(12.5), Inches(3.8))
            tf = b_box.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            for item in body_items:
                p = tf.add_paragraph()
                p.text = f"● {item}"
                p.font.name = FONT_SANS
                p.font.size = Pt(26)
                p.space_after = Pt(24)
                p.font.color.rgb = TEXT_DARK

    def add_diagram_slide(self, title, message, diagram_func):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = WHITE
        bg.line.visible = False

        self._add_base_design(slide)
        
        # Title
        t_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.6), Inches(12), Inches(1))
        p = t_box.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_SANS
        p.font.size = Pt(24)
        p.font.color.rgb = SUB_GREY
        
        # Message
        m_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.333), Inches(1.2))
        tf = m_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = message
        p.font.name = FONT_SANS
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        
        diagram_func(slide)

    # Simplified diagrams
    def draw_megabank_flow(self, slide):
        # Lots of banks -> Arrow -> 3 Megabanks
        y = Inches(3.5)
        
        # Left box
        b1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), y, Inches(3.5), Inches(1.5))
        b1.fill.solid()
        b1.fill.fore_color.rgb = LIGHT_GREY
        b1.line.color.rgb = ACCENT_LINE
        p = b1.text_frame.paragraphs[0]
        p.text = "昔の日本\n(たくさんの銀行があった)"
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        b1.text_frame.word_wrap = True
        
        # Arrow
        arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.5), y+Inches(0.2), Inches(2.3), Inches(1.1))
        arr.fill.solid()
        arr.fill.fore_color.rgb = TEXT_DARK
        arr.line.visible = False
        arr.text_frame.paragraphs[0].text = "お互いに合体！\n(合併)"
        arr.text_frame.paragraphs[0].font.size = Pt(16)
        arr.text_frame.paragraphs[0].font.color.rgb = WHITE
        
        # Right box
        b2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.3), y-Inches(0.5), Inches(3.5), Inches(2.5))
        b2.fill.solid()
        b2.fill.fore_color.rgb = TEXT_DARK
        b2.line.visible = False
        p = b2.text_frame.paragraphs[0]
        p.text = "今の3大メガバンク\n\n・三菱UFJ銀行\n・三井住友銀行\n・みずほ銀行"
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = WHITE
        b2.text_frame.word_wrap = True

    def draw_credit_creation(self, slide):
        y = Inches(4.0)
        
        shape1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.0), y, Inches(2), Inches(2))
        shape1.fill.solid()
        shape1.fill.fore_color.rgb = LIGHT_GREY
        shape1.line.color.rgb = ACCENT_LINE
        shape1.text_frame.paragraphs[0].text = "Aさんの預金\n(100円)"
        shape1.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
        shape1.text_frame.paragraphs[0].font.bold = True
        
        slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.2), y+Inches(0.8), Inches(0.8), Inches(0.4)).fill.solid()
        
        shape2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.2), y, Inches(2), Inches(2))
        shape2.fill.solid()
        shape2.fill.fore_color.rgb = LIGHT_GREY
        shape2.line.color.rgb = ACCENT_LINE
        shape2.text_frame.paragraphs[0].text = "Bさんに貸出\n(90円)"
        shape2.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
        shape2.text_frame.paragraphs[0].font.bold = True

        slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.4), y+Inches(0.8), Inches(0.8), Inches(0.4)).fill.solid()

        shape3 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.4), y, Inches(2), Inches(2))
        shape3.fill.solid()
        shape3.fill.fore_color.rgb = LIGHT_GREY
        shape3.line.color.rgb = ACCENT_LINE
        shape3.text_frame.paragraphs[0].text = "別の預金へ\n(90円)"
        shape3.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
        shape3.text_frame.paragraphs[0].font.bold = True
        
        slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.6), y+Inches(0.8), Inches(0.8), Inches(0.4)).fill.solid()
        
        shape4 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.6), y-Inches(0.25), Inches(2.5), Inches(2.5))
        shape4.fill.solid()
        shape4.fill.fore_color.rgb = TEXT_DARK
        shape4.line.visible = False
        shape4.text_frame.paragraphs[0].text = "世の中のお金\n= 190円に！"
        shape4.text_frame.paragraphs[0].font.color.rgb = WHITE
        shape4.text_frame.paragraphs[0].font.size = Pt(24)
        shape4.text_frame.paragraphs[0].font.bold = True
        
        lbl = slide.shapes.add_textbox(Inches(1.0), Inches(2.8), Inches(11), Inches(1))
        lbl.text_frame.paragraphs[0].text = "銀行を挟んで貸し借りを繰り返すだけで、世の中のお金が増える不思議な仕組み。"
        lbl.text_frame.paragraphs[0].font.size = Pt(22)

    def draw_interest_margin(self, slide):
        base_x = Inches(1.0)
        y = Inches(3.5)
        
        # Deposit
        b1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, base_x, y, Inches(3.3), Inches(2.3))
        b1.fill.solid()
        b1.fill.fore_color.rgb = LIGHT_GREY
        b1.line.color.rgb = ACCENT_LINE
        p1 = b1.text_frame.paragraphs[0]
        p1.text = "① みんなから預かる\n(利息 1円を払う)"
        p1.font.color.rgb = TEXT_DARK
        p1.font.size = Pt(22)
        p1.font.bold = True
        
        # Loan
        b2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, base_x + Inches(3.8), y, Inches(3.3), Inches(2.3))
        b2.fill.solid()
        b2.fill.fore_color.rgb = SUB_GREY
        b2.line.visible = False
        p2 = b2.text_frame.paragraphs[0]
        p2.text = "② 企業に貸す\n(利息 10円をもらう)"
        p2.font.color.rgb = WHITE
        p2.font.size = Pt(22)
        p2.font.bold = True
        
        # Margin
        b3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, base_x + Inches(7.6), y, Inches(3.5), Inches(2.3))
        b3.fill.solid()
        b3.fill.fore_color.rgb = TEXT_DARK
        b3.line.visible = False
        p3 = b3.text_frame.paragraphs[0]
        p3.text = "③ 銀行の利益\n(差額の 9円)"
        p3.font.color.rgb = WHITE
        p3.font.size = Pt(26)
        p3.font.bold = True
        
        lbl = slide.shapes.add_textbox(Inches(1.0), Inches(6.0), Inches(11.333), Inches(1))
        lbl.text_frame.paragraphs[0].text = "スーパーの「安く仕入れて、高く売る」と全く同じ『利ざや』の考え方です。"
        lbl.text_frame.paragraphs[0].font.size = Pt(24)


    def draw_balance_sheet(self, slide):
        y = Inches(3.5)
        w = Inches(4.5)
        h = Inches(3.0)
        center = self.prs.slide_width / 2
        
        left = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, center - w - Inches(0.2), y, w, h)
        left.fill.solid()
        left.fill.fore_color.rgb = LIGHT_GREY
        left.line.color.rgb = ACCENT_LINE
        left.text_frame.paragraphs[0].text = "■ お金の使い道 (資産)\n\n企業に貸しているお金など"
        left.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
        left.text_frame.paragraphs[0].font.size = Pt(22)
        left.text_frame.paragraphs[0].font.bold = True
        left.text_frame.word_wrap = True
        
        right = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, center + Inches(0.2), y, w, h)
        right.fill.solid()
        right.fill.fore_color.rgb = TEXT_DARK
        right.line.visible = False
        right.text_frame.paragraphs[0].text = "■ お金の集め方 (負債)\n\nみんなから預かっているお金"
        right.text_frame.paragraphs[0].font.color.rgb = WHITE
        right.text_frame.paragraphs[0].font.size = Pt(22)
        right.text_frame.paragraphs[0].font.bold = True
        right.text_frame.word_wrap = True
        
        lbl = slide.shapes.add_textbox(center - Inches(3), y - Inches(0.8), Inches(6), Inches(0.5))
        lbl.text_frame.paragraphs[0].text = "私たちが見ると逆！みんなの「預金」は銀行にとって「借金」"
        lbl.text_frame.paragraphs[0].font.size = Pt(20)
        lbl.text_frame.paragraphs[0].font.bold = True
        lbl.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    def draw_industry_map(self, slide):
        groups = [
            ("日本銀行", "全てのお金の元締め (紙幣を刷る)", Inches(0.6), Inches(3.0)),
            ("メガバンク", "全国展開・世界相手 (三菱UFJなど)", Inches(3.7), Inches(3.0)),
            ("地方銀行", "あなたの地元の企業を応援してる", Inches(6.8), Inches(3.0)),
            ("ネット銀行", "スマホで完結・店舗がない (楽天など)", Inches(9.9), Inches(3.0))
        ]
        for name, desc, x, y in groups:
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(2.8), Inches(2.5))
            box.fill.solid()
            if name == "日本銀行":
                box.fill.fore_color.rgb = TEXT_DARK
                box.text_frame.paragraphs[0].font.color.rgb = WHITE
            else:
                box.fill.fore_color.rgb = LIGHT_GREY
                box.line.color.rgb = ACCENT_LINE
                box.text_frame.paragraphs[0].font.color.rgb = TEXT_DARK
                
            p = box.text_frame.paragraphs[0]
            p.text = f"{name}\n\n{desc}"
            p.font.size = Pt(18)
            p.font.bold = True
            box.text_frame.word_wrap = True

    def add_summary_slide(self, items):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = WHITE
        bg.line.visible = False

        self._add_base_design(slide)
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.6), Inches(12), Inches(1))
        p = title_box.text_frame.paragraphs[0]
        p.text = "今日のまとめ"
        p.font.name = FONT_SANS
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        
        body_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(4.5))
        tf = body_box.text_frame
        tf.word_wrap = True
        for item in items:
            p = tf.add_paragraph()
            p.text = f"✔ {item}"
            p.font.name = FONT_SANS
            p.font.size = Pt(28)
            p.space_after = Pt(24)
            p.font.color.rgb = TEXT_DARK

    def save(self, filename):
        self.prs.save(filename)

# --- Generator Script ---
def generate_normal_presentation():
    gen = PresentationGenerator("みんなの銀行基礎知識", "〜 難しくない！お金と銀行のリアルな話 〜")
    
    # 1. 導入 (3)
    gen.add_cover_slide()
    gen.add_message_slide("導入", "Q. 銀行って、何をするところ？", [
        "お金を安全に預けておく「巨大な金庫」？",
        "毎月のお給料が振り込まれる「自分の口座」？",
        "実は、ただの「お金の置き場所」ではありません。",
        "社会中にお金をくまなく巡らせる『心臓』としての役割があります。"
    ])
    gen.add_transition_slide("1", "銀行ってそもそも何？")
    gen.add_message_slide("本日のアジェンダ", "今日お話しすること", [
        "1. 銀行の昔と今（歴史）",
        "2. 銀行はどうやって儲けてる？（仕組み）",
        "3. 今の銀行のリアル（現状）",
        "4. PayPayやスタバがライバル？（脅威）",
        "5. これからの銀行（未来）"
    ])
    
    # 2. 歴史 (2)
    gen.add_transition_slide("2", "銀行の歴史：合併の嵐")
    gen.add_diagram_slide("銀行の歴史", "昔はたくさんあった銀行が、合体して巨大化しました。", gen.draw_megabank_flow)
    
    # 3. 仕組み (8)
    gen.add_transition_slide("3", "銀行のお仕事と儲けの仕組み")
    gen.add_message_slide("銀行の三大業務（大切なお仕事）", "銀行が絶対にやっている3つのこと", [
        "預金 (よきん) ： みんなのお金を安全にお預かりする",
        "融資 (ゆうし) ： お金に困っている人や頑張る企業に貸す",
        "為替 (かわせ) ： 遠くの人へ、現金を直接運ばずにお金を送金する"
    ])
    gen.add_message_slide("お金のめぐり方", "銀行は「お金の橋渡し」", [
        "お金が余っている人（私たち）から預かる。",
        "お金が足りない人や企業に貸す。",
        "銀行という橋があるから、社会全体でお金がスムーズに回ります。"
    ])
    gen.add_diagram_slide("魔法の仕組み：信用創造", "銀行を挟むと、なぜかお金の総額が増える！？", gen.draw_credit_creation)
    gen.add_message_slide("信用創造のタネあかし", "「データ上のお金」が増えているだけ", [
        "実際に現金の札束や硬貨が増えているわけではありません。",
        "通帳の残高に「○○円ありますよ」と記録されることで、お金として使えます。",
        "この仕組みのおかげで、企業は事業を大きくすることができます。"
    ])
    gen.add_diagram_slide("どうやって儲ける？ (利ざや)", "普通の商売と同じ「安く仕入れて、高く売る」", gen.draw_interest_margin)
    gen.add_message_slide("他の儲け方 (手数料ビジネス)", "利ざや以外にもこんな稼ぎ方があります", [
        "振込手数料：お金を送る時にかかる「送料」のようなもの。",
        "ATM手数料：時間外や他行のお引き出しなどで発生するお金。",
        "投資信託・保険：他の金融商品を「代わりに売ってあげる」ことで貰えるお礼。"
    ])
    gen.add_diagram_slide("銀行のバランスシート (持ち物リスト)", "私たちのお金は、銀行にとって『預かりもの』", gen.draw_balance_sheet)

    # 4. 現状 (4)
    gen.add_transition_slide("4", "銀行業界のリアルな現状")
    gen.add_diagram_slide("銀行の仲間たち", "目的や規模に合わせていろいろな銀行があります。", gen.draw_industry_map)
    gen.add_message_slide("今の銀行の悩み：金利がほぼ0円", "「利ざや」が稼げなくて大ピンチ！", [
        "今は金利が低いため、企業に貸しても「利息」がほとんど貰えません。",
        "100万円を1年預けても、チロルチョコ1個買えない時代です。"
    ])
    gen.add_message_slide("地方銀行の「合体」と「統廃合」", "生き残りをかけて規模を大きくする", [
        "稼ぎにくくなったため、地方の銀行同士が協力・合体するニュースが増えています。",
        "また、人が来ない店舗（支店）を減らして、コストを下げる努力をしています。"
    ])

    # 5. ライバル (4)
    gen.add_transition_slide("5", "新しいライバルたちの登場")
    gen.add_message_slide("スマホ決済の台頭 (PayPayやメルペイ)", "銀行に行かなくても支払いができる！", [
        "お店で払う時、現金（銀行）を直接触らずにスマホで「ピッ」とするだけ。",
        "お金のやり取りに、銀行という「橋」を通らなくてもよい世界になっています。"
    ])
    gen.add_message_slide("ネット銀行の急成長", "「スマホの中にある便利な銀行」", [
        "店舗を持たないので、家賃や人件費がかからず、手数料が安いです。",
        "ポイントが貯まりやすかったり、アプリが使いやすかったりします。"
    ])
    gen.add_message_slide("超強力なライバル：IT企業やカフェ", "あのAppleやスタバが実質的な銀行に！？", [
        "Apple（Apple Cardなど）や、スターバックス（スタバカードへのチャージ）。",
        "みんなが先にお金を預けて使ってくれるので、「銀行のようなこと」ができています。",
        "銀行のライバルは、異業種の大企業になってきています。"
    ])

    # 6. 未来 (4)
    gen.add_transition_slide("6", "銀行のこれからの姿")
    gen.add_message_slide("DX (デジタル化) の推進", "紙とハンコからの卒業", [
        "「店舗に行って、書類を書いて、ハンコを押す」という昔のやり方を変える。",
        "スマホアプリひとつで全部完結する「超ベンリな銀行」へ進化中です。"
    ])
    gen.add_message_slide("これからのビジネスモデル", "「お金を貸すだけ」からの卒業", [
        "お金の貸し借りだけでなく、みんなの「生活の悩み」を解決するアドバイザーへ。",
        "他の便利なアプリの「裏側」に、銀行の仕組み・機能だけをこっそり提供します。"
    ])
    gen.add_message_slide("2030年の銀行像", "「気づかないうちに銀行を使っている」未来", [
        "わざわざ「銀行に行く」という言葉が、死語になるかもしれません。",
        "スマホやスマートウォッチ、生活のあらゆる場面に金融の仕組みが溶け込みます。"
    ])

    # 7. まとめ (1)
    gen.add_transition_slide("7", "まとめ")
    gen.add_summary_slide([
        "銀行は、預金・融資・為替を通じて社会にお金を巡らせる「心臓」。",
        "「安く仕入れて高く売る」のが基本で、今は金利が低くて大変。",
        "PayPayやスタバなど、他のサービスとの境界線がなくなりつつある。",
        "これからはスマホの奥に溶け込む「見えない銀行」へと進化していく。"
    ])

    output_path = "bank_normal_presentation.pptx"
    gen.save(output_path)
    print(f"Normal Presentation saved to {output_path}")

if __name__ == "__main__":
    generate_normal_presentation()
