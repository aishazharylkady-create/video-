import copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

TPL, OUT = sys.argv[1], sys.argv[2]
F = "Times New Roman"
RED = RGBColor(0xD0, 0x00, 0x2F)
WINE = RGBColor(0x7A, 0x00, 0x19)
INK = RGBColor(0x22, 0x1E, 0x1F)
GREY = RGBColor(0x6B, 0x63, 0x64)
CREAM = RGBColor(0xF7, 0xF1, 0xEA)
GHOST = RGBColor(0xF4, 0xE4, 0xE7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x7D, 0x4F)
LINE = RGBColor(0xE2, 0xD6, 0xD0)

prs = Presentation(TPL)
slides = list(prs.slides)

# replace template slide 10 (red title layout) with a white content slide
def drop_slide(prs, slide):
    sldIdLst = prs.slides._sldIdLst
    for sid in list(sldIdLst):
        if prs.part.related_part(sid.rId) is slide.part:
            prs.part.drop_rel(sid.rId)
            sldIdLst.remove(sid)
drop_slide(prs, slides[9])
blank = prs.slide_layouts[6]
s10 = prs.slides.add_slide(blank)
slides = list(prs.slides)

# remove stray "." textboxes
for s in slides[1:]:
    for sh in list(s.shapes):
        if sh.shape_type == 17 and sh.text_frame.text.strip() in (".", ""):
            sh._element.getparent().remove(sh._element)

def tb(s, x, y, w, h, paras, size=12, color=INK, bold=False, align=PP_ALIGN.LEFT,
       anchor=MSO_ANCHOR.TOP, italic=False, spacing=None, line=None):
    """paras: str | list of (str | list of runs). run = (text, {opts})"""
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        if spacing is not None:
            para.space_after = Pt(spacing)
        if line is not None:
            para.line_spacing = line
        runs = p if isinstance(p, list) else [(p, {})]
        for txt, o in runs:
            r = para.add_run()
            r.text = txt
            f = r.font
            f.name = F
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.color.rgb = o.get("color", color)
    return box

def rect(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, lw=0.75):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.shadow.inherit = False
    if fill is None:
        r.fill.background()
    else:
        r.fill.solid(); r.fill.fore_color.rgb = fill
    if line is None:
        r.line.fill.background()
    else:
        r.line.color.rgb = line; r.line.width = Pt(lw)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        r.adjustments[0] = 0.08
    return r

def header(s, ghost, title, kicker, sub=None):
    tb(s, 0.62, 0.05, 12.2, 1.6, ghost, size=110, color=GHOST, bold=True)
    tb(s, 0.75, 0.32, 11.5, 0.3, kicker.upper(), size=11, color=RED, bold=True)
    tb(s, 0.75, 0.55, 11.8, 0.8, title.upper(), size=34, color=WINE, bold=True)
    if sub:
        tb(s, 0.77, 1.22, 11.5, 0.35, sub, size=13, color=GREY, italic=True)
    ln = s.shapes.add_connector(1, Inches(0.78), Inches(1.6), Inches(12.55), Inches(1.6))
    ln.line.color.rgb = RED; ln.line.width = Pt(1.25)

def source(s, text):
    tb(s, 0.78, 7.02, 9.0, 0.3, "Дереккөз: " + text, size=9, color=GREY, italic=True)

def card(s, x, y, w, h, big, small, bigsize=30, fill=CREAM, bigcolor=RED):
    rect(s, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tb(s, x + 0.12, y + 0.08, w - 0.24, h * 0.55, big, size=bigsize, color=bigcolor, bold=True,
       anchor=MSO_ANCHOR.BOTTOM)
    tb(s, x + 0.12, y + h * 0.58, w - 0.24, h * 0.42, small, size=11, color=INK)

def style_chart(ch, size=11):
    ch.font.name = F; ch.font.size = Pt(size); ch.font.color.rgb = INK

def comma_labels(ch):
    for ser in ch.plots[0].series:
        for i, v in enumerate(ser.values):
            if v is None: continue
            dl = ser.points[i].data_label
            tf = dl.text_frame; tf.text = (("%g" % v).replace(".", ","))
            r = tf.paragraphs[0].runs[0]; r.font.name = F; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = INK
            dl.position = XL_LABEL_POSITION.OUTSIDE_END

def table(s, x, y, w, h, data, colw, header_fill=WINE, size=11, first_col_fill=None, row_h=None):
    rows, cols = len(data), len(data[0])
    gt = s.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    t = gt.table
    tblPr = gt._element.graphic.graphicData.tbl.tblPr
    tblPr.set("bandRow", "0"); tblPr.set("firstRow", "0")
    style = tblPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId')
    if style is None:
        from lxml import etree
        style = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId')
    style.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'  # No style, table grid
    for i, cw in enumerate(colw):
        t.columns[i].width = Inches(cw)
    for r in range(rows):
        if row_h:
            t.rows[r].height = Inches(row_h[r])
        for c in range(cols):
            cell = t.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.07)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            val = data[r][c]
            runs = val if isinstance(val, list) else [(val, {})]
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            for txt, o in runs:
                rr = p.add_run(); rr.text = txt
                rr.font.name = F; rr.font.size = Pt(o.get("size", size))
                rr.font.bold = o.get("bold", r == 0)
                rr.font.color.rgb = o.get("color", WHITE if r == 0 else INK)
            if r == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
            elif c == 0 and first_col_fill is not None:
                cell.fill.solid(); cell.fill.fore_color.rgb = first_col_fill
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if r % 2 else CREAM
    return t

# ---------------- 1. TITLE ----------------
s = slides[0]
for ph in list(s.placeholders):
    ph._element.getparent().remove(ph._element)
tb(s, 3.2, 1.75, 9.2, 0.4, "«МАРКЕТИНГ НЕГІЗДЕРІ» ПӘНІ · АКАДЕМИЯЛЫҚ ТАЛДАУ", size=13, color=WHITE, bold=True)
tb(s, 3.1, 2.05, 9.6, 1.6, "MAGNUM", size=96, color=WHITE, bold=True)
tb(s, 3.2, 3.55, 8.9, 1.4,
   "Компанияның маркетингтік жоспары мен маркетингтік мүмкіндіктерін 7P, PEST және SWOT талдау негізінде бағалау",
   size=22, color=WHITE, line=1.05)
ln = s.shapes.add_connector(1, Inches(3.25), Inches(5.05), Inches(11.6), Inches(5.05))
ln.line.color.rgb = WHITE; ln.line.width = Pt(1)
tb(s, 3.2, 5.2, 4.3, 1.0, [[("Орындаған:", {"bold": True})], "Жарылқады Айша Әсетқызы"], size=14, color=WHITE)
tb(s, 7.6, 5.2, 4.3, 1.0, [[("Қабылдаған:", {"bold": True})], "Нурахова Ботагөз"], size=14, color=WHITE)
tb(s, 3.2, 6.35, 8.5, 0.4, "Нархоз университеті · Алматы, 2026", size=12, color=WHITE, italic=True)

# ---------------- 2. COMPANY ----------------
s = slides[1]
header(s, "MAGNUM", "Компания сипаттамасы", "01 · Маркетинг жоспары · компания және нарықтағы орны",
       "Қазақстандағы ең ірі азық-түлік бөлшек сауда желісі")
cw, gap, x0 = 2.85, 0.13, 0.78
cards = [("2007", "Алматыда алғашқы дүкен ашылды"), ("200+", "филиал, Қазақстанның 14 қаласында"),
         ("453 млрд ₸", "табыс, Forbes-75 рейтингінде 7-орын"), ("1,4 млн+", "Magnum Club бонус бағдарламасының қатысушысы")]
for i, (b, t) in enumerate(cards):
    card(s, x0 + i * (cw + gap), 1.85, cw, 1.35, b, t, bigsize=30 if i != 2 else 26)
tb(s, 0.78, 3.4, 6.0, 0.35, "МИССИЯСЫ", size=12, color=RED, bold=True)
tb(s, 0.78, 3.7, 5.9, 1.0, "«Күн сайын отбасылық бюджетті сақтай отырып, өмір сапасын арттырамыз»",
   size=19, color=WINE, italic=True, line=1.0)
tb(s, 0.78, 4.75, 5.9, 0.35, "ФОРМАТТАРЫ", size=12, color=RED, bold=True)
fm = [("Hyper", "24/7 гипермаркет"), ("Super", "супермаркет"), ("Express", "үйге жақын"), ("Daily", "80–250 м²")]
for i, (a, b) in enumerate(fm):
    rect(s, 0.78 + i * 1.48, 5.1, 1.38, 0.85, WHITE, line=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tb(s, 0.78 + i * 1.48, 5.13, 1.38, 0.45, a, size=16, color=WINE, bold=True, align=PP_ALIGN.CENTER)
    tb(s, 0.78 + i * 1.48, 5.52, 1.38, 0.35, b, size=10.5, color=GREY, align=PP_ALIGN.CENTER)
tb(s, 0.78, 6.1, 5.9, 0.85,
   "Сауда түрі: азық-түлік және тұрмыстық тауарлар (FMCG), өзіндік брендтер, дайын өнім, Magnum Go онлайн-тапсырыс.",
   size=12, color=INK)
# timeline
tb(s, 7.15, 3.4, 5.4, 0.35, "ДАМУ КЕЗЕҢДЕРІ", size=12, color=RED, bold=True)
tl = [("2007", "Алматыда алғашқы Cash&Carry дүкені"),
      ("2019", "Magnum Club бонус бағдарламасы"),
      ("2021", "Daily шағын форматы іске қосылды"),
      ("2023", "Kaspi.kz-пен бірлескен E-Magnum (онлайн)"),
      ("2025", "Өзбекстан нарығынан шығу, Қазақстанға шоғырлану")]
bar = s.shapes.add_connector(1, Inches(7.32), Inches(3.9), Inches(7.32), Inches(6.75))
bar.line.color.rgb = LINE; bar.line.width = Pt(2)
for i, (yr, t) in enumerate(tl):
    y = 3.82 + i * 0.6
    rect(s, 7.24, y + 0.08, 0.17, 0.17, RED, shape=MSO_SHAPE.OVAL)
    tb(s, 7.55, y, 0.8, 0.35, yr, size=14, color=WINE, bold=True)
    tb(s, 8.35, y + 0.02, 4.3, 0.55, t, size=12, color=INK)
source(s, "magnum.kz/about; Forbes Kazakhstan, 2025; Forbes.kz, 2023, 2025; 10-дәріс материалы.")

# ---------------- 3. MARKET ----------------
s = slides[2]
header(s, "НАРЫҚ", "Нарықтық жағдай", "02 · Маркетинг жоспары · жағдайлық талдау",
       "Табыс төмендеген кезде тұтынушы бағаға және ыңғайлылыққа сезімтал")
steps = [["Жағдайлық", "талдау"], ["Нарықтық", "мүмкіндік"], "Мақсаттар", "Стратегия", "Бағдарлама", "Бюджет", "Бақылау"]
for i, st in enumerate(steps):
    shp = rect(s, 0.78 + i * 1.69, 1.8, 1.72, 0.72, RED if i < 2 else (WINE if i < 4 else CREAM), shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
    tb(s, 1.0 + i * 1.69, 1.84, 1.3, 0.65, st, size=10.5, color=WHITE if i < 4 else WINE, bold=True,
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line=0.9)
# chart retail growth
cd = CategoryChartData(); cd.categories = ["2024", "2025", "2026 (9 ай)"]
cd.add_series("Бөлшек сауда өсімі, %", (9.8, 7.5, None))
cd.add_series("Азық-түлік сатылымы өсімі, %", (None, 7.8, 10.1))
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.7), Inches(2.95), Inches(5.6), Inches(3.4), cd)
ch = gf.chart; style_chart(ch)
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
ch.value_axis.visible = False; ch.value_axis.has_major_gridlines = False
ch.category_axis.format.line.color.rgb = LINE
for ser, col in zip(ch.plots[0].series, (WINE, RED)):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
    ser.data_labels.show_value = True; ser.data_labels.font.size = Pt(12); ser.data_labels.font.bold = True
    ser.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
ch.plots[0].gap_width = 80
comma_labels(ch)
tb(s, 0.78, 2.65, 5.5, 0.3, "Қазақстанның бөлшек сауда динамикасы, жылдық өсім, %", size=12, color=RED, bold=True)
# right column
card(s, 6.6, 2.65, 1.9, 1.3, "35,1%", "азық-түлік үлесі, бөлшек сауда (2026, 9 ай)", bigsize=24)
card(s, 8.6, 2.65, 1.9, 1.3, "52%", "шығыстағы азық-түлік үлесі (2025)", bigsize=24)
card(s, 10.6, 2.65, 1.95, 1.3, "−3,4%", "нақты табыс, 2025 ж. III тоқсан", bigsize=24)
tb(s, 6.6, 4.05, 6.0, 0.3, "МАҚСАТТЫ АУДИТОРИЯ ЖӘНЕ БӘСЕКЕЛЕСТЕР", size=12, color=RED, bold=True)
table(s, 6.6, 4.38, 5.95, 2.4, [
    ["Сегмент / бәсекелес", "Сипаттамасы"],
    ["Қалалық отбасылар", "Апталық ірі себет, баға мен акцияға сезімтал"],
    ["Жастар (18–35)", "Онлайн-тапсырыс, жылдамдық, Kaspi экожүйесі"],
    ["Small / Skif / Spar", "150+ дүкен, үйге жақын формат"],
    ["Galmart · «Светофор»", "Премиум сегмент · дискаунтер"],
    ["Arbuz, Glovo, Wolt", "Жедел онлайн-жеткізу"],
], [2.1, 3.85], size=10.5)
source(s, "ҰСБ (stat.gov.kz), 2025–2026; Kursiv.media, 2026; INFOLine, 2025; 10-дәріс.")

# ---------------- 4. PEST ----------------
s = slides[3]
header(s, "PEST", "PEST-талдау", "03 · Маркетингтік мүмкіндіктер · макроорта",
       "Әр фактордың Magnum қызметіне нақты ықпалы және бағыты")
P = lambda t: [("+ ", {"bold": True, "color": GREEN}), (t, {})]
M = lambda t: [("− ", {"bold": True, "color": RED}), (t, {})]
data = [
    ["", "Нақты дерек (2025–2026)", "Magnum-ға ықпалы", "Бағыт"],
    [[("P", {"bold": True, "size": 26, "color": WHITE})], "ҚҚС 12% → 16% (2026); әлеуметтік маңызы бар 31 тауарға үстеме ≤ 15%",
     "Маржаға қысым; бағаны әділ ұстау арқылы сенім қалыптастыру мүмкіндігі", M("Қауіп")],
    [[("E", {"bold": True, "size": 26, "color": WHITE})], "Инфляция 12,3%, азық-түлік 13,5% (2025); базалық мөлшерлеме 18% → 17%",
     "Орташа чек тұрақсыз; арзан себет пен өзіндік брендке сұраныс өседі", M("Қауіп")],
    [[("S", {"bold": True, "size": 26, "color": WHITE})], "Шығыстың 52%-ы азық-түлікке; «үйге жақын» және жедел жеткізу",
     "Шағын форматтар мен Magnum Go-ға сұраныс; адалдық төмендейді", P("Мүмкіндік")],
    [[("T", {"bold": True, "size": 26, "color": WHITE})], "E-commerce 14,1% (2024) → 18,5% (2029); Kaspi e-Grocery 1,7 млн тұтынушы",
     "Омниарна және деректерге негізделген маркетинг", P("Мүмкіндік")],
]
t = table(s, 0.78, 1.85, 11.77, 4.3, data, [0.8, 4.6, 4.97, 1.4], size=12, first_col_fill=RED,
          row_h=[0.45, 0.96, 0.96, 0.96, 0.96])
for r in range(1, 5):
    t.cell(r, 0).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
rect(s, 0.78, 6.3, 11.77, 0.62, CREAM, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tb(s, 0.95, 6.33, 11.5, 0.56, [[("Қорытынды: ", {"bold": True, "color": WINE}),
   ("сыртқы ортада бағаны реттеу мен табыстың төмендеуі – негізгі қауіп, ал цифрландыру – Magnum үшін ең күшті мүмкіндік.", {})]],
   size=13, anchor=MSO_ANCHOR.MIDDLE)
source(s, "ҚР Ұлттық банкі, 2026; ҚР Салық кодексі; Zakon.kz, 2025; ҚР Сауда министрлігі; Kaspi.kz 9M 2026 есебі.")

# ---------------- 5. SWOT ----------------
s = slides[4]
header(s, "SWOT", "SWOT-талдау", "04 · Маркетингтік мүмкіндіктер · ішкі және сыртқы орта",
       "Әр тұжырым нақты деректермен негізделген")
quad = [
    ("S", "Күшті жақтар", WINE, WHITE, ["Нарықтағы ауқым: 200+ дүкен, 453 млрд ₸ табыс",
                                        "Magnum Club: 1,4 млн+ қатысушы",
                                        "Kaspi экожүйесі, E-Magnum dark store-лары",
                                        "Өзіндік брендтер: Kunde, Toibastar, «Всё в дом»"]),
    ("W", "Әлсіз жақтар", CREAM, INK, ["2024 ж. дүкен саны 13-ке азайды",
                                       "Өзбекстандағы экспансия сәтсіз аяқталды",
                                       "Гипермаркеттерге келушілер азаюда",
                                       "14 қаламен шектелген өңірлік қамту"]),
    ("O", "Мүмкіндіктер", CREAM, INK, ["E-commerce үлесі 18,5%-ға дейін өседі",
                                       "«Үйге жақын» сегмент: 2027 жылға дейін 300 дүкен жоспары",
                                       "Бағаға сезімтал тұтынушыға өзіндік бренд ұсыну",
                                       "Отандық өндірушілермен келісімшарттар"]),
    ("T", "Қауіптер", RED, WHITE, ["Инфляция 12,3%, ҚҚС 16%",
                                   "Әлеуметтік маңызы бар тауарларға 15% үстеме шегі",
                                   "Дискаунтерлер және Small желісінің кеңеюі",
                                   "Glovo, Wolt, Arbuz бәсекесі"]),
]
qw, qh = 5.83, 2.45
for i, (L, name, fill, fc, items) in enumerate(quad):
    x = 0.78 + (i % 2) * (qw + 0.11); y = 1.8 + (i // 2) * (qh + 0.11)
    rect(s, x, y, qw, qh, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tb(s, x + 0.15, y - 0.1, 1.2, 1.3, L, size=64, color=fc if fill != CREAM else RED, bold=True)
    tb(s, x + 1.25, y + 0.15, qw - 1.4, 0.4, name.upper(), size=14, color=fc if fill != CREAM else WINE, bold=True)
    tb(s, x + 1.25, y + 0.6, qw - 1.4, qh - 0.6, ["• " + it for it in items], size=14,
       color=fc, spacing=3)
source(s, "Forbes Kazakhstan, 2025; INFOLine, 2025; magnum.kz; Kapital.kz, 2026; ҚР Ұлттық банкі.")

# ---------------- 6. PRODUCT & PRICE ----------------
s = slides[5]
header(s, "7P", "Product және Price", "05 · 7P маркетинг кешені · 1–2",
       "Ассортимент пен баға «отбасылық бюджетті сақтау» миссиясына қызмет етеді")
tb(s, 0.78, 1.8, 5.8, 0.35, "PRODUCT · ТАУАРДЫҢ ДЕҢГЕЙЛЕРІ", size=13, color=RED, bold=True)
for i, (d, col) in enumerate([(3.3, CREAM), (2.4, GHOST), (1.5, RED)]):
    off = (3.3 - d) / 2
    rect(s, 0.85 + off, 2.25 + off, d, d, col, shape=MSO_SHAPE.OVAL)
tb(s, 1.85, 3.55, 1.3, 0.7, ["Нақты", "тауар"], size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
tb(s, 1.7, 2.53, 1.6, 0.4, "Кеңейтілген", size=11, color=WINE, bold=True, align=PP_ALIGN.CENTER)
tb(s, 1.7, 2.3, 1.6, 0.3, "Жинақталған", size=11, color=GREY, bold=True, align=PP_ALIGN.CENTER)
tb(s, 4.3, 2.25, 2.5, 3.4, [
    [("Нақты: ", {"bold": True, "color": WINE}), ("азық-түлік, тұрмыстық тауар, дайын өнім", {})],
    [("Кеңейтілген: ", {"bold": True, "color": WINE}), ("Magnum Go жеткізу, бонустар, 24/7 жұмыс", {})],
    [("Өзіндік брендтер: ", {"bold": True, "color": WINE}), ("Kunde, Toibastar, «Всё в дом», Ocean food", {})],
    [("Отандық өнім: ", {"bold": True, "color": WINE}), ("ассортименттің ~50%-ы (2018)", {})],
], size=11.5, spacing=6)
tb(s, 7.05, 1.8, 5.5, 0.35, "PRICE · БАҒА СТРАТЕГИЯСЫ", size=13, color=RED, bold=True)
table(s, 7.05, 2.2, 5.5, 3.45, [
    ["Стратегия (7-дәріс)", "Magnum-дағы көрінісі"],
    ["Психологиялық баға", "990 ₸, 1 490 ₸ сияқты «тартымды» бағалар"],
    ["Ассортименттік баға", "Эконом / орта / премиум баға деңгейлері"],
    ["Баға дифференциациясы", "Magnum Club жеке жеңілдіктері"],
    ["Акциялық баға", "Апталық каталог, «1+1» ұсыныстары"],
    ["Мемлекеттік реттеу", "Әлеуметтік маңызы бар тауарларға ≤ 15% үстеме"],
], [2.2, 3.3], size=11)
rect(s, 0.78, 5.85, 11.77, 1.05, CREAM, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tb(s, 0.95, 5.9, 11.5, 0.95, [[("Аналитикалық қорытынды: ", {"bold": True, "color": WINE}),
   ("инфляция жағдайында Magnum-ның басты бәсекелік құралы – өзіндік брендтер мен бонус арқылы берілетін жеке баға. "
    "Бірақ отандық өнім үлесі туралы ашық дерек ескірген (2018), бұл жеткізушілер саясатындағы ақпараттық олқылық.", {})]],
   size=12, anchor=MSO_ANCHOR.MIDDLE)
source(s, "magnum.kz; Foodovik, 2024; Wikipedia (2018 дерегі); ҚР Сауда министрлігі; 6- және 7-дәрістер.")

# ---------------- 7. PLACE & PROMOTION ----------------
s = slides[6]
header(s, "7P", "Place және Promotion", "06 · 7P маркетинг кешені · 3–4",
       "Омниарна: дүкен + қосымша + Kaspi экожүйесі")
tb(s, 0.78, 1.8, 5.8, 0.35, "PLACE · ӨТКІЗУ АРНАЛАРЫ", size=13, color=RED, bold=True)
nodes = [("Дүкендер", "200+ филиал · 14 қала", 0.78, 2.3), ("Magnum Go", "қосымша · жеткізу", 3.85, 2.3),
         ("Kaspi.kz", "e-Grocery · dark store", 2.3, 3.75)]
for n, d, x, y in nodes:
    rect(s, x, y, 2.7, 1.05, WINE if n == "Дүкендер" else (RED if n == "Kaspi.kz" else CREAM),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    fc = INK if n == "Magnum Go" else WHITE
    tb(s, x, y + 0.1, 2.7, 0.45, n, size=17, color=WINE if n == "Magnum Go" else WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(s, x, y + 0.55, 2.7, 0.4, d, size=11, color=fc, align=PP_ALIGN.CENTER)
for (x1, y1, x2, y2) in [(3.48, 2.82, 3.85, 2.82), (2.13, 3.35, 2.9, 3.75), (5.2, 3.35, 4.4, 3.75)]:
    c = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = GREY; c.line.width = Pt(1.5)
tb(s, 0.78, 4.95, 5.8, 0.9, "Тегін жеткізу 6 000 ₸-ден басталады. Daily форматы (80–250 м²) тұрғын аудандарға жақын орналасады.",
   size=11.5, color=INK)
# promotion chart
tb(s, 7.05, 1.8, 5.5, 0.35, "PROMOTION · ЖЫЛЖЫТУ", size=13, color=RED, bold=True)
cd = CategoryChartData(); cd.categories = ["2024 (факт)", "2025 (жоспар)", "2029 (мақсат)"]
cd.add_series("E-commerce үлесі, %", (14.1, 15.0, 18.5))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(7.0), Inches(2.15), Inches(5.6), Inches(1.95), cd)
ch = gf.chart; style_chart(ch); ch.has_legend = False
ch.value_axis.visible = False; ch.value_axis.has_major_gridlines = False
ch.category_axis.reverse_order = True; ch.category_axis.format.line.color.rgb = LINE
ser = ch.plots[0].series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = RED
ser.data_labels.show_value = True; ser.data_labels.font.bold = True; ser.data_labels.font.size = Pt(12)
ch.plots[0].gap_width = 60
ch.has_title = False
comma_labels(ch)
tb(s, 7.05, 4.05, 5.5, 0.3, "Қазақстан бөлшек саудасындағы e-commerce үлесі, %", size=10, color=GREY, italic=True)
promo = [("Magnum Club", "бонустар, жеке ұсыныстар"), ("SMM", "Instagram, TikTok"),
         ("Каталог", "апталық акциялар"), ("MUM Club", "аналарға арналған бағдарлама")]
for i, (a, b) in enumerate(promo):
    x = 7.05 + (i % 2) * 2.8; y = 4.45 + (i // 2) * 0.78
    rect(s, x, y, 2.7, 0.68, WHITE, line=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tb(s, x + 0.1, y + 0.03, 2.5, 0.32, a, size=13, color=WINE, bold=True)
    tb(s, x + 0.1, y + 0.34, 2.5, 0.3, b, size=10.5, color=GREY)
rect(s, 0.78, 6.12, 11.77, 0.8, CREAM, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tb(s, 0.95, 6.15, 11.5, 0.74, [[("Қорытынды: ", {"bold": True, "color": WINE}),
   ("Magnum онлайн-азық-түлік платформаларының үштігіне кіреді (Kaspi, Яндекс, Magnum). Омниарна оның басты бәсекелік артықшылығына айналуда.", {})]],
   size=12, anchor=MSO_ANCHOR.MIDDLE)
source(s, "magnum.kz/magnumgo; mama.magnum.kz; ҚР Сауда министрлігі (Ortcom.kz), 2025; Inform.kz.")

# ---------------- 8. PEOPLE & PROCESS ----------------
s = slides[7]
header(s, "7P", "People және Process", "07 · 7P маркетинг кешені · 5–6",
       "Қызметкерлер, жеткізушілер және тұтынушының сатып алу жолы")
card(s, 0.78, 1.85, 2.85, 1.4, "13,5 мың", "қызметкер (2021 ж.)", bigsize=28)
card(s, 3.76, 1.85, 2.85, 1.4, "1 100+", "жеткізуші (2020 ж.)", bigsize=28)
tb(s, 0.78, 3.45, 5.85, 0.35, "PEOPLE", size=13, color=RED, bold=True)
tb(s, 0.78, 3.8, 5.85, 2.2, [
    "• Кассирлер мен консультанттар – тұтынушымен байланыстың негізгі нүктесі",
    "• Сервис стандарттары: сыпайылық, кезек уақыты, тауардың сөреде болуы",
    "• Тұтынушымен кері байланыс: қосымша, колл-орталық, әлеуметтік желілер",
    "• Тәуекел: кадр тұрақсыздығы сервис сапасына әсер етеді",
], size=13.5, spacing=8)
tb(s, 6.95, 1.85, 5.6, 0.35, "PROCESS · ТҰТЫНУШЫ ЖОЛЫ", size=13, color=RED, bold=True)
jr = [("1", "Таңдау", "қосымша, каталог"), ("2", "Тапсырыс", "дүкен немесе Magnum Go"),
      ("3", "Жинау", "dark store, сөре"), ("4", "Төлем", "Kaspi Pay, касса"), ("5", "Бонус", "Magnum Club")]
for i, (n, a, b) in enumerate(jr):
    y = 2.3 + i * 0.75
    rect(s, 6.95, y, 0.58, 0.58, RED if i % 2 == 0 else WINE, shape=MSO_SHAPE.OVAL)
    tb(s, 6.95, y + 0.05, 0.58, 0.48, n, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb(s, 7.7, y - 0.02, 2.0, 0.35, a, size=15, color=WINE, bold=True)
    tb(s, 7.7, y + 0.3, 4.8, 0.3, b, size=11.5, color=GREY)
    if i < 4:
        c = s.shapes.add_connector(1, Inches(7.24), Inches(y + 0.58), Inches(7.24), Inches(y + 0.75))
        c.line.color.rgb = LINE; c.line.width = Pt(2)
rect(s, 0.78, 6.12, 11.77, 0.8, CREAM, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tb(s, 0.95, 6.15, 11.5, 0.74, [[("Қорытынды: ", {"bold": True, "color": WINE}),
   ("процестің цифрлануы ыңғайлылықты арттырады, бірақ дүкендегі сервис сапасы әлі тәуекел аймағы.", {})]],
   size=12, anchor=MSO_ANCHOR.MIDDLE)
source(s, "Spot.uz, 2020; Wikipedia / Wikiwand (2021 дерегі); magnum.kz. Жаңа деректер нақтыланады.")

# ---------------- 9. PHYSICAL EVIDENCE & RECOMMENDATIONS ----------------
s = slides[8]
header(s, "ҰСЫНЫС", "Physical Evidence және ұсыныстар", "08 · 7P · 7 және стратегиялық ұсыныстар",
       "Ұсыныстар PEST және SWOT талдауларынан туындайды")
pe = ["Дүкен дизайны", "Фирмалық түстер", "Навигация, сөре", "Қосымша интерфейсі", "Пакет, чек"]
tb(s, 0.78, 1.72, 2.6, 0.35, "PHYSICAL EVIDENCE:", size=12, color=RED, bold=True)
for i, p in enumerate(pe):
    rect(s, 3.0 + i * 1.93, 1.72, 1.83, 0.4, CREAM, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tb(s, 3.0 + i * 1.93, 1.74, 1.83, 0.36, p, size=11, color=WINE, bold=True, align=PP_ALIGN.CENTER,
       anchor=MSO_ANCHOR.MIDDLE)
table(s, 0.78, 2.35, 11.77, 4.55, [
    ["№", "Мәселе (дереккөзі)", "Ұсынылатын шешім", "Күтілетін нәтиже", "KPI"],
    [[("1", {"bold": True, "size": 20, "color": RED})], "Нақты табыс −3,4%, инфляция 12,3% (PEST-E, SWOT-T)",
     "Өзіндік брендтерді кеңейтіп, «қолжетімді себет» желісін ашу", "Бағаға сезімтал тұтынушыны ұстап қалу",
     "Өзіндік бренд үлесі, %; баға индексі"],
    [[("2", {"bold": True, "size": 20, "color": RED})], "Дүкен саны −13, гипермаркетке келушілер азаюда (SWOT-W, PEST-S)",
     "Daily/Express форматтарын тұрғын аудандарда дамытып, dark store-лармен байланыстыру",
     "Күнделікті сатып алуларды қайтару", "1 м² сатылым; шағын формат үлесі"],
    [[("3", {"bold": True, "size": 20, "color": RED})], "Онлайн бәсеке: Glovo, Wolt, Arbuz (PEST-T, SWOT-T)",
     "Magnum Club деректері негізінде жеке ұсыныстар", "Сатып алу жиілігі мен адалдықтың өсуі",
     "Белсенді мүше үлесі; retention; орташа чек"],
], [0.55, 3.1, 3.4, 2.4, 2.32], size=13, row_h=[0.45, 1.35, 1.35, 1.35])
source(s, "Авторлардың талдауы: PEST (4-слайд), SWOT (5-слайд) негізінде.")

# ---------------- 10. CONCLUSION ----------------
s = slides[9]
header(s, "ҚОРЫТЫНДЫ", "Қорытынды және дереккөздер", "09 · Қорытынды", None)
tb(s, 0.78, 1.8, 6.2, 0.35, "НЕГІЗГІ ТҰЖЫРЫМДАР", size=13, color=RED, bold=True)
concl = [("01", "Magnum – Қазақстандағы FMCG бөлшек саудасының көшбасшысы, бірақ гипермаркет моделінен шағын форматқа ауысу кезеңінде тұр."),
         ("02", "Негізгі сыртқы қауіптер – инфляция, ҚҚС-тың өсуі және үстемені шектеу."),
         ("03", "Ең күшті мүмкіндік – Kaspi экожүйесі арқылы омниарна мен цифрландыру."),
         ("04", "Ұсынылған үш шара бағаны, форматты және адалдықты KPI арқылы өлшеуге мүмкіндік береді.")]
for i, (n, t) in enumerate(concl):
    y = 2.25 + i * 1.12
    tb(s, 0.78, y - 0.1, 1.0, 0.8, n, size=34, color=RED, bold=True)
    tb(s, 1.75, y, 5.2, 1.05, t, size=12.5, color=INK)
rect(s, 7.25, 1.8, 5.3, 5.1, CREAM, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
tb(s, 7.45, 1.92, 5.0, 0.35, "ПАЙДАЛАНЫЛҒАН ӘДЕБИЕТТЕР", size=13, color=RED, bold=True)
tb(s, 7.45, 2.3, 4.95, 4.5, [
    "1. ҚР Ұлттық статистика бюросы. Ішкі сауда туралы, 2025–2026. stat.gov.kz",
    "2. ҚР Ұлттық банкі. Базалық мөлшерлеме туралы баспасөз релиздері, 2026. nationalbank.kz",
    "3. Forbes Kazakhstan. 75 крупнейших частных компаний Казахстана – 2025. forbes.kz",
    "4. Magnum Cash&Carry. Компания туралы. magnum.kz/about",
    "5. Kaspi.kz. 3Q & 9M 2026 Financial Results. ir.kaspi.kz",
    "6. ҚР Салық кодексі (№ 214-VIII), 2026.",
    "7. INFOLine. Қазақстанның ірі сауда желілерінің рейтингі, 2025.",
    "8. ҚР Сауда және интеграция министрлігі. Электрондық сауда дамуы, 2025.",
    "9. Kapital.kz. Kaspi.kz запустит до 300 магазинов у дома, 2026.",
    "10. Kursiv.media. Розничная торговля и реальные доходы, 2025–2026.",
    "11. «Маркетинг негіздері» дәрістері: 6, 7, 10. Нархоз университеті.",
], size=11.5, spacing=5)
prs.save(OUT)
print("saved", OUT)
