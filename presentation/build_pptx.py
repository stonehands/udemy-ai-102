"""
AIプロジェクトのリアル — 15スライド PowerPoint 生成スクリプト
BayCurrent / McKinsey / BCG スタイル
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.dml import MSO_THEME_COLOR
import copy
from lxml import etree

# ── Canvas: 16:9 widescreen ──────────────────────────────────────────────────
W, H = Inches(13.33), Inches(7.5)

# ── Color palette ────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x0F, 0x20, 0x44)
NAVY_MID   = RGBColor(0x1A, 0x32, 0x60)
BLUE       = RGBColor(0x15, 0x65, 0xC0)
BLUE_LIGHT = RGBColor(0x19, 0x76, 0xD2)
BLUE_PALE  = RGBColor(0xE3, 0xED, 0xF8)
TEAL       = RGBColor(0x00, 0x83, 0x8F)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE  = RGBColor(0xF7, 0xF9, 0xFC)
GRAY_100   = RGBColor(0xF0, 0xF4, 0xF8)
GRAY_200   = RGBColor(0xE2, 0xE8, 0xF0)
GRAY_400   = RGBColor(0x94, 0xA3, 0xB8)
GRAY_600   = RGBColor(0x47, 0x55, 0x69)
GRAY_800   = RGBColor(0x1E, 0x29, 0x3B)
RED        = RGBColor(0xC6, 0x28, 0x28)
ORANGE     = RGBColor(0xF5, 0x7C, 0x00)

# ── Helper utilities ─────────────────────────────────────────────────────────

def rgb(r, g, b): return RGBColor(r, g, b)

def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        if line_w:
            shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, x, y, w, h, text, size, bold=False, color=GRAY_800,
                align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Arial"
    return txb

def add_label(slide, x, y, text, color=BLUE):
    """Small ALL-CAPS spaced label"""
    add_textbox(slide, x, y, Inches(6), Pt(14),
                text.upper(), 9, bold=True, color=color)

def add_title(slide, x, y, text, w=None, color=NAVY):
    w = w or Inches(10)
    txb = slide.shapes.add_textbox(x, y, w, Inches(0.9))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = "Arial"
    return txb

def add_body(slide, x, y, w, h, text, size=13, color=GRAY_600, align=PP_ALIGN.LEFT, bold=False):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Arial"
    return txb

def add_multiline(slide, x, y, w, lines, size=13, color=GRAY_800,
                  bold_first=False, indent="  "):
    """lines: list of str. First line optionally bold."""
    txb = slide.shapes.add_textbox(x, y, w, Inches(3))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(2)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold_first and (line == lines[0])
        run.font.color.rgb = color
        run.font.name = "Arial"
    return txb

def top_rule(slide, color=NAVY, h=Pt(5)):
    add_rect(slide, 0, 0, W, h, fill=color)

def footer(slide, slide_num, total=15):
    """Navy footer bar"""
    fy = H - Inches(0.4)
    fh = Inches(0.4)
    add_rect(slide, 0, fy, W, fh, fill=NAVY)
    # Brand
    add_textbox(slide, Inches(0.6), fy + Pt(6), Inches(5), Pt(20),
                "BAYCURRENT CONSULTING", 8, bold=True,
                color=RGBColor(0xAA, 0xBB, 0xCC))
    # Page num
    add_textbox(slide, W - Inches(1.5), fy + Pt(6), Inches(1.2), Pt(20),
                f"{slide_num:02d} / {total}", 8, color=RGBColor(0x88, 0x99, 0xAA),
                align=PP_ALIGN.RIGHT)

def card(slide, x, y, w, h, title, body, top_color=None, bg=None):
    bg_c = bg or WHITE
    rect = add_rect(slide, x, y, w, h, fill=bg_c,
                    line=GRAY_200, line_w=Pt(1))
    if top_color:
        add_rect(slide, x, y, w, Pt(4), fill=top_color)
    add_textbox(slide, x + Inches(0.2), y + Inches(0.18), w - Inches(0.4), Pt(18),
                title, 11, bold=True, color=NAVY)
    add_textbox(slide, x + Inches(0.2), y + Inches(0.5), w - Inches(0.4),
                h - Inches(0.6), body, 11, color=GRAY_600)

def bullet_list(slide, x, y, items, size=13, color=GRAY_800, accent=BLUE):
    txb = slide.shapes.add_textbox(x, y, Inches(5), Inches(4))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(3)
        # arrow
        r1 = p.add_run()
        r1.text = "→ "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = accent
        r1.font.name = "Arial"
        r2 = p.add_run()
        r2.text = item
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
        r2.font.name = "Arial"
    return txb

def hl_box(slide, x, y, w, h, label, text, bg=NAVY, fg=WHITE, label_c=None):
    label_c = label_c or RGBColor(0xAA, 0xBB, 0xCC)
    add_rect(slide, x, y, w, h, fill=bg)
    add_textbox(slide, x + Inches(0.25), y + Inches(0.15), w - Inches(0.5), Pt(14),
                label.upper(), 8, bold=True, color=label_c)
    add_textbox(slide, x + Inches(0.25), y + Inches(0.42), w - Inches(0.5),
                h - Inches(0.5), text, 16, bold=True, color=fg)

def info_box(slide, x, y, w, h, label, text):
    add_rect(slide, x, y, w, h, fill=BLUE_PALE, line=BLUE_LIGHT, line_w=Pt(1))
    add_textbox(slide, x + Inches(0.2), y + Inches(0.12), w - Inches(0.4), Pt(14),
                label.upper(), 8, bold=True, color=BLUE)
    add_textbox(slide, x + Inches(0.2), y + Inches(0.38), w - Inches(0.4),
                h - Inches(0.5), text, 13, bold=True, color=NAVY)

def warn_box(slide, x, y, w, h, label, text):
    add_rect(slide, x, y, w, h, fill=RGBColor(0xFF, 0xF3, 0xE0),
             line=RGBColor(0xFF, 0xB7, 0x4D), line_w=Pt(1))
    add_rect(slide, x, y, Pt(5), h, fill=ORANGE)
    add_textbox(slide, x + Inches(0.2), y + Inches(0.12), w - Inches(0.4), Pt(14),
                label.upper(), 8, bold=True, color=ORANGE)
    add_textbox(slide, x + Inches(0.2), y + Inches(0.38), w - Inches(0.4),
                h - Inches(0.5), text, 13, color=RGBColor(0x5D, 0x40, 0x37))

def big_msg(slide, x, y, w, text):
    add_rect(slide, x, y, w, Inches(0.68),
             fill=BLUE_PALE, line=None)
    add_rect(slide, x, y, Pt(5), Inches(0.68), fill=BLUE)
    add_textbox(slide, x + Inches(0.2), y + Inches(0.1), w - Inches(0.3),
                Inches(0.55), text, 16, bold=True, color=NAVY)

# ── Build presentation ───────────────────────────────────────────────────────

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank_layout = prs.slide_layouts[6]  # completely blank

MX = Inches(0.7)   # left margin
MY = Inches(0.6)   # top margin (after rule)
CW = W - MX * 2   # content width

# ============================================================
# SLIDE 01 — TITLE  (dark navy)
# ============================================================
sl = prs.slides.add_slide(blank_layout)

# Background navy
add_rect(sl, 0, 0, W, H, fill=NAVY)
# Top accent stripe (blue→teal gradient approximated as blue)
add_rect(sl, 0, 0, W, Pt(6), fill=BLUE)

# Geometric circle outlines (light)
for r, op in [(Inches(4), RGBColor(0x1A,0x3A,0x70)),
              (Inches(2.8), RGBColor(0x12,0x2C,0x58))]:
    cx, cy = W - Inches(1.5), Inches(1)
    circ = sl.shapes.add_shape(9, cx - r, cy - r, r*2, r*2)  # oval
    circ.fill.background()
    circ.line.color.rgb = op
    circ.line.width = Pt(1)

# Company label
add_textbox(sl, MX, MY, Inches(6), Pt(16),
            "BAYCURRENT CONSULTING — AI PROJECT", 9, bold=True,
            color=RGBColor(0x88,0x99,0xBB))

# Main title
txb = sl.shapes.add_textbox(MX, MY + Inches(0.5), Inches(8), Inches(2.5))
txb.word_wrap = True
tf = txb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = "AIプロジェクトの"
r.font.size = Pt(52)
r.font.bold = True
r.font.color.rgb = WHITE
r.font.name = "Arial"
p2 = tf.add_paragraph()
p2.alignment = PP_ALIGN.LEFT
r2 = p2.add_run()
r2.text = "リアル"
r2.font.size = Pt(52)
r2.font.bold = True
r2.font.color.rgb = RGBColor(0x64, 0xB5, 0xF6)
r2.font.name = "Arial"

# Divider line
add_rect(sl, MX, MY + Inches(2.9), Inches(0.8), Pt(3), fill=BLUE_LIGHT)

# Subtitle
add_textbox(sl, MX, MY + Inches(3.1), Inches(8), Pt(22),
            "なぜ、今AIプロジェクトのPMを目指すのか", 16,
            color=RGBColor(0xAA,0xBB,0xCC))

# Main message box
add_rect(sl, MX, MY + Inches(3.7), Inches(7.5), Inches(1.3),
         fill=RGBColor(0x1A,0x30,0x5A))
add_textbox(sl, MX + Inches(0.2), MY + Inches(3.82), Inches(7), Inches(1.1),
            "AI競争の構造が変わり、企業の勝ち筋が実装力にシフトしている", 18,
            bold=True, color=WHITE)

# Footer
add_rect(sl, 0, H - Inches(0.4), W, Inches(0.4),
         fill=RGBColor(0x06,0x12,0x28))
add_textbox(sl, MX, H - Inches(0.38), Inches(5), Pt(20),
            "CONFIDENTIAL — FOR INTERNAL USE ONLY", 8,
            color=RGBColor(0x55,0x66,0x80))
add_textbox(sl, W - Inches(1.5), H - Inches(0.38), Inches(1.2), Pt(20),
            "01 / 15", 8, color=RGBColor(0x55,0x66,0x80), align=PP_ALIGN.RIGHT)

# ============================================================
# SLIDE 02 — 世界はもう動いている
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Context")
add_title(sl, MX, MY + Inches(0.28), "世界はもう動いている")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "競争軸はモデルから実装へ", 12, color=GRAY_600)

# Left: growth graph (simplified bar chart)
gx, gy = MX, MY + Inches(1.3)
gw, gh = Inches(5.2), Inches(3.4)
add_rect(sl, gx, gy, gw, gh, fill=GRAY_100, line=GRAY_200, line_w=Pt(1))
add_textbox(sl, gx + Inches(0.15), gy + Inches(0.1), gw - Inches(0.3), Pt(16),
            "AI Agent 市場規模（成長イメージ）", 9, bold=True, color=GRAY_600)
# Simple bars
bars = [0.3, 0.45, 0.65, 0.85, 1.0]
bar_labels = ["2021","2022","2023","2024","2025E"]
bar_w = Inches(0.55)
bar_base_y = gy + gh - Inches(0.45)
bar_max_h = Inches(2.3)
for i, (ratio, label) in enumerate(zip(bars, bar_labels)):
    bx = gx + Inches(0.4) + i * Inches(0.88)
    bh_val = bar_max_h * ratio
    by = bar_base_y - bh_val
    c = BLUE if i < 4 else NAVY
    add_rect(sl, bx, by, bar_w, bh_val, fill=c)
    add_textbox(sl, bx - Inches(0.05), bar_base_y + Pt(2), bar_w + Inches(0.1), Pt(14),
                label, 8, color=GRAY_600, align=PP_ALIGN.CENTER)
add_textbox(sl, gx + Inches(0.1), bar_base_y + Inches(0.3), gw - Inches(0.2), Pt(14),
            "市場は急拡大中", 10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# Right: 4-stage evolution
rx = MX + Inches(5.6)
ry = MY + Inches(1.3)
rw = CW - Inches(5.6)
stages = [
    ("Stage 1", "LLM競争", GRAY_200, GRAY_600),
    ("Stage 2", "Copilot競争", BLUE_PALE, BLUE),
    ("Stage 3", "AIエージェント競争", BLUE_PALE, BLUE),
    ("Stage 4", "全社AI運用競争", NAVY, WHITE),
]
sh = Inches(0.68)
sy = ry
for stage_label, stage_name, bg, fg in stages:
    add_rect(sl, rx, sy, rw, sh, fill=bg, line=GRAY_200, line_w=Pt(1))
    add_textbox(sl, rx + Inches(0.12), sy + Pt(4), Inches(1.2), Pt(14),
                stage_label.upper(), 8, bold=True, color=fg if bg == NAVY else BLUE)
    add_textbox(sl, rx + Inches(1.3), sy + Pt(4), rw - Inches(1.5), sh - Pt(8),
                stage_name, 14, bold=True, color=fg if bg == NAVY else NAVY)
    sy += sh
    if sy < ry + sh * 4:
        add_textbox(sl, rx + rw / 2 - Inches(0.2), sy, Inches(0.4), Pt(14),
                    "↓", 12, color=BLUE, align=PP_ALIGN.CENTER)
        sy += Inches(0.28)

# Bottom message
big_msg(sl, MX, H - Inches(1.15), CW,
        "競争軸はモデルから実装へ　— 今まさに実装力が問われている")
footer(sl, 2)

# ============================================================
# SLIDE 03 — AIファーストの時代
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Context")
add_title(sl, MX, MY + Inches(0.28), "AIファーストの時代")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "変化に適応できる企業が勝つ", 12, color=GRAY_600)

col_w = (CW - Inches(0.3)) / 2
ly = MY + Inches(1.3)
lh = Inches(3.2)

# Left card — 日本企業の課題
add_rect(sl, MX, ly, col_w, lh, fill=GRAY_100, line=GRAY_200, line_w=Pt(1))
add_rect(sl, MX, ly, col_w, Pt(4), fill=GRAY_400)
add_textbox(sl, MX + Inches(0.2), ly + Inches(0.1), col_w - Inches(0.4), Pt(18),
            "日本企業の課題", 11, bold=True, color=GRAY_600)
bullet_list(sl, MX + Inches(0.1), ly + Inches(0.5),
            ["人材不足", "生産性向上圧力", "レガシーシステム", "デジタル赤字"],
            size=13, color=GRAY_800, accent=GRAY_400)

# Arrow
add_textbox(sl, MX + col_w + Inches(0.05), ly + lh/2 - Inches(0.25),
            Inches(0.3), Inches(0.5), "→", 24, bold=True, color=BLUE,
            align=PP_ALIGN.CENTER)

# Right card — AIファースト企業
rx2 = MX + col_w + Inches(0.35)
add_rect(sl, rx2, ly, col_w, lh, fill=BLUE_PALE, line=BLUE_LIGHT, line_w=Pt(1))
add_rect(sl, rx2, ly, col_w, Pt(4), fill=BLUE)
add_textbox(sl, rx2 + Inches(0.2), ly + Inches(0.1), col_w - Inches(0.4), Pt(18),
            "AIファースト企業", 11, bold=True, color=BLUE)
bullet_list(sl, rx2 + Inches(0.1), ly + Inches(0.5),
            ["AI前提業務", "AIエージェント活用", "意思決定高速化", "全社最適"],
            size=13, color=NAVY, accent=BLUE)

big_msg(sl, MX, H - Inches(1.15), CW, "変化に適応できる企業が勝つ")
footer(sl, 3)

# ============================================================
# SLIDE 04 — AIエージェント基盤PJの全体像
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Overview")
add_title(sl, MX, MY + Inches(0.28), "AIエージェント基盤PJの全体像")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "全社をつなぐ共通基盤として推進", 12, color=GRAY_600)

# Hub center
import math
cx_hub = W / 2
cy_hub = MY + Inches(3.0)
hub_r = Inches(0.8)

hub = sl.shapes.add_shape(9, cx_hub - hub_r, cy_hub - hub_r, hub_r*2, hub_r*2)
hub.fill.solid()
hub.fill.fore_color.rgb = NAVY
hub.line.fill.background()
add_textbox(sl, cx_hub - hub_r, cy_hub - Inches(0.25), hub_r*2, Inches(0.5),
            "AIエージェント\n基盤", 11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Spokes
nodes = ["経営", "業務部門", "IT部門", "セキュリティ", "データ基盤", "インフラ"]
node_r = Inches(1.8)
node_colors = [NAVY, BLUE, BLUE, TEAL, TEAL, GRAY_600]
for i, (node, nc) in enumerate(zip(nodes, node_colors)):
    angle = math.radians(i * 60 - 90)
    nx = cx_hub + node_r * math.cos(angle)
    ny = cy_hub + node_r * math.sin(angle)
    nw, nh = Inches(1.3), Inches(0.6)

    # Line from hub to node
    line_shape = sl.shapes.add_connector(1,
        cx_hub, cy_hub, nx, ny)
    line_shape.line.color.rgb = GRAY_200
    line_shape.line.width = Pt(1.5)

    add_rect(sl, nx - nw/2, ny - nh/2, nw, nh, fill=nc)
    add_textbox(sl, nx - nw/2, ny - nh/2 + Inches(0.07), nw, nh,
                node, 12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

big_msg(sl, MX, H - Inches(1.15), CW,
        "全社をつなぐ共通基盤として推進 — PMが全体を統括する")
footer(sl, 4)

# ============================================================
# SLIDE 05 — 意思決定を動かしたプロセス
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Process")
add_title(sl, MX, MY + Inches(0.28), "意思決定を動かしたプロセス")

steps = ["経営課題の\n可視化", "価値シナリオ\n設計", "PoCでの\n効果検証", "投資判断"]
sw = (CW - Inches(1.2)) / 4
sy = MY + Inches(1.3)
sh_box = Inches(1.2)
sx = MX
for i, step in enumerate(steps):
    fc = NAVY if i == len(steps)-1 else BLUE_PALE
    tc = WHITE if i == len(steps)-1 else NAVY
    bc = NAVY if i == len(steps)-1 else BLUE
    add_rect(sl, sx, sy, sw, sh_box, fill=fc, line=bc, line_w=Pt(1.5))
    add_textbox(sl, sx + Inches(0.05), sy + Inches(0.05), sw - Inches(0.1),
                sh_box - Inches(0.1), f"0{i+1}  " + step, 13, bold=True,
                color=tc, align=PP_ALIGN.CENTER)
    sx += sw
    if i < len(steps) - 1:
        add_textbox(sl, sx, sy + Inches(0.35), Inches(0.3), Inches(0.5),
                    "→", 18, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
        sx += Inches(0.3)

# Bottom points
pts = [
    "経営課題に直結する価値設計",
    "小さく早く成功体験を創出",
    "リスクと対策をセットで提示",
    "中長期の競争優位につながる投資と説明",
]
add_textbox(sl, MX, MY + Inches(2.9), CW, Pt(18),
            "経営を動かしたポイント", 12, bold=True, color=NAVY)
bullet_list(sl, MX, MY + Inches(3.2), pts, size=13, color=GRAY_800)

big_msg(sl, MX, H - Inches(1.15), CW,
        "経営を動かす提案には、価値・リスク・競争優位の3点セットが必要")
footer(sl, 5)

# ============================================================
# SLIDE 06 — 多様な組織を巻き込む推進体制
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Organization")
add_title(sl, MX, MY + Inches(0.28), "多様な組織を巻き込む推進体制")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "組織横断でつなぎ、前に進める", 12, color=GRAY_600)

cx2 = W / 2
cy2 = MY + Inches(3.1)
hub_r2 = Inches(0.75)

hub2 = sl.shapes.add_shape(9, cx2 - hub_r2, cy2 - hub_r2, hub_r2*2, hub_r2*2)
hub2.fill.solid()
hub2.fill.fore_color.rgb = NAVY
hub2.line.fill.background()
add_textbox(sl, cx2 - hub_r2, cy2 - Inches(0.2), hub_r2*2, Inches(0.4),
            "AI PM", 14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

node_names = ["役員", "経営企画", "業務部門", "システム部門",
              "基盤開発", "技術検証", "インフラ", "セキュリティ", "データチーム"]
node_r2 = Inches(2.1)
node_col = [NAVY]*3 + [BLUE]*3 + [TEAL]*3

for i, (nn, nc) in enumerate(zip(node_names, node_col)):
    angle = math.radians(i * 40 - 90)
    nx = cx2 + node_r2 * math.cos(angle)
    ny = cy2 + node_r2 * math.sin(angle)
    nw2, nh2 = Inches(1.15), Inches(0.5)

    line_shape2 = sl.shapes.add_connector(1, cx2, cy2, nx, ny)
    line_shape2.line.color.rgb = GRAY_200
    line_shape2.line.width = Pt(1)

    add_rect(sl, nx - nw2/2, ny - nh2/2, nw2, nh2, fill=nc)
    add_textbox(sl, nx - nw2/2, ny - nh2/2 + Inches(0.05), nw2, nh2,
                nn, 11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

big_msg(sl, MX, H - Inches(1.15), CW,
        "PMは組織の境界を越えて合意を形成し、プロジェクトを前進させる")
footer(sl, 6)

# ============================================================
# SLIDE 07 — AIが変えた仕事のスピードと質
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Episode 1")
add_title(sl, MX, MY + Inches(0.28), "AIが変えた仕事のスピードと質")

col_w2 = (CW - Inches(0.4)) / 2
ly2 = MY + Inches(1.1)

for side, title, steps_list, bc, tc in [
    ("left",  "Before", ["考える", "持ち帰る", "作る", "レビュー"],
     GRAY_100, GRAY_600),
    ("right", "After",  ["考える", "その場で試作", "その場で議論", "その場で修正"],
     BLUE_PALE, BLUE),
]:
    sx2 = MX if side == "left" else MX + col_w2 + Inches(0.4)
    add_rect(sl, sx2, ly2, col_w2, Inches(3.8),
             fill=bc, line=GRAY_200, line_w=Pt(1))
    tc2 = GRAY_400 if side == "left" else BLUE
    add_textbox(sl, sx2 + Inches(0.2), ly2 + Inches(0.1), col_w2, Pt(18),
                title, 11, bold=True, color=tc2)
    item_y = ly2 + Inches(0.5)
    for item in steps_list:
        c_item = GRAY_600 if side == "left" else NAVY
        add_textbox(sl, sx2 + Inches(0.2), item_y, col_w2 - Inches(0.4),
                    Inches(0.55), item, 14, bold=(side == "right"), color=c_item)
        item_y += Inches(0.65)
        if item != steps_list[-1]:
            add_textbox(sl, sx2 + col_w2/2 - Inches(0.2), item_y - Inches(0.3),
                        Inches(0.4), Inches(0.3), "↓", 12, color=tc, align=PP_ALIGN.CENTER)

add_textbox(sl, MX + col_w2 + Inches(0.05), ly2 + Inches(1.5),
            Inches(0.4), Inches(0.5), "→", 22, bold=True, color=NAVY,
            align=PP_ALIGN.CENTER)

big_msg(sl, MX, H - Inches(1.15), CW, "AIは思考速度を変えた")
footer(sl, 7)

# ============================================================
# SLIDE 08 — AIは業務に組み込まれる
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Technology")
add_title(sl, MX, MY + Inches(0.28), "AIは業務に組み込まれる")

flow_steps = ["ChatGPT", "Copilot", "AI Agent", "Digital\nWorkforce"]
fw = (Inches(5.5) - Inches(0.6)) / 4
fy_fl = MY + Inches(1.2)
fh_fl = Inches(0.85)
fx = MX
for i, fs in enumerate(flow_steps):
    fc = NAVY if i == 3 else (BLUE if i == 2 else BLUE_PALE)
    tc = WHITE if i >= 2 else NAVY
    add_rect(sl, fx, fy_fl, fw, fh_fl, fill=fc, line=BLUE_LIGHT, line_w=Pt(1))
    add_textbox(sl, fx + Inches(0.05), fy_fl + Inches(0.1), fw - Inches(0.1),
                fh_fl - Inches(0.2), fs, 12, bold=True, color=tc, align=PP_ALIGN.CENTER)
    fx += fw
    if i < 3:
        add_textbox(sl, fx, fy_fl + Inches(0.25), Inches(0.15), Inches(0.4),
                    "→", 12, bold=True, color=BLUE)
        fx += Inches(0.15)

# Tech layer right
tx = MX + Inches(6.0)
add_textbox(sl, tx, MY + Inches(1.0), CW - Inches(6.0), Pt(16),
            "技術連携レイヤー", 11, bold=True, color=NAVY)
tech = ["API連携", "MCP (Model Context Protocol)", "RPA連携",
        "Event連携", "Agent Identity", "OBO認証"]
bullet_list(sl, tx, MY + Inches(1.4), tech, size=12, color=GRAY_800, accent=BLUE)

# Below flow
add_textbox(sl, MX, fy_fl + fh_fl + Inches(0.25), Inches(6), Pt(16),
            "業務実装の深度が競争優位を決める", 12, bold=True, color=NAVY)
add_textbox(sl, MX, fy_fl + fh_fl + Inches(0.6), Inches(6), Inches(1.5),
            "単なるツール導入ではなく、業務プロセスへの深い統合が\n真の競争力を生む。PMはこの実装深度を設計する。",
            13, color=GRAY_600)

big_msg(sl, MX, H - Inches(1.15), CW, "競争軸は業務実装 — 深く組み込んだ企業が勝つ")
footer(sl, 8)

# ============================================================
# SLIDE 09 — 経営価値を創出する
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Challenge 01")
add_title(sl, MX, MY + Inches(0.28), "経営価値を創出する")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "効率化だけでは数百億の投資は動かない", 12, color=GRAY_600)

col_w3 = (CW - Inches(1.0)) / 2
ly3 = MY + Inches(1.3)
lh3 = Inches(3.0)

add_rect(sl, MX, ly3, col_w3, lh3, fill=GRAY_100, line=GRAY_200, line_w=Pt(1))
add_rect(sl, MX, ly3, col_w3, Pt(4), fill=GRAY_400)
add_textbox(sl, MX + Inches(0.2), ly3 + Inches(0.1), col_w3, Pt(16),
            "現場視点", 10, bold=True, color=GRAY_400)
add_textbox(sl, MX + Inches(0.2), ly3 + Inches(0.42), col_w3 - Inches(0.4), Pt(24),
            "効率化", 28, bold=True, color=GRAY_600, align=PP_ALIGN.CENTER)
bullet_list(sl, MX + Inches(0.1), ly3 + Inches(1.2),
            ["仕事が楽になる", "ミスが減る", "時間短縮"],
            size=12, color=GRAY_600, accent=GRAY_400)

add_textbox(sl, MX + col_w3 + Inches(0.15), ly3 + lh3/2 - Inches(0.3),
            Inches(0.7), Inches(0.6), "≠", 28, bold=True, color=GRAY_400,
            align=PP_ALIGN.CENTER)

rx3 = MX + col_w3 + Inches(0.85)
add_rect(sl, rx3, ly3, col_w3, lh3, fill=BLUE_PALE, line=BLUE, line_w=Pt(1.5))
add_rect(sl, rx3, ly3, col_w3, Pt(4), fill=NAVY)
add_textbox(sl, rx3 + Inches(0.2), ly3 + Inches(0.1), col_w3, Pt(16),
            "経営視点", 10, bold=True, color=NAVY)
add_textbox(sl, rx3 + Inches(0.2), ly3 + Inches(0.42), col_w3 - Inches(0.4), Pt(24),
            "競争力・P/L", 22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
bullet_list(sl, rx3 + Inches(0.1), ly3 + Inches(1.2),
            ["売上向上", "利益拡大", "競争力強化"],
            size=12, color=NAVY, accent=BLUE)

warn_box(sl, MX, H - Inches(1.5), CW, Inches(0.65),
         "落とし穴",
         "「楽になります」だけでは数百億は動かない。経営言語への翻訳が必須。")
footer(sl, 9)

# ============================================================
# SLIDE 10 — システム接続の壁
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Challenge 02")
add_title(sl, MX, MY + Inches(0.28), "システム接続の壁")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "本当に難しいのは接続", 12, color=GRAY_600)

cols_data = [
    ("理想", "API 連携", BLUE_PALE, BLUE, ["APIでシームレス連携", "クリーンな実装", "エージェントが自律動作"]),
    ("現実", "壁が存在する", RGBColor(0xFF,0xF3,0xE0), ORANGE, ["APIが存在しない", "レガシーシステム", "接続方式がバラバラ", "データ所在不明"]),
    ("打ち手", "解決策", BLUE_PALE, NAVY, ["API新規構築", "MCP活用", "RPA連携", "ETL/データ連携"]),
]
cw3 = (CW - Inches(0.4)) / 3
cx3 = MX
ly4 = MY + Inches(1.2)
ch4 = Inches(3.5)

for title_c, sub_c, bg_c, ac_c, items_c in cols_data:
    add_rect(sl, cx3, ly4, cw3, ch4, fill=bg_c, line=ac_c, line_w=Pt(1.5))
    add_textbox(sl, cx3 + Inches(0.15), ly4 + Inches(0.1), cw3 - Inches(0.3), Pt(16),
                title_c, 10, bold=True, color=ac_c)
    add_textbox(sl, cx3 + Inches(0.15), ly4 + Inches(0.4), cw3 - Inches(0.3), Pt(18),
                sub_c, 13, bold=True, color=NAVY if bg_c != RGBColor(0xFF,0xF3,0xE0) else ORANGE)
    item_y4 = ly4 + Inches(0.9)
    for it in items_c:
        add_textbox(sl, cx3 + Inches(0.15), item_y4, cw3 - Inches(0.3), Inches(0.4),
                    "• " + it, 12, color=GRAY_800)
        item_y4 += Inches(0.45)
    cx3 += cw3 + Inches(0.2)

big_msg(sl, MX, H - Inches(1.15), CW,
        "PoC で動いても、本番に繋がらない — 接続設計こそがPMの真価")
footer(sl, 10)

# ============================================================
# SLIDE 11 — ネットワーク・認証・セキュリティの壁
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Challenge 03")
add_title(sl, MX, MY + Inches(0.28), "ネットワーク・認証・セキュリティの壁")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "AIエージェントは人の代わりに動く", 12, color=GRAY_600)

col_w4 = (CW - Inches(0.3)) / 2
ly5 = MY + Inches(1.2)
lh5 = Inches(3.2)

add_rect(sl, MX, ly5, col_w4, lh5, fill=BLUE_PALE, line=BLUE, line_w=Pt(1))
add_rect(sl, MX, ly5, col_w4, Pt(4), fill=BLUE)
add_textbox(sl, MX + Inches(0.2), ly5 + Inches(0.1), col_w4, Pt(16),
            "技術領域", 10, bold=True, color=BLUE)
bullet_list(sl, MX + Inches(0.1), ly5 + Inches(0.5),
            ["OA網 / 業務網 / 閉域網", "OAuth 2.0 / PKCE",
             "On-Behalf-Of (OBO)", "Microsoft Entra ID", "Agent Identity管理"],
            size=12, color=NAVY, accent=BLUE)

rx4 = MX + col_w4 + Inches(0.35)
add_rect(sl, rx4, ly5, col_w4, lh5, fill=GRAY_100, line=GRAY_200, line_w=Pt(1))
add_rect(sl, rx4, ly5, col_w4, Pt(4), fill=NAVY)
add_textbox(sl, rx4 + Inches(0.2), ly5 + Inches(0.1), col_w4, Pt(16),
            "ガバナンス領域", 10, bold=True, color=NAVY)
bullet_list(sl, rx4 + Inches(0.1), ly5 + Inches(0.5),
            ["権限管理・最小権限原則", "監査ログ・トレーサビリティ",
             "責任分界の設計", "エージェントの説明責任"],
            size=12, color=GRAY_800, accent=NAVY)

warn_box(sl, MX, H - Inches(1.5), CW, Inches(0.65),
         "根本的な難しさ",
         "人間は「自分でログインする」。エージェントは「権限を委任される」。この違いが設計を極めて複雑にする。")
footer(sl, 11)

# ============================================================
# SLIDE 12 — LLMコストと全社展開
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Challenge 04")
add_title(sl, MX, MY + Inches(0.28), "LLMコストと全社展開")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "コスト管理なくして全社展開はできない", 12, color=GRAY_600)

# Center stat
add_rect(sl, MX, MY + Inches(1.2), CW, Inches(0.75), fill=NAVY)
add_textbox(sl, MX, MY + Inches(1.32), CW, Inches(0.6),
            "数万人規模の全社利用", 20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

col_w5 = (CW - Inches(0.3)) / 2
ly6 = MY + Inches(2.1)
lh6 = Inches(2.6)

# Left costs
add_rect(sl, MX, ly6, col_w5, lh6, fill=GRAY_100, line=GRAY_200, line_w=Pt(1))
add_textbox(sl, MX + Inches(0.15), ly6 + Inches(0.1), col_w5, Pt(16),
            "コスト要因の積み上げ", 11, bold=True, color=NAVY)
cost_items = ["推論コスト（トークン単価）", "RAG検索コスト",
              "音声認識・合成", "画像・動画解析", "マルチエージェント並列実行"]
bullet_list(sl, MX + Inches(0.05), ly6 + Inches(0.5),
            cost_items, size=12, color=GRAY_800, accent=RED)

# Right optimization
rx5 = MX + col_w5 + Inches(0.35)
add_rect(sl, rx5, ly6, col_w5, lh6, fill=BLUE_PALE, line=BLUE, line_w=Pt(1))
add_textbox(sl, rx5 + Inches(0.15), ly6 + Inches(0.1), col_w5, Pt(16),
            "最適化戦略", 11, bold=True, color=NAVY)
opt_items = ["ROI最大化設計", "TCO最適化", "モデル戦略（適材適所）",
             "キャッシュ・圧縮活用", "コスト可視化ダッシュボード"]
bullet_list(sl, rx5 + Inches(0.05), ly6 + Inches(0.5),
            opt_items, size=12, color=NAVY, accent=BLUE)

big_msg(sl, MX, H - Inches(1.15), CW,
        "コストは将来下がる。今蓄積するノウハウこそが競争資産。")
footer(sl, 12)

# ============================================================
# SLIDE 13 — AI PMに求められる力
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "PM Competency")
add_title(sl, MX, MY + Inches(0.28), "AI PMに求められる力")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "技術 × 業務 × 人 をつなぐ総合力が必要", 12, color=GRAY_600)

# Hexagon-like using 6 rectangles arranged in circle
cx_hex = W / 2
cy_hex = MY + Inches(3.2)
hex_r = Inches(2.2)
hex_items = [
    ("AI", "AI技術理解 / 活用判断 / ガバナンス", NAVY),
    ("業務", "要件定義 / 業務改革 / 合意形成", BLUE),
    ("UX", "体験設計 / 利用定着 / プロトタイピング", TEAL),
    ("IT", "アーキテクチャ / 連携基盤 / 技術選定", BLUE),
    ("アジャイル", "高速検証 / MVP思考 / 改善サイクル", NAVY),
    ("ガバナンス", "リスク管理 / コンプライアンス / 責任分界", GRAY_600),
]

for i, (abbr, desc, nc) in enumerate(hex_items):
    angle = math.radians(i * 60 - 90)
    nx = cx_hex + hex_r * math.cos(angle)
    ny = cy_hex + hex_r * math.sin(angle)
    hw, hh = Inches(1.6), Inches(0.9)
    add_rect(sl, nx - hw/2, ny - hh/2, hw, hh, fill=nc)
    add_textbox(sl, nx - hw/2 + Inches(0.05), ny - hh/2 + Inches(0.04),
                hw - Inches(0.1), Pt(16), abbr, 11, bold=True,
                color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(sl, nx - hw/2 + Inches(0.05), ny - hh/2 + Inches(0.35),
                hw - Inches(0.1), Pt(26), desc, 8,
                color=RGBColor(0xCC,0xDD,0xEE) if nc == NAVY else RGBColor(0xDD,0xEE,0xFF),
                align=PP_ALIGN.CENTER)

# Center hub
ch_r = Inches(0.72)
ch_hub = sl.shapes.add_shape(9, cx_hex - ch_r, cy_hex - ch_r, ch_r*2, ch_r*2)
ch_hub.fill.solid()
ch_hub.fill.fore_color.rgb = WHITE
ch_hub.line.color.rgb = NAVY
ch_hub.line.width = Pt(2)
add_textbox(sl, cx_hex - ch_r, cy_hex - Inches(0.22), ch_r*2, Inches(0.44),
            "AI PM", 14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

big_msg(sl, MX, H - Inches(1.15), CW,
        "AI PMは6つの力を統合するオーケストレーター")
footer(sl, 13)

# ============================================================
# SLIDE 14 — この仕事の魅力
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl)
add_label(sl, MX, MY, "Value Proposition")
add_title(sl, MX, MY + Inches(0.28), "この仕事の魅力")
add_textbox(sl, MX, MY + Inches(0.85), CW, Pt(16),
            "圧倒的に成長できる環境がここにある", 12, color=GRAY_600)

col3_data = [
    ("幅広い学び", NAVY, ["AI", "業務改革", "UX", "クラウド", "経営"]),
    ("世界の最前線", BLUE, ["AI Agent", "MCP", "OBO", "Agent Identity", "AIガバナンス"]),
    ("実践の場", TEAL, ["Foundry", "Dify", "Copilot", "最新LLM", "実案件"]),
]
cw3b = (CW - Inches(0.4)) / 3
cx3b = MX
lh3b = Inches(3.6)
ly3b = MY + Inches(1.2)
for col_title, col_c, col_items in col3_data:
    add_rect(sl, cx3b, ly3b, cw3b, lh3b, fill=GRAY_100, line=col_c, line_w=Pt(1.5))
    add_rect(sl, cx3b, ly3b, cw3b, Inches(0.8), fill=col_c)
    add_textbox(sl, cx3b + Inches(0.1), ly3b + Inches(0.15), cw3b - Inches(0.2),
                Inches(0.55), col_title, 15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    iy = ly3b + Inches(1.0)
    for ci in col_items:
        add_textbox(sl, cx3b + Inches(0.15), iy, cw3b - Inches(0.3),
                    Inches(0.44), "→ " + ci, 13, color=NAVY)
        iy += Inches(0.46)
    cx3b += cw3b + Inches(0.2)

big_msg(sl, MX, H - Inches(1.15), CW,
        "圧倒的に成長できる環境がここにある — 最前線で学び続けよう")
footer(sl, 14)

# ============================================================
# SLIDE 15 — CLOSING
# ============================================================
sl = prs.slides.add_slide(blank_layout)
top_rule(sl, color=BLUE)
add_rect(sl, 0, Pt(5), W, H - Pt(5) - Inches(0.4), fill=OFF_WHITE)

add_label(sl, MX, MY, "Closing")
add_title(sl, MX, MY + Inches(0.35), "私の答え、そしてあなたへ",
          w=CW, color=NAVY)

# Divider
add_rect(sl, MX, MY + Inches(1.3), Inches(0.6), Pt(3), fill=BLUE)

col_w6 = (CW - Inches(0.4)) / 2
ly7 = MY + Inches(1.6)
lh7 = Inches(2.8)

# Left answer box
add_rect(sl, MX, ly7, col_w6, lh7, fill=NAVY)
add_textbox(sl, MX + Inches(0.2), ly7 + Inches(0.15), col_w6 - Inches(0.4), Pt(14),
            "MY ANSWER", 9, bold=True, color=RGBColor(0x88,0x99,0xBB))
add_textbox(sl, MX + Inches(0.2), ly7 + Inches(0.5), col_w6 - Inches(0.4),
            lh7 - Inches(0.65),
            "世界で戦えるように、\n日本企業を\n応援したい", 22, bold=True, color=WHITE)

# Right question box
rx6 = MX + col_w6 + Inches(0.45)
add_rect(sl, rx6, ly7, col_w6, lh7, fill=WHITE, line=GRAY_200, line_w=Pt(1.5))
add_textbox(sl, rx6 + Inches(0.2), ly7 + Inches(0.15), col_w6 - Inches(0.4), Pt(14),
            "YOUR TURN", 9, bold=True, color=BLUE)
add_textbox(sl, rx6 + Inches(0.2), ly7 + Inches(0.5), col_w6 - Inches(0.4),
            lh7 - Inches(0.65),
            "なぜ\nAIプロジェクト PM を\n目指しますか？", 20, bold=True, color=NAVY)

# Closing quote
add_textbox(sl, MX, ly7 + lh7 + Inches(0.25), CW, Inches(0.6),
            "AIエージェントは、まだ完成された世界ではありません。だからこそ、面白い。",
            13, color=GRAY_600, italic=True)

footer(sl, 15)

# ── Save ─────────────────────────────────────────────────────────────────────
out_path = "/home/user/udemy-ai-102/presentation/ai-project-frontline.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
