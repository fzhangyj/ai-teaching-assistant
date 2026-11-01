from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "技术型助教自我介绍_tasteskill优化版.pptx"
ASSET = ROOT / "assets" / "product-preview.png"
GRID_BG = ROOT / "assets" / "ppt-grid-bg.png"

W, H = Inches(13.333), Inches(7.5)

INK = RGBColor(20, 22, 20)
MUTED = RGBColor(102, 112, 106)
BG = RGBColor(244, 245, 240)
SOFT = RGBColor(233, 236, 228)
LINE = RGBColor(204, 212, 202)
ACCENT = RGBColor(15, 125, 87)
ACCENT_DARK = RGBColor(8, 61, 44)
WHITE = RGBColor(255, 255, 255)


def set_fill(shape, color, transparency=0):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.fill.transparency = transparency


def set_line(shape, color=LINE, width=1):
    shape.line.color.rgb = color
    shape.line.width = Pt(width)


def add_text(slide, text, x, y, w, h, size=24, color=INK, bold=False, font="Aptos", align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_label(slide, text, x=0.72, y=0.58):
    add_text(slide, text.upper(), x, y, 3.2, 0.28, size=9.5, color=ACCENT, bold=True, font="Consolas")


def add_bg(slide, label):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = BG
    slide.shapes.add_picture(str(GRID_BG), 0, 0, width=W, height=H)
    mark = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.72), Inches(6.85), Inches(0.42), Inches(0.32))
    set_fill(mark, ACCENT_DARK)
    mark.line.fill.background()
    add_text(slide, "TA", 0.82, 6.92, 0.24, 0.12, size=8, color=WHITE, bold=True)
    add_text(slide, label, 1.24, 6.87, 4.2, 0.2, size=8.5, color=MUTED, font="Consolas")


def add_card(slide, x, y, w, h, title, body, index=None, dark=False):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(shape, ACCENT_DARK if dark else WHITE, 0 if dark else 6)
    set_line(shape, ACCENT_DARK if dark else LINE, 1)
    if index:
        add_text(slide, index, x + 0.26, y + 0.22, 0.48, 0.2, size=10, color=WHITE if dark else ACCENT, bold=True, font="Consolas")
    add_text(slide, title, x + 0.26, y + 0.58, w - 0.52, 0.38, size=18, color=WHITE if dark else INK, bold=True)
    body_h = h - 1.22
    if body and body_h > 0.18:
        add_text(slide, body, x + 0.26, y + 1.1, w - 0.52, body_h, size=12.5, color=RGBColor(202, 226, 216) if dark else MUTED)
    return shape


def add_step(slide, x, y, w, h, state, title, index):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(shape, WHITE, 6)
    set_line(shape, LINE, 1)
    add_text(slide, index, x + 0.24, y + 0.22, 0.46, 0.18, size=9.5, color=ACCENT, bold=True, font="Consolas")
    add_text(slide, state, x + 0.88, y + 0.2, 1.1, 0.2, size=10.5, color=MUTED)
    add_text(slide, title, x + 0.88, y + 0.48, w - 1.12, 0.24, size=16, color=INK, bold=True)
    return shape


def add_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "PRE LAUNCH PRODUCT")
    add_label(slide, "Pre Launch Product", 0.72, 0.72)
    add_text(slide, "一个已经做出来，\n等待公开发布的\nAI 助教产品。", 0.72, 1.38, 5.45, 3.8, size=46, bold=True)
    add_text(
        slide,
        "不是泛泛聊天工具，而是围绕课程项目、代码理解、\n问题定位和交付推进设计的工作台。",
        0.78,
        5.42,
        5.5,
        0.64,
        size=15,
        color=MUTED,
    )
    add_text(slide, "技术型助教 · 小学生 · 会用 Codex、Bolt 等 AI 工具", 0.78, 6.08, 5.4, 0.28, size=13.5, color=ACCENT_DARK, bold=True)
    frame = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(6.55), Inches(1.32), Inches(5.8), Inches(4.5))
    set_fill(frame, WHITE, 8)
    set_line(frame, LINE)
    slide.shapes.add_picture(str(ASSET), Inches(6.75), Inches(1.54), width=Inches(5.4))
    add_text(slide, "核心版本已完成，处于预发布和体验校准阶段。", 6.74, 5.92, 4.7, 0.34, size=10.5, color=MUTED)


def add_positioning(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "POSITIONING")
    add_label(slide, "Positioning")
    add_text(slide, "产品不是陪聊，\n而是把项目推进到下一步。", 0.72, 1.18, 5.6, 1.7, size=35, bold=True)
    add_text(slide, "定位要清楚：它不是做一个更会说话的 AI，\n而是做一个更像工程助教的产品。", 0.76, 3.12, 5.35, 0.82, size=15, color=MUTED)
    add_card(slide, 6.65, 1.12, 2.8, 2.05, "读懂代码结构", "拆开组件、接口、状态和数据流，让学习者知道代码为什么这样组织。", "01")
    add_card(slide, 9.72, 1.12, 2.8, 2.05, "定位真实问题", "从报错、复现、依赖和环境入手，缩小问题范围。", "02")
    add_card(slide, 6.65, 3.52, 2.8, 2.05, "推进项目交付", "关注页面完成度、接口连通性、部署准备和展示效果。", "03")
    add_card(slide, 9.72, 3.52, 2.8, 2.05, "保留学习路径", "留下判断过程，让下一次同类问题能自己处理。", "04", dark=True)


def add_about(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "ABOUT ME")
    add_label(slide, "About Me")
    add_text(slide, "我是技术型助教，小学生，\n正在用 AI 把想法做成产品。", 0.72, 1.1, 6.0, 1.7, size=35, bold=True)
    add_text(slide, "我会使用多种 AI 工具，比如 Codex 和 Bolt。\n对我来说，AI 不是用来炫技的，而是帮助我更快理解代码、搭建原型、修正问题和完成作品。", 0.76, 3.12, 5.75, 1.0, size=14.5, color=MUTED)
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(1.12), Inches(4.72), Inches(4.62))
    set_fill(panel, WHITE, 6)
    set_line(panel, LINE, 1)
    mark = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(7.35), Inches(1.52), Inches(0.9), Inches(0.72))
    set_fill(mark, ACCENT_DARK)
    mark.line.fill.background()
    add_text(slide, "ZF", 7.58, 1.72, 0.46, 0.22, size=16, color=WHITE, bold=True)
    rows = [
        ("姓名", "技术型助教"),
        ("年龄", "小学生"),
        ("AI 工具", "Codex / Bolt / 以及更多 AI 工具"),
        ("当前方向", "用 AI 辅助学习编程，并把项目推进到可演示状态"),
    ]
    for i, (label, value) in enumerate(rows):
        y = 2.55 + i * 0.72
        add_text(slide, label, 7.35, y, 0.9, 0.22, size=10, color=ACCENT, bold=True, font="Consolas")
        add_text(slide, value, 8.38, y - 0.02, 2.95, 0.32, size=13.2, color=INK, bold=True)


def add_progress(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "CURRENT STATE")
    add_label(slide, "Progress")
    add_text(slide, "已经做出产品，\n不是只停留在想法。", 0.72, 1.16, 5.2, 1.45, size=36, bold=True)
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(0.78), Inches(3.1), Inches(5.3), Inches(2.1))
    set_fill(panel, ACCENT_DARK)
    panel.line.fill.background()
    add_text(slide, "当前版本已完成核心体验：\n输入项目上下文后，产品可以协助理解代码、拆解问题、给出修改方向。\n一次次排查会沉淀成更清楚的学习路径。", 1.08, 3.42, 4.55, 1.42, size=13.8, color=RGBColor(230, 249, 240))
    steps = [("完成", "核心流程"), ("完成", "产品界面"), ("进行中", "体验打磨"), ("下一步", "公开发布")]
    x = 6.55
    for i, (state, title) in enumerate(steps):
        add_step(slide, x, 1.2 + i * 1.18, 5.7, 0.88, state, title, f"0{i+1}")


def add_method(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "METHOD")
    add_label(slide, "Method")
    add_text(slide, "工程化、清楚、有边界。", 0.72, 1.16, 7.8, 0.72, size=39, bold=True)
    add_text(slide, "这个产品需要给人的感觉不是热闹，而是可靠。\n每一页都应该让人知道：它能处理什么、不能承诺什么、下一步在哪里。", 0.76, 2.14, 7.2, 0.82, size=15, color=MUTED)
    add_card(slide, 0.78, 3.42, 3.76, 1.8, "少一点拟人，多一点判断", "重点不是像朋友聊天，而是帮助用户看清问题结构和下一步行动。", "01")
    add_card(slide, 4.88, 3.42, 3.76, 1.8, "少一点热闹，多一点证据", "用明确状态、过程记录和交付检查建立信任。", "02")
    add_card(slide, 8.98, 3.42, 3.76, 1.8, "少一点模板，多一点完成度", "从真实课程项目出发，优先解决代码、调试、部署和展示。", "03", dark=True)


def add_product_preview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "PRODUCT PREVIEW")
    add_label(slide, "Preview")
    add_text(slide, "核心界面：\n项目诊断工作台。", 0.72, 1.12, 4.5, 1.5, size=35, bold=True)
    add_text(slide, "这一页承担证据作用：\n让听众知道产品已经有界面、有流程、有明确的使用场景。", 0.76, 2.92, 4.34, 0.88, size=15, color=MUTED)
    slide.shapes.add_picture(str(ASSET), Inches(5.18), Inches(0.82), width=Inches(7.32))
    add_card(slide, 0.78, 4.4, 4.15, 1.16, "视觉策略", "浅底、深绿、细网格、真实产品图，避免紫蓝渐变和聊天气泡。", None, dark=True)


def add_support(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "SUPPORT")
    add_label(slide, "Support")
    add_text(slide, "我能具体帮你什么？", 0.72, 1.1, 6.4, 0.8, size=38, bold=True)
    items = [
        ("解释代码逻辑", "把函数、组件、接口调用和数据流拆开讲清楚。"),
        ("定位 bug", "一起看报错、复现问题、缩小范围，找到真实原因。"),
        ("完善项目", "把作业变成结构完整、体验更好、可以展示的作品。"),
        ("理解 AI / API / 部署", "讲清模型调用、接口参数、环境变量和上线前检查。"),
    ]
    for i, (title, body) in enumerate(items):
        add_card(slide, 0.78 + (i % 2) * 6.0, 2.38 + (i // 2) * 1.72, 5.42, 1.28, title, body, f"0{i+1}", dark=(i == 3))


def add_close(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, "READY")
    add_label(slide, "Ready")
    add_text(slide, "如果你想把项目做成\n不只是能交作业的作品，\n可以从一次项目诊断开始。", 0.72, 1.18, 7.1, 2.7, size=36, bold=True)
    add_text(slide, "技术型助教 · 预发布产品演示", 0.78, 5.1, 4.2, 0.34, size=14, color=ACCENT, bold=True)
    badge = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(9.35), Inches(4.8), Inches(2.6), Inches(0.56))
    set_fill(badge, ACCENT_DARK)
    badge.line.fill.background()
    add_text(slide, "Core Build Ready", 9.78, 4.96, 1.8, 0.16, size=11, color=WHITE, bold=True, font="Consolas", align=PP_ALIGN.CENTER)


def build():
    make_grid_bg()
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    add_title_slide(prs)
    add_positioning(prs)
    add_about(prs)
    add_progress(prs)
    add_method(prs)
    add_product_preview(prs)
    add_support(prs)
    add_close(prs)
    prs.save(OUT)
    print(OUT)


def make_grid_bg():
    width, height = 1920, 1080
    img = Image.new("RGB", (width, height), (244, 245, 240))
    draw = ImageDraw.Draw(img, "RGBA")
    for x in range(0, width, 104):
        draw.line((x, 0, x, height), fill=(204, 212, 202, 58), width=1)
    for y in range(0, height, 104):
        draw.line((0, y, width, y), fill=(204, 212, 202, 52), width=1)
    draw.ellipse((1320, -260, 2180, 560), fill=(15, 125, 87, 20))
    GRID_BG.parent.mkdir(exist_ok=True)
    img.save(GRID_BG)


if __name__ == "__main__":
    build()
