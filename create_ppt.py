from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── Color Palette ──
DARK_BG       = RGBColor(15, 23, 42)
CARD_BG       = RGBColor(30, 41, 59)
ACCENT        = RGBColor(99, 102, 241)
ACCENT_LIGHT  = RGBColor(168, 85, 247)
WHITE         = RGBColor(248, 250, 252)
MUTED         = RGBColor(148, 163, 184)
GREEN         = RGBColor(52, 211, 153)
ORANGE        = RGBColor(251, 146, 60)
RED           = RGBColor(248, 113, 113)
CYAN          = RGBColor(34, 211, 238)


# ─── Helper Functions ────────────────────────────────────────────────

def add_dark_bg(slide):
    """Fill entire slide with dark background"""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = DARK_BG


def add_accent_bar(slide, left=0, top=0, width=None, height=Inches(0.06)):
    """Add a thin gradient-like accent bar"""
    if width is None:
        width = prs.slide_width
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT
    shape.line.fill.background()


def add_side_accent(slide):
    """Add a vertical accent strip on the left"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.15), prs.slide_height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT
    shape.line.fill.background()


def add_slide_number(slide, num, total=10):
    """Add slide number at bottom right"""
    txBox = slide.shapes.add_textbox(
        Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.4)
    )
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"{num} / {total}"
    p.font.size = Pt(11)
    p.font.color.rgb = MUTED
    p.alignment = PP_ALIGN.RIGHT


def add_card(slide, left, top, width, height, text_lines, icon="", highlight_color=ACCENT):
    """Add a card-like rounded rectangle with text"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = CARD_BG
    shape.line.color.rgb = RGBColor(51, 65, 85)
    shape.line.width = Pt(1)
    shape.shadow.inherit = False

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.25)
    tf.margin_bottom = Inches(0.2)

    for i, line in enumerate(text_lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        if i == 0 and icon:
            p.text = f"{icon}  {line}"
        else:
            p.text = line

        if i == 0:
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = highlight_color
            p.space_after = Pt(8)
        else:
            p.font.size = Pt(14)
            p.font.color.rgb = WHITE
            p.space_after = Pt(4)


def add_title_block(slide, title, subtitle="", top=Inches(0.8)):
    """Add the main title for content slides"""
    # Title
    txBox = slide.shapes.add_textbox(Inches(0.8), top, Inches(11), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = WHITE

    # Underline bar
    add_accent_bar(slide, Inches(0.8), top + Inches(0.75), Inches(2.5), Inches(0.05))

    if subtitle:
        txBox2 = slide.shapes.add_textbox(Inches(0.8), top + Inches(0.9), Inches(10), Inches(0.5))
        tf2 = txBox2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle
        p2.font.size = Pt(16)
        p2.font.color.rgb = MUTED


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 1 — TITLE SLIDE
# ═══════════════════════════════════════════════════════════════════════
slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # blank
add_dark_bg(slide1)

# Decorative circle (top-right)
circle1 = slide1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10), Inches(-1.5), Inches(5), Inches(5))
circle1.fill.solid()
circle1.fill.fore_color.rgb = RGBColor(30, 27, 75)
circle1.line.fill.background()

# Decorative circle (bottom-left)
circle2 = slide1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(4.5), Inches(6), Inches(6))
circle2.fill.solid()
circle2.fill.fore_color.rgb = RGBColor(30, 27, 75)
circle2.line.fill.background()

# Accent bar at top
add_accent_bar(slide1, height=Inches(0.08))

# Main Title
txBox = slide1.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10), Inches(1.5))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "State & District Explorer"
p.font.size = Pt(54)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

# Subtitle
txBox2 = slide1.shapes.add_textbox(Inches(2), Inches(3.3), Inches(9), Inches(0.8))
tf2 = txBox2.text_frame
p2 = tf2.paragraphs[0]
p2.text = "A Web-Based Tool to Explore Districts of Indian States & Union Territories"
p2.font.size = Pt(22)
p2.font.color.rgb = MUTED
p2.alignment = PP_ALIGN.CENTER

# Divider line
add_accent_bar(slide1, Inches(4.5), Inches(4.3), Inches(4), Inches(0.04))

# Presented by
txBox3 = slide1.shapes.add_textbox(Inches(2), Inches(4.8), Inches(9), Inches(1.2))
tf3 = txBox3.text_frame
p3 = tf3.paragraphs[0]
p3.text = "Presented By"
p3.font.size = Pt(14)
p3.font.color.rgb = MUTED
p3.alignment = PP_ALIGN.CENTER

p4 = tf3.add_paragraph()
p4.text = "Harshit  &  Golu"
p4.font.size = Pt(28)
p4.font.bold = True
p4.font.color.rgb = ACCENT
p4.alignment = PP_ALIGN.CENTER

p5 = tf3.add_paragraph()
p5.text = "Maa Bharti PG College  •  BCA  •  2025-26"
p5.font.size = Pt(14)
p5.font.color.rgb = MUTED
p5.alignment = PP_ALIGN.CENTER

add_slide_number(slide1, 1)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 2 — TABLE OF CONTENTS
# ═══════════════════════════════════════════════════════════════════════
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide2)
add_side_accent(slide2)
add_title_block(slide2, "Table of Contents", "Overview of the presentation")

contents = [
    ("01", "Introduction", "Project ka overview aur uddeshya"),
    ("02", "Problem Statement", "Kya problem solve kar rahe hain"),
    ("03", "Key Features", "Project ki main features"),
    ("04", "Technology Stack", "Kaunsi technologies use hui hain"),
    ("05", "Project Structure", "Files aur folders ka layout"),
    ("06", "System Architecture", "Project kaise kaam karta hai"),
    ("07", "Code Walkthrough", "Important code snippets"),
    ("08", "Output / Demo", "Project ka final output"),
    ("09", "Future Scope", "Aage kya improvements ho sakti hain"),
    ("10", "Thank You", ""),
]

for i, (num, title, desc) in enumerate(contents):
    row = i // 2
    col = i % 2
    left = Inches(1.0) + col * Inches(6)
    top = Inches(2.2) + row * Inches(0.95)

    txBox = slide2.shapes.add_textbox(left, top, Inches(5.5), Inches(0.8))
    tf = txBox.text_frame

    p = tf.paragraphs[0]
    run1 = p.add_run()
    run1.text = f"{num}   "
    run1.font.size = Pt(18)
    run1.font.bold = True
    run1.font.color.rgb = ACCENT

    run2 = p.add_run()
    run2.text = title
    run2.font.size = Pt(18)
    run2.font.color.rgb = WHITE

    if desc:
        run3 = p.add_run()
        run3.text = f"   —  {desc}"
        run3.font.size = Pt(12)
        run3.font.color.rgb = MUTED

add_slide_number(slide2, 2)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 3 — INTRODUCTION
# ═══════════════════════════════════════════════════════════════════════
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide3)
add_side_accent(slide3)
add_title_block(slide3, "Introduction", "Project ka parichay aur uddeshya")

add_card(slide3, Inches(0.8), Inches(2.3), Inches(5.8), Inches(2.0),
    ["What is this Project?",
     "State & District Explorer ek web-based",
     "application hai jo users ko India ke kisi",
     "bhi state ke saare districts instantly",
     "search karne ki suvidha deta hai."],
    icon="🔍", highlight_color=ACCENT
)

add_card(slide3, Inches(6.8), Inches(2.3), Inches(5.8), Inches(2.0),
    ["Why was it built?",
     "India mein 700+ districts hain aur unhe",
     "yaad rakhna mushkil hai. Yeh tool ek",
     "fast, offline-capable directory provide",
     "karta hai bina kisi complex database ke."],
    icon="💡", highlight_color=ACCENT_LIGHT
)

add_card(slide3, Inches(0.8), Inches(4.6), Inches(11.8), Inches(2.2),
    ["Project Objective",
     "✦  Indian states aur union territories ke districts ko ek jagah available karna.",
     "✦  User-friendly web interface banana jo mobile aur desktop dono par kaam kare.",
     "✦  Python aur Web Technologies ka practical implementation seekhna aur dikhana.",
     "✦  JSON data format ka real-world use case demonstrate karna."],
    icon="🎯", highlight_color=GREEN
)

add_slide_number(slide3, 3)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 4 — PROBLEM STATEMENT
# ═══════════════════════════════════════════════════════════════════════
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide4)
add_side_accent(slide4)
add_title_block(slide4, "Problem Statement", "Yeh project kyun zaroori hai?")

problems = [
    ("❌", "Manual Search is Slow", "Har baar Google par state ke districts search karna time-consuming hai aur internet zaroori hai.", RED),
    ("❌", "No Single Offline Source", "Koi aisi simple tool nahi thi jo bina internet ke saare districts ek jagah dikha sake.", ORANGE),
    ("✅", "Our Solution", "Ek lightweight web app jo locally run hota hai, instantly search results deta hai, aur kisi bhi internet connection ki zaroorat nahi.", GREEN),
]

for i, (icon, title, desc, color) in enumerate(problems):
    top = Inches(2.3) + i * Inches(1.6)
    add_card(slide4, Inches(1.0), top, Inches(11.5), Inches(1.35),
        [title, desc], icon=icon, highlight_color=color
    )

add_slide_number(slide4, 4)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 5 — KEY FEATURES
# ═══════════════════════════════════════════════════════════════════════
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide5)
add_side_accent(slide5)
add_title_block(slide5, "Key Features", "Project ki mukhya visheshatayen")

features = [
    ("⚡", "Instant Search", "State ka naam type karte hi\ndistricts turant screen par\naate hain, koi delay nahi.", CYAN),
    ("🎨", "Modern Dark UI", "Glassmorphism effects,\nsmooth animations aur\nresponsive design.", ACCENT_LIGHT),
    ("📱", "Dual Mode", "Web Mode (Browser) aur\nTerminal Mode (CLI) dono\nmein kaam karta hai.", GREEN),
    ("🔒", "Offline Capable", "Ek baar run karne ke baad\ninternet ki zaroorat nahi,\nsab local data se chalta hai.", ORANGE),
]

for i, (icon, title, desc, color) in enumerate(features):
    left = Inches(0.8) + i * Inches(3.15)
    add_card(slide5, left, Inches(2.3), Inches(2.9), Inches(4.0),
        [title, "", desc], icon=icon, highlight_color=color
    )

add_slide_number(slide5, 5)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 6 — TECHNOLOGY STACK
# ═══════════════════════════════════════════════════════════════════════
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide6)
add_side_accent(slide6)
add_title_block(slide6, "Technology Stack", "Project mein use hui technologies")

techs = [
    ("Frontend", [
        ("HTML5", "Web page ka structure aur layout define karta hai."),
        ("CSS3", "Styling, animations, glassmorphism effects ke liye."),
        ("JavaScript", "Search logic, DOM manipulation, aur data fetching."),
    ], ACCENT),
    ("Backend", [
        ("Python 3", "Local HTTP server host karne ke liye (http.server module)."),
        ("JSON", "States aur districts ka data store karne ke liye lightweight format."),
    ], GREEN),
]

col_left = Inches(0.8)
for j, (category, items, color) in enumerate(techs):
    left = col_left + j * Inches(6.3)

    # Category heading
    txBox = slide6.shapes.add_textbox(left, Inches(2.3), Inches(5.8), Inches(0.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"▸  {category}"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = color

    for i, (tech, desc) in enumerate(items):
        top = Inches(3.0) + i * Inches(1.35)
        add_card(slide6, left, top, Inches(5.8), Inches(1.15),
            [tech, desc], highlight_color=color
        )

add_slide_number(slide6, 6)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 7 — PROJECT STRUCTURE
# ═══════════════════════════════════════════════════════════════════════
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide7)
add_side_accent(slide7)
add_title_block(slide7, "Project Structure", "Files aur unka kaam")

files_info = [
    ("📄  index.html", "Main web page — UI design,\nsearch box, aur JavaScript logic\nsab isi file mein hai. (8 KB)", ACCENT),
    ("🗂  states-and-districts.json", "Data file — India ke saare states\naur unke districts ka data JSON\nformat mein store hai. (13 KB)", GREEN),
    ("⚙  server.py", "Python script — Ek local web\nserver banata hai port 5000 par\ntaaki webpage browser mein chale.", ORANGE),
    ("💻  data.py", "CLI script — Terminal se seedha\nstate search karne ke liye.\nBina browser ke kaam karta hai.", CYAN),
]

for i, (title, desc, color) in enumerate(files_info):
    col = i % 2
    row = i // 2
    left = Inches(0.8) + col * Inches(6.2)
    top = Inches(2.3) + row * Inches(2.3)
    add_card(slide7, left, top, Inches(5.8), Inches(2.0),
        [title, "", desc], highlight_color=color
    )

add_slide_number(slide7, 7)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 8 — SYSTEM ARCHITECTURE / HOW IT WORKS
# ═══════════════════════════════════════════════════════════════════════
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide8)
add_side_accent(slide8)
add_title_block(slide8, "How It Works?", "System Architecture aur Data Flow")

steps = [
    ("Step 1", "Server Start", "python server.py\ncommand se local HTTP\nserver port 5000 par\nstart hota hai.", ACCENT),
    ("Step 2", "Page Load", "Browser mein localhost:5000\nkholne par index.html load\nhota hai aur fetch() se\nJSON data read hota hai.", ACCENT_LIGHT),
    ("Step 3", "User Search", "User search box mein\nstate ka naam type\nkarta hai aur Search\nbutton click karta hai.", GREEN),
    ("Step 4", "Show Results", "JavaScript JSON data\nmein match dhoondh kar\ndistricts ki list screen\npar animate karke dikhata hai.", ORANGE),
]

for i, (step, title, desc, color) in enumerate(steps):
    left = Inches(0.6) + i * Inches(3.2)

    # Step number badge
    badge = slide8.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.9), Inches(2.2), Inches(1.2), Inches(0.45)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = color
    badge.line.fill.background()
    btf = badge.text_frame
    btf.paragraphs[0].text = step
    btf.paragraphs[0].font.size = Pt(14)
    btf.paragraphs[0].font.bold = True
    btf.paragraphs[0].font.color.rgb = WHITE
    btf.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Arrow between steps (except last)
    if i < len(steps) - 1:
        arrow = slide8.shapes.add_shape(
            MSO_SHAPE.RIGHT_ARROW, left + Inches(2.85), Inches(3.8), Inches(0.5), Inches(0.35)
        )
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = MUTED
        arrow.line.fill.background()

    add_card(slide8, left, Inches(2.8), Inches(2.9), Inches(3.5),
        [title, "", desc], highlight_color=color
    )

add_slide_number(slide8, 8)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 9 — CODE WALKTHROUGH
# ═══════════════════════════════════════════════════════════════════════
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide9)
add_side_accent(slide9)
add_title_block(slide9, "Code Walkthrough", "Important code snippets")

# server.py code block
code1_title = slide9.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(5.5), Inches(0.4))
tf = code1_title.text_frame
p = tf.paragraphs[0]
p.text = "▸  server.py  —  Local Server"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT

code1 = slide9.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.7), Inches(5.8), Inches(4.0)
)
code1.fill.solid()
code1.fill.fore_color.rgb = RGBColor(17, 24, 39)
code1.line.color.rgb = RGBColor(51, 65, 85)
code1.line.width = Pt(1)

ctf = code1.text_frame
ctf.margin_left = Inches(0.3)
ctf.margin_top = Inches(0.25)
ctf.word_wrap = True

code_lines_1 = [
    "import http.server",
    "import socketserver",
    "",
    "PORT = 5000",
    "",
    "with socketserver.TCPServer(",
    '    ("", PORT),',
    "    http.server.SimpleHTTPRequestHandler",
    ") as server:",
    '    print(f"Server started at',
    '        http://localhost:{PORT}")',
    "    server.serve_forever()",
]

for i, line in enumerate(code_lines_1):
    if i == 0:
        p = ctf.paragraphs[0]
    else:
        p = ctf.add_paragraph()
    p.text = line
    p.font.size = Pt(13)
    p.font.name = "Consolas"
    p.font.color.rgb = RGBColor(190, 242, 100) if line.startswith("import") else WHITE
    p.space_after = Pt(2)

# JavaScript search logic
code2_title = slide9.shapes.add_textbox(Inches(6.9), Inches(2.2), Inches(5.5), Inches(0.4))
tf2 = code2_title.text_frame
p2 = tf2.paragraphs[0]
p2.text = "▸  index.html  —  Search Logic (JS)"
p2.font.size = Pt(16)
p2.font.bold = True
p2.font.color.rgb = GREEN

code2 = slide9.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.7), Inches(5.8), Inches(4.0)
)
code2.fill.solid()
code2.fill.fore_color.rgb = RGBColor(17, 24, 39)
code2.line.color.rgb = RGBColor(51, 65, 85)
code2.line.width = Pt(1)

ctf2 = code2.text_frame
ctf2.margin_left = Inches(0.3)
ctf2.margin_top = Inches(0.25)
ctf2.word_wrap = True

code_lines_2 = [
    "// JSON data load karna",
    "fetch('states-and-districts.json')",
    "  .then(res => res.json())",
    "  .then(data => {",
    "    statesData = data.states;",
    "  });",
    "",
    "// Search function",
    "function performSearch() {",
    "  const query = searchInput.value",
    "    .trim().toLowerCase();",
    "",
    "  const found = statesData",
    "    .find(s =>",
    "      s.state.toLowerCase() === query",
    "    );",
    "}",
]

for i, line in enumerate(code_lines_2):
    if i == 0:
        p = ctf2.paragraphs[0]
    else:
        p = ctf2.add_paragraph()
    p.text = line
    p.font.size = Pt(13)
    p.font.name = "Consolas"
    if line.startswith("//"):
        p.font.color.rgb = MUTED
    elif "function" in line or "const" in line:
        p.font.color.rgb = RGBColor(125, 211, 252)
    else:
        p.font.color.rgb = WHITE
    p.space_after = Pt(2)

add_slide_number(slide9, 9)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 10 — OUTPUT / DEMO
# ═══════════════════════════════════════════════════════════════════════
slide10 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide10)
add_side_accent(slide10)
add_title_block(slide10, "Output / Demo", "Project ka final output")

# Web Output description
add_card(slide10, Inches(0.8), Inches(2.3), Inches(5.8), Inches(4.2),
    ["🌐  Web Output (Browser)",
     "",
     "Browser mein http://localhost:5000 kholne par:",
     "",
     "✦  Ek dark-themed search interface dikhta hai",
     "✦  Search box mein state naam type karein",
     "✦  Districts animated list mein dikhte hain",
     "✦  Har district left-border highlight ke saath",
     "✦  Total districts ka count bhi dikhta hai",
     "✦  Hover effects aur smooth transitions"],
    highlight_color=ACCENT
)

# Terminal Output description
add_card(slide10, Inches(6.9), Inches(2.3), Inches(5.8), Inches(4.2),
    ["💻  Terminal Output (CLI)",
     "",
     "python data.py chalane par:",
     "",
     "Enter State Name: Maharashtra",
     "",
     "Districts of Maharashtra:",
     "1: Ahmednagar",
     "2: Akola",
     "3: Amravati",
     "4: Aurangabad",
     "... aur baaki saare districts"],
    highlight_color=GREEN
)

add_slide_number(slide10, 10)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 11 — FUTURE SCOPE
# ═══════════════════════════════════════════════════════════════════════
slide11 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide11)
add_side_accent(slide11)
add_title_block(slide11, "Future Scope", "Aage ki improvements aur possibilities")

future_items = [
    ("🗺️", "Map Integration", "Interactive India map add karna jisme\nclick karne par state highlight ho aur\nuske districts dikhe.", ACCENT),
    ("📊", "More Data", "Har district ka population, area,\nlanguage, aur literacy rate jaisi\naur information add karna.", GREEN),
    ("🔎", "Auto-Suggest", "Search box mein type karte waqt\nautomatically state names suggest\nhone chahiye (autocomplete feature).", ORANGE),
    ("📱", "Mobile App", "Is project ko ek Android/iOS app\nmein convert karna taaki log\napne phone par bhi use kar sakein.", CYAN),
]

for i, (icon, title, desc, color) in enumerate(future_items):
    col = i % 2
    row = i // 2
    left = Inches(0.8) + col * Inches(6.2)
    top = Inches(2.3) + row * Inches(2.3)
    add_card(slide11, left, top, Inches(5.8), Inches(2.0),
        [title, "", desc], icon=icon, highlight_color=color
    )

add_slide_number(slide11, 11)


# ═══════════════════════════════════════════════════════════════════════
#   SLIDE 12 — THANK YOU
# ═══════════════════════════════════════════════════════════════════════
slide12 = prs.slides.add_slide(prs.slide_layouts[6])
add_dark_bg(slide12)
add_accent_bar(slide12, height=Inches(0.08))

# Decorative circles
c1 = slide12.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.5), Inches(-2), Inches(6), Inches(6))
c1.fill.solid()
c1.fill.fore_color.rgb = RGBColor(30, 27, 75)
c1.line.fill.background()

c2 = slide12.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2.5), Inches(4), Inches(7), Inches(7))
c2.fill.solid()
c2.fill.fore_color.rgb = RGBColor(30, 27, 75)
c2.line.fill.background()

# Thank You text
txBox = slide12.shapes.add_textbox(Inches(2), Inches(2.0), Inches(9), Inches(1.2))
tf = txBox.text_frame
p = tf.paragraphs[0]
p.text = "Thank You!"
p.font.size = Pt(60)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.CENTER

# Subtitle
txBox2 = slide12.shapes.add_textbox(Inches(2), Inches(3.5), Inches(9), Inches(0.8))
tf2 = txBox2.text_frame
p2 = tf2.paragraphs[0]
p2.text = "Questions & Feedback are Welcome"
p2.font.size = Pt(24)
p2.font.color.rgb = MUTED
p2.alignment = PP_ALIGN.CENTER

# Divider
add_accent_bar(slide12, Inches(5), Inches(4.5), Inches(3), Inches(0.04))

# Contact info placeholder
txBox3 = slide12.shapes.add_textbox(Inches(2), Inches(5.0), Inches(9), Inches(1.5))
tf3 = txBox3.text_frame
p3 = tf3.paragraphs[0]
p3.text = "Harshit  &  Golu"
p3.font.size = Pt(22)
p3.font.bold = True
p3.font.color.rgb = ACCENT
p3.alignment = PP_ALIGN.CENTER

p4 = tf3.add_paragraph()
p4.text = "Maa Bharti PG College"
p4.font.size = Pt(14)
p4.font.color.rgb = MUTED
p4.alignment = PP_ALIGN.CENTER

add_slide_number(slide12, 12)


# ─── SAVE ────────────────────────────────────────────────────────────
output_path = r"D:\project\State_District_Explorer_Presentation.pptx"
prs.save(output_path)
print(f"\nPresentation saved successfully!")
print(f"Location: {output_path}")
print(f"Total Slides: {len(prs.slides)}")
