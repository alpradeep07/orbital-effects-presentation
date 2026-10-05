from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ---------- Theme helpers ----------
NAVY = RGBColor(18, 42, 74)
BLUE = RGBColor(31, 96, 160)
TEAL = RGBColor(42, 167, 155)
ORANGE = RGBColor(239, 125, 49)
LIGHT = RGBColor(242, 246, 250)
TEXT = RGBColor(24, 32, 38)
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(120, 129, 140)


def set_bg(slide, color=LIGHT):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_title(slide, title, subtitle=None):
    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(12), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.runs[0].font.size = Pt(26)
    p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = NAVY

    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.15), Inches(11.5), Inches(0.5))
        tf2 = sub_box.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle
        p2.runs[0].font.size = Pt(12)
        p2.runs[0].font.color.rgb = GRAY


def add_bullets(slide, items, left, top, width, height, font_size=18, color=TEXT):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.bullet = True
        p.space_after = Pt(9)
        p.alignment = PP_ALIGN.LEFT
        p.runs[0].font.size = Pt(font_size)
        p.runs[0].font.color.rgb = color


def add_formula_box(slide, formula, left, top, width, height, title=None):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = BLUE
    box.line.width = Pt(1.2)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0.08
    tf.margin_right = 0.08
    tf.margin_top = 0.08
    tf.margin_bottom = 0.08
    if title:
        p = tf.paragraphs[0]
        p.text = title
        p.runs[0].font.size = Pt(12)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = NAVY
    p = tf.add_paragraph() if title else tf.paragraphs[0]
    p.text = formula
    p.runs[0].font.size = Pt(18)
    p.runs[0].font.color.rgb = NAVY
    p.alignment = PP_ALIGN.CENTER


# Slide 1: Title
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
# banner
banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.8))
banner.fill.solid(); banner.fill.fore_color.rgb = NAVY
banner.line.fill.background()

# Title text
textbox = slide.shapes.add_textbox(Inches(0.7), Inches(1.1), Inches(12), Inches(1.0))
tf = textbox.text_frame
p = tf.paragraphs[0]
p.text = "Orbital Effects: Eclipse and Sun Transit Outage"
p.runs[0].font.size = Pt(28)
p.runs[0].font.bold = True
p.runs[0].font.color.rgb = NAVY

sub = slide.shapes.add_textbox(Inches(0.7), Inches(2.0), Inches(10), Inches(0.5))
tf2 = sub.text_frame
p2 = tf2.paragraphs[0]
p2.text = "Satellite communications disruptions caused by solar geometry and Earth shadow"
p2.runs[0].font.size = Pt(18)
p2.runs[0].font.color.rgb = GRAY

# diagram: Earth + Sun + satellite
# Earth
earth = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.5), Inches(3.8), Inches(1.6), Inches(1.6))
earth.fill.solid(); earth.fill.fore_color.rgb = BLUE
earth.line.color.rgb = BLUE
# Sun
sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.7), Inches(3.3), Inches(1.5), Inches(1.5))
sun.fill.solid(); sun.fill.fore_color.rgb = RGBColor(255, 196, 0)
sun.line.color.rgb = RGBColor(255, 196, 0)
# satellite
sat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.8), Inches(4.1), Inches(0.6), Inches(0.4))
sat.fill.solid(); sat.fill.fore_color.rgb = ORANGE
sat.line.color.rgb = ORANGE
# lines
line1 = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(3.3), Inches(4.7), Inches(5.9), Inches(4.3))
line1.line.color.rgb = GRAY
line2 = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(5.9), Inches(4.3), Inches(8.9), Inches(4.1))
line2.line.color.rgb = GRAY

# labels
for label, x, y in [
    ("Earth", 2.7, 5.8),
    ("Sun", 9.0, 5.1),
    ("Satellite", 5.3, 5.0),
]:
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(1.8), Inches(0.4))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = label
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(14)
    p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = TEXT

# footer
footer = slide.shapes.add_textbox(Inches(0.7), Inches(6.8), Inches(12), Inches(0.4))
ft = footer.text_frame
p3 = ft.paragraphs[0]
p3.text = "Prepared for orbital mechanics / communications systems review"
p3.runs[0].font.size = Pt(12)
p3.runs[0].font.color.rgb = GRAY


# Slide 2: Why it matters
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "Why Eclipse and Sun Transit Outage Matter")
items = [
    "Satellite links are highly sensitive to the geometry between Earth, satellite, and Sun.",
    "During eclipse, a satellite may lose direct solar power and switch to battery operation.",
    "During sun transit outage, the Sun acts as a strong radio-noise source behind the satellite.",
    "Both events are predictable and occur near seasonal orbital alignments.",
]
add_bullets(slide, items, 0.8, 1.8, 7.4, 4.5, font_size=20)

# simple diagram on right
right = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.3), Inches(2.0), Inches(3.7), Inches(3.8))
right.fill.solid(); right.fill.fore_color.rgb = WHITE
right.line.color.rgb = BLUE
right.line.width = Pt(1.5)
# draw geometry
# earth
earth = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.1), Inches(3.3), Inches(1.2), Inches(1.2))
earth.fill.solid(); earth.fill.fore_color.rgb = BLUE
earth.line.color.rgb = BLUE
# sat
sat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.7), Inches(3.6), Inches(0.5), Inches(0.3))
sat.fill.solid(); sat.fill.fore_color.rgb = ORANGE
sat.line.color.rgb = ORANGE
# sun
sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.0), Inches(2.1), Inches(0.8), Inches(0.8))
sun.fill.solid(); sun.fill.fore_color.rgb = RGBColor(255, 196, 0)
sun.line.color.rgb = RGBColor(255, 196, 0)
# connector line
ln = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(9.7), Inches(3.9), Inches(10.7), Inches(3.8))
ln.line.color.rgb = GRAY
ln.line.width = Pt(1.3)

# labels
for label, x, y in [("Earth", 8.9, 5.0), ("Sun", 10.0, 1.4), ("Satellite", 10.2, 4.5)]:
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(2.1), Inches(0.4))
    p = tb.text_frame.paragraphs[0]
    p.text = label
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(12)
    p.runs[0].font.color.rgb = TEXT


# Slide 3: GEO geometry
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "Geostationary Orbit Geometry")
# left text
add_bullets(slide, [
    "A GEO satellite remains nearly fixed above the equator at about 35,786 km.",
    "The apparent path of the Sun moves on the ecliptic, so seasonal alignment occurs.",
    "Outage and eclipse are most likely near equinox when the Sun crosses the orbital plane.",
    "The geometry is determined by Earth radius, orbital radius, and solar declination."
], 0.8, 1.8, 5.8, 3.6, font_size=20)

# right diagram
# Earth
earth = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.5), Inches(3.1), Inches(1.7), Inches(1.7))
earth.fill.solid(); earth.fill.fore_color.rgb = BLUE
earth.line.color.rgb = BLUE
# sun
sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.5), Inches(2.1), Inches(1.1), Inches(1.1))
sun.fill.solid(); sun.fill.fore_color.rgb = RGBColor(255, 196, 0)
sun.line.color.rgb = RGBColor(255, 196, 0)
# satellite
sat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.9), Inches(3.8), Inches(0.8), Inches(0.4))
sat.fill.solid(); sat.fill.fore_color.rgb = ORANGE
sat.line.color.rgb = ORANGE
# lines
ln1 = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(8.3), Inches(3.7), Inches(8.9), Inches(3.9))
ln1.line.color.rgb = GRAY
ln2 = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(8.9), Inches(3.9), Inches(10.6), Inches(3.1))
ln2.line.color.rgb = GRAY

# shadow cone
shadow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(5.8), Inches(2.6), Inches(2.1), Inches(2.6))
shadow.fill.solid(); shadow.fill.fore_color.rgb = RGBColor(190, 215, 255)
shadow.line.color.rgb = BLUE
# labels
for label, x, y in [("Earth", 7.3, 5.2), ("Sun", 10.4, 1.4), ("GEO satellite", 8.7, 5.2)]:
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(2.0), Inches(0.4))
    p = tb.text_frame.paragraphs[0]
    p.text = label
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(12)
    p.runs[0].font.color.rgb = TEXT

# add formula box
add_formula_box(slide, "T = 2π√(a³/μ)", 0.8, 5.8, 3.2, 1.1, title="Kepler: orbital period")
add_formula_box(slide, "a ≈ 42164 km (GEO)", 4.3, 5.8, 2.6, 1.1, title="Typical radius")


# Slide 4: Eclipse
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "Eclipse: Satellite Enters Earth Shadow")
add_bullets(slide, [
    "When the satellite passes through the Earth’s umbra, direct sunlight is blocked.",
    "The spacecraft loses solar power and must rely on batteries or reduced load operations.",
    "The effect is seasonal and strongest near the equinoxes.",
    "Eclipse duration depends on Earth radius, orbit radius, and solar declination."
], 0.8, 1.8, 5.8, 3.8, font_size=20)

# geometry diagram
# earth
earth = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.3), Inches(2.8), Inches(1.8), Inches(1.8))
earth.fill.solid(); earth.fill.fore_color.rgb = BLUE
earth.line.color.rgb = BLUE
# satellite path
path = slide.shapes.add_shape(MSO_SHAPE.LINE, Inches(8.8), Inches(3.7), Inches(10.9), Inches(3.7))
# using shape line isn't straightforward, so create simple polygon
# draw shadow cone as polygon
shadow = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(5.7), Inches(2.2), Inches(2.4), Inches(3.0))
shadow.fill.solid(); shadow.fill.fore_color.rgb = RGBColor(220, 230, 255)
shadow.line.color.rgb = BLUE
# sat in shadow at right
sat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.7), Inches(3.3), Inches(0.7), Inches(0.4))
sat.fill.solid(); sat.fill.fore_color.rgb = ORANGE
sat.line.color.rgb = ORANGE
# sun
sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.5), Inches(1.5), Inches(1.3), Inches(1.3))
sun.fill.solid(); sun.fill.fore_color.rgb = RGBColor(255, 196, 0)
sun.line.color.rgb = RGBColor(255, 196, 0)
# labels
for label, x, y in [("Earth", 7.3, 5.0), ("Sun", 10.5, 0.8), ("Satellite in eclipse", 10.2, 4.3)]:
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(2.6), Inches(0.4))
    p = tb.text_frame.paragraphs[0]
    p.text = label
    p.runs[0].font.size = Pt(12)
    p.runs[0].font.color.rgb = TEXT

# formulas
add_formula_box(slide, "t_ecl = (2/ω) arccos((R_e/R_s) cos α)", 1.0, 5.8, 5.0, 1.2, title="Eclipse duration (approx.)")
add_formula_box(slide, "β = arcsin(R_e / d)", 6.5, 5.8, 2.8, 1.2, title="Shadow half-angle")


# Slide 5: Sun Transit Outage
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "Sun Transit Outage: Sun Behind the Satellite")
add_bullets(slide, [
    "The sun appears behind the satellite as seen from the Earth station.",
    "The Sun is a bright radio-noise source and overwhelms the weak satellite signal.",
    "This creates temporary signal degradation or communication blackout.",
    "The event is highly seasonal and depends on the angular separation between the Sun and satellite."
], 0.8, 1.8, 5.8, 3.8, font_size=20)

# diagram
# earth station
station = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.3), Inches(3.0), Inches(1.0), Inches(1.0))
station.fill.solid(); station.fill.fore_color.rgb = BLUE
station.line.color.rgb = BLUE
# satellite
sat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.8), Inches(3.3), Inches(0.8), Inches(0.4))
sat.fill.solid(); sat.fill.fore_color.rgb = ORANGE
sat.line.color.rgb = ORANGE
# sun
sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.4), Inches(2.0), Inches(1.2), Inches(1.2))
sun.fill.solid(); sun.fill.fore_color.rgb = RGBColor(255, 196, 0)
sun.line.color.rgb = RGBColor(255, 196, 0)
# line to station
line = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(3.3), Inches(3.5), Inches(7.8), Inches(3.5))
line.line.color.rgb = GRAY
# sun line behind satellite
line2 = slide.shapes.add_connector(MSO_SHAPE.LINE, Inches(7.8), Inches(3.5), Inches(9.7), Inches(2.8))
line2.line.color.rgb = GRAY
# labels
for label, x, y in [("Earth station", 2.0, 5.0), ("Satellite", 7.4, 4.6), ("Sun", 9.3, 0.9)]:
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(2.2), Inches(0.4))
    p = tb.text_frame.paragraphs[0]
    p.text = label
    p.runs[0].font.size = Pt(12)
    p.runs[0].font.color.rgb = TEXT

# formula box
add_formula_box(slide, "θ = arccos[sin El_s sin El_Sun + cos El_s cos El_Sun cos(Az_s − Az_Sun)]", 0.7, 5.7, 7.2, 1.4, title="Angular separation between satellite and Sun")


# Slide 6: RF link impact
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "RF Link Impact and Noise Increase")
add_bullets(slide, [
    "The Sun contributes additional noise temperature to the receiving system.",
    "As the noise power rises, the carrier-to-noise ratio drops sharply.",
    "A small angular separation can result in strong degradation even when the satellite itself is healthy.",
    "This effect is why the outage window is accurately predicted and scheduled."
], 0.8, 1.8, 5.8, 3.8, font_size=20)

# diagram: C/N relationship
chart = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.2), Inches(2.3), Inches(3.7), Inches(3.1))
chart.fill.solid(); chart.fill.fore_color.rgb = WHITE
chart.line.color.rgb = BLUE
# simple line graph inside
line_shape = slide.shapes.add_shape(MSO_SHAPE.LINE, Inches(8.6), Inches(4.8), Inches(11.4), Inches(2.8))
line_shape.line.color.rgb = ORANGE
line_shape.line.width = Pt(2.5)
# formula box
add_formula_box(slide, "T_sys = T_antenna + T_sky + T_sun", 0.8, 5.8, 4.5, 1.1, title="System noise temperature")
add_formula_box(slide, "C/N = C / (N + N_sun)", 5.9, 5.8, 3.2, 1.1, title="Carrier-to-noise")


# Slide 7: Comparison table
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "Eclipse vs Sun Transit Outage")
# table background
left = Inches(0.8)
top = Inches(1.8)
width = Inches(11.8)
height = Inches(4.0)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
shape.fill.solid(); shape.fill.fore_color.rgb = WHITE
shape.line.color.rgb = BLUE
shape.line.width = Pt(1.5)
# table text
box = slide.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(11.2), Inches(3.4))
tf = box.text_frame
rows = [
    ["Aspect", "Eclipse", "Sun Transit Outage"],
    ["Primary cause", "Satellite enters Earth shadow", "Sun is behind satellite as seen from Earth station"],
    ["Main impact", "Power loss / battery use", "RF noise / signal degradation"],
    ["Time window", "Seasonal near equinox", "Seasonal near equinox"],
    ["Typical outcome", "Reduced power availability", "Temporary outage / link fade"],
]
for r_idx, row in enumerate(rows):
    p = tf.paragraphs[0] if r_idx == 0 else tf.add_paragraph()
    p.text = " | ".join(row)
    p.runs[0].font.size = Pt(16 if r_idx == 0 else 15)
    p.runs[0].font.bold = (r_idx == 0)
    p.runs[0].font.color.rgb = NAVY


# Slide 8: References and conclusion
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "Conclusion and References")
add_bullets(slide, [
    "Eclipse and sun transit outage are both caused by specific orbital alignments of the Sun, Earth, and satellite.",
    "Eclipse affects satellite power; sun transit outage affects Earth-station communication quality.",
    "Both are predictable using orbital geometry, solar declination, and link-budget calculations.",
    "Satellite operators use seasonal forecasts to plan outage windows and power margins.",
], 0.8, 1.8, 6.8, 3.6, font_size=20)

ref_box = slide.shapes.add_textbox(Inches(8.1), Inches(2.2), Inches(4.1), Inches(3.0))
ref_tf = ref_box.text_frame
for i, item in enumerate([
    "Dennis Roddy, Satellite Communications",
    "Spacecraft systems engineering texts",
    "ITU and operator eclipse prediction data",
    "Orbital mechanics / link budget references"
]):
    p = ref_tf.paragraphs[0] if i == 0 else ref_tf.add_paragraph()
    p.text = item
    p.runs[0].font.size = Pt(13)
    p.runs[0].font.color.rgb = TEXT


# Save
out_path = "Orbital_Effects_Eclipse_and_Sun_Transit_Outage.pptx"
prs.save(out_path)
print(f"PowerPoint created: {out_path}")
