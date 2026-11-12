"""Build the Tideline project report on the SRM template (front pages kept) as a Word document.

Run from the tideline folder:  python report/build_report.py
Input:  report/template_srm.docx (copy of the SRM case-study template with its logos and front pages)
Output: report/Tideline_Project_Report.docx
Formatting follows the SRMIST specification in the template: Times New Roman; chapter headings 16 pt bold,
section 14 pt bold, sub-section 12 pt bold, body 12 pt; 1.5 line spacing; margins L 1.5", R/T/B 1.0";
Roman page numbers for the front matter and Arabic from the first chapter.
"""
import copy
import csv
import json
import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX  # noqa: F401
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "outputs")
SCR = os.path.join(HERE, "screens_small")
D = json.load(open(os.path.join(OUT, "web_data.json"), encoding="utf-8"))
M = D["meta"]
TNR = "Times New Roman"

TITLE = "TIDELINE: SEA-LEVEL RISE AND COASTAL PROTECTION"
COURSE = "21CEO310T – Global Warming and Climate Change Working Model Project"
STUDENT, REGNO = "AKSHANSH SHARMA", "RA2411003012268"

doc = Document(os.path.join(HERE, "template_srm.docx"))
body = doc.element.body


# ------------------------------------------------------------------ front matter: fill the template
def para(i):
    return Paragraph(list(body.iterchildren())[i], doc)


def set_text(p, text, keep_highlight=False):
    """Put text in the first run (keeping its font) and empty the others."""
    runs = p.runs
    runs[0].text = text
    if not keep_highlight:
        runs[0].font.highlight_color = None
    for r in runs[1:]:
        r.text = ""


els = list(body.iterchildren())
P0 = {i: Paragraph(e, doc) for i, e in enumerate(els) if e.tag == qn("w:p")}
set_text(P0[0], TITLE)
P0[0].runs[0].font.size = Pt(22)
set_text(P0[1], COURSE)
set_text(P0[3], f"{STUDENT} [{REGNO}]")
# faculty name, designation and department stay highlighted for the student to fill in
bon = P0[27]
for r in bon.runs:
    r.text = ""
r0 = bon.runs[0]
r0.text = (f"This is to certify that the 21CEO310T – Global Warming and Climate Change working model report titled "
           f"“Tideline: Sea-Level Rise and Coastal Protection” is the bonafide work of {STUDENT} ({REGNO}), who undertook the task of "
           f"completing the working model and this report within the allotted time.")
r0.font.highlight_color = None
abstract_p = P0[44]
ABSTRACT = (
    "Tideline is a software working model that forecasts sea-level rise and compares how well mangroves, raised houses and "
    "embankments protect a low-lying Indian coastal town up to 2100. Four official datasets are used: the CSIRO tide-gauge "
    "reconstruction of global mean sea level (1880–2019), NOAA satellite altimetry (1993–2025), the PSMSL tide gauge at Mumbai and "
    "NASA GISS global temperature. After quality checks they are joined into one annual record of 145 years. Linear Regression, "
    "Support Vector Regression and Random Forest are trained on 1880–1999 and tested on 2000–2024, years the models never see. The "
    "novelty is physics-informed machine learning: the input features are the terms of the semi-empirical sea-level law, in which the "
    "rate of rise follows global warming. With these features the models reach R² = 0.99 and a mean absolute error of about 2 mm, "
    "while a plain Random Forest scores R² = −2.88 because it cannot predict beyond its training range. The learned coefficients are "
    "physically meaningful, giving a present-day rate of 4.0 mm per year against 3.7 mm per year measured by satellites. Projections "
    "to 2100 give +0.43, +0.50 and +0.62 m under the SSP1-2.6, SSP2-4.5 and SSP5-8.5 pathways, consistent with the IPCC AR6 for low "
    "and middle emissions. The forecast is localised to Mumbai and used to flood a model town of 1,200 homes under storm surges. "
    "Mangroves give the best benefit-cost ratio (about 29) and capture carbon, but drown when the sea rises faster than 7 mm per "
    "year; embankments keep the town dry until they are overtopped and then fail all at once; a hybrid of mangroves and a lower "
    "wall is the most robust. The results are presented in an interactive 3D website.")
assert len(ABSTRACT.split()) <= 300, len(ABSTRACT.split())
for r in abstract_p.runs:
    r.text = ""
ra = abstract_p.runs[0]
ra.text = ABSTRACT; ra.font.highlight_color = None; ra.font.name = TNR; ra.font.size = Pt(14)
ra._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), TNR)
abstract_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
abstract_p.paragraph_format.line_spacing = 1.5

# remove the template's spacer paragraphs and its guidance pages; start new pages explicitly instead
idx_abs_title, idx_toc_title, idx_toc_tbl = 43, 67, 68
for i, e in enumerate(els):
    if i > idx_toc_tbl and e.tag != qn("w:sectPr"):
        body.remove(e)
for i in list(range(45, 67)) + list(range(33, 43)):
    if els[i].tag == qn("w:p") and not P0[i].text.strip() and not els[i].xpath(".//w:sectPr|.//w:drawing"):
        body.remove(els[i])
P0[idx_abs_title].paragraph_format.page_break_before = True
P0[idx_toc_title].paragraph_format.page_break_before = True
toc = doc.tables[-1]

# section break after the contents: front matter in Roman numerals, body in Arabic from 1
body_sect = body.find(qn("w:sectPr"))
front_sect = copy.deepcopy(body_sect)


def pgnum(sect, fmt, start=None):
    for old in sect.findall(qn("w:pgNumType")):
        sect.remove(old)
    pn = OxmlElement("w:pgNumType"); pn.set(qn("w:fmt"), fmt)
    if start is not None:
        pn.set(qn("w:start"), str(start))
    # schema order: pgNumType goes before cols/titlePg/docGrid
    anchor = sect.find(qn("w:cols"))
    (anchor.addprevious(pn) if anchor is not None else sect.append(pn))


pgnum(front_sect, "lowerRoman")
first_sect = els[19].find(".//" + qn("w:sectPr"))
if first_sect is not None:
    pgnum(first_sect, "lowerRoman")
brk = OxmlElement("w:p"); ppr = OxmlElement("w:pPr"); ppr.append(front_sect); brk.append(ppr)
toc._tbl.addnext(brk)
pgnum(body_sect, "decimal", 1)
s = doc.sections[-1]
s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = Inches(1.5), Inches(1.0), Inches(1.0), Inches(1.0)

# ------------------------------------------------------------------ body styles (SRMIST specification)
body_style = doc.styles.add_style("TL Body", 1)
body_style.base_style = doc.styles["Normal"]
body_style.font.name = TNR; body_style.font.size = Pt(12)
body_style.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), TNR)
pf = body_style.paragraph_format
pf.line_spacing = 1.5; pf.space_after = Pt(6); pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for lvl, size in [(1, 16), (2, 14), (3, 12)]:
    h = doc.styles[f"Heading {lvl}"]
    h.font.name = TNR; h.font.size = Pt(size); h.font.bold = True; h.font.italic = False
    h.font.color.rgb = RGBColor(0, 0, 0)
    h.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), TNR)
    rf = h.element.get_or_add_rPr().get_or_add_rFonts()
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        rf.attrib.pop(qn(a), None)
    h.paragraph_format.space_before = Pt(18 if lvl == 1 else 12); h.paragraph_format.space_after = Pt(8)
    h.paragraph_format.line_spacing = 1.5
for name in ("List Bullet", "List Number", "Caption"):
    st = doc.styles[name]; st.font.name = TNR; st.font.size = Pt(12 if name != "Caption" else 10.5)
    st.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), TNR)
    if name != "Caption":
        st.paragraph_format.line_spacing = 1.5; st.paragraph_format.space_after = Pt(3)

FIG = [0]
TAB = [0]
TOCROWS = []          # (number, title, level, bookmark)
BM = [100]


def bookmark(p, name):
    BM[0] += 1
    s_ = OxmlElement("w:bookmarkStart"); s_.set(qn("w:id"), str(BM[0])); s_.set(qn("w:name"), name)
    e_ = OxmlElement("w:bookmarkEnd"); e_.set(qn("w:id"), str(BM[0]))
    p._p.insert(1 if p._p.pPr is not None else 0, s_); p._p.append(e_)


def H(text, lvl=1, new_page=None):
    h = doc.add_heading(text, level=lvl)
    if lvl == 1 and (new_page if new_page is not None else True) and sum(1 for t in TOCROWS if t[2] == 1) > 1:
        h.paragraph_format.page_break_before = True
    if lvl <= 2:
        m = re.match(r"^(\d+(?:\.\d+)?)\.?\s+(.*)$", text)
        num, title = (m.group(1), m.group(2)) if m else ("", text)
        name = f"_TL{len(TOCROWS)}"
        bookmark(h, name); TOCROWS.append((num, title, lvl, name))
    return h


def P(text, bold_lead=None, italic=False, align=None):
    p = doc.add_paragraph(style="TL Body")
    if bold_lead:
        r = p.add_run(bold_lead); r.bold = True
    r = p.add_run(text); r.italic = italic
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def B(items, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if isinstance(it, tuple):
            r = p.add_run(it[0]); r.bold = True; p.add_run(it[1])
        else:
            p.add_run(it)


def EQ(text):
    p = doc.add_paragraph(style="TL Body"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.font.name = "Cambria Math"; r.italic = True


def shade(cell, hexcol):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hexcol); tcPr.append(sh)


def cell_text(c, text, bold=False, size=10):
    c.text = ""; p = c.paragraphs[0]; r = p.add_run(str(text)); r.bold = bold; r.font.size = Pt(size); r.font.name = TNR
    p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(0)
    return r


def T(header, rows, caption, widths=None, bold_rows=()):
    TAB[0] += 1
    cp = doc.add_paragraph(style="TL Body"); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cp.add_run(f"Table {TAB[0]}: {caption}"); r.bold = True; r.font.size = Pt(11)
    t = doc.add_table(rows=1, cols=len(header)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        cell_text(t.rows[0].cells[i], h, bold=True); shade(t.rows[0].cells[i], "D9E8E6")
    for k, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cell_text(cells[i], v, bold=k in bold_rows)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph(style="TL Body")


def F(path, caption, width=14.5):
    if not os.path.exists(path):
        print("  (missing figure, skipped):", os.path.basename(path)); return
    FIG[0] += 1
    pic = doc.add_paragraph(style="TL Body"); pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.paragraph_format.keep_with_next = True; pic.paragraph_format.line_spacing = 1.0
    pic.add_run().add_picture(path, width=Cm(min(width, 14.5)))
    cp = doc.add_paragraph(style="TL Body"); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cp.add_run(f"Figure {FIG[0]}: {caption}"); r.italic = True; r.font.size = Pt(10.5)


REPO = "https://github.com/Akinc5/tideline-sea-level"
NB_GITHUB = REPO + "/blob/main/Tideline_Colab.ipynb"
NB_COLAB = "https://colab.research.google.com/github/Akinc5/tideline-sea-level/blob/main/Tideline_Colab.ipynb"


def LINK(label, url):
    """A paragraph 'label: url' where the url is a clickable hyperlink."""
    p = doc.add_paragraph(style="TL Body")
    p.add_run(label + ": ").bold = True
    rid = doc.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                             is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1"); rpr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    r.append(rpr); t = OxmlElement("w:t"); t.text = url; r.append(t); h.append(r); p._p.append(h)


def fo(name):
    return os.path.join(OUT, name)


ALIAS = {"web01_hero_2024": "s1_hero_2024", "web02_1880": "s2_1880", "web04_2100_forecast": "s3_2100_calm",
         "web05_2100_storm": "s4_2100_storm", "web06_2050_defended": "s5_2050_defended"}


def sc(name):
    p = os.path.join(SCR, name + ".jpg")
    if not os.path.exists(p) and name in ALIAS:   # fall back to the direct canvas captures
        p = os.path.join(SCR, ALIAS[name] + ".jpg")
    return p


def rows_csv(name):
    with open(fo(name), encoding="utf-8") as f:
        return list(csv.reader(f))


f2 = lambda v, d=2: f"{float(v):.{d}f}"

# the abstract page is section 2; register it first so it heads the contents
bookmark(P0[idx_abs_title], "_TLabs")
TOCROWS.append(("2", "Abstract of the Working Model", 1, "_TLabs"))

# ================================================================== 3. AIM / INTRODUCTION
H("3. Aim of the Working Model / Introduction")
H("3.1 Introduction", 2)
P("Global mean sea level has risen by about 0.20 m since 1901, and the rate of rise has roughly tripled, from about 1.3 mm per year "
  "in 1901–1971 to 3.7 mm per year in 2006–2018 (IPCC AR6 WG1). Two processes, both caused by global warming, explain most of this rise. "
  "As the ocean absorbs heat, sea water expands (thermal expansion). Melting glaciers and the Greenland and Antarctic ice sheets also add "
  "water to the ocean. The IPCC estimates that between 1971 and 2018 thermal expansion contributed about 50% of the rise, glaciers 22%, "
  "ice sheets 20% and changes in land water storage 8%.", align="j")
P("India has a coastline of more than 7,500 km, and cities such as Mumbai, Kolkata and Chennai sit only a few metres above high tide. "
  "For these coasts the danger is not the slow rise alone: a higher sea lets every storm surge and high tide reach further inland. A "
  "flood that used to happen once a century can become a once-a-decade event. Coastal planners therefore need two things: a trustworthy "
  "forecast of the rise, and an honest comparison of the ways to protect a town against it.", align="j")
P("There are three main families of coastal defence. Nature-based (\"green\") defences such as mangrove forests absorb wave energy and "
  "trap sediment. Engineered (\"grey\") defences such as embankments and sea walls block the water. Accommodation measures such as "
  "raising houses on plinths let the water come but reduce the damage it causes. Each behaves very differently as the sea keeps rising, "
  "and a single snapshot comparison can hide how and when each one fails.", align="j")
H("3.2 Aim", 2)
P("To build an explainable machine-learning working model that forecasts global and local sea-level rise from real observational data, "
  "and uses that forecast to simulate and compare the performance, cost and failure point of mangroves, raised structures and "
  "embankments for a low-lying Indian coastal town up to the year 2100.", align="j")
H("3.3 Objectives", 2)
B([
    "Collect sea-level and temperature data from official sources (CSIRO, NOAA, PSMSL, NASA GISS) and report the quality of every row "
    "and column used.",
    "Train and compare Linear Regression, Support Vector Regression and Random Forest, and evaluate them with MAE, MSE, RMSE and R² on "
    "years they have never seen.",
    "Make the model explainable by building its features from the physics of sea-level rise, and check that the learned coefficients "
    "make physical sense.",
    "Project sea level to 2100 under three IPCC warming pathways with uncertainty, and compare the results with the IPCC AR6 assessment.",
    "Convert the global forecast into a local one for Mumbai, then flood a model town under storms of different return periods.",
    "Compare mangroves, an embankment, raised houses and a hybrid using homes flooded, expected annual damage, cost, benefit-cost ratio, "
    "carbon captured and the year each defence fails.",
    "Present everything in an interactive 3D website so that anyone can explore the results.",
], "List Number")
H("3.4 Novelty", 2)
B([
    ("Physics-informed ML. ", "The features are the terms of the semi-empirical sea-level law, so the Linear Regression model is the "
     "published scientific law with coefficients learned from data, and the ML models only learn the small residual."),
    ("Honest validation. ", "A chronological train/test split, rolling-origin windows and a deliberate random-split experiment show how "
     "shuffled time-series data makes models look far better than they are."),
    ("From global numbers to houses. ", "The forecast is localised with the Mumbai tide gauge and used to flood individual homes, "
     "joining climate science, machine learning and engineering economics in one model."),
    ("Time to failure, not a snapshot. ", "Each defence is run year by year to 2100 to find when it stops working; for mangroves this "
     "includes drowning when the sea rises faster than they can build up sediment."),
    ("Interactive 3D. ", "A scroll-driven 3D coastal diorama and a live simulator let viewers change the defence, pathway, storm and "
     "year and see the town flood."),
])

# ================================================================== 4. MATERIALS / METHODS
H("4. Materials / Methods Used")
P("Tideline is a fully software-based working model: no physical hardware is needed. The materials are datasets, software tools and "
  "scientific methods.")
H("4.1 Software and tools", 2)
T(["Tool / library", "Purpose in the model"], [
    ["Google Colab / Jupyter Notebook", "Cloud environment where the whole model runs (Runtime → Run all)"],
    ["Python 3", "Programming language for data processing, modelling and simulation"],
    ["pandas, NumPy", "Reading, cleaning, joining and transforming the datasets"],
    ["scikit-learn", "Linear Regression, SVR, Random Forest, TimeSeriesSplit, GridSearchCV, MAE/MSE/R² metrics"],
    ["Matplotlib", "Figures 1–8 of the analysis"],
    ["three.js (WebGL)", "3D coastal diorama: terrain, water shader, houses, mangroves, embankment"],
    ["D3.js", "Interactive charts on the website"],
    ["HTML, CSS, JavaScript", "Website, parallax scroll story and the in-browser flood simulator"],
    ["Netlify", "Hosting the public website (tideline-sea-level.netlify.app)"],
    ["GitHub", "Source code and notebook (github.com/Akinc5/tideline-sea-level), opened in Colab from there"],
], "Software used.", widths=[5, 11])
H("4.2 Datasets", 2)
dq = rows_csv("data_quality_report.csv")
T(dq[0][:9], [r[:9] for r in dq[1:]], "Data sources and data-quality report (rows, columns, coverage and missing years).")
B([
    ("CSIRO GMSL reconstruction (Church & White): ", "global mean sea level from tide gauges with 1σ uncertainty, 1880–2019."),
    ("NOAA Laboratory for Satellite Altimetry: ", "global mean sea level from TOPEX/Poseidon, Jason-1/2/3 and Sentinel-6 satellites, "
     "1993–2025, averaged to annual values."),
    ("PSMSL station 43, Mumbai (Apollo Bandar): ", "annual relative sea level (Revised Local Reference). 25 years are flagged by PSMSL "
     "for too many missing days and are excluded."),
    ("NASA GISS GISTEMP v4: ", "global annual surface temperature anomaly, used as the driver of sea-level rise."),
])
P(f"The final modelling table has {M['rows']} rows (years 1880–2024) and {M['columns']} columns: year, global mean sea level, "
  f"temperature T, 3-year smoothed temperature T3, cumulative warming ∫T dt, centred year, dT/dt, and the Mumbai level.", align="j")
H("4.3 Methods", 2)
B([
    ("Data splicing: ", f"tide gauges are used up to 1992 and satellites from 1993. The satellite series is shifted by "
     f"{M['offset_mm']:.0f} mm, its mean difference from CSIRO over the 1993–2019 overlap, so the record has no jump "
     f"(overlap RMSE {M['splice_rmse']:.1f} mm). All levels are relative to the 1995–2014 mean, the IPCC AR6 baseline."),
    ("Semi-empirical sea-level law: ", "the physical basis of the features (Section 5)."),
    ("Machine-learning algorithms: ", "Linear Regression, Support Vector Regression (RBF kernel, tuned C, ε, γ) and Random Forest "
     "(tuned depth and leaf size), each trained three ways (strategies A, B and C)."),
    ("Validation: ", "chronological split (train 1880–1999, test 2000–2024), TimeSeriesSplit cross-validation for tuning, "
     "rolling-origin windows (1960–79, 1980–99, 2000–24) and a random-split leakage test."),
    ("Evaluation metrics: ", "Mean Absolute Error (MAE), Mean Squared Error (MSE), Root Mean Squared Error (RMSE) and R² score."),
    ("Projections: ", "integration of the fitted law under SSP temperature pathways, with a moving-block bootstrap (500 resamples, "
     "10-year blocks) for 90% uncertainty bands."),
    ("Flood and defence simulation: ", "Gumbel extreme-value storm surge, bathtub flooding, exponential wave attenuation by mangroves, "
     "embankment overtopping, the JRC depth-damage curve, expected annual damage and discounted cost-benefit analysis."),
])

# ================================================================== 5. WORKING PRINCIPLE
H("5. Working Principle")
P("Tideline works in four linked stages: physics tells the model what to learn, machine learning learns how much from real data, the "
  "forecast is localised to one coast, and a flood simulator turns the rise into homes flooded and money saved by each defence.",
  align="j")
H("5.1 Stage 1: why the sea rises (physics)", 2)
P("Warmer water expands and warmer air melts ice, so the speed of sea-level rise depends on how much warmer the planet is than an "
  "equilibrium temperature. Rahmstorf (2007) and Vermeer & Rahmstorf (2009) wrote this as:", align="j")
EQ("dH/dt = a · (T − T₀) + b · dT/dt")
P("where H is global mean sea level, T is the global temperature anomaly, T₀ is the temperature at which the sea would stop rising, "
  "a is the sensitivity (mm per year of rise per °C), and b captures a fast short-term response. Integrating over time gives the "
  "sea level itself:", align="j")
EQ("H(t) = a · ∫T dt − a · T₀ · t + b · T + c")
P("So the sea level depends on the accumulated warming ∫T dt, not just the current temperature. These terms (cumulative warming, "
  "centred year and temperature) become the input features of the machine-learning models.", align="j")
H("5.2 Stage 2: machine learning that follows the physics", 2)
B([
    ("Strategy A, pure ML: ", "inputs are only year and temperature. The model must discover the physics by itself."),
    ("Strategy B, physics features: ", "inputs are ∫T dt, centred year and smoothed temperature. Linear Regression on these "
     "features is exactly the semi-empirical law, with a, T₀ and b learned from data."),
    ("Strategy C, physics-guided hybrid: ", "Linear Regression provides the physical backbone, and SVR or Random Forest learns only "
     "the small residual left over, using temperature and its rate of change."),
])
P("Random Forest cannot predict values larger than anything in its training data, because every prediction is an average of training "
  "targets. The sea keeps reaching new record highs, so a plain Random Forest fails on the future. The hybrid fixes this: the physics "
  "backbone extrapolates and the forest only corrects small wiggles.", align="j")
P("Evaluation metrics, where yᵢ is the observed and ŷᵢ the predicted sea level for n test years:", align="j")
EQ("MAE = (1/n) Σ |yᵢ − ŷᵢ|        MSE = (1/n) Σ (yᵢ − ŷᵢ)²        R² = 1 − Σ(yᵢ − ŷᵢ)² / Σ(yᵢ − ȳ)²")
H("5.3 Stage 3: from global to local", 2)
P("Future temperatures for each SSP pathway are fed into the fitted law and integrated year by year to 2100. Uncertainty is estimated "
  "with a moving-block bootstrap: the model is refitted 500 times on resampled 10-year blocks of the record. The Mumbai tide gauge rises "
  f"{M['mumbai_trend']:.2f} mm/yr against a global {M['global_trend']:.2f} mm/yr over the same years, because of vertical land motion "
  f"and local ocean effects. The local rise is therefore:", align="j")
EQ(f"local rise = global rise + β · (year − 2024),   β = {M['beta']:.2f} mm/yr")
H("5.4 Stage 4: flood and defence simulator", 2)
P("For each year and storm, the total water level reaching the town is:", align="j")
EQ("water level = high tide (MHHW) + local sea-level rise + storm surge(T) + wave set-up − mangrove surge reduction")
B([
    ("Storm surge: ", "a T-year storm (for example 1-in-100, a 1% chance each year) comes from a Gumbel extreme-value distribution: "
     "surge(T) = μ − β · ln(−ln(1 − 1/T))."),
    ("Flooding: ", "a home floods if the water level is above the ground under it (bathtub model). Water depth = level − ground."),
    ("Damage: ", "the JRC depth-damage curve gives the fraction of a home's value lost at each depth. Expected annual damage (EAD) is "
     "the damage averaged over storms of all return periods."),
    ("Mangroves: ", "wave height decays exponentially with forest width W (H = H₀·e^(−kW)) and the surge drops by a fixed amount per "
     "km of forest. If local sea level rises faster than 7 mm/yr, the forest cannot keep up by trapping sediment and starts to drown "
     "(5% health lost per year). Mangroves also capture about 6.4 t of CO₂ per hectare per year."),
    ("Embankment: ", "blocks all water until the water level plus wave run-up exceeds its crest height, then the town floods as if "
     "there were no wall (overtopping)."),
    ("Raised houses: ", "each home is lifted on a plinth, so the damaging depth is reduced by the plinth height; the street still floods."),
    ("Economics: ", "benefit-cost ratio BCR = present value of damage avoided ÷ (build cost + present value of maintenance), "
     "discounted at 4% per year over 2025–2100. BCR > 1 means the defence is worth its cost."),
])
F(sc("web01_hero_2024"), "The working model: a 3D coastal town on India's west coast, 1,200 homes with mangroves, jetty, "
  "lighthouse and fishing boats. The sea level shown is driven by the ML forecast.")

# ================================================================== 6. CONSTRUCTION / STEPS
H("6. Construction / Steps Used / Screenshots of the ML Model")
P("The model was constructed step by step in a single Google Colab notebook (Tideline_Colab.ipynb), whose outputs are then exported "
  "to the website. The steps below follow the notebook cells in order. The complete code, with every executed output, can be opened "
  "directly in Google Colab from the project's GitHub repository (Runtime → Run all reproduces every result in this report):",
  align="j")
LINK("Open in Google Colab", NB_COLAB)
LINK("Notebook on GitHub", NB_GITHUB)
LINK("Project repository", REPO)
steps = [
    ("Step 1: Set up the environment.", "Import pandas, NumPy, Matplotlib and scikit-learn, fix the random seed, and set the key "
     "options: USE_LIVE_DATA, the training cut-off TRAIN_END = 1999 and the 1995–2014 baseline.", ["nb00_intro", "nb01_setup"]),
    ("Step 2: Acquire the data.", "Download the four datasets live from CSIRO (GitHub mirror), NOAA STAR, PSMSL and NASA GISS. If a "
     "server is offline, the notebook falls back to the snapshot saved in the data/ folder.", ["nb02_fetch"]),
    ("Step 3: Parse and check data quality.", "Parse each file, convert units, drop flagged years and print a data-quality report "
     "with rows, columns, coverage, missing years and ranges (Table 2).", ["nb03_data_quality"]),
    ("Step 4: Build one sea-level record.", "Join the tide-gauge and satellite records, aligning the satellites to CSIRO over their "
     "1993–2019 overlap, and re-reference everything to the 1995–2014 mean.", ["nb04_splice"]),
    ("Step 5: Engineer physics-informed features.", "Compute smoothed temperature T3, cumulative warming ∫T dt, centred year and "
     "dT/dt, the terms of the semi-empirical law.", ["nb05_features", "nb06_features_plot"]),
    ("Step 6: Train the models.", "Train Linear Regression, SVR and Random Forest under strategies A, B and C. SVR and Random Forest "
     "hyper-parameters are tuned with GridSearchCV and TimeSeriesSplit inside the training years only. The hybrid class adds an ML "
     "residual to the physics backbone.", ["nb07_models"]),
    ("Step 7: Evaluate with MAE, MSE, RMSE and R².", "Score every model on the unseen 2000–2024 test years and plot the comparison.",
     ["nb08_metrics"]),
    ("Step 8: Robustness checks.", "Rolling-origin validation on three windows and a random-split experiment that shows the leakage "
     "problem of shuffling time-series data.", ["nb09_robustness"]),
    ("Step 9: Explain the model.", "Read the physical constants a, T₀ and b from the Linear Regression coefficients and compare the "
     "implied rate of rise with satellite observations.", ["nb10_explainability"]),
    ("Step 10: Localise to Mumbai.", "Compare the Mumbai tide-gauge trend with the global trend over the same years and derive the "
     "local offset β.", ["nb11_mumbai"]),
    ("Step 11: Project to 2100.", "Integrate the law under SSP1-2.6, SSP2-4.5 and SSP5-8.5 temperature paths with 500 block-bootstrap "
     "resamples and compare with IPCC AR6.", ["nb12_projections", "nb13_projection_plot"]),
    ("Step 12: Build the defence simulator.", "Define the town, the sea and storm statistics, the four defences and their costs; run "
     "every defence for every year and pathway.", ["nb14_simulator", "nb15_simulator_plot"]),
    ("Step 13: Export and build the website.", "Export all results to web_data.json, which build_web.py inlines into the 3D website. "
     "The website re-implements the flood model in JavaScript and reproduces the notebook's numbers exactly.", []),
]
for title, text, shots in steps:
    H(title, 3)
    P(text, align="j")
H("6.1 Screenshots of the working website", 2)
for name, cap in [
    ("web02_1880", "Scroll story, 1880: the sea 27 cm lower than in 2024, from the CSIRO tide-gauge record."),
    ("web03_2024_rate", "Scroll story, 2024: satellites measure 3.7 mm of rise per year."),
    ("web04_2100_forecast", "Scroll story, 2100: the ML forecast for a high-emissions path."),
    ("web05_2100_storm", "Scroll story, 2100 with a 1-in-100-year storm and no defences: flooded homes turn red."),
    ("web06_2050_defended", "Scroll story, 2050: a mangrove belt and a 3.8 m embankment hold the same storm."),
    ("web08_gmsl_chart", "Website: the joined global sea-level record, 1880–2024."),
    ("web10_law_metrics", "Website: the semi-empirical law and the test-set metrics table."),
    ("web12_projections", "Website: projections to 2100 against IPCC AR6 likely ranges."),
    ("web13_sim_hybrid", "Website simulator: hybrid defence in 2070 under SSP2-4.5, embankment holding."),
    ("web14_sim_embankment_2092", "Website simulator: the embankment overtopped in 2092, homes flooded."),
    ("web15_sim_mangrove_ssp585", "Website simulator: mangroves drowned by 2100 under SSP5-8.5."),
    ("web17_defence_cards", "Website: comparison cards for the five defence options."),
]:
    F(sc(name), cap, width=15.5)

# ================================================================== 7. OUTCOME / RESULTS
H("7. Outcome / Results")
H("7.1 Model performance on unseen years (2000–2024)", 2)
mm = rows_csv("model_metrics_test.csv")
T(["Model", "MAE (mm)", "MSE (mm²)", "RMSE (mm)", "R² (test)", "R² (train)"],
  [[r[0], f2(r[1]), f2(r[2], 1), f2(r[3]), f2(r[4], 3), f2(r[5], 3)] for r in mm[1:]],
  "Test-set metrics of all eight models. A = pure ML, B = physics features, C = physics-guided hybrid.", bold_rows=(3, 6, 7))
P("The physics-guided hybrids (C) give the lowest errors, MAE ≈ 1.96 mm and R² = 0.989, followed closely by Linear Regression on the "
  "physics features (B), MAE 2.07 mm, R² = 0.988. The pure ML models (A) score well on the training years (R² ≈ 0.98–0.999) but fail on "
  "the future (R² ≤ −0.06), and Random Forest is the worst (R² = −2.88) because it cannot predict beyond its training range. This is "
  "the central lesson of the model: for a quantity that keeps rising, physical knowledge matters more than algorithmic complexity.",
  align="j")
F(fo("fig03_model_comparison.png"), "Comparison of test-set errors for all models (notebook Figure 3).")
F(fo("fig04_predicted_vs_observed.png"), "Predicted vs observed sea level on the unseen test years (notebook Figure 4).")
H("7.2 Robustness", 2)
trap = D["trap"]
T(["Validation split", "MAE (mm)", "RMSE (mm)", "R²"], [[t["split"], f2(t["MAE"]), f2(t["RMSE"]), f2(t["R²"], 3)] for t in trap],
  "Leakage test for the pure Random Forest: a shuffled split hides the failure that an honest chronological split reveals.")
P("With a random split, test years sit between training years and the Random Forest appears excellent (R² = 0.996). With an honest "
  "chronological split the same model gives R² = −2.88. Rolling-origin windows (1960–79, 1980–99, 2000–24) confirm that the physics "
  "models are consistently better than the pure ML models.", align="j")
H("7.3 Explainability", 2)
T(["Quantity", "Learned by the model", "Reference / observation"], [
    ["Sensitivity a", f"{M['a']:.2f} mm/yr per °C", "3.4 mm/yr per °C (Rahmstorf 2007): same order, lower"],
    ["Equilibrium temperature T₀", f"{M['T0']:.2f} °C (vs 1951–80)", "Negative: the sea was already rising in the late 1800s"],
    ["Present-day rate of rise", f"{M['rate_now']:.2f} mm/yr", f"{M['rate_obs']:.2f} mm/yr measured by satellites"],
    ["20th-century average rate", f"{D['rate_20c']:.2f} mm/yr (observed)", "IPCC AR6: about 1.3–1.9 mm/yr"],
], "Physical constants recovered from the Linear Regression coefficients.")
H("7.4 Projections to 2100", 2)
pj = rows_csv("projections_2100.csv")
T(["Pathway", "T in 2100 (°C)", "Rise by 2050 (m)", "Rise by 2100 (m)", "90% range (m)", "IPCC AR6 median (m)", "AR6 likely range (m)"],
  [[r[0], r[1], f2(r[2]), f2(r[3]), f"{f2(r[4])}–{f2(r[5])}", r[6], r[7]] for r in pj[1:]],
  "Tideline sea-level projections relative to 1995–2014, compared with IPCC AR6.")
P("Trained only on observations, the model lands inside the IPCC AR6 likely range for the low and middle pathways. Under SSP5-8.5 it "
  "is below the IPCC range, because a law calibrated on the 20th century cannot anticipate the faster ice-sheet loss that the IPCC "
  "expects at high warming. This limitation is stated openly rather than hidden.", align="j")
F(fo("fig06_projections.png"), "Projected global mean sea level to 2100 under three pathways with 90% bootstrap bands (notebook Figure 6).")
F(fo("fig05_mumbai.png"), "Mumbai relative sea level against the global mean (notebook Figure 5).")
H("7.5 Comparison of coastal defences", 2)
dc = rows_csv("defence_comparison.csv")
rows = []
for r in dc[1:]:
    if r[0] != "SSP2-4.5":
        continue
    rows.append([r[1], r[2], r[3], f2(r[4]), f2(r[5]), f2(r[6], 1), f2(r[7], 1) if r[7] else "—", f2(r[8], 1)])
T(["Defence", "Homes flooded 2050", "Homes flooded 2100", "EAD 2100 ($M)", "Cost ($M)", "PV benefit ($M)", "BCR", "CO₂ (kt)"], rows,
  "Defence comparison under SSP2-4.5 for a 1-in-100-year storm (1,200 homes, 1 km of coast).")
rows5 = [[r[1], r[2], r[3], f2(r[7], 1) if r[7] else "—", (f"{float(r[9]) * 100:.0f}%" if r[9] else "—")] for r in dc[1:] if r[0] == "SSP5-8.5"]
T(["Defence", "Homes flooded 2050", "Homes flooded 2100", "BCR", "Mangrove health 2100"], rows5,
  "The same defences under high emissions (SSP5-8.5): the mangroves drown.")
B([
    ("Mangrove belt: ", "the best value for money (BCR ≈ 29) and it captures about 19 kt of CO₂ by 2100. It reduces but does not "
     "stop flooding (321 homes in 2100 under SSP2-4.5). Under SSP5-8.5 the sea rises faster than 7 mm/yr and the belt drowns, "
     "so its benefit disappears."),
    ("Embankment: ", "keeps every home dry in 2050, then is overtopped from about 2071 under SSP2-4.5 and the town floods all at once "
     "(467 homes in 2100). This 'cliff edge' is the main risk of grey defences."),
    ("Raised houses: ", "homes still get wet, but expected annual damage falls from $1.81M to $0.06M a year by 2100; BCR ≈ 2.0."),
    ("Hybrid (300 m mangroves + 3.8 m wall): ", "mangroves weaken the waves so a lower, cheaper wall holds a storm longer; zero homes "
     "flooded in 2050 and 341 in 2100. It is the most robust option overall."),
])
F(fo("fig07_ead_over_time.png"), "Expected annual damage over time for each defence (notebook Figure 7).")
F(fo("fig08_mangrove_health.png"), "Mangrove health under the three pathways: drowning above 7 mm/yr of rise (notebook Figure 8).")
H("7.6 Summary of outcomes", 2)
B([
    "An explainable ML model that forecasts sea level with ≈2 mm error on unseen years and recovers physically meaningful constants.",
    "Projections to 2100 consistent with IPCC AR6 for low and middle emissions.",
    "A quantitative comparison showing that green, grey and accommodation defences fail in different ways and at different times.",
    "A public interactive 3D website where the flood model runs live and gives the same numbers as the notebook.",
])

# ================================================================== 8. APPLICATIONS
H("8. Applications")
B([
    ("Coastal and urban planning: ", "municipal bodies such as the Brihanmumbai Municipal Corporation can screen protection options "
     "for a ward and see when each one would stop working."),
    ("Climate adaptation policy: ", "supports state coastal zone management plans and India's National Action Plan on Climate "
     "Change by comparing nature-based and engineered adaptation with costs and benefits."),
    ("Mangrove conservation and blue carbon: ", "quantifies both the protection and the CO₂ captured by mangroves, supporting "
     "restoration funding and carbon-credit projects."),
    ("Disaster management: ", "estimating homes and people exposed to 1-in-10 to 1-in-500-year storms helps plan evacuation, cyclone "
     "shelters and early warning."),
    ("Infrastructure design: ", "the time-to-failure analysis indicates how high embankments, ports, roads and coastal power plants "
     "need to be to stay safe over their design life."),
    ("Insurance and finance: ", "expected annual damage is the basis for flood insurance premiums and for the economic case for "
     "climate-resilient investment."),
    ("Education and public awareness: ", "the scroll-driven 3D website explains sea-level rise, uncertainty and adaptation to "
     "students and the public without needing any software."),
    ("Teaching machine learning: ", "a clear case study of why time-series data must be split chronologically and how physical "
     "knowledge improves extrapolation."),
])

# ================================================================== 9. FUTURE SCOPE
H("9. Future Scope")
B([
    ("Real town data: ", "replace the illustrative town with a high-resolution digital elevation model (for example CartoDEM or "
     "LiDAR), real building footprints and census population for a specific Mumbai ward."),
    ("Dynamic flood routing: ", "replace the static bathtub with a hydrodynamic model (for example LISFLOOD-FP) that moves water "
     "inland over time and includes drainage and rainfall, as in the July 2005 Mumbai floods."),
    ("Ice-sheet processes: ", "add process-based or emulated ice-sheet contributions so that high-emission projections capture "
     "possible Antarctic instability."),
    ("More predictors and deep learning: ", "add ocean heat content, CO₂, ENSO and land-water storage; try Gaussian Processes and "
     "LSTM networks with physics-informed loss functions, plus SHAP explanations."),
    ("Land subsidence from satellites: ", "use InSAR and GNSS measurements of vertical land motion to localise the forecast to "
     "every part of the coast."),
    ("Ecology of mangroves: ", "model sediment supply, salinity and species, and the space mangroves need to migrate inland."),
    ("More defences: ", "beach nourishment, sand dunes, living shorelines, floodgates, managed retreat and early-warning systems."),
    ("Other cities: ", "extend to Kolkata, Chennai, Visakhapatnam, Kochi and Surat using their PSMSL tide gauges."),
    ("Live dashboard: ", "connect the website to monthly updates from NOAA and NASA so the forecast refreshes automatically."),
    ("Mobile and VR: ", "a mobile app or virtual-reality walk-through of the flooded town for community engagement."),
])

H("References", 1)
for ref in [
    "Church, J. A. & White, N. J. (2011). Sea-level rise from the late 19th to the early 21st century. Surveys in Geophysics, 32, 585–602.",
    "Rahmstorf, S. (2007). A semi-empirical approach to projecting future sea-level rise. Science, 315, 368–370.",
    "Vermeer, M. & Rahmstorf, S. (2009). Global sea level linked to global temperature. PNAS, 106(51), 21527–21532.",
    "IPCC (2021). Climate Change 2021: The Physical Science Basis. Working Group I contribution to AR6, Chapter 9 and SPM.",
    "NOAA Laboratory for Satellite Altimetry. Global mean sea level time series. star.nesdis.noaa.gov/socd/lsa/SeaLevelRise",
    "Permanent Service for Mean Sea Level (PSMSL). Station 43, Mumbai/Bombay (Apollo Bandar). psmsl.org",
    "GISTEMP Team (2025). GISS Surface Temperature Analysis (GISTEMP v4). NASA Goddard Institute for Space Studies.",
    "Huizinga, J., de Moel, H. & Szewczyk, W. (2017). Global flood depth-damage functions. JRC Technical Report, European Commission.",
    "McIvor, A., Möller, I., Spencer, T. & Spalding, M. (2012). Reduction of wind and swell waves by mangroves. The Nature Conservancy.",
    "Saintilan, N. et al. (2020). Thresholds of mangrove survival under rapid sea level rise. Science, 368, 1118–1121.",
    "Narayan, S. et al. (2016). The effectiveness, costs and coastal protection benefits of natural and nature-based defences. PLoS ONE.",
    "Alongi, D. M. (2014). Carbon cycling and storage in mangrove forests. Annual Review of Marine Science, 6, 195–219.",
    "Pedregosa, F. et al. (2011). Scikit-learn: Machine Learning in Python. JMLR, 12, 2825–2830.",
]:
    doc.add_paragraph(ref, style="List Number")


# ------------------------------------------------------------------ table of contents (page numbers are PAGEREF fields)
def pageref(cell, name):
    cell.text = ""; p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)

    def run(child):
        r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr"); f = OxmlElement("w:rFonts")
        f.set(qn("w:ascii"), TNR); f.set(qn("w:hAnsi"), TNR); rpr.append(f)
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "20"); rpr.append(sz); r.append(rpr); r.append(child); p._p.append(r)

    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin"); run(b)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" PAGEREF {name} \\h "; run(it)
    sp = OxmlElement("w:fldChar"); sp.set(qn("w:fldCharType"), "separate"); run(sp)
    t = OxmlElement("w:t"); t.text = "#"; run(t)
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); run(e)


rows = toc.rows
proto_main, proto_sub = copy.deepcopy(rows[1]._tr), copy.deepcopy(rows[2]._tr)
for r in list(rows)[1:]:
    toc._tbl.remove(r._tr)
for num, title, lvl, name in TOCROWS:
    tr = copy.deepcopy(proto_main if lvl == 1 else proto_sub)
    toc._tbl.append(tr)
    row = toc.rows[-1]
    for c, txt in ((row.cells[0], num), (row.cells[1], title)):
        prs = c.paragraphs
        runs = prs[0].runs
        if runs:
            runs[0].text = txt
            for x in runs[1:]:
                x.text = ""
        else:
            prs[0].add_run(txt)
    pageref(row.cells[2], name)

# ask Word to refresh the contents page numbers when the file is opened
settings = doc.settings.element
uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true"); settings.append(uf)

path = os.path.join(HERE, "Tideline_Project_Report.docx")
doc.save(path)


def strip_embedded_fonts(p):
    """The SRM template embeds ~6 MB of fonts; Times New Roman is on every PC, so drop them."""
    import re, shutil, zipfile
    tmp = p + ".tmp"
    with zipfile.ZipFile(p) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            if it.filename.startswith("word/fonts/"):
                continue
            data = zi.read(it.filename)
            if it.filename == "word/fontTable.xml":
                data = re.sub(rb"<w:embed(Regular|Bold|Italic|BoldItalic)\b[^>]*/>", b"", data)
            elif it.filename == "word/_rels/fontTable.xml.rels":
                data = re.sub(rb"<Relationship [^>]*Target=\"fonts/[^>]*/>", b"", data)
            elif it.filename == "[Content_Types].xml":
                data = re.sub(rb"<Override [^>]*PartName=\"/word/fonts/[^>]*/>", b"", data)
            elif it.filename == "word/settings.xml":
                data = re.sub(rb"<w:(embedTrueTypeFonts|saveSubsetFonts|embedSystemFonts)\b[^>]*/>", b"", data)
            zo.writestr(it, data)
    shutil.move(tmp, p)


strip_embedded_fonts(path)
print("wrote", path, f"({FIG[0]} figures, {TAB[0]} tables, {len(TOCROWS)} contents rows)")
