"""Generate CISC7204 Preliminary Project Proposal Word document."""

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import os

OUTPUT = os.path.join(
    os.path.dirname(__file__),
    "CISC7204-Assgn01-MC649355-PreliminaryProjectProposal.docx",
)

doc = Document()

# --- Styles ---
style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(12)

for level in range(1, 4):
    h = doc.styles[f"Heading {level}"]
    h.font.name = "Times New Roman"
    h.font.bold = True

# --- Front Page ---
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("CISC7204\nData Science and Data Visualization\n\n")
run.bold = True
run.font.size = Pt(14)

run2 = title.add_run("Preliminary Project Proposal\n(Assignment 01 – Ideation)\n\n")
run2.bold = True
run2.font.size = Pt(13)

doc.add_paragraph()
info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
info.add_run("Student Name: CHAN CHENG I\n").bold = True
info.add_run("Student ID: MC649355\n\n")
info.add_run("Course Code: CISC7204\n")
info.add_run("Course Title: Data Science and Data Visualization\n")
info.add_run("Coursework Item: Individual Project Proposal – Ideation\n")
info.add_run("Submission Date: 20 September 2026\n")

doc.add_page_break()

# --- Project Title ---
doc.add_heading("1. Project Title and Research Question", level=1)

doc.add_paragraph(
    "Project Title: Trends and Disparities in Hong Kong's Private Residential "
    "Property Market: A Data-Driven Analysis of Price Evolution by Housing Class "
    "and Region (1999–2025)"
)

doc.add_paragraph(
    "Research Question: How have average private residential property prices in "
    "Hong Kong evolved across different housing classes (A–E) and geographic "
    "regions (Hong Kong Island, Kowloon, and the New Territories) since 1999, "
    "and to what extent do these trends reveal widening or narrowing affordability "
    "gaps among housing segments?"
)

# --- Region and Domain (Ideation) ---
doc.add_heading("2. Region, Domain, and Ideation", level=1)

doc.add_paragraph(
    "Region: Hong Kong Special Administrative Region (HKSAR), China"
)
doc.add_paragraph(
    "Domain Category: Housing and Real Estate (Private Residential Property Market)"
)

doc.add_heading("Research Question (Ideation)", level=2)
doc.add_paragraph(
    "Hong Kong was the least affordable housing market in the Demographia International "
    "Housing Affordability 2025 Edition, with a median house-price-to-household-income "
    "ratio (median multiple) of 14.4 based on third-quarter 2024 data (Cox, 2025). "
    "While headlines often report aggregate price movements, less attention is paid to "
    "how different housing size classes and districts diverge over time. Following the "
    "Rating and Valuation Department (RVD) classification, Class A–E refer to saleable "
    "area bands, not social class: Class A (less than 40 m²), Class B (40–69.9 m²), "
    "Class C (70–99.9 m²), Class D (100–159.9 m²), and Class E (160 m² or above) "
    "(Rating and Valuation Department, 2025). This project investigates whether smaller "
    "units (Class A and B) have experienced price growth at the same rate as larger "
    "units (Class D and E), and whether geographic disparities between Hong Kong Island, "
    "Kowloon, and the New Territories have widened or narrowed since 1999."
)

doc.add_heading("Data Sources (at least two links)", level=2)

p = doc.add_paragraph()
p.add_run("Dataset 1 – ").bold = True
p.add_run(
    "Rating and Valuation Department (RVD), Private Domestic – Average Prices "
    "by Class, Monthly (from 1999)\n"
)
p.add_run("Direct CSV: ").italic = True
p.add_run("http://www.rvd.gov.hk/datagovhk/1.2M.csv\n")
p.add_run("Portal page: ").italic = True
p.add_run(
    "https://data.gov.hk/en-data/dataset/hk-rvd-tsinfo_rvd-property-market-statistics/"
    "resource/79db448f-46a0-4bea-8163-e9bff613e5cb\n"
)
p.add_run(
    "Note: Average prices are published in Hong Kong dollars per square metre of "
    "saleable area (HK$/m²), not as total transaction prices.\n"
)

p2 = doc.add_paragraph()
p2.add_run("Dataset 2 – ").bold = True
p2.add_run(
    "Rating and Valuation Department (RVD), Private Domestic – Price Indices "
    "by Class (Territory-wide), Monthly (from 1993)\n"
)
p2.add_run("Direct CSV: ").italic = True
p2.add_run("http://www.rvd.gov.hk/datagovhk/1.4M.csv\n")
p2.add_run("Portal page: ").italic = True
p2.add_run(
    "https://data.gov.hk/en-data/dataset/hk-rvd-tsinfo_rvd-property-market-statistics/"
    "resource/6a54eb2b-cf5a-482d-b82c-ee57b1814aae\n"
)
p2.add_run(
    "Note: The official price index series is expressed with 1999 = 100 "
    "(Rating and Valuation Department, n.d.).\n"
)

doc.add_heading("Visualization Justification", level=2)
doc.add_paragraph(
    "A preliminary line chart (see accompanying Jupyter Notebook and figure) plots "
    "territory-wide price indices for Classes A through E from 1999 to the present. "
    "This visualization directly addresses the research question by making long-term "
    "trends and cross-class disparities immediately visible: if affordability gaps "
    "are widening, lower-class indices will diverge upward more steeply from "
    "higher-class indices (or vice versa if the luxury segment outpaces the mass "
    "market). Comparing index trajectories before and after major economic events "
    "(e.g., the 2008 global financial crisis, 2019 social unrest, and the COVID-19 "
    "pandemic) allows an initial assessment of which housing segments are most "
    "sensitive to macroeconomic shocks. The chart serves as a foundation for deeper "
    "statistical analysis in the full project."
)

doc.add_paragraph(
    "A second exploratory chart compares average prices across Hong Kong Island, "
    "Kowloon, and the New Territories for a selected housing class, revealing "
    "regional price gaps that complement the territory-wide index analysis."
)

# --- Purpose ---
doc.add_page_break()
doc.add_heading("3. Project Purpose", level=1)

doc.add_heading("Outcomes", level=2)
doc.add_paragraph(
    "This project connects directly to CISC7204 learning outcomes in data wrangling, "
    "exploratory data analysis, statistical modelling, and data-driven storytelling. "
    "By analysing official Hong Kong government housing data, the project bridges "
    "classroom techniques (pandas, matplotlib, regression) with a real-world policy "
    "issue that affects residents daily. The findings will deepen understanding of "
    "how data science can inform public discourse on housing affordability—a topic "
    "relevant to urban planning, finance, and social policy. For future career "
    "planning, the project builds a portfolio piece demonstrating end-to-end "
    "analytical capability applicable to roles in consulting, financial services, "
    "government analytics, or real estate research."
)

doc.add_heading("Skills", level=2)
skills = [
    "Data acquisition and cleaning of government CSV datasets with irregular headers and missing values",
    "Exploratory data analysis (EDA) using pandas, matplotlib, and seaborn",
    "Time-series visualisation and interpretation of long-term trends",
    "Statistical inference (correlation, regression, hypothesis testing) to compare housing classes and regions",
    "Feature engineering for temporal and categorical variables (housing class, region)",
    "Model development and evaluation (linear/polynomial regression, regularisation) following Module 06 practices",
    "Communicating analytical findings through charts, summary tables, and written interpretation",
    "Critical evaluation of data limitations (e.g., aggregated averages masking within-district variation)",
]
for s in skills:
    doc.add_paragraph(s, style="List Bullet")

doc.add_heading("Knowledge", level=2)
knowledge = [
    "Hong Kong private residential property market structure and RVD housing classification (Classes A–E)",
    "Time-series analysis concepts: trends, seasonality, and structural breaks",
    "Housing affordability metrics and their policy implications in high-density urban economies",
    "Regression modelling for continuous outcome variables (property prices and price indices)",
    "Data governance and open-data practices of Hong Kong government agencies (RVD, C&SD)",
    "Ethical considerations when presenting housing data that affects public policy perceptions",
]
for k in knowledge:
    doc.add_paragraph(k, style="List Bullet")

# --- Tasks ---
doc.add_heading("4. Project Tasks", level=1)

doc.add_paragraph(
    "The project will follow the data science lifecycle taught in CISC7204 Modules 01–06. "
    "Specific steps and deliverables are outlined below."
)

tasks = [
    ("Phase 1 – Data Collection and Preparation",
     [
         "Download and inspect RVD CSV files (average prices and price indices)",
         "Clean headers, parse date columns (MM-YYYY format), handle missing entries and remark flags",
         "Merge datasets on time period and housing class where appropriate",
         "Deliverable: Cleaned datasets saved as CSV; data dictionary documenting variables",
     ]),
    ("Phase 2 – Exploratory Data Analysis",
     [
         "Produce line charts of price indices and average prices by class and region over time",
         "Calculate summary statistics (mean, median, growth rates) for each housing class",
         "Identify structural breaks and compare pre/post major economic events",
         "Deliverable: EDA notebook with annotated visualisations and initial observations",
     ]),
    ("Phase 3 – Statistical Analysis and Modelling",
     [
         "Conduct correlation analysis between housing classes and regions",
         "Build regression models to predict price indices from time, class, and region variables",
         "Evaluate models using R², RMSE, and cross-validation (following Module 05–06 methods)",
         "Test hypotheses (e.g., whether Class A price growth significantly differs from Class E)",
         "Deliverable: Model comparison table and diagnostic plots",
     ]),
    ("Phase 4 – Interpretation and Reporting",
     [
         "Synthesise findings into answers to the research question",
         "Discuss limitations (aggregated data, no individual transaction records)",
         "Prepare final report with executive summary, methodology, results, and conclusions",
         "Deliverable: Final Jupyter Notebook and written project report (PDF)",
     ]),
]

for phase_title, items in tasks:
    doc.add_heading(phase_title, level=2)
    for item in items:
        doc.add_paragraph(item, style="List Bullet")

doc.add_heading("Guides, Worksheets, and Requirements", level=2)
doc.add_paragraph(
    "This project will follow the analytical workflow established in CISC7204 course "
    "modules, particularly Module 06 (Final Project: House Sales in King County), which "
    "provides a template for data import, wrangling, EDA, model development, and "
    "evaluation. Course lab worksheets on pandas data frames, matplotlib visualisation, "
    "and scikit-learn regression (Modules 02–05) will serve as reference guides. "
    "The RVD data dictionary and DATA.GOV.HK metadata will be consulted for variable "
    "definitions. All Python code will be written in Python 3 using Jupyter Notebook, "
    "consistent with course requirements."
)

# --- Criteria for Success ---
doc.add_heading("5. Criteria for Success", level=1)

doc.add_paragraph(
    "A successful finished project will exhibit the following characteristics:"
)

criteria = [
    "Research question clearly answered with evidence from data analysis, not speculation",
    "At least two government datasets integrated and cleaned reproducibly in Python",
    "Minimum of four well-labelled visualisations (time-series trends, regional comparison, model diagnostics, summary chart)",
    "At least one regression model with reported R² and RMSE, compared against a baseline",
    "Statistical findings interpreted in the context of Hong Kong housing policy and affordability",
    "Limitations of the data and analysis explicitly acknowledged",
    "Reproducible Jupyter Notebook with comments explaining each analytical step",
    "Final written report (1,500–2,500 words) with introduction, methodology, results, discussion, and conclusion",
]
for c in criteria:
    doc.add_paragraph(c, style="List Bullet")

doc.add_heading("Assessment Support Materials", level=2)
doc.add_paragraph(
    "To evaluate success, I will prepare the following supporting materials:"
)
support = [
    "A self-assessment checklist mapping each deliverable to the criteria above",
    "Annotated chart examples showing good vs. poor visualisation practices (axis labels, legends, titles)",
    "A model evaluation rubric with thresholds (e.g., R² > 0.80 for time-series regression as a target)",
    "A final summary dashboard (multi-panel figure) telling the data story from 1999 to present",
]
for s in support:
    doc.add_paragraph(s, style="List Bullet")

# --- Collaboration ---
doc.add_heading("6. Collaboration Issues", level=1)

doc.add_paragraph(
    "This is an individual project. No team-based division of work applies. "
    "However, the following personal work practices will be adopted to ensure "
    "consistent progress:"
)

practices = [
    "Weekly self-review against the phase deliverables listed in Section 4",
    "Maintain a project log (date, tasks completed, issues encountered) in the Jupyter Notebook",
    "Seek instructor or peer feedback during lab sessions on EDA and model design",
    "Use version control (Git) to track notebook changes and enable rollback if needed",
    "Cite all data sources (RVD, DATA.GOV.HK) and any referenced publications in the final report",
]
for pr in practices:
    doc.add_paragraph(pr, style="List Bullet")

doc.add_paragraph(
    "If future coursework involves group work, the experience gained from managing "
    "this individual end-to-end project will inform my approach to role allocation "
    "(coordinator, scribe, progress tracker, liaison) in collaborative settings."
)

# --- References ---
doc.add_heading("References", level=1)
refs = [
    "Cox, W. (2025). Demographia International Housing Affordability, 2025 Edition. "
    "Center for Demographics and Policy, Chapman University. "
    "https://www.chapman.edu/communication/_files/Demographia-International-Housing-Affordability-2025-Edition.pdf",
    "Rating and Valuation Department. (2026). Property Market Statistics [Dataset]. "
    "DATA.GOV.HK. "
    "https://data.gov.hk/en-data/dataset/hk-rvd-tsinfo_rvd-property-market-statistics",
    "Rating and Valuation Department. (2025). Technical Notes (Hong Kong Property Review 2025). "
    "https://www.rvd.gov.hk/doc/en/hkpr25/07_Technical_Notes_TC.pdf",
    "Rating and Valuation Department. (n.d.). Property Market Statistics. "
    "https://www.rvd.gov.hk/en/publications/property_market_statistics.html",
    "Rating and Valuation Department. (n.d.). Private Domestic – Average Prices by Class – "
    "Monthly (from 1999) [CSV]. http://www.rvd.gov.hk/datagovhk/1.2M.csv",
    "Rating and Valuation Department. (n.d.). Private Domestic – Price Indices by Class "
    "(Territory-wide) – Monthly (from 1993) [CSV]. http://www.rvd.gov.hk/datagovhk/1.4M.csv",
]
for i, ref in enumerate(refs, 1):
    doc.add_paragraph(f"[{i}] {ref}")

doc.save(OUTPUT)
print(f"Saved: {OUTPUT}")
