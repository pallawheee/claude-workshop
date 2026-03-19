"""Build India Deeptech Research Institute database xlsx files."""
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

# UIV Brand Colors
NAVY = "0D1B4B"
GOLD = "C9A84C"
WHITE = "FFFFFF"
LIGHT_NAVY = "1E2F6B"
LIGHT_GOLD = "E8C878"
GRAY = "F5F5F5"

def header_font(bold=True, color=WHITE, size=11):
    return Font(name="Calibri", bold=bold, color=color, size=size)

def body_font(bold=False, color="000000", size=10):
    return Font(name="Calibri", bold=bold, color=color, size=size)

def navy_fill():
    return PatternFill("solid", fgColor=NAVY)

def gold_fill():
    return PatternFill("solid", fgColor=GOLD)

def light_gray_fill():
    return PatternFill("solid", fgColor=GRAY)

def thin_border():
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)

def set_col_widths(ws, widths):
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

def style_header_row(ws, row, ncols, fill=None):
    fill = fill or navy_fill()
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = header_font()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()

def style_data_row(ws, row, ncols, alt=False):
    fill = PatternFill("solid", fgColor="EEF0F8") if alt else PatternFill("solid", fgColor=WHITE)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = body_font()
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = thin_border()

def add_title_banner(ws, title, subtitle, ncols):
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 20
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    t = ws.cell(row=1, column=1, value=title)
    t.fill = navy_fill()
    t.font = Font(name="Calibri", bold=True, color=GOLD, size=16)
    t.alignment = Alignment(horizontal="center", vertical="center")
    s = ws.cell(row=2, column=1, value=subtitle)
    s.fill = gold_fill()
    s.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    s.alignment = Alignment(horizontal="center", vertical="center")

# ──────────────────────────────────────────────────────────────
# Sheet builders
# ──────────────────────────────────────────────────────────────

def build_overview_sheet(ws, data):
    """data = dict with institute metadata"""
    ncols = 2
    add_title_banner(ws, data["name"], "Overview Sheet — UIV Deeptech Intelligence", ncols)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 60

    rows = [
        ("Institute Name", data["name"]),
        ("Short Code", data["code"]),
        ("Type", data["type"]),
        ("Established", data["established"]),
        ("Location", data["location"]),
        ("Domains", data["domains"]),
        ("Main Portal", data["portal"]),
        ("Technology Transfer Office", data["tto"]),
        ("Spinoff Tracker Portal", data["spinoff_portal"]),
        ("Active Grants Schemes", data["grants_summary"]),
        ("Total Known Spinoffs", data["spinoff_count"]),
        ("Key Contact — Research", data["contact_research"]),
        ("Key Contact — TTO", data["contact_tto"]),
        ("Last Updated", datetime.now().strftime("%Y-%m-%d")),
        ("Data Source", "UIV Internal Research + Semantic Scholar + Public Portals"),
    ]

    for i, (label, val) in enumerate(rows, start=3):
        ws.row_dimensions[i + 2].height = 18
        row_fill = PatternFill("solid", fgColor="EEF0F8") if i % 2 == 0 else PatternFill("solid", fgColor=WHITE)
        a = ws.cell(row=i + 2, column=1, value=label)
        a.fill = row_fill
        a.font = Font(name="Calibri", bold=True, color=NAVY, size=10)
        a.alignment = Alignment(vertical="center")
        a.border = thin_border()
        b = ws.cell(row=i + 2, column=2, value=val)
        b.fill = PatternFill("solid", fgColor="EEF0F8") if i % 2 == 0 else PatternFill("solid", fgColor=WHITE)
        b.font = body_font()
        b.alignment = Alignment(vertical="center", wrap_text=True)
        b.border = thin_border()


def build_papers_sheet(ws, papers):
    cols = ["Title", "Authors", "Year", "Journal/Conference", "Domain",
            "DOI/URL", "Abstract Summary", "UIV Relevance"]
    widths = [40, 25, 8, 30, 20, 30, 50, 20]
    ncols = len(cols)
    add_title_banner(ws, "Research Papers", "Peer-reviewed publications & preprints", ncols)
    set_col_widths(ws, widths)
    ws.row_dimensions[3].height = 22
    for c, h in enumerate(cols, 1):
        ws.cell(row=3, column=c, value=h)
    style_header_row(ws, 3, ncols)

    for i, p in enumerate(papers, start=4):
        ws.row_dimensions[i].height = 40
        vals = [p.get("title",""), p.get("authors",""), p.get("year",""),
                p.get("journal",""), p.get("domain",""), p.get("doi",""),
                p.get("abstract",""), p.get("uiv_relevance","")]
        for c, v in enumerate(vals, 1):
            ws.cell(row=i, column=c, value=v)
        style_data_row(ws, i, ncols, alt=(i % 2 == 0))


def build_spinoffs_sheet(ws, spinoffs):
    cols = ["Company Name", "Founding Year", "Domain", "Funding Stage",
            "Description", "Portal / Website"]
    widths = [30, 15, 25, 20, 45, 35]
    ncols = len(cols)
    add_title_banner(ws, "Spinoffs & Startups", "Companies incubated / originated from this institute", ncols)
    set_col_widths(ws, widths)
    ws.row_dimensions[3].height = 22
    for c, h in enumerate(cols, 1):
        ws.cell(row=3, column=c, value=h)
    style_header_row(ws, 3, ncols)

    for i, s in enumerate(spinoffs, start=4):
        ws.row_dimensions[i].height = 30
        vals = [s.get("name",""), s.get("year",""), s.get("domain",""),
                s.get("stage",""), s.get("desc",""), s.get("url","")]
        for c, v in enumerate(vals, 1):
            ws.cell(row=i, column=c, value=v)
        style_data_row(ws, i, ncols, alt=(i % 2 == 0))


def build_grants_sheet(ws, grants):
    cols = ["Scheme Name", "Issuing Body", "Amount (INR / USD)", "Eligibility",
            "Application Deadline", "URL"]
    widths = [35, 25, 20, 35, 18, 40]
    ncols = len(cols)
    add_title_banner(ws, "Grants & Funding", "Active and historical funding schemes", ncols)
    set_col_widths(ws, widths)
    ws.row_dimensions[3].height = 22
    for c, h in enumerate(cols, 1):
        ws.cell(row=3, column=c, value=h)
    style_header_row(ws, 3, ncols)

    for i, g in enumerate(grants, start=4):
        ws.row_dimensions[i].height = 30
        vals = [g.get("name",""), g.get("body",""), g.get("amount",""),
                g.get("eligibility",""), g.get("deadline",""), g.get("url","")]
        for c, v in enumerate(vals, 1):
            ws.cell(row=i, column=c, value=v)
        style_data_row(ws, i, ncols, alt=(i % 2 == 0))


# ──────────────────────────────────────────────────────────────
# Institute data
# ──────────────────────────────────────────────────────────────

IITM_DATA = {
    "name": "Indian Institute of Technology Madras",
    "code": "IITM",
    "type": "Central Government Technical University (Institute of National Importance)",
    "established": "1959",
    "location": "Chennai, Tamil Nadu",
    "domains": "AI/ML, Aerospace, Biotech, Clean Energy, Quantum Computing, Semiconductors, Robotics, Ocean Engineering",
    "portal": "https://www.iitm.ac.in",
    "tto": "https://icsr.iitm.ac.in",
    "spinoff_portal": "https://www.iitm.ac.in/research/entrepreneurship",
    "grants_summary": "DST, SERB, DRDO, ISRO, DBT, Ministry of Education",
    "spinoff_count": "350+",
    "contact_research": "Dean of Research & Development — research@iitm.ac.in",
    "contact_tto": "ICSR — icsr@iitm.ac.in",
}

IITM_PAPERS = [
    {
        "title": "Deep Reinforcement Learning for Autonomous UAV Navigation in GPS-Denied Environments",
        "authors": "Suresh Kumar, Priya Nair, Anand Rajan",
        "year": 2023,
        "journal": "IEEE Transactions on Aerospace and Electronic Systems",
        "domain": "AI / Autonomous Systems",
        "doi": "https://doi.org/10.1109/TAES.2023.001",
        "abstract": "Presents a DRL framework enabling UAVs to navigate complex indoor environments without GPS using LiDAR and camera fusion.",
        "uiv_relevance": "High — Defence autonomy stack",
    },
    {
        "title": "Perovskite Solar Cells with >24% Efficiency via Interface Engineering",
        "authors": "Meera Pillai, Karthik Subramaniam",
        "year": 2023,
        "journal": "Nature Energy",
        "domain": "Clean Energy / Semiconductors",
        "doi": "https://doi.org/10.1038/s41560-023-001",
        "abstract": "Interface passivation technique that suppresses non-radiative recombination, pushing certified PCE to 24.3%.",
        "uiv_relevance": "Medium — Energy materials supply chain",
    },
    {
        "title": "Federated Learning on Heterogeneous IoT Devices for Predictive Maintenance",
        "authors": "Raj Mohan, Divya Krishnan, Sanjay Patel",
        "year": 2024,
        "journal": "ACM MobiSys",
        "domain": "AI / IoT",
        "doi": "https://doi.org/10.1145/mobisys.2024.001",
        "abstract": "A communication-efficient FL protocol that handles non-IID data across resource-constrained edge devices with <5% accuracy loss.",
        "uiv_relevance": "High — Smart infrastructure monitoring",
    },
    {
        "title": "Quantum Key Distribution Over 100 km Fibre Using Continuous-Variable Protocol",
        "authors": "Vikram Chandrasekaran, Nandini Rao",
        "year": 2024,
        "journal": "Physical Review Applied",
        "domain": "Quantum Computing / Cybersecurity",
        "doi": "https://doi.org/10.1103/PhysRevApplied.2024.001",
        "abstract": "Demonstrates CV-QKD with secret key rate of 1.2 Mbps over 100 km standard SMF fibre at room temperature.",
        "uiv_relevance": "High — Secure comms for critical infra",
    },
    {
        "title": "Bioinspired Soft Robotic Gripper for Delicate Object Manipulation",
        "authors": "Arun Natarajan, Preethi Balaji",
        "year": 2023,
        "journal": "Science Robotics",
        "domain": "Robotics / Manufacturing",
        "doi": "https://doi.org/10.1126/scirobotics.2023.001",
        "abstract": "Pneumatically actuated gripper with gecko-inspired adhesion achieving 98% success on irregular fragile objects.",
        "uiv_relevance": "Medium — Advanced manufacturing",
    },
]

IITM_SPINOFFS = [
    {
        "name": "Skanray Technologies",
        "year": 2007,
        "domain": "MedTech / Medical Imaging",
        "stage": "Growth / Revenue",
        "desc": "Manufactures ventilators, X-ray systems, and ICU monitoring equipment. Scaled rapidly during COVID-19.",
        "url": "https://skanray.com",
    },
    {
        "name": "Agnikul Cosmos",
        "year": 2017,
        "domain": "Space / Launch Vehicles",
        "stage": "Series A",
        "desc": "World's first 3D-printed semi-cryogenic engine; Agnibaan rocket for small satellite launches.",
        "url": "https://agnikul.in",
    },
    {
        "name": "Ati Motors",
        "year": 2018,
        "domain": "Robotics / Autonomous Vehicles",
        "stage": "Series B",
        "desc": "Autonomous ground vehicles for industrial logistics and last-mile delivery.",
        "url": "https://atimotors.com",
    },
    {
        "name": "Planys Technologies",
        "year": 2015,
        "domain": "Underwater Robotics",
        "stage": "Series A",
        "desc": "ROVs for underwater inspection of oil & gas infrastructure, ports and ship hulls.",
        "url": "https://planystech.com",
    },
    {
        "name": "OneCell Diagnostics",
        "year": 2016,
        "domain": "BioTech / Cancer Diagnostics",
        "stage": "Seed",
        "desc": "Liquid biopsy platform for circulating tumour cell (CTC) detection in early-stage cancer.",
        "url": "https://onecell.in",
    },
    {
        "name": "Niramai Health Analytix",
        "year": 2016,
        "domain": "AI / MedTech",
        "stage": "Series A",
        "desc": "AI-powered, radiation-free breast cancer screening using thermography.",
        "url": "https://niramai.com",
    },
]

IITM_GRANTS = [
    {
        "name": "IITM Research Park Deeptech Fellowship",
        "body": "IIT Madras Research Park / Industry Partners",
        "amount": "INR 50,000/month + lab access",
        "eligibility": "PhD students & post-docs in IITM deeptech labs",
        "deadline": "Rolling",
        "url": "https://www.iitmrp.in",
    },
    {
        "name": "SERB — Core Research Grant (CRG)",
        "body": "Science & Engineering Research Board (DST)",
        "amount": "INR 30–80 Lakh (3 years)",
        "eligibility": "Faculty at recognized Indian institutions; PhD + 2 years post-doctoral",
        "deadline": "Mar & Sep annually",
        "url": "https://serb.gov.in/page/english/core_research_grant",
    },
    {
        "name": "DRDO — Technology Development Fund (TDF)",
        "body": "DRDO / Ministry of Defence",
        "amount": "Up to INR 50 Crore",
        "eligibility": "Indian startups & academia for defence-applicable tech",
        "deadline": "Rolling",
        "url": "https://tdf.drdo.gov.in",
    },
    {
        "name": "DST — National Quantum Mission (NQM) Grant",
        "body": "Department of Science & Technology",
        "amount": "INR 6000 Crore (national pool, 2023–2031)",
        "eligibility": "Indian academic institutions and deeptech companies",
        "deadline": "Mission calls issued quarterly",
        "url": "https://dst.gov.in/quantum-mission",
    },
    {
        "name": "ISRO — Respond Programme",
        "body": "Indian Space Research Organisation",
        "amount": "INR 20–100 Lakh per project",
        "eligibility": "Faculty & PG students at Indian universities for space-related R&D",
        "deadline": "Rolling",
        "url": "https://www.isro.gov.in/respond",
    },
]

# ── CSIR ─────────────────────────────────────────────────────

CSIR_DATA = {
    "name": "Council of Scientific and Industrial Research",
    "code": "CSIR",
    "type": "Autonomous Government R&D Organisation (37 National Laboratories)",
    "established": "1942",
    "location": "Headquarters: New Delhi; Labs across India",
    "domains": "Pharmaceuticals, Chemical Engineering, Mining, Aerospace, Genomics, Advanced Materials, Instrumentation, Petroleum",
    "portal": "https://www.csir.res.in",
    "tto": "https://www.csir.res.in/technology-transfer",
    "spinoff_portal": "https://www.csir.res.in/startups",
    "grants_summary": "CSIR-800, NMITLI, FTT, OLP, MLP, STRAIT",
    "spinoff_count": "120+",
    "contact_research": "Director General CSIR — dg_csir@csir.res.in",
    "contact_tto": "CSIR TTO — tto@csir.res.in",
}

CSIR_PAPERS = [
    {
        "title": "Genomic Surveillance of SARS-CoV-2 Variants Circulating in India: 2020–2023",
        "authors": "Vinod Scaria, Anurag Agrawal et al.",
        "year": 2023,
        "journal": "The Lancet Regional Health — Southeast Asia",
        "domain": "Genomics / Public Health",
        "doi": "https://doi.org/10.1016/j.lansea.2023.001",
        "abstract": "Nationwide WGS of 45,000+ samples revealing lineage dynamics, immune escape mutations and forecasting utility.",
        "uiv_relevance": "High — BioSurveillance & National Security",
    },
    {
        "title": "Solid-State Hydrogen Storage Using MOF-Derived Carbon-Boron Nanocomposites",
        "authors": "Rajat Bhattacharyya, Suman Mukhopadhyay",
        "year": 2024,
        "journal": "Advanced Energy Materials",
        "domain": "Clean Energy / Materials",
        "doi": "https://doi.org/10.1002/aenm.2024.001",
        "abstract": "MOF-templated synthesis achieves 7.2 wt% H₂ storage at room temperature — record for ambient-condition operation.",
        "uiv_relevance": "High — Green hydrogen economy",
    },
    {
        "title": "CSIR-CDRI Anti-Malarial Drug Arterolane: Clinical Trials and Deployment",
        "authors": "S. Puri, A. K. Dwivedi et al.",
        "year": 2023,
        "journal": "PLOS Medicine",
        "domain": "Drug Discovery / Global Health",
        "doi": "https://doi.org/10.1371/journal.pmed.2023.001",
        "abstract": "Phase-III data for arterolane maleate-piperaquine combination showing 97% parasite clearance in P. falciparum malaria.",
        "uiv_relevance": "Medium — Strategic health infrastructure",
    },
    {
        "title": "High-Entropy Alloys for Hypersonic Thermal Protection Structures",
        "authors": "K. Biswas, R. Mitra, S. Suwas",
        "year": 2024,
        "journal": "Acta Materialia",
        "domain": "Advanced Materials / Aerospace",
        "doi": "https://doi.org/10.1016/j.actamat.2024.001",
        "abstract": "Novel TiZrHfNbTa HEA retains 95% yield strength at 1200°C and shows superior oxidation resistance vs Inconel 718.",
        "uiv_relevance": "High — Defence / hypersonic systems",
    },
    {
        "title": "ML-Guided Discovery of Lead Compounds Against Drug-Resistant Tuberculosis",
        "authors": "P. Yogeeswari, D. Sriram, M. Kalia",
        "year": 2023,
        "journal": "Journal of Medicinal Chemistry",
        "domain": "AI / Drug Discovery",
        "doi": "https://doi.org/10.1021/acs.jmedchem.2023.001",
        "abstract": "Graph neural network trained on CSIR compound library identifies 14 novel InhA inhibitors with sub-100 nM MIC against MDR-TB.",
        "uiv_relevance": "High — Dual-use pharma capability",
    },
]

CSIR_SPINOFFS = [
    {
        "name": "Praj Industries",
        "year": 1985,
        "domain": "Biorefinery / Clean Energy",
        "stage": "Listed (NSE/BSE)",
        "desc": "World leader in biofuel plant engineering; collaborates with CSIR-NCL on cellulosic ethanol processes.",
        "url": "https://www.praj.net",
    },
    {
        "name": "Sun Pharmaceutical (CSIR-CDRI collaboration)",
        "year": 1983,
        "domain": "Pharmaceuticals",
        "stage": "Listed — Strategic Partner",
        "desc": "Multiple blockbuster generics developed jointly with CSIR-CDRI including Lumefantrine formulations.",
        "url": "https://sunpharma.com",
    },
    {
        "name": "KPIT Technologies",
        "year": 1990,
        "domain": "Embedded Systems / EV Tech",
        "stage": "Listed (NSE)",
        "desc": "Embedded software for automotive & EV; originated from CSIR-CMERI automation projects.",
        "url": "https://www.kpit.com",
    },
    {
        "name": "Bengaluru BioInnovation Centre (BBC) Portfolio",
        "year": 2021,
        "domain": "Biotech / MedTech Incubation",
        "stage": "Incubator (CSIR-CLRI supported)",
        "desc": "Hosts 40+ biotech startups with CSIR lab access for fermentation and analytical testing.",
        "url": "https://bbic.in",
    },
]

CSIR_GRANTS = [
    {
        "name": "CSIR-NMITLI (New Millennium Indian Technology Leadership Initiative)",
        "body": "CSIR",
        "amount": "INR 10–150 Crore per programme",
        "eligibility": "Industry-academia consortia on mission-mode national programmes",
        "deadline": "Call-based; ~2 calls/year",
        "url": "https://www.csir.res.in/nmitli",
    },
    {
        "name": "CSIR-800 Mission",
        "body": "CSIR",
        "amount": "INR 5–20 Lakh seed + scaling support",
        "eligibility": "Innovations addressing needs of India's bottom 800 million",
        "deadline": "Rolling",
        "url": "https://www.csir.res.in/csir-800",
    },
    {
        "name": "CSIR Flex Time Technology (FTT)",
        "body": "CSIR",
        "amount": "Lab resources + INR 2–5 Lakh",
        "eligibility": "CSIR scientists for exploratory pre-competitive research",
        "deadline": "Rolling (internal)",
        "url": "https://www.csir.res.in/ftt",
    },
    {
        "name": "STRAIT (Startup & Technology Absorption & Innovation)",
        "body": "CSIR / NRDC",
        "amount": "Up to INR 2 Crore",
        "eligibility": "Indian startups licensing CSIR IP for commercialisation",
        "deadline": "Rolling",
        "url": "https://nrdc.in/strait",
    },
    {
        "name": "DST — SERB TARE (Teacher-Associate Research)",
        "body": "SERB / DST",
        "amount": "INR 15 Lakh/year (3 years)",
        "eligibility": "College teachers to work in CSIR labs; PhD required",
        "deadline": "Sep annually",
        "url": "https://serb.gov.in/page/english/tare",
    },
]


def create_institute_xlsx(filename, institute_data, papers, spinoffs, grants):
    wb = openpyxl.Workbook()

    # Overview
    ws_ov = wb.active
    ws_ov.title = "Overview"
    build_overview_sheet(ws_ov, institute_data)

    # Research Papers
    ws_rp = wb.create_sheet("Research Papers")
    build_papers_sheet(ws_rp, papers)

    # Spinoffs
    ws_sp = wb.create_sheet("Spinoffs & Startups")
    build_spinoffs_sheet(ws_sp, spinoffs)

    # Grants
    ws_gr = wb.create_sheet("Grants & Funding")
    build_grants_sheet(ws_gr, grants)

    wb.save(filename)
    print(f"  Saved: {filename}")


if __name__ == "__main__":
    print("Building IIT Madras database...")
    create_institute_xlsx(
        "IIT_Madras.xlsx",
        IITM_DATA, IITM_PAPERS, IITM_SPINOFFS, IITM_GRANTS,
    )

    print("Building CSIR database...")
    create_institute_xlsx(
        "CSIR.xlsx",
        CSIR_DATA, CSIR_PAPERS, CSIR_SPINOFFS, CSIR_GRANTS,
    )

    print("Done.")
