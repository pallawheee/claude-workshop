"""Build Master_Index.xlsx linking all institute files with summary stats."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

NAVY = "0D1B4B"
GOLD  = "C9A84C"
WHITE = "FFFFFF"

def navy_fill():
    return PatternFill("solid", fgColor=NAVY)

def gold_fill():
    return PatternFill("solid", fgColor=GOLD)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(bold=True, color=WHITE, size=11):
    return Font(name="Calibri", bold=bold, color=color, size=size)

def body(bold=False, color="000000", size=10):
    return Font(name="Calibri", bold=bold, color=color, size=size)

# ── Registry of all 11 institutes ───────────────────────────────────────────

INSTITUTES = [
    {
        "name": "IIT Madras",
        "code": "IITM",
        "file": "IIT_Madras.xlsx",
        "type": "IIT",
        "location": "Chennai, TN",
        "domains": "AI, Aerospace, Quantum, Biotech, Energy",
        "papers_count": 5,
        "spinoffs_count": 6,
        "grants_count": 5,
        "status": "Active",
        "last_refreshed": "2026-03-19",
        "notes": "Research Park model; IITM #1 NIRF 2023",
    },
    {
        "name": "CSIR",
        "code": "CSIR",
        "file": "CSIR.xlsx",
        "type": "National Lab Network",
        "location": "Pan-India (37 labs)",
        "domains": "Pharma, Materials, Genomics, Aerospace, Chemicals",
        "papers_count": 5,
        "spinoffs_count": 4,
        "grants_count": 5,
        "status": "Active",
        "last_refreshed": "2026-03-19",
        "notes": "Largest publicly funded R&D org in India",
    },
    {
        "name": "IIT Bombay",
        "code": "IITB",
        "file": "IIT_Bombay.xlsx",
        "type": "IIT",
        "location": "Mumbai, MH",
        "domains": "Fintech, AI, Photonics, Nano, Manufacturing",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "SINE incubator; proximity to Mumbai startup ecosystem",
    },
    {
        "name": "IIT Delhi",
        "code": "IITD",
        "file": "IIT_Delhi.xlsx",
        "type": "IIT",
        "location": "New Delhi",
        "domains": "Policy-Tech, AI, Textiles, Smart Cities, Health",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "Strong policy-research linkage; FITT TTO",
    },
    {
        "name": "IIT Hyderabad",
        "code": "IITH",
        "file": "IIT_Hyderabad.xlsx",
        "type": "IIT",
        "location": "Hyderabad, TS",
        "domains": "AI/ML, Pharma, Semiconductors, EVs",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "Located in Hyderabad pharma/tech hub",
    },
    {
        "name": "IIT Kharagpur",
        "code": "IITKgp",
        "file": "IIT_Kharagpur.xlsx",
        "type": "IIT",
        "location": "Kharagpur, WB",
        "domains": "Mining, Agriculture-Tech, Railways, AI, Cyber",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "Oldest IIT; Rajendra Misra Centre for Cybersecurity",
    },
    {
        "name": "IISc",
        "code": "IISc",
        "file": "IISc.xlsx",
        "type": "Research University",
        "location": "Bengaluru, KA",
        "domains": "Aerospace, Materials, Quantum, Biotech, AI",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "India's top research university; Society for Innovation & Development",
    },
    {
        "name": "DRDO",
        "code": "DRDO",
        "file": "DRDO.xlsx",
        "type": "Defence R&D",
        "location": "Pan-India (50+ labs)",
        "domains": "Defence Systems, Missiles, Electronics, Materials, Cyber",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "TDF programme for startup & academia collaboration",
    },
    {
        "name": "ISRO",
        "code": "ISRO",
        "file": "ISRO.xlsx",
        "type": "Space Agency",
        "location": "Bengaluru (HQ) + centres",
        "domains": "Launch, Satellites, Earth Obs, Quantum Comms",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "IN-SPACe enables private sector collaboration",
    },
    {
        "name": "NRDC",
        "code": "NRDC",
        "file": "NRDC.xlsx",
        "type": "Technology Transfer Organisation",
        "location": "New Delhi",
        "domains": "IP Licensing, Technology Commercialisation (all sectors)",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "National Research Development Corporation — DSIR under MoST",
    },
    {
        "name": "BARC",
        "code": "BARC",
        "file": "BARC.xlsx",
        "type": "Nuclear R&D",
        "location": "Mumbai, MH",
        "domains": "Nuclear Energy, Radiation Tech, Isotopes, Materials",
        "papers_count": 0,
        "spinoffs_count": 0,
        "grants_count": 0,
        "status": "Pending",
        "last_refreshed": "—",
        "notes": "Department of Atomic Energy; dual-use civilian/strategic",
    },
]


def build_master_index():
    wb = openpyxl.Workbook()

    # ── Sheet 1: Institute Directory ─────────────────────────────────────────
    ws = wb.active
    ws.title = "Institute Directory"
    ws.freeze_panes = "A4"

    # Title banner
    ncols = 12
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    t = ws.cell(row=1, column=1, value="UIV India Deeptech Research Institute Intelligence")
    t.fill = navy_fill()
    t.font = Font(name="Calibri", bold=True, color=GOLD, size=18)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    s = ws.cell(row=2, column=1, value=f"Master Index — {len(INSTITUTES)} Institutes | Generated {datetime.now().strftime('%Y-%m-%d')}")
    s.fill = gold_fill()
    s.font = Font(name="Calibri", bold=True, color=NAVY, size=11)
    s.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    headers = ["#", "Institute", "Code", "File", "Type", "Location",
               "Key Domains", "Papers", "Spinoffs", "Grants", "Status", "Last Refreshed"]
    col_widths = [4, 28, 8, 22, 22, 18, 40, 8, 9, 7, 10, 14]
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.column_dimensions[get_column_letter(c)].width = w
        cell = ws.cell(row=3, column=c, value=h)
        cell.fill = navy_fill()
        cell.font = hdr()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[3].height = 22

    for i, inst in enumerate(INSTITUTES, start=4):
        alt = PatternFill("solid", fgColor="EEF0F8") if i % 2 == 0 else PatternFill("solid", fgColor=WHITE)
        status_fill = PatternFill("solid", fgColor="D4EDDA") if inst["status"] == "Active" else PatternFill("solid", fgColor="FFF3CD")
        vals = [i - 3, inst["name"], inst["code"], inst["file"], inst["type"],
                inst["location"], inst["domains"], inst["papers_count"],
                inst["spinoffs_count"], inst["grants_count"], inst["status"],
                inst["last_refreshed"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.fill = status_fill if c == 11 else alt
            cell.font = body()
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="center" if c in (1,7,8,9,10,11,12) else "left")
            cell.border = thin_border()
        ws.row_dimensions[i].height = 22

    # ── Sheet 2: Summary Stats ────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary Stats")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 20

    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    t2 = ws2.cell(row=1, column=1, value="UIV Deeptech Database — Summary Statistics")
    t2.fill = navy_fill()
    t2.font = Font(name="Calibri", bold=True, color=GOLD, size=14)
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 30

    total_papers = sum(i["papers_count"] for i in INSTITUTES)
    total_spinoffs = sum(i["spinoffs_count"] for i in INSTITUTES)
    total_grants = sum(i["grants_count"] for i in INSTITUTES)
    active = sum(1 for i in INSTITUTES if i["status"] == "Active")

    stats = [
        ("Total Institutes Tracked", len(INSTITUTES)),
        ("Active (data populated)", active),
        ("Pending (data TBD)", len(INSTITUTES) - active),
        ("Total Research Papers", total_papers),
        ("Total Spinoffs / Startups", total_spinoffs),
        ("Total Grant Schemes", total_grants),
        ("Database Version", "1.0.0"),
        ("Generated On", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Primary Brand Color — Navy", "#0D1B4B"),
        ("Primary Brand Color — Gold", "#C9A84C"),
        ("Maintained By", "UIV Intelligence Team"),
        ("Paper Refresh Tool", "refresh.py (Semantic Scholar API)"),
    ]

    for row, (label, val) in enumerate(stats, start=3):
        alt_fill = PatternFill("solid", fgColor="EEF0F8") if row % 2 == 0 else PatternFill("solid", fgColor=WHITE)
        a = ws2.cell(row=row, column=1, value=label)
        a.font = Font(name="Calibri", bold=True, color=NAVY, size=10)
        a.fill = alt_fill
        a.border = thin_border()
        a.alignment = Alignment(vertical="center")
        b = ws2.cell(row=row, column=2, value=val)
        b.font = body()
        b.fill = alt_fill
        b.border = thin_border()
        b.alignment = Alignment(vertical="center", horizontal="center")
        ws2.row_dimensions[row].height = 18

    # ── Sheet 3: Domain Matrix ────────────────────────────────────────────────
    ws3 = wb.create_sheet("Domain Coverage Matrix")
    domains = ["AI / ML", "Aerospace", "Quantum", "Biotech / Health", "Clean Energy",
               "Semiconductors", "Robotics", "Defence / Dual-Use", "Pharmaceuticals",
               "Advanced Materials", "Space", "Nuclear", "Cyber"]
    institute_domains = {
        "IITM": ["AI / ML", "Aerospace", "Quantum", "Biotech / Health", "Clean Energy", "Semiconductors", "Robotics"],
        "CSIR": ["Biotech / Health", "Clean Energy", "Advanced Materials", "Pharmaceuticals", "Defence / Dual-Use"],
        "IITB": ["AI / ML", "Semiconductors", "Advanced Materials"],
        "IITD": ["AI / ML", "Biotech / Health", "Cyber"],
        "IITH": ["AI / ML", "Pharmaceuticals", "Semiconductors"],
        "IITKgp": ["AI / ML", "Cyber", "Advanced Materials"],
        "IISc": ["Aerospace", "Quantum", "Biotech / Health", "Advanced Materials"],
        "DRDO": ["Defence / Dual-Use", "Aerospace", "Cyber", "Advanced Materials"],
        "ISRO": ["Space", "Aerospace", "Quantum"],
        "NRDC": ["AI / ML", "Biotech / Health", "Clean Energy"],
        "BARC": ["Nuclear", "Advanced Materials", "Defence / Dual-Use"],
    }
    codes = [i["code"] for i in INSTITUTES]

    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(codes)+1)
    t3 = ws3.cell(row=1, column=1, value="Domain Coverage Matrix — ● = Active Coverage")
    t3.fill = navy_fill()
    t3.font = Font(name="Calibri", bold=True, color=GOLD, size=13)
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 28

    ws3.column_dimensions["A"].width = 22
    ws3.cell(row=2, column=1, value="Domain").fill = gold_fill()
    ws3.cell(row=2, column=1).font = hdr(color=NAVY)
    ws3.cell(row=2, column=1).border = thin_border()
    ws3.cell(row=2, column=1).alignment = Alignment(horizontal="center", vertical="center")

    for c, code in enumerate(codes, 2):
        ws3.column_dimensions[get_column_letter(c)].width = 8
        cell = ws3.cell(row=2, column=c, value=code)
        cell.fill = gold_fill()
        cell.font = hdr(color=NAVY, size=9)
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[2].height = 18

    for r, domain in enumerate(domains, start=3):
        alt_fill = PatternFill("solid", fgColor="EEF0F8") if r % 2 == 0 else PatternFill("solid", fgColor=WHITE)
        d_cell = ws3.cell(row=r, column=1, value=domain)
        d_cell.font = Font(name="Calibri", bold=True, color=NAVY, size=10)
        d_cell.fill = alt_fill
        d_cell.border = thin_border()
        d_cell.alignment = Alignment(vertical="center")
        ws3.row_dimensions[r].height = 18
        for c, code in enumerate(codes, 2):
            covered = domain in institute_domains.get(code, [])
            v = "●" if covered else ""
            cell = ws3.cell(row=r, column=c, value=v)
            cell.fill = PatternFill("solid", fgColor="D4EDDA") if covered else alt_fill
            cell.font = Font(name="Calibri", color="1A7A3C" if covered else "888888", size=12, bold=covered)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border()

    wb.save("Master_Index.xlsx")
    print("  Saved: Master_Index.xlsx")


if __name__ == "__main__":
    build_master_index()
