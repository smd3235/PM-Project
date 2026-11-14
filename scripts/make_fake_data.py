"""Generate a fake raw-data sample: 12 managers (incl. Rhonda Adomson for exclusion
testing), ~120 data rows across 2024-2026, mixed Active / Pending Pickup statuses.
Keeps the 4 real rows from the real export, then appends generated rows.
"""
import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

random.seed(42)

MANAGERS = [
    ("John Murphy", ["Tony Charles", "Steve Smith"]),
    ("Paul Atkin", ["Steve Smith", "Maria Lopez", "Devon Pierce"]),
    ("Mich Peters", ["Barb Rogers", "Colin Hart"]),
    ("Karen Whitfield", ["Priya Nair", "Owen Blake", "Tessa Grant"]),
    ("Derek Owens", ["Luis Vega", "Hannah Sood"]),
    ("Sandra Kim", ["Marcus Webb", "Elena Petrova", "Jordan Lee"]),
    ("Victor Ramos", ["Aisha Bello", "Tom Kaczmarek"]),
    ("Alice Fontaine", ["Nina Rossi", "Sam Whitby", "Grace Ofori"]),
    ("Brian Cully", ["Dan Kowalski", "Mira Chandra"]),
    ("Naomi Osei", ["Kofi Mensah", "Julia Brandt", "Ravi Patel"]),
    ("Frank Delgado", ["Sofia Marino", "Ethan Cole"]),
    ("Rhonda Adomson", ["Adam Clarkson", "Bethany Fisk"]),
]

FIRST = ["James", "Mary", "Robert", "Linda", "Michael", "Patricia", "David", "Barbara",
         "William", "Elizabeth", "Richard", "Jennifer", "Joseph", "Maria", "Thomas",
         "Susan", "Charles", "Margaret", "Christopher", "Dorothy", "Daniel", "Lisa",
         "Matthew", "Sandra", "Anthony", "Ashley", "Mark", "Kimberly", "Donald", "Emily"]
LAST = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
        "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson",
        "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson",
        "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson"]
PHYSICIANS = ["Smith, Paul", "Jordan, Kyle", "Lemnon, James", "Ingish, Francine",
              "Bakare, Tunde", "Voss, Ingrid", "Hale, Cynthia", "Marsh, Doyle",
              "Nakamura, Ken", "Oyelaran, Bisi", "Petrov, Alexei", "Quinn, Siobhan"]
CITIES = ["Dallas", "New York", "Nashville", "Chicago", "Houston", "Phoenix",
          "Atlanta", "Denver", "Seattle", "Boston", "Miami", "Detroit",
          "Minneapolis", "Orlando", "Cleveland", "Portland", "Austin", "Tampa"]
TECHS = ["Tony Charles", "Steve Smith", "Barb Rogers", "Adam Clarkson", "Maria Lopez",
         "Owen Blake", "Priya Nair", "Marcus Webb", "Luis Vega", "Hannah Sood",
         "Nina Rossi", "Sam Whitby", "Dan Kowalski", "Kofi Mensah"]

def rand_date(start: date, end: date) -> date:
    return start + timedelta(days=random.randint(0, (end - start).days))

def gen_rows(n: int):
    rows = []
    for i in range(n):
        mgr, reps = random.choice(MANAGERS)
        rep = random.choice(reps)
        status = "Pending Pickup" if random.random() < 0.12 else "Active"
        start = rand_date(date(2024, 6, 1), date(2026, 11, 1))
        pm = rand_date(date(2025, 1, 1), date(2027, 6, 30)) if random.random() < 0.85 else None
        end = rand_date(pm, date(2027, 9, 30)) if (pm and random.random() < 0.7) else None
        setup = rand_date(start, date(2026, 12, 15)) if random.random() < 0.25 else None
        name = f"{random.choice(LAST)}, {random.choice(FIRST)} {start:%Y/%m/%d}"
        rows.append([
            rep, mgr, name, random.choice(CITIES), random.choice(PHYSICIANS),
            random.choice(CITIES), status, start, pm, end, setup,
            random.choice(TECHS) if random.random() < 0.9 else None,
        ])
    return rows

# 4 real rows preserved from the real export
REAL = [
    ["Tony Charles", "John Murphy", "Sektor, Peter 2025/08/13", "Chiago", "Smith, Paul",
     "Chiago", "Active", date(2025, 8, 18), None, None, None, "Tony Charles"],
    ["Steve Smith", "Paul Atkin", "Oxeran, Danille 2025/12/11", "Dallas", "Jordan, Kyle",
     "Dallas", "Active", date(2025, 12, 12), None, None, None, "Steve Smith"],
    ["Barb Rogers", "Mich Peters", "Abraja, Nacy 2024/10/29", "New York", "Lemnon, James",
     "New York", "Active", date(2024, 11, 19), date(2025, 12, 17), None, None, "Barb Rogers"],
    ["Adam Clarkson", "Rhonda Adomson", "Bahdor, Kyle 2025/02/09", "Nashville", "Ingish, Francine",
     "Nashville", "Pending Pickup", date(2025, 6, 27), None, date(2026, 4, 7), None, "Adam Clarkson"],
]

HEADERS = ["Assigned Sales Representative", "Assigned Sales Manager", "Name", "Patient",
           "Treatment Physician", "Customer", "Treatment Status", "Treatment Start Date",
           "Treatment PM Date", "Treatment End Date", "Equipment Setup Date",
           "Assigned Setup Technician"]

wb = Workbook()
ws = wb.active
ws.title = "AOTIOpsTreatmentPMDatesResults"

hdr_font = Font(name="Arial", size=7, bold=True)
hdr_fill = PatternFill("solid", fgColor="FFD0D0D0")
for c, h in enumerate(HEADERS, 1):
    cell = ws.cell(1, c, h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(vertical="center", wrap_text=True)

widths = [24, 22, 30, 14, 22, 14, 16, 14, 14, 14, 14, 24]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

rows = REAL + gen_rows(120)
for r, row in enumerate(rows, 2):
    for c, v in enumerate(row, 1):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Arial", size=10)
        if c in (8, 9, 10, 11) and v:
            cell.number_format = "mm-dd-yy"

out = r"C:\Users\smd32\Downloads\PM Project\AOTIOpsTreatmentPMDatesResults_FAKE.xlsx"
wb.save(out)

from collections import Counter
mgrs = Counter(r[1] for r in rows)
stats = Counter(r[6] for r in rows)
print("saved:", out)
print("total rows:", len(rows))
print("managers:", len(mgrs), dict(mgrs))
print("statuses:", dict(stats))
print("pending-pickup rows (incl. Rhonda):", stats["Pending Pickup"])