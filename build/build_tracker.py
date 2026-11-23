from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from pathlib import Path
import sys

# `python build/build_tracker.py` writes the blank template.
# `python build/build_tracker.py --sample` writes a copy filled with HYPOTHETICAL people.
SAMPLE_MODE = "--sample" in sys.argv
OUT_DIR = Path(__file__).resolve().parent.parent / "out"
OUT_DIR.mkdir(exist_ok=True)
OUT = OUT_DIR / ("Skill_Gap_Tracker_Project_Consultant"
                 + ("_SAMPLE_hypothetical" if SAMPLE_MODE else "") + ".xlsx")

# HYPOTHETICAL sample data: invented people and ratings, not real employees.
# Levels follow the SKILLS order below; None = not rated yet.
SAMPLE_PEOPLE = [
    ("Hypothetical person A", "Sample only",
     [2, 1, 2, 3, 2, 2, 1, 3], [3, 2, 2, 3, 3, 2, 2, 3]),
    ("Hypothetical person B", "Sample only",
     [3, 2, 1, 2, 3, 1, 2, 2], [3, 3, 2, 3, 3, 2, 3, 2]),
    ("Hypothetical person C", "Sample only (not re-rated yet)",
     [1, 2, 2, 2, 2, 3, 2, 1], [None] * 8),
]

NAVY = "1F2F6B"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF8E1")
thin = Side(style="thin", color="BBBBBB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

# Skills taken from a Project Consultant JD (aimed at Project Consultant roles at Godrej Consumer Products).
# Required levels are a proposal based on how the JD describes each responsibility; adjust them.
# KEEP IN SYNC: web/index.html (DEFAULT_SKILLS) has the same skill names and required levels.
SKILLS = [
    ("Scope & planning",
     "Defines project scope, delivery plans and scope statements (JD: scope, delivery plans, tasks).",
     3,
     "Can describe a scope and a delivery plan, but needs help on what is out of scope.",
     "Writes a scope and delivery plan; identifies what is out of scope.",
     "Keeps scope stable through changes; re-plans when scope moves.",
     "Sets scope for several projects at once and resolves conflicts between them."),
    ("Timelines & resources",
     "Builds timelines and assigns resources, including dates and people (JD: timelines, resources).",
     3,
     "Lists tasks but dates are rough or missing.",
     "Builds a timeline with dates and owners for each task.",
     "Updates the timeline when dependencies change; flags delays early.",
     "Balances resources across several projects and reshapes timelines to match."),
    ("Budget tracking",
     "Plans budgets and tracks planned against actual cost (JD: budgets).",
     2,
     "Knows the budget but does not track actuals.",
     "Tracks planned against actual cost for a project.",
     "Explains variances and proposes corrective actions.",
     "Owns budgets across several projects and reports to senior management."),
    ("Stakeholder management",
     "Aligns internal and external stakeholders on timelines, deliverables and quality (JD: stakeholder alignment).",
     3,
     "Knows who the stakeholders are.",
     "Keeps stakeholders informed with regular updates.",
     "Gets agreement from stakeholders with differing views.",
     "Manages senior and external stakeholders through difficult decisions."),
    ("Risk management",
     "Identifies project risks, develops mitigations and manages issues (JD: risks, mitigation).",
     3,
     "Can name risks when asked.",
     "Keeps a risk list with a mitigation for each risk.",
     "Reviews risks regularly and acts on early warning signs.",
     "Anticipates risks across several projects and reports them to senior management."),
    ("Presentation & communication",
     "Presents project status and recommendations to stakeholders and senior management (JD: presentation skills, communicate status).",
     3,
     "Presents content but the message is not always clear to the audience.",
     "Presents clearly to a defined audience with a clear recommendation.",
     "Adapts the presentation to different audiences, including senior management.",
     "Presents to senior leaders and gets decisions."),
    ("Workshop facilitation & feedback",
     "Runs workshops and sessions and collects, collates and publishes feedback (JD: feedback for each module).",
     3,
     "Attends or assists in sessions.",
     "Runs a session with a plan and collects feedback.",
     "Designs sessions and collates feedback into a published summary.",
     "Runs a programme of sessions and uses feedback to change later sessions."),
    ("Tools (MS Office, G Suite, Monday.com, Zoom)",
     "Uses the tools the JD lists for project work (JD: skill set).",
     2,
     "Has used one or two of the listed tools.",
     "Has used all the listed tools for project work.",
     "Sets up boards and templates in the listed tools for others.",
     "Designs project processes around the tools and trains others."),
]

wb = Workbook()

# --- Sheet 1: Skill matrix -------------------------------------------------
ws = wb.active
ws.title = "Skill matrix"
ws["A1"] = "Project Consultant skill matrix (from a Project Consultant JD)"
ws["A1"].font = Font(bold=True, size=13, color=NAVY)
ws["A2"] = ("Required levels are a proposal from the JD wording; change the yellow cells. "
            "Levels are 1 to 4. The Ratings sheet reads these cells.")
ws["A2"].font = Font(italic=True, color="555555")

headers = ["Skill", "What it means", "Required level", "Level 1", "Level 2", "Level 3", "Level 4"]
for c, h in enumerate(headers, 1):
    cell = ws.cell(row=4, column=c, value=h)
    cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="center")

for r, (skill, meaning, req, l1, l2, l3, l4) in enumerate(SKILLS, start=5):
    vals = [skill, meaning, req, l1, l2, l3, l4]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.alignment = WRAP
        cell.border = BORDER
    ws.cell(row=r, column=3).fill = INPUT_FILL
    ws.cell(row=r, column=3).alignment = Alignment(horizontal="center", vertical="top")

for col, width in zip("ABCDEFG", [30, 42, 12, 34, 34, 34, 34]):
    ws.column_dimensions[col].width = width
ws.row_dimensions[4].height = 30
ws.freeze_panes = "B5"

LAST_SKILL_ROW = 4 + len(SKILLS)

# --- Sheet 2: Ratings (confirmed by the manager) ---------------------------
rs = wb.create_sheet("Ratings")
rs["A1"] = "Confirmed levels per person (manager's rating is the one used)"
rs["A1"].font = Font(bold=True, size=13, color=NAVY)
rs["A2"] = ("Enter a level 1 to 4 in each yellow cell. Gap = required level minus confirmed level, never below 0. "
            "Leave a cell blank until it is confirmed.")
rs["A2"].font = Font(italic=True, color="555555")

N_PEOPLE = 10  # room for ten people; add rows by copying formulas down
FIRST = 6      # first data row
LAST = FIRST + N_PEOPLE - 1

# Columns: A person, B role/team, C..J confirmed level per skill, K..R gap per skill, S open gaps
conf_start = 3
gap_start = conf_start + len(SKILLS)  # column index of first gap column

rs.cell(row=4, column=conf_start, value="Confirmed level (1–4)")
rs.cell(row=4, column=gap_start, value="Gap (required − confirmed)")
rs.merge_cells(start_row=4, start_column=conf_start, end_row=4, end_column=conf_start + len(SKILLS) - 1)
rs.merge_cells(start_row=4, start_column=gap_start, end_row=4, end_column=gap_start + len(SKILLS) - 1)
for col in range(conf_start, gap_start + len(SKILLS)):
    rs.cell(row=4, column=col).fill = HDR_FILL
    rs.cell(row=4, column=col).font = HDR_FONT
    rs.cell(row=4, column=col).alignment = Alignment(horizontal="center")

skill_names = [s[0] for s in SKILLS]
heads = ["Person", "Role / team"] + skill_names + skill_names + ["Open gaps"]
for c, h in enumerate(heads, 1):
    cell = rs.cell(row=5, column=c, value=h)
    cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
rs.row_dimensions[5].height = 58

open_col = gap_start + len(SKILLS)

for r in range(FIRST, LAST + 1):
    rs.cell(row=r, column=1).fill = INPUT_FILL
    rs.cell(row=r, column=2).fill = INPUT_FILL
    rs.cell(row=r, column=1).border = BORDER
    rs.cell(row=r, column=2).border = BORDER
    for i in range(len(SKILLS)):
        cc = conf_start + i
        rs.cell(row=r, column=cc).fill = INPUT_FILL
        rs.cell(row=r, column=cc).border = BORDER
        rs.cell(row=r, column=cc).alignment = Alignment(horizontal="center")
        # Gap formula reads the required level from the Skill matrix row for this skill.
        gc = gap_start + i
        conf_ref = f"{L(cc)}{r}"
        req_ref = f"'Skill matrix'!$C${5 + i}"
        rs.cell(row=r, column=gc, value=f'=IF({conf_ref}="","",MAX(0,{req_ref}-{conf_ref}))')
        rs.cell(row=r, column=gc).alignment = Alignment(horizontal="center")
        rs.cell(row=r, column=gc).border = BORDER
    gap_range = f"{L(gap_start)}{r}:{L(gap_start + len(SKILLS) - 1)}{r}"
    rs.cell(row=r, column=open_col, value=f'=COUNTIF({gap_range},">0")')
    rs.cell(row=r, column=open_col).alignment = Alignment(horizontal="center")
    rs.cell(row=r, column=open_col).border = BORDER

# Heatmap on the gap columns: 0 green, 1 amber, 2+ red.
gap_block = f"{L(gap_start)}{FIRST}:{L(gap_start + len(SKILLS) - 1)}{LAST}"
rs.conditional_formatting.add(gap_block, CellIsRule(operator="equal", formula=["0"],
                              fill=PatternFill("solid", fgColor="C8E6C9")))
rs.conditional_formatting.add(gap_block, CellIsRule(operator="equal", formula=["1"],
                              fill=PatternFill("solid", fgColor="FFE082")))
rs.conditional_formatting.add(gap_block, CellIsRule(operator="greaterThanOrEqual", formula=["2"],
                              fill=PatternFill("solid", fgColor="EF9A9A")))

rs.column_dimensions["A"].width = 22
rs.column_dimensions["B"].width = 18
for i in range(len(SKILLS) * 2):
    rs.column_dimensions[L(conf_start + i)].width = 14
rs.column_dimensions[L(open_col)].width = 12
rs.freeze_panes = "C6"

# --- Sheet 3: Re-assessment after training ---------------------------------
rr = wb.create_sheet("Re-assessment")
rr["A1"] = "After training: re-rated level per person, and gap closed"
rr["A1"].font = Font(bold=True, size=13, color=NAVY)
rr["A2"] = ("Enter the re-rated level after training. Closed = gap before (from Ratings) minus gap after, never below 0. "
            "Only rows with both a before and an after level count.")
rr["A2"].font = Font(italic=True, color="555555")

rr.cell(row=4, column=1, value="Person")
rr.cell(row=4, column=2, value="Role / team")
for i, (skill, *_rest) in enumerate(SKILLS):
    rr.cell(row=4, column=3 + i, value=f"Re-rated: {skill}")
for i, (skill, *_rest) in enumerate(SKILLS):
    rr.cell(row=4, column=3 + len(SKILLS) + i, value=f"Gap after: {skill}")
for i, (skill, *_rest) in enumerate(SKILLS):
    rr.cell(row=4, column=3 + 2 * len(SKILLS) + i, value=f"Closed: {skill}")
for c in range(1, 3 + 3 * len(SKILLS) + 1):
    cell = rr.cell(row=4, column=c)
    cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
rr.row_dimensions[4].height = 58

after_start = 3 + len(SKILLS)
closed_start = 3 + 2 * len(SKILLS)
for r in range(FIRST, LAST + 1):
    rr.cell(row=r, column=1, value=f"=IF(Ratings!A{r}=\"\",\"\",Ratings!A{r})")
    rr.cell(row=r, column=2, value=f"=IF(Ratings!B{r}=\"\",\"\",Ratings!B{r})")
    for i in range(len(SKILLS)):
        re_col = 3 + i
        rr.cell(row=r, column=re_col).fill = INPUT_FILL
        rr.cell(row=r, column=re_col).border = BORDER
        rr.cell(row=r, column=re_col).alignment = Alignment(horizontal="center")
        req_ref = f"'Skill matrix'!$C${5 + i}"
        after_col = after_start + i
        re_ref = f"{L(re_col)}{r}"
        rr.cell(row=r, column=after_col,
                value=f'=IF({re_ref}="","",MAX(0,{req_ref}-{re_ref}))')
        rr.cell(row=r, column=after_col).alignment = Alignment(horizontal="center")
        rr.cell(row=r, column=after_col).border = BORDER
        before_ref = f"Ratings!{L(gap_start + i)}{r}"
        after_ref = f"{L(after_col)}{r}"
        rr.cell(row=r, column=closed_start + i,
                value=f'=IF(OR({before_ref}="",{after_ref}=""),"",MAX(0,{before_ref}-{after_ref}))')
        rr.cell(row=r, column=closed_start + i).alignment = Alignment(horizontal="center")
        rr.cell(row=r, column=closed_start + i).border = BORDER

rr.column_dimensions["A"].width = 22
rr.column_dimensions["B"].width = 18
for c in range(3, 3 + 3 * len(SKILLS)):
    rr.column_dimensions[L(c)].width = 14
rr.freeze_panes = "C5"

# --- Sheet 4: Summary -------------------------------------------------------
sm = wb.create_sheet("Summary")
sm["A1"] = "Summary by skill"
sm["A1"].font = Font(bold=True, size=13, color=NAVY)
sm["A2"] = "All figures update from the Ratings and Re-assessment sheets."
sm["A2"].font = Font(italic=True, color="555555")
cols = ["Skill", "People with an open gap (before)", "Total gap before", "Total gap after",
        "Total gap closed", "% of gap closed (people re-rated)"]
for c, h in enumerate(cols, 1):
    cell = sm.cell(row=4, column=c, value=h)
    cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
sm.row_dimensions[4].height = 44

for i, (skill, *_rest) in enumerate(SKILLS):
    r = 5 + i
    gcol = L(gap_start + i)
    acol = L(after_start + i)
    ccol = L(closed_start + i)
    sm.cell(row=r, column=1, value=skill)
    sm.cell(row=r, column=2, value=f'=COUNTIF(Ratings!{gcol}{FIRST}:{gcol}{LAST},">0")')
    sm.cell(row=r, column=3, value=f'=SUM(Ratings!{gcol}{FIRST}:{gcol}{LAST})')
    sm.cell(row=r, column=4, value=f"=SUM('Re-assessment'!{acol}{FIRST}:{acol}{LAST})")
    sm.cell(row=r, column=5, value=f'=SUM(\'Re-assessment\'!{ccol}{FIRST}:{ccol}{LAST})')
    # % closed only counts the gap-before of people who were re-rated,
    # otherwise people not re-assessed yet would drag the percentage down.
    before_rerated = (f"SUMPRODUCT((Ratings!{gcol}{FIRST}:{gcol}{LAST}<>\"\")"
                      f"*('Re-assessment'!{acol}{FIRST}:{acol}{LAST}<>\"\"),"
                      f"Ratings!{gcol}{FIRST}:{gcol}{LAST})")
    sm.cell(row=r, column=6, value=f'=IF({before_rerated}=0,"",E{r}/{before_rerated})')
    sm.cell(row=r, column=6).number_format = "0%"
    for c in range(1, 7):
        sm.cell(row=r, column=c).border = BORDER
        sm.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top",
                                                       horizontal="left" if c == 1 else "center")

sm.column_dimensions["A"].width = 40
for col in "BCDEF":
    sm.column_dimensions[col].width = 18

tot = 5 + len(SKILLS)
sm.cell(row=tot, column=1, value="Total").font = Font(bold=True)
for c in range(2, 6):
    col = L(c)
    sm.cell(row=tot, column=c, value=f"=SUM({col}5:{col}{tot - 1})").font = Font(bold=True)
for c in range(1, 7):
    sm.cell(row=tot, column=c).border = BORDER
    sm.cell(row=tot, column=c).alignment = Alignment(horizontal="left" if c == 1 else "center")

# Only whole levels 1 to 4 (or blank) can be typed into the yellow level cells.
def levels_only(sheet, ref):
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="4",
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="Level 1 to 4", error="Enter a whole number from 1 to 4, or leave it blank.")
    sheet.add_data_validation(dv)
    dv.add(ref)

levels_only(ws, f"C5:C{LAST_SKILL_ROW}")
levels_only(rs, f"{L(conf_start)}{FIRST}:{L(conf_start + len(SKILLS) - 1)}{LAST}")
levels_only(rr, f"C{FIRST}:{L(2 + len(SKILLS))}{LAST}")

if SAMPLE_MODE:
    ws["A3"] = "SAMPLE FILE: every person and rating in this file is HYPOTHETICAL."
    ws["A3"].font = Font(bold=True, color="B71C1C")
    for n, (name, role, before, after) in enumerate(SAMPLE_PEOPLE):
        r = FIRST + n
        rs.cell(row=r, column=1, value=name)
        rs.cell(row=r, column=2, value=role)
        for i, (b, a) in enumerate(zip(before, after)):
            rs.cell(row=r, column=conf_start + i, value=b)
            rr.cell(row=r, column=3 + i, value=a)

wb.save(OUT)
print("saved", OUT)
