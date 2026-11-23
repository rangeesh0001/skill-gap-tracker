# Skill-gap tracker: Project Consultant

A small tool for finding and closing skill gaps for a Project Consultant role. It comes in two forms that use the same rules: an Excel workbook and a single-page web tool.

## The problem

Companies train heavily, but they cannot show which skill gaps exist in each role or whether training closed them, so training spend is hard to target and justify. Supporting facts: 59% of enterprise leaders say their organisation has an AI skills gap in 2026, even though most are already investing in AI training (DataCamp, 2026). L&D leaders are prioritising skills identification and gap analysis (SHRM, 2026).

## The solution

A tracker that, for one role:

1. Lists the required skills and what each level means.
2. Records a manager-confirmed level per person per skill (the employee's self-rating is not used for the gap).
3. Calculates gap = max(0, required level − confirmed level).
4. Records a re-rating after training.
5. Calculates gap closed = gap before − gap after.
6. Summarises by skill: people with a gap, total gap before and after, gap closed, and % closed.

## Who it's for

- Team leads and managers who rate their team against the role.
- L&D or HR people who plan training and want to see whether it worked.

## The skills

Eight skills taken from a Project Consultant job description, prepared with Project Consultant roles at Godrej Consumer Products in mind. The required levels are a proposal based on the JD wording, not a published standard. Change them to fit your team.

| Skill | Required level (1–4) |
|---|---|
| Scope & planning | 3 |
| Timelines & resources | 3 |
| Budget tracking | 2 |
| Stakeholder management | 3 |
| Risk management | 3 |
| Presentation & communication | 3 |
| Workshop facilitation & feedback | 3 |
| Tools (MS Office, G Suite, Monday.com, Zoom) | 2 |

The workbook's **Skill matrix** sheet describes what each level looks like for each skill.

## How the numbers are calculated

- **Gap** = required level − confirmed level, never below 0. A person above the required level has a gap of 0, not a negative one.
- **Gap after** = required level − re-rated level after training, never below 0.
- **Gap closed** = gap before − gap after, never below 0. It only counts for people who have both a before and an after rating.
- **% of gap closed** = gap closed ÷ gap before, counting only people who were re-rated. People not re-assessed yet don't lower the percentage.
- **Open gaps** for a person = the number of skills where their gap is above 0.

Example: required level 3, confirmed level 1 → gap 2. Re-rated at 3 after training → gap after 0, gap closed 2 (100%).

Note: "Total gap after" only includes people who were re-rated, so *gap before − gap after* equals *gap closed* only when everyone has been re-rated.

## How to use the spreadsheet

```
pip install openpyxl
python build/build_tracker.py            # blank template
python build/build_tracker.py --sample   # copy filled with hypothetical people
```

Both files are written to `out/`:

- `Skill_Gap_Tracker_Project_Consultant.xlsx`: blank template.
- `Skill_Gap_Tracker_Project_Consultant_SAMPLE_hypothetical.xlsx`: the same workbook with three invented people.

In the workbook, type only in the yellow cells:

1. **Skill matrix**: adjust the required levels.
2. **Ratings**: one row per person (room for 10), with the manager-confirmed level per skill. Leave a cell blank until it is confirmed. Gaps are colour-coded: green 0, amber 1, red 2 or more.
3. **Re-assessment**: the re-rated level after training. Names copy over from Ratings.
4. **Summary**: per-skill totals. Fills in by itself.

Level cells only accept whole numbers from 1 to 4.

## How to use the web tool

Open `web/index.html` in a browser. Nothing is installed and no data leaves the page.

1. Adjust required levels in section 1 if needed.
2. Add a person in section 2 and pick their confirmed levels.
3. Pick re-rated levels in section 3 after training.
4. Read the summary in section 4.

**Save in this browser** keeps the data in that browser only (it doesn't sync to other devices). **Load saved** brings it back. **Copy summary as text** copies the summary as tab-separated text you can paste into a sheet or email.

## Worked example: FMCG shop floor

`web/fmcg.html` applies the same method to a 10-person FMCG production team (filling, packing, mixing, maintenance, quality, stores and utilities) on 8 shop-floor skills, including EHS & safe work practices and quality & food safety.

| | Result |
|---|---|
| Open skill gaps before training | 45 |
| Levels of gap, before → after | 55 → 26 |
| Gap closed | 29 levels (53%) |
| Largest gap still open | EHS & safe work practices (9 levels, 40% closed) |

The employee ratings and post-training scores are **illustrative**, not measured, and names are replaced with employee numbers. The skills and priorities are backed by shop-floor process records (downtime Pareto, line output against benchmark, manpower deployment), shown on the page.

Rebuild the page from the workbook:

```
python build/build_fmcg_site.py path/to/FMCG_Shop_Floor_Skill_Gap_Tracker_Completed.xlsx
```

The script recalculates every gap and stops if its totals don't match the workbook's own Summary sheet.

## Limitations

- **Sample data only.** Every person and rating in this repo is hypothetical. No real employee data is included.
- **Manager confirmation needed.** The tracker is only as good as the ratings. Levels should be confirmed by a manager against the level descriptions, not self-rated.
- **No connection to an HR system.** Data is typed in by hand. Nothing syncs with an HRIS, LMS or training records.
- The required levels are a proposal from one job description and haven't been validated with a hiring team.
- The web tool saves in one browser only. Clearing browser data deletes it.

## Open placeholders

These details aren't confirmed yet, so they're marked `[add: …]` in `case-study/CASE_STUDY.md` ("Where this comes from"):

- LOTO workshop: your role, dates, number of attendees, whether feedback was collected
- Management seminar: your role, dates, audience, outcome

## Repo layout

```
README.md
build/build_tracker.py     builds the Excel workbook
web/index.html             the web tool
web/fmcg.html              FMCG shop-floor results page
build/build_fmcg_site.py   builds web/fmcg.html from the FMCG workbook
out/                       generated workbooks
case-study/CASE_STUDY.md   a worked run on the hypothetical sample
```
