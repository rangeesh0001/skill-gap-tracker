# Skill-gap tracker: Project Consultant

A small tool for finding and closing skill gaps for a Project Consultant role. It comes in two forms that use the same rules: an Excel workbook and a single-page web tool.

## The problem

A team lead needs to know which skills each person on a project team still has to build for the Project Consultant role, and whether training actually closed those gaps. That usually lives in people's heads or in scattered notes, so nobody can say which skills need training first or show what the training changed.

This tracker sets a required level for each skill, records a manager-confirmed level for each person, and calculates the gap before and after training.

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

## Limitations

- **Sample data only.** Every person and rating in this repo is hypothetical. No real employee data is included.
- **Manager confirmation needed.** The tracker is only as good as the ratings. Levels should be confirmed by a manager against the level descriptions, not self-rated.
- **No connection to an HR system.** Data is typed in by hand. Nothing syncs with an HRIS, LMS or training records.
- The required levels are a proposal from one job description and haven't been validated with a hiring team.
- The web tool saves in one browser only. Clearing browser data deletes it.

## Repo layout

```
README.md
build/build_tracker.py     builds the Excel workbook
web/index.html             the web tool
out/                       generated workbooks
case-study/CASE_STUDY.md   a worked run on the hypothetical sample
```
