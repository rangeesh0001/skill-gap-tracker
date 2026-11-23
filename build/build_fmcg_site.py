"""Build web/fmcg.html, a results page for the FMCG shop-floor skill-gap workbook.

Usage: python build/build_fmcg_site.py <path to FMCG_Shop_Floor_Skill_Gap_Tracker_Completed.xlsx>

Names in the workbook are replaced with "Employee 01..10" so no person's name reaches the repo.
Gaps are recalculated here and checked against the workbook's own Summary sheet.
"""
import sys
from html import escape
from pathlib import Path
import openpyxl

SRC = Path(sys.argv[1])
OUT = Path(__file__).resolve().parent.parent / "web" / "fmcg.html"
N_SKILLS = 8

wb = openpyxl.load_workbook(SRC, data_only=True)
sm, rt, ra, su, ev = (wb[n] for n in ["Skill matrix", "Ratings", "Re-assessment", "Summary", "Operational evidence"])

skills = [dict(name=sm.cell(r, 1).value, meaning=sm.cell(r, 2).value, req=sm.cell(r, 3).value,
               levels=[sm.cell(r, c).value for c in range(4, 8)], evidence=sm.cell(r, 8).value)
          for r in range(5, 5 + N_SKILLS)]

people = []
for r in range(6, 16):
    if not rt.cell(r, 1).value:
        continue
    people.append(dict(id=f"Employee {len(people) + 1:02d}", role=rt.cell(r, 2).value,
                       before=[rt.cell(r, 3 + i).value for i in range(N_SKILLS)],
                       after=[ra.cell(r, 3 + i).value for i in range(N_SKILLS)]))

gap = lambda req, lvl: None if lvl is None else max(0, req - lvl)
for p in people:
    p["gb"] = [gap(s["req"], v) for s, v in zip(skills, p["before"])]
    p["ga"] = [gap(s["req"], v) for s, v in zip(skills, p["after"])]
    p["closed"] = [None if b is None or a is None else max(0, b - a) for b, a in zip(p["gb"], p["ga"])]

rows = []
for i, s in enumerate(skills):
    gb = [p["gb"][i] for p in people if p["gb"][i] is not None]
    rows.append(dict(name=s["name"], with_gap=sum(g > 0 for g in gb), before=sum(gb),
                     after=sum(p["ga"][i] or 0 for p in people),
                     closed=sum(p["closed"][i] or 0 for p in people)))
tot = {k: sum(r[k] for r in rows) for k in ("with_gap", "before", "after", "closed")}

# Self-check: our numbers must match the workbook's Summary sheet (B..E, rows 5-12).
for i, r in enumerate(rows):
    xl = [su.cell(5 + i, c).value for c in range(2, 6)]
    assert xl == [r["with_gap"], r["before"], r["after"], r["closed"]], (r["name"], xl, r)

evidence = [[ev.cell(r, c).value for c in range(1, 4)] for r in range(4, 20)]
limitation = ev.cell(20, 2).value

pct = lambda c, b: f"{int(c / b * 100 + 0.5)}%" if b else "–"  # round half up, like Excel
e = lambda v: escape(str(v))


def cell(g, lvl=None):
    if g is None:
        return '<td class="c">–</td>'
    label = f"{lvl}" if lvl is not None else ""
    return f'<td class="c g{min(g, 2)}" title="level {lvl}, gap {g}">{label}<small>gap {g}</small></td>'


skill_heads = "".join(f'<th class="c">{e(s["name"])}<br><small>req {s["req"]}</small></th>' for s in skills)


def heatmap(key_lvl, key_gap):
    body = "".join(
        f'<tr><td><b>{e(p["id"])}</b><br><small>{e(p["role"])}</small></td>'
        + "".join(cell(g, l) for g, l in zip(p[key_gap], p[key_lvl]))
        + f'<td class="c">{sum((g or 0) > 0 for g in p[key_gap])}</td></tr>'
        for p in people)
    return f'<div class="scroll"><table><tr><th>Employee</th>{skill_heads}<th class="c">Open gaps</th></tr>{body}</table></div>'


max_before = max(r["before"] for r in rows) or 1
summary_rows = "".join(
    f'<tr><td>{e(r["name"])}</td><td class="c">{r["with_gap"]}</td>'
    f'<td><div class="bar"><span class="b" style="width:{r["before"] / max_before * 100:.0f}%"></span>'
    f'<span class="a" style="width:{r["after"] / max_before * 100:.0f}%"></span></div>'
    f'<small>{r["before"]} → {r["after"]}</small></td>'
    f'<td class="c">{r["closed"]}</td><td class="c"><b>{pct(r["closed"], r["before"])}</b></td></tr>'
    for r in rows)

ranked = sorted(rows, key=lambda r: (-r["after"], r["name"]))
priorities = "".join(
    f'<li><b>{e(r["name"])}</b>: {r["after"]} levels still open after training '
    f'({pct(r["closed"], r["before"])} of the gap closed)</li>'
    for r in ranked if r["after"] > 0)

focus_rows = "".join(
    f'<tr><td><b>{e(p["id"])}</b><br><small>{e(p["role"])}</small></td><td>'
    + (", ".join(f'{e(skills[i]["name"])} (gap {g})'
                 for i, g in sorted(enumerate(p["ga"]), key=lambda x: -(x[1] or 0)) if g)
       or "No open gaps")
    + f'</td><td class="c">{sum(c or 0 for c in p["closed"])}</td></tr>'
    for p in people)

evidence_rows = "".join(f"<tr><td>{e(a)}</td><td>{e(b)}</td><td>{e(c)}</td></tr>" for a, b, c in evidence)

matrix_rows = "".join(
    f'<tr><td><b>{e(s["name"])}</b><br><small>{e(s["meaning"])}</small></td><td class="c">{s["req"]}</td>'
    + "".join(f"<td><small>{e(l)}</small></td>" for l in s["levels"]) + "</tr>"
    for s in skills)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FMCG Shop-Floor Skill Gaps</title>
<style>
  :root {{ --navy:#1F2F6B; --line:#d7dbe6; --muted:#5b6275; --bg:#fff; --fg:#222; --card:#f6f7fb; }}
  @media (prefers-color-scheme: dark) {{
    :root {{ --navy:#9fb3ff; --line:#33384a; --muted:#a3a9bb; --bg:#14161d; --fg:#e6e8ef; --card:#1c1f29; }}
  }}
  body {{ font-family: Georgia, "Liberation Serif", serif; color:var(--fg); background:var(--bg); margin:0; padding:20px 16px; max-width:1200px; margin-inline:auto; }}
  h1 {{ color:var(--navy); font-size:24px; margin:0 0 4px; }}
  h2 {{ color:var(--navy); font-size:17px; border-bottom:1px solid var(--navy); padding-bottom:3px; margin:28px 0 10px; }}
  .note {{ color:var(--muted); font-size:13px; margin:0 0 10px; }}
  .banner {{ background:var(--card); border-left:4px solid #e0a000; padding:8px 12px; font-size:13px; margin:12px 0; }}
  .kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:10px; }}
  .kpi {{ background:var(--card); border:1px solid var(--line); padding:10px 12px; }}
  .kpi b {{ display:block; font-size:26px; color:var(--navy); }}
  .kpi span {{ font-size:12px; color:var(--muted); }}
  table {{ border-collapse:collapse; width:100%; font-size:13px; }}
  th, td {{ border:1px solid var(--line); padding:5px 6px; vertical-align:top; text-align:left; }}
  th {{ background:#1F2F6B; color:#fff; font-weight:normal; }}
  .c {{ text-align:center; }}
  td small, th small {{ display:block; font-size:11px; color:var(--muted); }}
  th small {{ color:#cfd6f5; }}
  td.g0 {{ background:#c8e6c9; color:#222; }} td.g1 {{ background:#ffe082; color:#222; }} td.g2 {{ background:#ef9a9a; color:#222; }}
  td.g0 small, td.g1 small, td.g2 small {{ color:#444; }}
  .bar {{ position:relative; height:10px; background:var(--card); border:1px solid var(--line); margin:3px 0; }}
  .bar span {{ position:absolute; left:0; top:0; bottom:0; }}
  .bar .b {{ background:#ef9a9a; }} .bar .a {{ background:#1F2F6B; }}
  .scroll {{ overflow-x:auto; }}
  ul {{ font-size:14px; }}
  a {{ color:var(--navy); }}
</style>
</head>
<body>
<h1>FMCG shop-floor skill gaps: before and after training</h1>
<p class="note">{len(people)} employees across filling, packing, mixing, maintenance, quality, stores and utilities, rated on {N_SKILLS} shop-floor skills (levels 1–4) before and after training.</p>
<p class="banner"><b>Illustrative ratings.</b> Employee ratings and post-training scores are sample data, not measured assessments, and names are replaced with employee numbers. Operational evidence is taken from shop-floor process records. {e(limitation)}</p>

<div class="kpis">
  <div class="kpi"><b>{tot["with_gap"]}</b><span>open skill gaps before training ({len(people)} people × {N_SKILLS} skills)</span></div>
  <div class="kpi"><b>{tot["before"]} → {tot["after"]}</b><span>levels of gap, before → after</span></div>
  <div class="kpi"><b>{tot["closed"]}</b><span>levels of gap closed</span></div>
  <div class="kpi"><b>{pct(tot["closed"], tot["before"])}</b><span>of the total gap closed</span></div>
</div>

<h2>1. Results by skill</h2>
<p class="note">Bar: red = gap before, navy = gap still open after. Gap = required level − confirmed level, never below 0. Gap closed = gap before − gap after.</p>
<div class="scroll"><table>
<tr><th>Skill</th><th class="c">People with a gap</th><th>Gap before → after</th><th class="c">Gap closed</th><th class="c">% closed</th></tr>
{summary_rows}
<tr><td><b>Total</b></td><td class="c"><b>{tot["with_gap"]}</b></td><td><b>{tot["before"]} → {tot["after"]}</b></td><td class="c"><b>{tot["closed"]}</b></td><td class="c"><b>{pct(tot["closed"], tot["before"])}</b></td></tr>
</table></div>

<h2>2. Where to focus next</h2>
<p class="note">Skills ranked by the gap still open after training.</p>
<ul>{priorities}</ul>

<h2>3. Next focus per employee</h2>
<p class="note">Only the skills each person still has a gap in, largest first. This is the starting point for a personal study plan instead of training everyone on everything.</p>
<div class="scroll"><table><tr><th>Employee</th><th>Still to build</th><th class="c">Levels closed</th></tr>{focus_rows}</table></div>

<h2>4. Skill levels before training</h2>
<p class="note">Number = confirmed level. Green = no gap, amber = 1 level short, red = 2 or more short.</p>
{heatmap("before", "gb")}

<h2>5. Skill levels after training</h2>
{heatmap("after", "ga")}

<h2>6. Operational evidence behind the skill choice</h2>
<p class="note">Process observations from shop-floor records, used to choose the skills and priorities. These are not individual employee ratings.</p>
<div class="scroll"><table><tr><th>Area / metric</th><th>Value</th><th>What it means for the tracker</th></tr>{evidence_rows}</table></div>

<h2>7. Skill matrix</h2>
<div class="scroll"><table><tr><th>Skill</th><th class="c">Required</th><th>Level 1</th><th>Level 2</th><th>Level 3</th><th>Level 4</th></tr>{matrix_rows}</table></div>

<p class="note" style="margin-top:24px">Generated by <code>build/build_fmcg_site.py</code>. Interactive tracker: <a href="index.html">index.html</a>.</p>
</body>
</html>
"""
OUT.write_text(html, encoding="utf-8")
print(f"saved {OUT}")
print(f"people={len(people)} open gaps={tot['with_gap']} gap {tot['before']} -> {tot['after']} closed={tot['closed']} ({pct(tot['closed'], tot['before'])})")
for r in ranked:
    print(f"  {r['name']}: {r['with_gap']} people, {r['before']} -> {r['after']}, closed {r['closed']} ({pct(r['closed'], r['before'])})")
