"""Batch exports and instructor benchmark comparisons."""
import csv
from html import escape
import json
from pathlib import Path
from statistics import mean

from grader import COURSE, MAXIMUM, RUBRIC, TITLE, VERSION


def csv_value(value):
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@", "\t", "\r")):
        return "'" + value
    return value


def item_feedback(item):
    return f"• {item['name']} ({item['score']}/{item['maximum']}) · {item['location']}\n  {item['feedback']}"


def flag_labels(result):
    labels = []
    if result.get("marker"):
        labels.append(result["marker"])
    if result.get("similarity"):
        labels.append("Similar submission")
    if result.get("comparison_error"):
        labels.append("Comparison unavailable")
    return ", ".join(labels)


def flag_details(result):
    lines = []
    if result.get("marker"):
        lines.append(f"Assignment marker found: {result['marker']}.")
    if result.get("comparison_error"):
        lines.append(f"Comparison unavailable: {result['comparison_error']}")
    for match in result.get("similarity", []):
        peer = match["peer"]
        lines.append(f"Compared with {peer}")
        lines.append(f"Matching transaction rows: {match['matching_rows']}/{match['row_count']} "
                     f"({match['overlap']:.1f}%) and {match['matching_rows']}/{match['peer_row_count']} "
                     f"({match['peer_overlap']:.1f}%).")
        if match["row_pairs"]:
            lines.append(f"Matched rows in {match['sheet']} ↔ {match['peer_sheet']}: " +
                         ", ".join(f"{a} ↔ {b}" for a, b in match["row_pairs"]))
    if lines:
        lines.append("Instructor review only. These signals carry zero points and do not establish authorship or AI use.")
    return "\n".join(lines)


def feedback(result):
    if "error" in result:
        return "The workbook could not be graded. Submit a readable .xlsx or .xlsm workbook for this assignment."
    lines = [f"{result['filename']} — {result['total']}/{MAXIMUM} ({result['level']})"]
    for item in result["items"]:
        if item["score"] < item["maximum"]:
            lines.append(item_feedback(item))
    if len(lines) == 1:
        lines.append("Full marks on the scored rubric items.")
    lines.extend(f"• Assignment advisory · {note}" for note in result["advisories"])
    if any(item["review"] for item in result["items"]):
        lines.append("Some checks may need instructor verification.")
    return "\n\n".join(lines)


def details(result, index=None):
    flags = flag_details(result) if index is None else ""
    if "error" in result:
        return f"{result['source_path']}\n{result['error']}\n{flags}"
    items = result["items"] if index is None else [result["items"][index]]
    return (flags + "\n\n" if flags else "") + "\n\n".join(
        f"{item['name']} — {item['score']}/{item['maximum']}\nLocation: {item['location']}\n"
        f"Expected: {item['expected']}\nDetected: {item['detected']}"
        + (f"\nInstructor review: {item['review']}" if item["review"] else "")
        for item in items
    )


def score_rows(results):
    for result in results:
        yield [result["filename"], result.get("total", ""), result["status"], flag_labels(result),
               *[item["score"] for item in result.get("items", [])],
               *([""] * len(RUBRIC) if "error" in result else [])]


def export_scores(path, results):
    with Path(path).open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["Filename", f"Total /{MAXIMUM}", "Status", "Instructor flags", *[f"{name} /{points}" for name, points in RUBRIC]])
        writer.writerows([csv_value(value) for value in row] for row in score_rows(results))


WORD_REPORT_CSS = r'''
:root{--ink:#1d2327;--muted:#5f6b72;--line:#dde3e6;--bg:#f6f8f9;--card:#fff;--teal:#00796B;
--full:#e3f3ee;--full-ink:#0b5d4b;--mid:#fff3d6;--mid-ink:#7a5300;--low:#fde4e1;--low-ink:#9b1c12;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif}
header{background:var(--teal);color:#fff;padding:20px 28px}
header h1{margin:0;font-size:22px} header p{margin:2px 0 0;opacity:.9}
main{max-width:1100px;margin:0 auto;padding:20px 16px 60px}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:20px 22px;margin:0 0 20px}
h2{margin:0 0 4px;font-size:19px} h3{margin:18px 0 8px;font-size:15px;color:var(--teal)}
.stats{display:flex;gap:12px;margin:10px 0 14px;flex-wrap:wrap}
.stats div{background:var(--bg);border-radius:8px;padding:8px 14px;min-width:120px}
.stats b{display:block;font-size:20px} .stats span{color:var(--muted);font-size:12px}
.tablewrap{overflow-x:auto}
table{border-collapse:collapse;width:100%}
.grid th,.grid td{border-bottom:1px solid var(--line);padding:6px 6px;text-align:center;white-space:nowrap}
.grid th{font-size:12px;color:var(--muted);font-weight:600}
.grid td.name{text-align:left;font-weight:600} .grid td.total{font-weight:700}
.sc{display:inline-block;min-width:26px;border-radius:5px;padding:1px 6px;font-weight:600}
td.sc{display:table-cell}
.full{background:var(--full);color:var(--full-ink)} .mid{background:var(--mid);color:var(--mid-ink)} .low{background:var(--low);color:var(--low-ink)}
.legend{color:var(--muted);font-size:12px;margin:10px 0 0} .legend .sc{margin:0 4px 0 10px}
.badge{border-radius:10px;padding:1px 9px;font-size:12px;font-weight:600}
.flag-yes{background:#ffe9a8;color:#6b4a00} .flag-no{background:#eef1f2;color:var(--muted)} .flag-check{background:#e3ecfb;color:#1f4f9a}
a{color:var(--teal)}
.s-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;border-bottom:1px solid var(--line);padding-bottom:10px}
.file{color:var(--muted);font-size:12px;word-break:break-all}
.s-total{font-size:14px;white-space:nowrap} .s-total b{font-size:30px;color:var(--teal)}
.est{font-size:11px;color:var(--muted);border:1px solid var(--line);border-radius:8px;padding:0 6px;margin-left:4px}
.sentence{border-radius:8px;padding:10px 14px;margin:14px 0 4px;border:1px solid var(--line)}
.sentence.flag-yes{background:#fff6d6;border-color:#f0d47a;color:var(--ink)}
.sentence.flag-no{background:var(--bg);color:var(--ink)} .sentence.flag-check{background:#eef4fd;color:var(--ink)}
.sentence p{margin:4px 0} .sentence .note{float:right;font-size:11px;color:var(--muted)}
blockquote{margin:6px 0 0;padding:6px 12px;border-left:3px solid #e0b93a;background:#fff;font-style:italic}
.s-bar{display:flex;align-items:center;gap:10px} .s-bar h3{flex:1}
.copy{border:1px solid var(--teal);background:#fff;color:var(--teal);border-radius:6px;padding:4px 10px;cursor:pointer;font:inherit;font-size:12px}
.top{font-size:12px;white-space:nowrap}
.crit{border-left:4px solid;border-radius:6px;padding:8px 12px;margin:0 0 10px;background:var(--card)}
.crit.mid{border-color:#e0a800;background:#fffbf0;color:var(--ink)} .crit.low{border-color:#d0453a;background:#fff7f6;color:var(--ink)}
.crit-head{display:flex;gap:12px;align-items:baseline;flex-wrap:wrap}
.crit-name{font-weight:700;flex:1} .crit-score{font-weight:700} .crit-lost{color:var(--low-ink);font-size:12px;font-weight:600}
.ref{color:var(--muted);font-size:12px} .band{margin-left:8px;font-style:italic}
.crit ul{margin:6px 0 2px;padding-left:20px} .crit li{margin:3px 0}
.allfull{background:var(--full);color:var(--full-ink);padding:8px 12px;border-radius:6px}
.fulllist{color:var(--muted);font-size:13px}
details{border:1px solid var(--line);border-radius:6px;margin:0 0 6px}
summary{cursor:pointer;padding:6px 10px;font-weight:600} .muted{color:var(--muted);font-weight:400}
.ev{font-size:12px} .ev th,.ev td{border-top:1px solid var(--line);padding:5px 8px;text-align:left;vertical-align:top}
.ev th{color:var(--muted)} .loc{color:var(--muted);font-size:11px}
.st{border-radius:4px;padding:0 6px;font-weight:600;white-space:nowrap}
.st.ok{background:var(--full);color:var(--full-ink)} .st.bad{background:var(--low);color:var(--low-ink)}
.st.unk{background:#e3ecfb;color:#1f4f9a} .st.info{background:#eef1f2;color:var(--muted)}
.err{color:var(--low-ink)}
.limits{color:var(--muted);font-size:12px}
@media print{body{background:#fff} header{color:#000;background:#fff;border-bottom:2px solid var(--teal)}
section{border:none;padding:0} .student{page-break-before:always} .copy,.top{display:none}
details{border:none} details>summary{list-style:none} .evidence{display:none}}
@media (max-width:640px){header{padding:16px} section{padding:14px} .s-head{flex-direction:column}}
'''

def score_color(score, maximum):
    return "full" if score == maximum else "mid" if score else "low"


def html_submission(index, result):
    filename = escape(result["filename"])
    status = escape(result["status"])
    score = result.get("total")
    marker = bool(result.get("marker"))
    flag = "yes" if marker else "no"
    answer = "Yes" if marker else "No"
    summary = (f'<div class="s-head"><div><h2>{filename}</h2><div class="file">{status}</div></div>'
               f'<div class="s-total"><b>{score if score is not None else "—"}</b>/{MAXIMUM}</div></div>'
               f'<div class="sentence flag-{flag}"><div class="sentence-head">Hidden-instruction sentence: '
               f'<b>{answer}</b><span class="note">Instructor only · 0 rubric points</span></div></div>')
    feedback_text = feedback(result)
    if "error" in result:
        marks = f'<p class="err">{escape(feedback_text)}</p>'
        evidence = f'<details><summary>Grading error</summary><p class="err">{escape(result["error"])}</p></details>'
    else:
        deductions, full, evidence_rows = [], [], []
        for item in result["items"]:
            name, tally = escape(item["name"]), f'{item["score"]}/{item["maximum"]}'
            color = score_color(item["score"], item["maximum"])
            lost = item["maximum"] - item["score"]
            location = escape(item["location"])
            if lost:
                unit = "mark" if lost == 1 else "marks"
                deductions.append(f'<div class="crit {color}"><div class="crit-head"><span class="crit-name">{name}</span>'
                                  f'<span class="crit-score">{tally}</span><span class="crit-lost">−{lost} {unit}</span></div>'
                                  f'<div class="ref">{location}</div><ul><li>{escape(item["feedback"])}</li></ul></div>')
            else:
                full.append(name)
            expected = escape(item["expected"])
            detected = escape(item["detected"]).replace("\n", "<br>")
            review = f'<p class="limits">{escape(item["review"])}</p>' if item["review"] else ""
            evidence_rows.append(f'<details><summary>{name} <span class="muted">· {tally}</span></summary>'
                                 '<div class="tablewrap"><table class="ev"><thead><tr>'
                                 '<th>Location</th><th>Expected</th><th>Found</th></tr></thead>'
                                 f'<tbody><tr><td class="loc">{location}</td><td>{expected}</td><td>{detected}</td>'
                                 f'</tr></tbody></table></div>{review}</details>')
        marks = "".join(deductions) or '<p class="allfull">No marks lost.</p>'
        if full:
            marks += '<p class="fulllist">Full marks: ' + ", ".join(full) + "</p>"
        if result["advisories"]:
            marks += '<ul class="limits">' + "".join(f"<li>{escape(note)}</li>" for note in result["advisories"]) + "</ul>"
        evidence = "".join(evidence_rows)
    if flags := flag_details(result):
        evidence += '<details><summary>Instructor flags</summary><p class="limits">' + escape(flags).replace("\n", "<br>") + "</p></details>"
    return (f'<section class="student" id="s{index}">{summary}<div class="s-bar"><h3>Marks lost</h3>'
            '<button class="copy" type="button">Copy feedback</button><a class="top" href="#top">↑ Summary</a></div>'
            f'<div class="feedback">{marks}</div><pre class="copy-text" hidden>{escape(feedback_text)}</pre>'
            f'<div class="evidence"><h3>Technical evidence</h3>{evidence}</div></section>')


def html_report(results):
    totals = [result["total"] for result in results if "total" in result]
    average = f"{mean(totals):.1f}" if totals else "—"
    marked = sum(bool(result.get("marker")) for result in results)
    headings = "".join(f'<th title="{escape(name)}">{escape(name)}</th>' for name, _ in RUBRIC)
    rows, sections = [], []
    for index, result in enumerate(results):
        cells = "".join(f'<td class="sc {score_color(item["score"], item["maximum"])}">{item["score"]}</td>'
                        for item in result.get("items", []))
        if "error" in result:
            cells = f'<td class="err" colspan="{len(RUBRIC)}">Grading error</td>'
        flag, answer = ("yes", "Yes") if result.get("marker") else ("no", "No")
        rows.append(f'<tr><td class="name"><a href="#s{index}">{escape(result["filename"])}</a></td>'
                    f'<td class="total">{result.get("total", "—")}</td>{cells}'
                    f'<td><span class="badge flag-{flag}">{answer}</span></td></tr>')
        sections.append(html_submission(index, result))
    script = """
for (const button of document.querySelectorAll('.copy')) {
  button.addEventListener('click', async () => {
    const text = button.closest('section').querySelector('.copy-text').textContent;
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const field = document.createElement('textarea');
      field.value = text;
      field.style.position = 'fixed';
      field.style.opacity = '0';
      document.body.append(field);
      field.select();
      const copied = document.execCommand('copy');
      field.remove();
      if (!copied) return;
    }
    button.textContent = 'Copied';
    setTimeout(() => button.textContent = 'Copy feedback', 1500);
  });
}
"""
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Excel Project Grading Report</title><style>{WORD_REPORT_CSS}</style></head><body>
<header id="top"><h1>BCIT Excel Project Grader — Grading Report</h1>
<p>{COURSE} · Instructor Grading Tool · Engine {VERSION}</p></header>
<main><section class="summary"><h2>Class summary</h2><div class="stats">
<div><b>{len(results)}</b><span>submissions</span></div><div><b>{average}</b><span>average / {MAXIMUM}</span></div>
<div><b>{marked}</b><span>added sentence flagged</span></div></div>
<div class="tablewrap"><table class="grid"><thead><tr><th>Student</th><th>Total</th>{headings}<th>Added sentence</th>
</tr></thead><tbody>{''.join(rows)}</tbody></table></div>
<p class="legend"><span class="sc full">Full</span> full marks <span class="sc mid">Partial</span> some issues
<span class="sc low">0</span> major issues. Click a name to jump to that student.</p></section>
{''.join(sections)}</main><script>{script}</script></body></html>"""


def export_report(path, results, kind):
    path = Path(path)
    if kind == "csv":
        export_scores(path, results)
        return
    if kind == "feedback":
        content = "\n\n" + ("\n\n" + "─" * 50 + "\n\n").join(feedback(result) for result in results)
        content = content.strip() + "\n"
    elif kind == "html":
        content = html_report(results)
    elif kind == "json":
        content = json.dumps({"course": COURSE, "version": VERSION, "maximum": MAXIMUM,
                              "submissions": results}, ensure_ascii=False, indent=2)
    else:
        content = f"{TITLE} · {COURSE} · v{VERSION}\n\n" + "\n\n".join(
            f"{feedback(result)}\n\nINSTRUCTOR EVIDENCE\n{details(result)}" for result in results)
    path.write_text(content, encoding="utf-8")


def benchmark(manifest, grade):
    """Compare automatic results with instructor-entered item scores and keep the evidence."""
    manifest = Path(manifest)
    entries = json.loads(manifest.read_text(encoding="utf-8"))
    output = []
    maxima = dict(RUBRIC)
    for entry in entries:
        human = entry["scores"]
        if human.keys() != maxima.keys() or any(type(score) is not int or not 0 <= score <= maxima[name]
                                                for name, score in human.items()):
            raise ValueError("Every benchmark must provide a valid instructor score for each rubric item.")
        result = grade(manifest.parent / entry["workbook"])
        deltas = [{"criterion": item["name"], "instructor": human[item["name"]], "automatic": item["score"]}
                  for item in result.get("items", []) if human[item["name"]] != item["score"]]
        output.append({"instructor_total": sum(human.values()), "notes": entry.get("notes", ""),
                       "differences": deltas, "result": result})
    return output
