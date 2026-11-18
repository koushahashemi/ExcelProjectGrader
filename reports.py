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


def feedback(result):
    if "error" in result:
        return "The workbook could not be graded. Submit a readable .xlsx workbook for this assignment."
    lines = [f"{result['filename']} — {result['total']}/{MAXIMUM} ({result['level']})"]
    for item in result["items"]:
        if item["score"] < item["maximum"]:
            lines.append(f"• {item['name']} ({item['score']}/{item['maximum']}) · {item['location']}\n  {item['feedback']}")
    if len(lines) == 1:
        lines.append("Full marks on the scored rubric items.")
    lines.extend(f"• Assignment advisory · {note}" for note in result["advisories"])
    if any(item["review"] for item in result["items"]):
        lines.append("Some checks may need instructor verification.")
    return "\n\n".join(lines)


def details(result, index=None):
    if "error" in result:
        return f"{result['source_path']}\n{result['error']}"
    items = result["items"] if index is None else [result["items"][index]]
    return "\n\n".join(
        f"{item['name']} — {item['score']}/{item['maximum']}\nLocation: {item['location']}\n"
        f"Expected: {item['expected']}\nDetected: {item['detected']}"
        + (f"\nInstructor review: {item['review']}" if item["review"] else "")
        for item in items
    )


def score_rows(results):
    for result in results:
        yield [result["filename"], result.get("total", ""), result["status"],
               *[item["score"] for item in result.get("items", [])],
               *([""] * len(RUBRIC) if "error" in result else [])]


def export_scores(path, results):
    with Path(path).open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["Filename", f"Total /{MAXIMUM}", "Status", *[f"{name} /{points}" for name, points in RUBRIC]])
        writer.writerows([csv_value(value) for value in row] for row in score_rows(results))


def html_report(results):
    totals = [result["total"] for result in results if "total" in result]
    average = f"{mean(totals):.1f}/{MAXIMUM}" if totals else "Unavailable"
    rows = []
    sections = []
    for index, (row, result) in enumerate(zip(score_rows(results), results)):
        rows.append(f'<tr><th scope="row"><a href="#submission-{index}">{escape(str(row[0]))}</a></th>'
                    + "".join(f"<td>{escape(str(cell))}</td>" for cell in row[1:]) + "</tr>")
        sections.append(f'<section id="submission-{index}"><h2>{escape(result["filename"])}</h2>'
                        f'<label for="feedback-{index}">Student feedback</label>'
                        f'<button type="button" data-feedback="feedback-{index}">Copy feedback</button>'
                        f'<textarea id="feedback-{index}" readonly rows="7">{escape(feedback(result))}</textarea>'
                        f'<details><summary>Instructor evidence</summary><pre>{escape(details(result))}</pre></details></section>')
    headings = "".join(f'<th scope="col">{escape(name)}</th>' for name in
                       ["Submission", "Total", "Status", *[f"{name} /{maximum}" for name, maximum in RUBRIC]])
    return f"""<!doctype html>
<html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>{TITLE} · Instructor report</title>
<style>
body {{ font: 16px system-ui; margin: 2rem auto; max-width: 1100px; padding: 0 1rem; color: #1f302b; }}
h1, h2, a {{ color: #006454; }} .scroll {{ overflow: auto; }}
table {{ border-collapse: collapse; }} td, th {{ padding: .45rem; border: 1px solid #b9c9c3; text-align:left; }}
section {{ border-top: 1px solid #bacac4; margin-top: 2rem; }} textarea {{ display:block; width: 98%; margin: .5rem 0; }}
pre {{ white-space: pre-wrap; overflow-wrap: anywhere; }} button {{ margin-left: 1rem; }}
</style>
<h1>{TITLE}</h1><p>{COURSE} · v{VERSION} · Instructor report</p>
<p>{len(results)} submissions · {len(totals)} graded · {len(results)-len(totals)} errors · Mean {average}</p>
<div class="scroll"><table><thead><tr>{headings}</tr></thead><tbody>{''.join(rows)}</tbody></table></div>
{''.join(sections)}
<script>
for (const button of document.querySelectorAll('button[data-feedback]')) {{
  button.onclick = async () => {{
    const text = document.getElementById(button.dataset.feedback);
    text.select();
    try {{ await navigator.clipboard.writeText(text.value); button.textContent = 'Copied'; }}
    catch {{ button.textContent = 'Press Ctrl+C to copy'; }}
  }};
}}
</script></html>"""


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
