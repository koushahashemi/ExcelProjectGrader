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
        writer.writerow(["Submission", *[name for name, _ in RUBRIC], "Total", "Maximum", "Status", "Instructor flags"])
        for result in results:
            if "error" in result:
                row = [result["filename"], *([""] * len(RUBRIC)), "", "", f"Error: {result['error']}"]
            else:
                row = [result["filename"], *[item["score"] for item in result["items"]],
                       result["total"], MAXIMUM, result["status"]]
            writer.writerow([csv_value(value) for value in [*row, flag_labels(result)]])


def student_report(result):
    if "error" in result:
        return f"{result['filename']}\n{feedback(result)}"
    lines = [TITLE, COURSE, f"Feedback on lost marks — {result['filename']}", f"Total: {result['total']}/{MAXIMUM}", ""]
    deductions = [item for item in result["items"] if item["score"] < item["maximum"]]
    if not deductions:
        lines.append("No marks lost in this submission.")
    for item in deductions:
        lost = item["maximum"] - item["score"]
        unit = "mark" if lost == 1 else "marks"
        lines.extend([f"{item['name']}: {item['score']}/{item['maximum']} ({lost} {unit} lost)",
                      f"• {item['feedback']} Location: {item['location']}", ""])
    lines.extend(f"• Assignment advisory · {note}" for note in result["advisories"])
    return "\n".join(lines).rstrip()


def instructor_report(result):
    if "error" in result:
        return "\n".join(filter(None, [result["filename"], result["error"], flag_details(result)]))
    estimated = " (estimated)" if result["status"] == "Review suggested" else ""
    lines = [TITLE, COURSE, f"{result['filename']}: {result['total']}/{MAXIMUM}{estimated}",
             f"Engine {result['version']}", f"Scoring: Automatic · {result['level']}", ""]
    for item in result["items"]:
        outcome = ("FAIL" if item["score"] < item["maximum"] else
                   "UNVERIFIED" if item["review"] else "PASS")
        lines.extend([f"{item['name']}: {item['score']}/{item['maximum']}",
                      f"  [{outcome}] {item['name']} — {item['location']}",
                      f"    Expected: {item['expected']}",
                      "    Found: " + item["detected"].replace("\n", "\n    ")])
        if item["review"]:
            lines.append(f"    Instructor review: {item['review']}")
        lines.append("")
    if result["advisories"]:
        lines.extend(["Assignment advisories (no separate rubric points)", *result["advisories"], ""])
    if flags := flag_details(result):
        lines.extend(["Instructor flags (no separate rubric points)", flags])
    return "\n".join(lines).rstrip()


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
                        f'<textarea id="feedback-{index}" readonly rows="7">{escape(student_report(result))}</textarea>'
                        f'<details><summary>Instructor evidence</summary><pre>{escape(instructor_report(result))}</pre></details></section>')
    headings = "".join(f'<th scope="col">{escape(name)}</th>' for name in
                       ["Submission", "Total", "Status", "Instructor flags", *[f"{name} /{maximum}" for name, maximum in RUBRIC]])
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
        content = "\n\n".join(student_report(result) for result in results)
    elif kind == "html":
        content = html_report(results)
    elif kind == "json":
        content = json.dumps([{"maximum": MAXIMUM, **result} for result in results], ensure_ascii=False, indent=2)
    else:
        content = "\n\n".join(instructor_report(result) for result in results)
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
