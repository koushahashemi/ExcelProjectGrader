"""Review aggregated Excel grader exports from one or more grading sets."""
from dataclasses import dataclass
import csv
import json
import math
from pathlib import Path
import re
from statistics import mean, median
import tkinter as tk
from tkinter import filedialog, ttk

from dashboard_charts import BG, FULL, INK, LOW, MID, MUTED, TEAL, histogram
from dashboard_page import DashboardPage
from grader import FORMAT, FORMULAS, GROUPS, MAXIMUM, RUBRIC, SETUP

CATEGORIES = {
    "Worksheet Setup": SETUP,
    "Worksheet Formatting": FORMAT,
    "Formulae": FORMULAS,
    **{f"{group} Chart": tuple(name for name, _ in RUBRIC if name.startswith(f"{group} · ")) for group in GROUPS},
}
POINTS = dict(RUBRIC)
ALIASES = {"DetailTable": "Excel Table Creation"}


@dataclass
class Submission:
    set_name: str
    filename: str
    total: float | None
    maximum: float
    scores: dict[str, float]

    def percent(self, items):
        if len(items) == len(POINTS) and self.total is not None:
            return 100 * self.total / self.maximum
        if items and self.scores.keys() >= set(items):
            return 100 * sum(self.scores[name] for name in items) / sum(POINTS[name] for name in items)
        return None


def score(value, maximum):
    try:
        number = float(value)
    except (ValueError, TypeError, OverflowError):
        return None
    return number if math.isfinite(number) and 0 <= number <= maximum else None


def unsanitize(value):
    return value[1:] if value.startswith("'") and value[1:].lstrip().startswith(("=", "+", "-", "@")) else value


def read_report(path):
    """Load current and legacy CSV/JSON grade exports."""
    path = Path(path)
    fallback = path.stem
    submissions = []
    if path.suffix.lower() == ".json":
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(payload, dict):
            rows = payload.get("submissions")
            fallback = payload.get("set_name") or fallback
            default_maximum = score(payload.get("maximum"), math.inf) or MAXIMUM
        else:
            rows, default_maximum = payload, MAXIMUM
        if not isinstance(rows, list) or any(not isinstance(row, dict) or not row.get("filename") for row in rows):
            raise ValueError("Expected a grader report with a submissions list.")
        for row in rows:
            values = {}
            for item in row.get("items", []):
                name = ALIASES.get(item["name"], item["name"])
                if name in POINTS and (value := score(item.get("score"), POINTS[name])) is not None:
                    values[name] = value
            maximum = score(row.get("maximum"), math.inf) or default_maximum
            total = None if row.get("error") else score(row.get("total"), maximum)
            submissions.append(Submission(str(row.get("set_name") or fallback), str(row["filename"]), total,
                                          maximum, values if not row.get("error") else {}))
        return submissions

    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        headers = {head.strip().casefold(): head for head in reader.fieldnames or []}
        filename = headers.get("filename") or headers.get("submission")
        total_header = next((head for key, head in headers.items() if re.fullmatch(r"total(?:\s*/\s*\d+)?", key)), None)
        rubrics = {}
        for head in reader.fieldnames or []:
            name = re.sub(r"\s*/\s*\d+$", "", head.strip())
            name = ALIASES.get(name, name)
            if name in POINTS:
                rubrics[name] = head
        if not filename or not (total_header or rubrics):
            raise ValueError("Expected grade CSV columns for submission and total or rubric scores.")
        for row in reader:
            if not row.get(filename):
                continue
            default = re.search(r"/\s*(\d+)$", total_header or "")
            maximum = score(row.get(headers.get("maximum")), math.inf) or (int(default[1]) if default else MAXIMUM) or MAXIMUM
            failed = str(row.get(headers.get("status")) or "").casefold().startswith("error")
            values = {name: value for name, head in rubrics.items()
                      if (value := score(row.get(head), POINTS[name])) is not None}
            total = None if failed else score(row.get(total_header), maximum)
            set_name = unsanitize(row.get(headers.get("set name")) or "") or fallback
            submissions.append(Submission(set_name, unsanitize(row[filename]), total, maximum,
                                          values if not failed else {}))
    return submissions


def rubric_stats(submissions, names):
    rows = []
    for name in names:
        values = [entry.scores[name] for entry in submissions if name in entry.scores]
        if not values:
            continue
        maximum = POINTS[name]
        full = values.count(maximum)
        zero = values.count(0)
        rows.append((name, maximum, len(values), mean(values), min(values), max(values), full,
                     len(values) - full - zero, zero, 100 * (len(values) - full) / len(values),
                     sum(maximum - value for value in values)))
    return sorted(rows, key=lambda row: (-row[9], -row[10], row[0]))


def dashboard(parent=None, paths=()):
    window = tk.Toplevel(parent) if parent is not None else tk.Tk()
    window.title("Excel Project Grader · Results Dashboard")
    window.geometry("1380x820")
    window.minsize(880, 620)
    window.configure(background=BG)
    banner = tk.Canvas(window, height=91, background=TEAL, highlightthickness=0)
    banner.pack(fill="x")

    def paint_banner(event):
        banner.delete("all")
        width = event.width
        start, end = (6, 65, 64), (0, 121, 107)
        for index in range(64):
            color = "#" + "".join(f"{round(a + (b-a)*index/63):02x}" for a, b in zip(start, end))
            banner.create_rectangle(width * index / 64, 0, width * (index + 1) / 64 + 1, 91,
                                    fill=color, outline="")
        for offset in (0, 38, 76):
            banner.create_oval(width - 245 + offset, -80, width + 20 + offset, 175, outline="#378d7c")
        banner.create_text(24, 35, text="Results Dashboard", fill="white", anchor="w",
                           font=("Segoe UI", 22, "bold"))
        banner.create_text(26, 66, text="Excel Project Grader", fill="#c3e9de", anchor="w",
                           font=("Segoe UI", 10))

    banner.bind("<Configure>", paint_banner)

    filters = tk.Frame(window, background=BG, padx=16, pady=10)
    filters.pack(fill="x")
    loaded = {}
    status = tk.StringVar(window, value="Load Excel grader JSON or CSV exports to compare results.")
    set_filter, category_filter, item_filter = (tk.StringVar(window, value="All") for _ in range(3))
    ttk.Button(filters, text="Load reports…", command=lambda: choose()).pack(side="left", padx=(0, 5))
    ttk.Button(filters, text="Clear", command=lambda: clear()).pack(side="left", padx=(0, 10))
    controls = []
    for text, variable, size in (("Set", set_filter, 21), ("Category", category_filter, 19),
                                 ("Rubric item", item_filter, 25)):
        tk.Label(filters, text=text, background=BG, foreground=INK).pack(side="left", padx=(5, 3))
        control = ttk.Combobox(filters, textvariable=variable, state="readonly", width=size)
        control.pack(side="left")
        control.bind("<<ComboboxSelected>>", lambda event: refresh())
        controls.append(control)
    tk.Label(window, textvariable=status, anchor="w", background=BG, foreground=MUTED,
             padx=16, wraplength=1000, justify="left").pack(fill="x")

    zoom_controls = ttk.Frame(window, padding=(16, 0, 16, 4))
    zoom_controls.pack(fill="x")
    page = DashboardPage(window, zoom_controls)

    def percentage(value):
        return "—" if value is None else f"{value:.1f}%"

    def summary(values):
        return [percentage(fn(values)) for fn in (min, max, mean, median)] if values else ["—"] * 4

    def refresh():
        entries = [entry for batch in loaded.values() for entry in batch]
        # A named CSV and JSON for the same student can be loaded together. Prefer
        # the copy with more usable rubric scores and show the duplicate count.
        unique = {}
        for entry in entries:
            identity = (entry.set_name.casefold(), entry.filename.casefold())
            previous = unique.get(identity)
            if previous is None or (len(entry.scores), entry.total is not None) > (len(previous.scores), previous.total is not None):
                unique[identity] = entry
        controls[0]["values"] = ["All"] + sorted({entry.set_name for entry in unique.values()}, key=str.casefold)
        controls[1]["values"] = ["All", *CATEGORIES]
        available = CATEGORIES.get(category_filter.get(), tuple(POINTS))
        controls[2]["values"] = ["All", *available]
        if item_filter.get() not in controls[2]["values"]:
            item_filter.set("All")
        selected = (item_filter.get(),) if item_filter.get() != "All" else available
        shown = [entry for entry in unique.values() if set_filter.get() == "All" or entry.set_name == set_filter.get()]
        scored = [(entry, entry.percent(selected)) for entry in shown]
        values = [value for _, value in scored if value is not None]
        metrics = [f"{len(values)} / {len(shown)}", *summary(values)]
        stats = rubric_stats(shown, selected)
        tables = {title: [] for title in ("Rubric statistics", "Set comparison", "Score distribution", "Submissions")}
        for name, maximum, count, avg, low, high, full, partial, zero, below, lost in stats:
            tint = FULL if below == 0 else LOW if below >= 50 else MID
            tables["Rubric statistics"].append(((name, maximum, count, f"{avg:.2f}", f"{low:g}", f"{high:g}",
                                                  full, partial, zero, percentage(below), f"{lost:g}"), tint))
        if stats and stats[0][9]:
            insight = f"Most frequent deduction · {stats[0][0]} ({stats[0][9]:.1f}% below full marks)"
        else:
            insight = ("No deductions in the selected rubric results." if stats
                       else "No rubric scores for the current filters.")
        if duplicate_count := len(entries) - len(unique):
            insight += f" · {duplicate_count} duplicate record(s) collapsed"
        for name in sorted({entry.set_name for entry in shown}, key=str.casefold):
            batch = [value for entry, value in scored if entry.set_name == name]
            valid = [value for value in batch if value is not None]
            tables["Set comparison"].append(((name, len(batch), len(valid), *summary(valid)), None))
        for index, count in enumerate(histogram(values)):
            lower, upper = 10 * index, 10 * (index + 1)
            tint = LOW if index < 5 else MID if index < 8 else FULL
            label = f"{lower}% to under {upper}%" if index < 9 else "90% to 100%"
            tables["Score distribution"].append(((label, count, percentage(100 * count / len(values) if values else None)), tint))
        for entry, value in sorted(scored, key=lambda pair: (pair[0].set_name.casefold(), pair[0].filename.casefold())):
            total = "—" if entry.total is None else f"{entry.total:g}/{entry.maximum:g}"
            tables["Submissions"].append(((entry.filename, entry.set_name, percentage(value), total), None))
        page.update(metrics, insight, scored, stats, selected, tables)

    def open_paths(paths):
        errors = []
        for value in paths:
            path = Path(value).resolve()
            try:
                loaded[path] = read_report(path)
            except (OSError, ValueError, TypeError, KeyError, csv.Error) as exc:
                errors.append(f"{path.name} ({exc})")
        status.set(f"{len(loaded)} report(s) loaded" + (f" · Could not load {', '.join(errors)}" if errors else ""))
        if set_filter.get() != "All" and not any(
                entry.set_name == set_filter.get() for batch in loaded.values() for entry in batch):
            set_filter.set("All")
        refresh()

    def choose():
        paths = filedialog.askopenfilenames(parent=window, title="Select grading reports",
                                            filetypes=[("Grading reports", ("*.json", "*.csv"))])
        if paths:
            open_paths(paths)

    def clear():
        loaded.clear()
        for variable in (set_filter, category_filter, item_filter):
            variable.set("All")
        status.set("Load Excel grader JSON or CSV exports to compare results.")
        refresh()

    refresh()
    if paths:
        open_paths(paths)
    return window


if __name__ == "__main__":
    import sys

    dashboard(paths=sys.argv[1:]).mainloop()
