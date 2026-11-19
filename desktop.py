"""Instructor desktop interface for Excel Project Grader."""
import argparse
from colorsys import hls_to_rgb
import json
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from batch_tools import compare_workbooks, similarity_groups, trim_filenames, workbook_file
from grader import COURSE, MAXIMUM, RUBRIC, TITLE, VERSION, grade
from reports import benchmark, details, export_report, feedback, flag_labels, item_feedback, score_rows

GROUP_COLORS = ("#d7ebff", "#dcf4dc", "#f1ddff", "#ffdcca", "#d1f4f2", "#ffd9e6",
                "#e8efc7", "#dce0ff", "#f5dfce", "#d8e9e1", "#e9dce9", "#d5e4ec")


def group_color(group):
    if group <= len(GROUP_COLORS):
        return GROUP_COLORS[group - 1]
    hue = ((group - len(GROUP_COLORS)) * 0.61803398875) % 1
    rgb = hls_to_rgb(hue, 0.83, 0.70)
    return "#" + "".join(f"{round(channel * 255):02x}" for channel in rgb)


EXCEL_FILES = [("Excel workbooks", ("*.xlsx", "*.xlsm"))]


class App:
    def __init__(self, root, drop_type):
        self.root = root
        self.queue = queue.Queue()
        self.results = []
        self.pending = set()
        self.remaining = 0
        self.comparing = False
        self.comparison_output = queue.Queue()
        root.title(f"{TITLE} {VERSION} · {COURSE}")
        root.geometry("1240x790")
        root.minsize(880, 580)
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), foreground="#00796B")
        style.configure("Treeview", rowheight=26)
        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"{COURSE} · v{VERSION}").pack(anchor="w", pady=(2, 12))

        actions = ttk.Frame(frame)
        actions.pack(fill="x")
        ttk.Button(actions, text="Add Excel files…", command=self.choose).pack(side="left")
        ttk.Button(actions, text="Add folder…", command=self.folder).pack(side="left", padx=5)
        self.clear_button = ttk.Button(actions, text="Clear batch", command=self.clear)
        self.clear_button.pack(side="left")
        ttk.Button(actions, text="Export grades CSV…", command=lambda: self.save("csv")).pack(side="right")
        ttk.Button(actions, text="Save feedback…", command=lambda: self.save("feedback")).pack(side="right", padx=5)
        ttk.Button(actions, text="Save full report…", command=lambda: self.save("report")).pack(side="right")

        utilities = ttk.Frame(frame)
        utilities.pack(fill="x", pady=(5, 0))
        self.compare_button = ttk.Button(utilities, text="Check for similar submissions…", command=self.compare)
        self.compare_button.pack(side="left")
        self.rename_button = ttk.Button(utilities, text="Trim filenames…", command=self.rename_dialog)
        self.rename_button.pack(side="left", padx=5)

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(10, 5))
        self.status = tk.StringVar(value="Drop .xlsx or .xlsm submissions here, or select Add Excel files. Double-click a filename to open it.")
        ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=(0, 8))
        panes = ttk.Panedwindow(frame, orient="vertical")
        panes.pack(fill="both", expand=True)
        upper = ttk.Frame(panes)
        panes.add(upper, weight=2)
        self.batch = self.table(upper, ["Submission", f"Total /{MAXIMUM}", "Status", "Instructor flags", *[f"{name} /{points}" for name, points in RUBRIC]],
                                [240, 85, 130, 210, *[145] * len(RUBRIC)])
        self.batch.tag_configure("flag", background="#fff0b3", foreground="#673500")
        self.batch.bind("<<TreeviewSelect>>", self.select_submission)
        self.batch.bind("<Double-1>", self.open_workbook)
        self.batch.bind("<Return>", self.open_workbook)
        lower = ttk.Frame(panes, padding=(0, 10, 0, 0))
        panes.add(lower, weight=3)
        self.selection = ttk.Label(lower, text="Select a submission to inspect its rubric scores.")
        self.selection.pack(anchor="w", pady=(0, 7))
        review = ttk.Panedwindow(lower, orient="horizontal")
        review.pack(fill="both", expand=True)
        criteria = ttk.Frame(review)
        review.add(criteria, weight=1)
        self.rubric = self.table(criteria, ["Rubric item", "Score", "Review"], [190, 60, 60])
        self.rubric.bind("<<TreeviewSelect>>", self.select_criterion)
        views = self.views = ttk.Notebook(review)
        review.add(views, weight=2)
        self.feedback = ScrolledText(views, wrap="word", padx=10, pady=8, state="disabled")
        self.details = ScrolledText(views, wrap="word", padx=10, pady=8, state="disabled")
        self.feedback.tag_configure("focused", background="#fff0b3")
        views.add(self.feedback, text="Feedback")
        views.add(self.details, text="Technical details")
        self.register_drop(root, drop_type)
        root.after(75, self.poll)

    def register_drop(self, widget, drop_type):
        widget.drop_target_register(drop_type)
        widget.dnd_bind("<<Drop>>", self.drop)
        for child in widget.winfo_children():
            if child.winfo_class() != "Menu":
                self.register_drop(child, drop_type)

    @staticmethod
    def table(parent, labels, widths):
        columns = [f"column{i}" for i in range(len(labels))]
        tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        for i, (column, label, width) in enumerate(zip(columns, labels, widths)):
            tree.heading(column, text=label)
            tree.column(column, width=width, stretch=False, anchor="w" if i == 0 else "center")
        tree.tag_configure("deduction", foreground="#a1241b")
        tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        return tree

    def choose(self):
        self.add(filedialog.askopenfilenames(filetypes=EXCEL_FILES, title="Select submissions"))

    def folder(self):
        directory = filedialog.askdirectory(title="Select a folder of submissions")
        if directory:
            self.add(sorted(Path(directory).iterdir()))

    def drop(self, event):
        try:
            paths = self.root.tk.splitlist(event.data)
        except tk.TclError:
            self.status.set("Could not read the dropped files.")
            return "refuse_drop"
        self.add(paths)
        return "copy"

    def add(self, paths):
        if self.comparing:
            self.status.set("Wait for the submission comparison to finish before adding files.")
            return
        candidates = dict.fromkeys(Path(value).resolve() for value in paths)
        workbooks = [path for path in candidates if workbook_file(path) and path not in self.pending
                     and str(path) not in {result["source_path"] for result in self.results}]
        if not workbooks:
            if candidates:
                self.status.set("No new .xlsx or .xlsm submissions found.")
            return
        if not self.remaining:
            self.progress.start()
        self.remaining += len(workbooks)
        self.pending.update(workbooks)
        self.clear_button.configure(state="disabled")
        self.status.set(f"Grading {self.remaining} workbook(s)…")
        output = self.queue

        def worker():
            for path in workbooks:
                output.put(grade(path))

        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                result = self.queue.get_nowait()
                index = len(self.results)
                self.results.append(result)
                self.pending.discard(Path(result["source_path"]))
                self.remaining -= 1
                self.batch.insert("", "end", iid=str(index), values=next(score_rows([result])),
                                  tags=self.result_tags(result))
                if not self.batch.selection():
                    self.batch.selection_set(str(index))
                if self.remaining:
                    self.status.set(f"Grading {self.remaining} workbook(s)…")
                else:
                    self.progress.stop()
                    self.clear_button.configure(state="normal")
                    errors = sum("error" in item for item in self.results)
                    self.status.set(f"Batch complete · {len(self.results)-errors} graded · {errors} errors")
        except queue.Empty:
            pass
        self.poll_comparison()
        self.root.after(75, self.poll)

    @staticmethod
    def result_tags(result, group=None):
        if group is not None:
            return (f"similarity-{group}",)
        if flag_labels(result):
            return ("flag",)
        return ("deduction",) if "error" in result else ()

    def refresh_rows(self):
        groups = similarity_groups(self.results)
        for group in sorted(set(groups.values())):
            self.batch.tag_configure(f"similarity-{group}", background=group_color(group), foreground="#202020")
        for index, result in enumerate(self.results):
            group = groups.get(result["source_path"])
            values = next(score_rows([result]))
            if group is not None:
                values[3] += f" · Group {group}"
            self.batch.item(str(index), values=values, tags=self.result_tags(result, group))
        self.select_submission()

    def compare(self):
        if self.remaining or self.comparing:
            self.status.set("Wait for the current batch operation to finish.")
            return
        if len(self.results) < 2:
            messagebox.showinfo("Compare submissions", "Add at least two workbooks to compare.")
            return
        paths = [result["source_path"] for result in self.results]
        self.comparing = True
        for button in (self.clear_button, self.compare_button, self.rename_button):
            button.configure(state="disabled")
        self.progress.start()
        self.status.set(f"Comparing {len(paths)} workbooks…")
        output = self.comparison_output

        def worker():
            try:
                output.put((compare_workbooks(paths), None))
            except Exception as error:
                output.put((None, str(error)))

        threading.Thread(target=worker, daemon=True).start()

    def poll_comparison(self):
        try:
            comparison, failure = self.comparison_output.get_nowait()
        except queue.Empty:
            return
        self.comparing = False
        for button in (self.clear_button, self.compare_button, self.rename_button):
            button.configure(state="normal")
        self.progress.stop()
        if failure:
            self.status.set("Comparison could not be completed.")
            messagebox.showerror("Comparison failed", failure)
            return
        flags, errors = comparison
        for result in self.results:
            path = result["source_path"]
            result["similarity"] = flags[path]
            result.pop("comparison_error", None)
            if path in errors:
                result["comparison_error"] = errors[path]
        self.refresh_rows()
        similarity_count = sum(bool(matches) for matches in flags.values())
        self.status.set(f"Comparison complete · {similarity_count} workbooks with matching transaction rows · {len(errors)} errors")

    def rename_dialog(self):
        if self.remaining or self.comparing:
            self.status.set("Wait for the current batch operation to finish.")
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("Trim Excel filenames")
        dialog.transient(self.root)
        panel = ttk.Frame(dialog, padding=16)
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text="Remove the prefix through the first underscore from .xlsx and .xlsm filenames.",
                  wraplength=470).pack(anchor="w")
        output = ScrolledText(panel, width=66, height=12, wrap="word", state="disabled")
        output.pack(fill="both", expand=True, pady=(10, 0), side="bottom")
        buttons = ttk.Frame(panel)
        buttons.pack(fill="x", pady=(8, 0))

        def apply(paths):
            paths = [path for path in paths if workbook_file(Path(path))]
            if not paths:
                return
            if not messagebox.askyesno("Rename workbooks", f"Apply filename trimming to {len(paths)} workbook(s)?",
                                       parent=dialog):
                return
            renamed, skipped = trim_filenames(paths)
            if renamed:
                for result in self.results:
                    old = result["source_path"]
                    if old in renamed:
                        result["source_path"] = renamed[old]
                        result["filename"] = Path(renamed[old]).name
                    for match in result.get("similarity", []):
                        match["peer"] = renamed.get(match["peer"], match["peer"])
                self.refresh_rows()
            lines = [f"Renamed {len(renamed)} · Skipped {len(skipped)}"]
            lines.extend(f"{Path(before).name} → {Path(after).name}" for before, after in renamed.items())
            self.show(output, "\n".join(lines + skipped))

        def choose_folder():
            folder = filedialog.askdirectory(parent=dialog)
            if folder:
                apply(sorted(Path(folder).iterdir()))

        ttk.Button(buttons, text="Select files…", command=lambda: apply(
            filedialog.askopenfilenames(parent=dialog, filetypes=EXCEL_FILES))).pack(side="left")
        ttk.Button(buttons, text="Select folder…", command=choose_folder).pack(side="left", padx=5)
        ttk.Button(buttons, text="Close", command=dialog.destroy).pack(side="right")
        dialog.grab_set()

    @staticmethod
    def show(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def selected(self):
        selection = self.batch.selection()
        return self.results[int(selection[0])] if selection else None

    def select_submission(self, _event=None):
        result = self.selected()
        if not result:
            return
        self.selection.configure(text=f"{result['filename']}  ·  {result.get('total', '—')}/{MAXIMUM}  ·  {result['status']}")
        self.rubric.delete(*self.rubric.get_children())
        for index, item in enumerate(result.get("items", [])):
            self.rubric.insert("", "end", iid=str(index),
                               values=(item["name"], f"{item['score']}/{item['maximum']}", "Yes" if item["review"] else ""),
                               tags=("deduction",) if item["score"] < item["maximum"] else ())
        self.show(self.feedback, feedback(result))
        self.show(self.details, details(result))

    def select_criterion(self, _event=None):
        selection = self.rubric.selection()
        result = self.selected()
        if result and selection:
            index = int(selection[0])
            self.show(self.details, details(result, index))
            self.feedback.tag_remove("focused", "1.0", "end")
            item = result["items"][index]
            if self.views.select() == str(self.feedback) and item["score"] < item["maximum"]:
                preceding = [item_feedback(other) for other in result["items"][:index]
                             if other["score"] < other["maximum"]]
                line = 3 + sum(block.count("\n") + 2 for block in preceding)
                start = f"{line}.0"
                end = f"{line + item_feedback(item).count(chr(10))}.end"
                self.feedback.tag_add("focused", start, end)
                self.feedback.yview(start)

    def open_workbook(self, event):
        if event.keysym != "Return":
            row = self.batch.identify_row(event.y)
            if not row or self.batch.identify_column(event.x) != "#1":
                return
            self.batch.selection_set(row)
        result = self.selected()
        if not result:
            return
        try:
            path = Path(result["source_path"])
            if not path.is_file():
                raise FileNotFoundError(path)
            if sys.platform == "win32":
                os.startfile(str(path))
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])
        except OSError as error:
            messagebox.showerror("Could not open workbook", str(error))

    def clear(self):
        self.results.clear()
        self.batch.delete(*self.batch.get_children())
        self.rubric.delete(*self.rubric.get_children())
        self.selection.configure(text="Select a submission to inspect its rubric scores.")
        self.show(self.feedback, "")
        self.show(self.details, "")
        self.status.set("Batch cleared.")

    def save(self, kind):
        if not self.results:
            messagebox.showinfo("No submissions", "Add submissions before exporting.")
            return
        export_options = {
            "feedback": ("Feedback", ".txt", [("Plain-language feedback", "*.txt")]),
            "csv": ("Grades", ".csv", [("CSV grade table", "*.csv")]),
            "report": ("Report", ".txt", [("Detailed text report", "*.txt"),
                                          ("Machine-readable JSON", "*.json"),
                                          ("Formatted HTML", "*.html")]),
        }
        label, suffix, types = export_options[kind]
        path = filedialog.asksaveasfilename(defaultextension=suffix, filetypes=types,
                                            initialfile=f"BCIT_BSYS1000_Fall2026_{label}{suffix}")
        if path:
            try:
                report_kind = Path(path).suffix.lower().lstrip(".") if kind == "report" else kind
                export_report(path, self.results, report_kind)
                self.status.set(f"Saved {Path(path).name}")
            except OSError as error:
                messagebox.showerror("Could not save report", str(error))


def main():
    parser = argparse.ArgumentParser(description=TITLE)
    parser.add_argument("workbooks", nargs="*")
    parser.add_argument("--benchmark", type=Path)
    parser.add_argument("--json", type=Path, help="Grade without opening the desktop and save JSON evidence.")
    args = parser.parse_args()
    if args.benchmark:
        output = json.dumps(benchmark(args.benchmark, grade), ensure_ascii=False, indent=2)
        if args.json:
            args.json.write_text(output, encoding="utf-8")
        else:
            print(output)
        return
    if args.json:
        export_report(args.json, [grade(path) for path in args.workbooks], "json")
        return
    from tkinterdnd2 import DND_FILES, TkinterDnD
    root = TkinterDnD.Tk()
    app = App(root, DND_FILES)
    app.add(args.workbooks)
    root.mainloop()


if __name__ == "__main__":
    main()
