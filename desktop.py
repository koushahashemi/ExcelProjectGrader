"""Instructor desktop interface for Excel Project Grader."""
import argparse
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

from grader import COURSE, MAXIMUM, RUBRIC, TITLE, VERSION, grade
from reports import benchmark, details, export_report, feedback, score_rows


class App:
    def __init__(self, root, drop_type):
        self.root = root
        self.queue = queue.Queue()
        self.results = []
        self.pending = set()
        self.remaining = 0
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
        ttk.Button(actions, text="Save student feedback…", command=lambda: self.save("feedback")).pack(side="right", padx=5)
        report = ttk.Menubutton(actions, text="Save instructor report")
        menu = tk.Menu(report, tearoff=False)
        for title, kind in (("Formatted HTML", "html"), ("Plain text", "txt"), ("JSON evidence", "json")):
            menu.add_command(label=title, command=lambda kind=kind: self.save(kind))
        report["menu"] = menu
        report.pack(side="right")

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(10, 5))
        self.status = tk.StringVar(value="Drop .xlsx submissions here, or select Add Excel files. Double-click a filename to open it.")
        ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=(0, 8))
        panes = ttk.Panedwindow(frame, orient="vertical")
        panes.pack(fill="both", expand=True)
        upper = ttk.Frame(panes)
        panes.add(upper, weight=2)
        self.batch = self.table(upper, ["Submission", f"Total /{MAXIMUM}", "Status", *[f"{name} /{points}" for name, points in RUBRIC]],
                                [240, 85, 130, *[145] * len(RUBRIC)])
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
        views = ttk.Notebook(review)
        review.add(views, weight=2)
        self.feedback = ScrolledText(views, wrap="word", padx=10, pady=8, state="disabled")
        self.details = ScrolledText(views, wrap="word", padx=10, pady=8, state="disabled")
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
        self.add(filedialog.askopenfilenames(filetypes=[("Excel workbook", "*.xlsx")], title="Select submissions"))

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
        candidates = dict.fromkeys(Path(value).resolve() for value in paths)
        workbooks = [path for path in candidates if path.is_file() and path.suffix.lower() == ".xlsx"
                     and not path.name.startswith("~$") and path not in self.pending
                     and str(path) not in {result["source_path"] for result in self.results}]
        if not workbooks:
            if candidates:
                self.status.set("No new .xlsx submissions found.")
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
                                  tags=("deduction",) if "error" in result else ())
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
        self.root.after(75, self.poll)

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
            self.show(self.details, details(result, int(selection[0])))

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
        suffix = "txt" if kind == "feedback" else kind
        path = filedialog.asksaveasfilename(defaultextension=f".{suffix}", filetypes=[(suffix.upper(), f"*.{suffix}")],
                                            initialfile=f"excel-grader-{kind}.{suffix}")
        if path:
            try:
                export_report(path, self.results, kind)
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
