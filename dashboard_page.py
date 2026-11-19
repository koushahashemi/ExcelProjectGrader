"""A zoomable, two-axis canvas for the dashboard's charts and tables."""
import tkinter as tk
from tkinter import font, ttk

from dashboard_charts import CHART_TYPES, DashboardCharts, BG, FULL, INK, LINE, LOW, MID, MUTED, TEAL

PAGE = "#eef4f2"
CARD = "#ffffff"
ACCENT = "#d7ebe4"
ROW_HEIGHT = 29
METRICS = ("Scored / loaded", "Minimum", "Maximum", "Average", "Median")
ZOOM_LEVELS = (50, 65, 75, 85, 100, 115, 125, 150, 175, 200)
TABLES = {
    "Score distribution": (("Score range", 240), ("Submissions", 145), ("Share", 110)),
    "Rubric statistics": (("Rubric item", 187), ("Max pts", 66), ("Graded", 62), ("Mean", 59),
                          ("Min", 54), ("Max", 54), ("Full", 51), ("Partial", 56),
                          ("Zero", 50), ("Below full", 77), ("Pts lost", 70)),
    "Set comparison": (("Set", 240), ("Loaded", 73), ("Scored", 73), ("Minimum", 83),
                       ("Maximum", 85), ("Average", 83), ("Median", 81)),
    "Submissions": (("Filename", 530), ("Set", 250), ("Filtered score", 170), ("Overall score", 170)),
}


class DashboardPage:
    def __init__(self, parent, controls):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, padx=12, pady=(3, 10))
        self.canvas = tk.Canvas(frame, background=PAGE, highlightthickness=0, takefocus=True)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=self.canvas.xview)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=horizontal.set, yscrollcommand=vertical.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.canvas.bind("<Button-1>", lambda _event: self.canvas.focus_set())
        for sequence, direction in (("<Up>", -1), ("<Down>", 1)):
            self.canvas.bind(sequence, lambda _event, steps=direction: self.canvas.yview_scroll(steps, "units"))
        for sequence, direction in (("<Left>", -1), ("<Right>", 1)):
            self.canvas.bind(sequence, lambda _event, steps=direction: self.canvas.xview_scroll(steps, "units"))
        for button in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind(button, self.wheel)

        self.zoom = tk.StringVar(parent, value="Fit width")
        self.factor = 1
        self.fonts = {}
        self.measure_font = font.Font(root=self.canvas, family="Segoe UI", size=11)
        self.metric_values = ()
        self.insight = ""
        self.data = ([], [], ())
        self.tables = {}
        self.charts = DashboardCharts(self.canvas)
        self.pending_draw = None
        ttk.Label(controls, text="Zoom").pack(side="left", padx=(0, 6))
        ttk.Button(controls, text="−", width=3, command=lambda: self.change_zoom(-1)).pack(side="left")
        picker = ttk.Combobox(controls, textvariable=self.zoom, state="readonly", width=10,
                              values=["Fit width", *[f"{z}%" for z in ZOOM_LEVELS]])
        picker.pack(side="left", padx=4)
        picker.bind("<<ComboboxSelected>>", lambda _event: self.redraw())
        ttk.Button(controls, text="+", width=3, command=lambda: self.change_zoom(1)).pack(side="left")
        ttk.Button(controls, text="Fit width", command=self.fit).pack(side="left", padx=6)
        self.canvas.bind("<Configure>", self.resize)

    def resize(self, _event):
        if self.pending_draw is not None:
            self.canvas.after_cancel(self.pending_draw)
        self.pending_draw = self.canvas.after_idle(self.redraw)

    def fit(self):
        self.zoom.set("Fit width")
        self.redraw()
        self.canvas.xview_moveto(0)

    def change_zoom(self, direction):
        percent = round(self.factor * 100)
        options = ZOOM_LEVELS if direction > 0 else reversed(ZOOM_LEVELS)
        level = next((value for value in options if (value - percent) * direction > 0), None)
        if level is not None:
            self.zoom.set(f"{level}%")
            self.redraw()

    def wheel(self, event):
        if event.num in (4, 5):
            steps = -1 if event.num == 4 else 1
        elif event.delta:
            steps = -int(event.delta / 120 or (1 if event.delta > 0 else -1))
        else:
            return "break"
        if event.state & 0x4:
            self.change_zoom(-1 if steps > 0 else 1)
        elif event.state & 0x1:
            self.canvas.xview_scroll(steps, "units")
        else:
            self.canvas.yview_scroll(steps, "units")
        return "break"

    def update(self, metrics, insight, scored, rubric_stats, selected, tables):
        self.metric_values, self.insight = metrics, insight
        self.data = scored, rubric_stats, selected
        self.tables = tables
        self.redraw()

    def rounded(self, x, y, width, height, fill, radius=18):
        right, bottom = x + width, y + height
        points = (x, y + radius, x, y, x + radius, y, right - radius, y, right, y,
                  right, y + radius, right, bottom - radius, right, bottom,
                  right - radius, bottom, x + radius, bottom, x, bottom, x, bottom - radius)
        self.canvas.create_polygon(points, smooth=True, splinesteps=16, fill=fill, outline="")

    def card(self, x, y, width, height, fill=CARD):
        self.rounded(x, y + 4, width, height, "#e0ebe7")
        self.rounded(x, y, width, height, fill)

    def text(self, x, y, value, *, size=11, color=INK, bold=False, anchor="w"):
        self.canvas.create_text(x, y, text=value, anchor=anchor, fill=color,
                                font=("Segoe UI", size, "bold" if bold else "normal"))

    def table_height(self, title):
        return 76 + max(1, len(self.tables.get(title, []))) * ROW_HEIGHT

    def clipped(self, value, width):
        value = str(value)
        if self.measure_font.measure(value) <= width:
            return value
        low, high = 0, len(value)
        while low + 1 < high:
            middle = (low + high) // 2
            if self.measure_font.measure(value[:middle] + "…") <= width:
                low = middle
            else:
                high = middle
        return value[:low] + "…"

    def draw_table(self, title, x, y, width, height):
        self.card(x, y, width, height)
        self.text(x + 18, y + 22, title, size=13, color=TEAL, bold=True)
        columns = TABLES[title]
        first_width = width - 32 - sum(col_width for _, col_width in columns[1:])
        columns = [(columns[0][0], first_width), *columns[1:]]
        self.rounded(x + 12, y + 41, width - 24, 26, ACCENT, radius=9)
        for index, (label, _) in enumerate(columns):
            left = x + 18 + sum(cell_width for _, cell_width in columns[:index])
            self.text(left, y + 54, label, size=10, color=TEAL, bold=True)
        rows = self.tables.get(title, [])
        if not rows:
            self.text(x + 18, y + 84, "No results for the current filters.", color=MUTED)
        for index, (values, tint) in enumerate(rows):
            current = y + 69 + index * ROW_HEIGHT
            if index % 2 == 0:
                self.rounded(x + 12, current, width - 24, 27, BG, radius=7)
            left = x + 18
            for column_index, (value, (_, cell_width)) in enumerate(zip(values, columns)):
                if tint and column_index >= len(columns) - 2:
                    self.rounded(left - 3, current + 3, cell_width - 5, 21, tint, radius=7)
                self.text(left, current + 13, self.clipped(value, cell_width - 11))
                left += cell_width

    def draw_chart(self, kind, x, y, width, height, chart_height):
        self.card(x, y, width, height)
        self.rounded(x + 16, y + 13, 4, 15, TEAL, radius=2)
        before = set(self.canvas.find_all())
        self.charts.render(kind, width - 16, chart_height, *self.data)
        for item in self.canvas.find_all():
            if item not in before:
                self.canvas.move(item, x + 9, y)

    def scale_fonts(self):
        for item in self.canvas.find_withtag("all"):
            if self.canvas.type(item) != "text":
                continue
            original = tuple(self.canvas.tk.splitlist(self.canvas.itemcget(item, "font")))
            family, size, *style = original
            key = (family, max(1, round(int(size) * self.factor)), *style)
            if key not in self.fonts:
                self.fonts[key] = font.Font(root=self.canvas, family=family, size=key[1],
                                            weight="bold" if "bold" in style else "normal")
            width = float(self.canvas.itemcget(item, "width"))
            self.canvas.itemconfigure(item, font=self.fonts[key], width=round(width * self.factor))

    def redraw(self):
        self.pending_draw = None
        old_x, old_y = self.canvas.xview()[0], self.canvas.yview()[0]
        self.canvas.delete("all")
        gap, margin = 18, 18
        pairs = ((CHART_TYPES[0], "Score distribution"), (CHART_TYPES[1], "Rubric statistics"),
                 (CHART_TYPES[2], "Set comparison"), (CHART_TYPES[3], CHART_TYPES[4]))
        sizes = {kind: self.charts.measure(kind, 680, *self.data) for kind in CHART_TYPES}
        left_width = max(sizes[kind][0] for kind, _ in pairs) + 16
        right_width = 854
        width = 2 * margin + left_width + gap + right_width
        span = width - 2 * margin
        metric_width = (span - 4 * gap) / 5
        for index, (name, value) in enumerate(zip(METRICS, self.metric_values)):
            x = margin + index * (metric_width + gap)
            self.card(x, margin, metric_width, 88, "#dcefe8" if index == 3 else CARD)
            self.text(x + 18, margin + 24, name, color=MUTED)
            self.text(x + 18, margin + 60, value, size=25, color=TEAL, bold=True)
        y = margin + 103
        self.rounded(margin, y, span, 35, MID, radius=12)
        self.text(margin + 17, y + 17, self.clipped(self.insight, span - 32), color=INK)
        y += 51
        for left, right in pairs:
            left_height = sizes[left][1]
            right_height = sizes[right][1] if right in sizes else self.table_height(right)
            height = max(left_height, right_height)
            self.draw_chart(left, margin, y, left_width, height, left_height)
            right_x = margin + left_width + gap
            if right in sizes:
                self.draw_chart(right, right_x, y, right_width, height, right_height)
            else:
                self.draw_table(right, right_x, y, right_width, height)
            y += height + gap
        last_height = self.table_height("Submissions")
        self.draw_table("Submissions", margin, y, span, last_height)
        height = y + last_height + margin
        self.factor = (max(1, self.canvas.winfo_width()) / width if self.zoom.get() == "Fit width"
                       else int(self.zoom.get().rstrip("%")) / 100)
        self.canvas.scale("all", 0, 0, self.factor, self.factor)
        self.scale_fonts()
        self.canvas.configure(scrollregion=(0, 0, round(width * self.factor), round(height * self.factor)),
                              xscrollincrement=18, yscrollincrement=24)
        self.canvas.xview_moveto(0 if self.zoom.get() == "Fit width" else old_x)
        self.canvas.yview_moveto(old_y)
