"""Native Tk charts for the filtered grading dashboard."""
from collections import defaultdict
from statistics import mean, quantiles
import tkinter as tk
from tkinter import ttk

TEAL, BG, INK, MUTED, LINE = "#00796B", "#f6f8f9", "#1d2327", "#5f6b72", "#dde3e6"
FULL, MID, LOW = "#e3f3ee", "#fff3d6", "#fde4e1"
CHART_TYPES = ("Score histogram", "Rubric outcomes", "Set × rubric heatmap", "Box plots by set", "Points lost")
FONT = ("Segoe UI", 10)


def histogram(scores):
    bins = [0] * 10
    for score in scores:
        bins[min(9, int(score // 10))] += 1
    return bins


def box_summary(values):
    if len(values) == 1:
        low = center = high = values[0]
    else:
        low, center, high = quantiles(values, n=4, method="inclusive")
    spread = 1.5 * (high - low)
    inside = [value for value in values if low - spread <= value <= high + spread]
    if not inside:
        inside = values
    outliers = [value for value in values if not min(inside) <= value <= max(inside)]
    return min(inside), low, center, high, max(inside), outliers


class DashboardCharts:
    def __init__(self, parent):
        controls = ttk.Frame(parent, padding=8)
        controls.pack(fill="x")
        ttk.Label(controls, text="Chart").pack(side="left", padx=(0, 5))
        self.kind = tk.StringVar(parent, value=CHART_TYPES[0])
        picker = ttk.Combobox(controls, textvariable=self.kind, values=CHART_TYPES, state="readonly", width=26)
        picker.pack(side="left")
        picker.bind("<<ComboboxSelected>>", self.draw)
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(frame, background="white", highlightthickness=0)
        x = ttk.Scrollbar(frame, orient="horizontal", command=self.canvas.xview)
        y = ttk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=x.set, yscrollcommand=y.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        x.grid(row=1, column=0, sticky="ew")
        y.grid(row=0, column=1, sticky="ns")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.canvas.bind("<Configure>", self.draw)
        self.data = ([], [], ())

    def update(self, scored, stats, selected):
        self.data = scored, stats, selected
        self.draw()

    def text(self, x, y, text, anchor="w", color=INK, bold=False):
        self.canvas.create_text(x, y, text=text, anchor=anchor, fill=color,
                                font=(*FONT, "bold") if bold else FONT)

    def rect(self, x0, y0, x1, y1, fill):
        self.canvas.create_rectangle(x0, y0, x1, y1, fill=fill, outline=LINE)

    def percent_axis(self, left, right, y):
        self.canvas.create_line(left, y, right, y, fill=LINE)
        for value in range(0, 101, 20):
            x = left + (right - left) * value / 100
            self.canvas.create_line(x, y, x, y + 5, fill=MUTED)
            self.text(x, y + 16, f"{value}%", "center", MUTED)

    def draw(self, _event=None):
        scored, stats, selected = self.data
        self.canvas.delete("all")
        width = max(680, self.canvas.winfo_width())
        height = max(265, self.canvas.winfo_height())
        kind = self.kind.get()
        self.text(20, 20, kind, color=TEAL, bold=True)
        sets = defaultdict(list)
        for submission, result in scored:
            sets[submission.set_name].append((submission, result))
        names = sorted(sets, key=str.casefold)
        values = [value for _, value in scored if value is not None]
        if kind == CHART_TYPES[0] and values:
            self.draw_histogram(values, width, height)
        elif kind == CHART_TYPES[1] and stats:
            height = max(height, 90 + 30 * len(stats))
            self.draw_bars(stats, width, False)
        elif kind == CHART_TYPES[2] and stats:
            width = max(width, 285 + 120 * len(names))
            height = max(height, 90 + 30 * len(selected))
            self.draw_heatmap(sets, names, stats, selected)
        elif kind == CHART_TYPES[3] and values:
            height = max(height, 93 + 37 * len(names))
            self.draw_boxes(sets, names, width, height)
        elif kind == CHART_TYPES[4] and stats:
            height = max(height, 90 + 30 * len(stats))
            self.draw_bars(sorted(stats, key=lambda row: (-row[10], row[0])), width, True)
        else:
            self.text(20, 52, "No scored results for the current filters.", color=MUTED)
        self.canvas.configure(scrollregion=(0, 0, width, height))

    def draw_histogram(self, values, width, height):
        bins = histogram(values)
        self.text(20, 45, f"{len(values)} scored submission(s) · Bin width 10 percentage points · Final bin includes 100%", color=MUTED)
        left, right, top, bottom = 48, width - 28, 80, height - 45
        bar_width = (right - left) / len(bins)
        for index, count in enumerate(bins):
            x0 = left + bar_width * index
            y = bottom - (bottom - top) * count / max(bins)
            if count:
                fill = LOW if index < 5 else MID if index < 8 else FULL
                self.rect(x0, y, x0 + bar_width, bottom, fill)
                self.text(x0 + bar_width / 2, y - 11, str(count), "center")
        self.percent_axis(left, right, bottom)

    def draw_bars(self, stats, width, by_loss):
        left, right = 250, width - 130
        capacity = max((row[10] for row in stats), default=0) or 1
        description = ("Ranked by total rubric points deducted" if by_loss else
                       "Green · Full marks     Amber · Partial marks     Red · Zero marks")
        self.text(20, 45, description, color=MUTED)
        for index, row in enumerate(stats):
            name, _, n, _, _, _, full, partial, zero, _, lost = row
            y = 75 + index * 30
            self.text(20, y, name)
            lengths = ((lost / capacity, LOW),) if by_loss else ((full / n, FULL), (partial / n, MID), (zero / n, LOW))
            x0 = left
            for fraction, color in lengths:
                if fraction:
                    x1 = x0 + (right - left) * fraction
                    self.rect(x0, y - 9, x1, y + 9, color)
                    x0 = x1
            label = f"{lost:g} pts · n={n}" if by_loss else f"{full} / {partial} / {zero} · n={n}"
            self.text(right + 8, y, label, color=MUTED)

    def draw_heatmap(self, sets, names, stats, selected):
        self.text(20, 43, "Cell values show average percent of available marks and number graded.", color=MUTED)
        for col, name in enumerate(names):
            x = 250 + col * 120 + 59
            self.canvas.create_text(x, 73, text=name, fill=TEAL, width=112, font=(*FONT, "bold"))
        maxima = {row[0]: row[1] for row in stats}
        for row, item in enumerate(selected):
            y = 104 + row * 30
            self.text(20, y, item)
            for col, name in enumerate(names):
                x = 250 + col * 120
                values = [entry.scores[item] for entry, _ in sets[name] if item in entry.scores]
                value = 100 * mean(values) / maxima[item] if values else None
                fill = BG if value is None else FULL if value >= 85 else MID if value >= 50 else LOW
                self.rect(x, y - 13, x + 118, y + 13, fill)
                label = "—" if value is None else f"{value:.1f}% · n={len(values)}"
                self.text(x + 59, y, label, "center")

    def draw_boxes(self, sets, names, width, height):
        left, right = 235, width - 42
        self.text(20, 45, "Box · middle 50%     Line · median     Whiskers · within 1.5× IQR     Dots · outliers", color=MUTED)

        def position(value):
            return left + (right - left) * value / 100

        for index, name in enumerate(names):
            values = [value for _, value in sets[name] if value is not None]
            y = 75 + index * 37
            self.text(20, y, f"{name} · n={len(values)}")
            if not values:
                continue
            low, first, center, third, high, outliers = box_summary(values)
            self.canvas.create_line(position(low), y, position(high), y, fill=TEAL)
            for cap in (low, high):
                x = position(cap)
                self.canvas.create_line(x, y - 6, x, y + 6, fill=TEAL)
            self.rect(position(first), y - 10, position(third), y + 10, FULL)
            x = position(center)
            self.canvas.create_line(x, y - 10, x, y + 10, fill=TEAL, width=2)
            for point in outliers:
                x = position(point)
                self.canvas.create_oval(x - 2, y - 2, x + 2, y + 2, fill=TEAL, outline="")
        self.percent_axis(left, right, height - 33)
