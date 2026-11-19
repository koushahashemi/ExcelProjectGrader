"""Read saved Excel workbooks without launching Excel or executing formulas."""
from collections import Counter, defaultdict
from datetime import date, timedelta
import csv
from difflib import SequenceMatcher
import math
from pathlib import Path
import re
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.styles.numbers import is_date_format

VERSION = "1.5.0"
WORKBOOK_SUFFIXES = {".xlsx", ".xlsm"}
TITLE = "Excel Project Grader"
COURSE = "BSYS-1000 · Individual Project (Excel) · Fall 2026"
SETUP = ("CountryCodes import", "Hide Parameters", "Unique countries", "Adjacent dates")
FORMAT = ("Column widths", "Data types", "Excel Table Creation")
FORMULAS = ("Total Revenue", "Total Cost", "Total Profit", "Days to Fulfil", "Order Status", "ISO Codes")
GROUPS = ("Bar", "Line", "Combo")
RUBRIC = [(name, 1) for name in SETUP] + [(name, 2) for name in FORMAT + FORMULAS]
RUBRIC += [(f"{group} · {item}", 2) for group in GROUPS
           for item in ("Aggregation", "Chart", "Titles", "PivotTable")]
MAXIMUM = sum(points for _, points in RUBRIC)

with Path(__file__).with_name("country_codes.csv").open(encoding="utf-8", newline="") as source:
    COUNTRY_ROWS = list(csv.reader(source))[1:]


def key(value):
    return re.sub(r"[^\w]|_", "", str(value or "").casefold())


COUNTRIES = {key(row[0]): row[2] for row in COUNTRY_ROWS}
HEADERS = ("Region", "Country", "ISO Codes", "Item Type", "Sales Channel", "Order Priority",
           "Order Date", "Order ID", "Ship Date", "Units Sold", "Unit Price", "Unit Cost", *FORMULAS[:-1])
MONEY = {"Unit Price", "Unit Cost", "Total Revenue", "Total Cost", "Total Profit"}


def similar(actual, expected):
    a, b = key(actual), key(expected)
    if a == b:
        return True
    return (len(b) >= 7 and re.findall(r"\d+", a) == re.findall(r"\d+", b)
            and SequenceMatcher(None, a, b).ratio() >= .9)


def header_map(sheet):
    populated = [(cell.row, cell.value) for cell in sheet._cells.values() if cell.value is not None]
    first_rows = sorted({row for row, _ in populated})[:12]
    best = (0, 1, {})
    for row in first_rows:
        found = {}
        for cell in (cell for cell in sheet._cells.values() if cell.row == row and cell.value is not None):
            matches = [name for name in HEADERS if similar(cell.value, name)]
            if len(matches) == 1:
                found[matches[0]] = cell.column
        if len(found) > best[0]:
            best = len(found), row, found
    return best[1:]


def sheet_named(book, name):
    return next((s for s in book if similar(s.title, name)), None)


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def equal(actual, expected):
    if numeric(expected):
        if isinstance(actual, timedelta):
            actual = actual.total_seconds() / 86400
        return numeric(actual) and math.isclose(actual, expected, rel_tol=1e-7, abs_tol=.011)
    return similar(actual, expected)


def day(value, epoch):
    if numeric(value):
        try:
            value = from_excel(value, epoch)
        except (OverflowError, ValueError):
            return None
    return value if isinstance(value, date) else None


def currency(fmt):
    # A locale-only marker such as [$-409] carries no currency symbol.
    fmt = re.sub(r"\[\$-.*?\]", "", fmt)
    return bool(re.search(r"[$€£¥₹₩₽₫₪₱₺₴₦฿¢]|\b(?:USD|CAD|AUD|EUR|GBP|JPY|CHF)\b", fmt, re.I))


def format_ok(cell, name, value=None):
    fmt = cell.number_format
    code = re.sub(r'"[^"]*"|\\.|\[[^]]*\]', "", fmt)
    if re.search(r"[0#?]E[+-][0#?]", code, re.I):
        return False
    if name in MONEY:
        return currency(fmt) and bool(re.search(r"[0#?]", code))
    if name in {"Order Date", "Ship Date"}:
        return is_date_format(fmt) and (bool(re.search(r"[dy]", code, re.I)) or ("m" in code.casefold() and not re.search(r"[hs]", code, re.I)))
    if name in {"Order ID", "Units Sold", "Days to Fulfil"} and fmt.casefold() != "general":
        sections = code.split(";")
        number = r"[+-]?\(?[0#?]+(?:,[0#?]+)*(?:\.[0#?]+)?\)?[+-]?"
        numeric_sections = [re.sub(r"_.|\*.", "", section).replace(" ", "") for section in sections[:3]]
        return (bool(numeric_sections[0]) and not currency(fmt)
                and all(re.fullmatch(number, section) for section in numeric_sections if section)
                and (name != "Order ID" or all("." not in section for section in numeric_sections)))
    if fmt.casefold() not in {"general", "@"}:
        return False
    if fmt.casefold() == "general" and numeric(value) and value != 0:
        magnitude = abs(value)
        available = width(cell, cell.parent.sheet_format.defaultColWidth or 8.43)
        digits = len(str(abs(int(value))))
        if magnitude >= 1e11 or magnitude < 1e-4 or (digits >= 7 and digits > available + 1):
            return False  # Likely General-format scientific display.
    return True


def rubric_item(index, score, location, expected, detected, feedback="", review=""):
    name, maximum = RUBRIC[index]
    return dict(name=name, maximum=maximum, score=score, location=location, expected=expected,
                detected=detected, feedback=feedback, review=review)


def bands(errors):
    return 2 if not errors else 1 if len(errors) <= 2 else 0


def dependencies(formula, row, columns):
    """Resolve common A1 and structured references for arithmetic checks."""
    def structured(match):
        token = match[1].strip("@")
        return next((f"{{{name}}}" for name in columns if key(name) == key(token)), match[0])

    table = r"(?:[A-Z_][\w.]*)?"
    row_ref = table + r"\[\[#This\s+Row\]\s*,\s*\[([^\[\]]+)\]\]"
    short_ref = table + r"\[@(?:\[([^\[\]]+)\]|([^\[\]]+))\]"
    formula = re.sub(row_ref, structured, formula, flags=re.I)
    formula = re.sub(short_ref, lambda m: next(
        (f"{{{name}}}" for name in columns if key(name) == key(m[1] or m[2])), m[0]), formula, flags=re.I)
    formula = re.sub(r"(?<![\w\[\]])\[([^\[\]]+)\](?![\]:])", structured, formula)
    by_column = {get_column_letter(col): name for name, col in columns.items()}
    formula = re.sub(r"(?<![\w}])\$?([A-Z]{1,3})\$?(\d+)(?!\w)",
                     lambda m: f"{{{by_column[m[1].upper()]}}}" if int(m[2]) == row
                     and m[1].upper() in by_column else m[0], formula, flags=re.I)
    return formula


def formula_matches(formula, name, row, columns):
    expression = dependencies(formula, row, columns).replace(" ", "").upper()
    body = expression.removeprefix("=").replace("_XLFN.", "")
    simple = re.sub(r"[()]", "", body)
    operands = {
        "Total Revenue": ("{UNITSSOLD}", "{UNITPRICE}", "*"),
        "Total Cost": ("{UNITSSOLD}", "{UNITCOST}", "*"),
        "Total Profit": ("{TOTALREVENUE}", "{TOTALCOST}", "-"),
        "Days to Fulfil": ("{SHIPDATE}", "{ORDERDATE}", "-"),
    }
    if name in operands:
        a, b, op = operands[name]
        if simple == a + op + b or (op == "*" and simple == b + op + a):
            return True
        if op == "*" and body in {f"PRODUCT({a},{b})", f"PRODUCT({b},{a})"}:
            return True
        if name == "Days to Fulfil" and body in {f"DAYS({a},{b})", f'DATEDIF({b},{a},"D")'}:
            return True
        if name == "Total Profit":
            return simple == "{UNITSSOLD}*{UNITPRICE}-{UNITSSOLD}*{UNITCOST}" or body == "{UNITSSOLD}*({UNITPRICE}-{UNITCOST})"
        return False
    if name == "Order Status":
        required = ('{ORDERPRIORITY}="C"', '"C"={ORDERPRIORITY}')
        days = ('{DAYSTOFULFIL}>14', '14<{DAYSTOFULFIL}', '{DAYSTOFULFIL}>=15', '15<={DAYSTOFULFIL}')
        return any(re.fullmatch(r'IF\(AND\(' + re.escape(a + "," + b) + r'\),"DELAYED","ONTIME"\)', body)
                   for first in required for second in days for a, b in ((first, second), (second, first)))
    table_range = r"CODESTABLE(?:\[(?:#ALL|#DATA)?\]|\[\[COUNTRY\]:\[ISOCODES\]\])?"
    sheet = r"(?:'COUNTRYCODES'|COUNTRYCODES)!"
    country_range = r"\$?A\$?\d+:\$?A\$?\d+"
    iso_range = r"\$?C\$?\d+:\$?C\$?\d+"
    vlookup = rf"VLOOKUP\(\{{COUNTRY\}},(?:{table_range}|{sheet}\$?A\$?\d+:\$?[C-F]\$?\d+),3,(?:FALSE(?:\(\))?|0)\)"
    xlookup = (rf"XLOOKUP\(\{{COUNTRY\}},(?:CODESTABLE\[COUNTRY\]|{sheet}{country_range}),"
               rf'(?:CODESTABLE\[ISOCODES\]|{sheet}{iso_range})(?:,(?:[^,()"]|"[^"]*")*(?:,0?(?:,[+-]?[12])?)?)?\)')
    return bool(re.fullmatch(vlookup + "|" + xlookup, body))


def load_pair(path):
    # Inspect the ZIP directory before decoding the workbook. No external links or macros are loaded.
    with ZipFile(path) as archive:
        if sum(entry.file_size for entry in archive.infolist()) > 150_000_000:
            raise ValueError("Workbook exceeds the grader's 150 MB expanded size limit.")
    book = load_workbook(path, keep_links=False)
    try:
        values = load_workbook(path, data_only=True, keep_links=False)
    except Exception:
        book.close()
        raise
    return book, values


class Detail:
    def __init__(self, sheet, cached):
        self.sheet, self.cached = sheet, cached
        self.header, self.columns = header_map(sheet)
        region = self.columns.get("Region")
        country = self.columns.get("Country")
        order_id = self.columns.get("Order ID")
        indicators = {column for column in (region, country, order_id) if column}
        self.rows = sorted({cell.row for cell in sheet._cells.values() if cell.row > self.header
                            and cell.column in indicators and cell.value is not None})

    def cell(self, row, name):
        col = self.columns.get(name)
        return self.sheet.cell(row, col) if col else None

    def value(self, row, name):
        cell = self.cell(row, name)
        return self.cached[cell.coordinate].value if cell else None

    def location(self, name, rows):
        column = self.columns.get(name)
        if not column:
            return f"{self.sheet.title} · {name}"
        letter = get_column_letter(column)
        addresses = ", ".join(f"{letter}{row}" for row in rows[:12])
        if len(rows) > 12:
            addresses += f" (+{len(rows)-12} rows)"
        return f"{self.sheet.title}!{addresses or letter + str(self.header)}"

    def expected(self, row, name):
        def get(field):
            return self.value(row, field)

        def multiply(a, b):
            return a * b if numeric(a) and numeric(b) else None

        if name == "Total Revenue":
            return multiply(get("Units Sold"), get("Unit Price"))
        if name == "Total Cost":
            return multiply(get("Units Sold"), get("Unit Cost"))
        if name == "Total Profit":
            revenue, cost = get("Total Revenue"), get("Total Cost")
            if revenue is None and self.cell(row, "Total Revenue") and self.cell(row, "Total Revenue").data_type == "f":
                revenue = self.expected(row, "Total Revenue")
            if cost is None and self.cell(row, "Total Cost") and self.cell(row, "Total Cost").data_type == "f":
                cost = self.expected(row, "Total Cost")
            return revenue - cost if numeric(revenue) and numeric(cost) else None
        if name == "Days to Fulfil":
            start = day(get("Order Date"), self.sheet.parent.epoch)
            end = day(get("Ship Date"), self.sheet.parent.epoch)
            return (end - start).days if start and end else None
        if name == "Order Status":
            elapsed = get("Days to Fulfil")
            if isinstance(elapsed, timedelta):
                elapsed = elapsed.days
            if elapsed is None:
                elapsed = self.expected(row, "Days to Fulfil")
            return ("Delayed" if key(get("Order Priority")) == "c" and elapsed > 14 else "On time") if numeric(elapsed) else None
        return COUNTRIES.get(key(get("Country")))


def setup_items(book, detail):
    codes = sheet_named(book, "CountryCodes")
    if codes is None:
        codes = next((s for s in book.worksheets if any(key(t) == "codestable" for t in s.tables)), None)
    provided = set()
    if codes is not None:
        for cell in list(codes._cells.values()):
            if cell.column == 1 and cell.value:
                provided.add((key(cell.value), key(codes.cell(cell.row, 3).value)))
    missing = {(country, key(iso)) for country, iso in COUNTRIES.items()} - provided
    parameters = sheet_named(book, "Parameters")
    found = defaultdict(list)
    for row in detail.rows:
        if key(detail.value(row, "Country")):
            found[key(detail.value(row, "Country"))].append(row)
    duplicate = sorted(row for rows in found.values() if len(rows) > 1 for row in rows)
    dates = [detail.columns.get(name) for name in ("Order Date", "Ship Date")]
    adjacent = all(dates) and abs(dates[0] - dates[1]) == 1
    return [
        rubric_item(0, int(not missing), codes.title if codes else "Workbook", "Imported CountryCodes worksheet",
                    f"{len(COUNTRIES) - len(missing)}/{len(COUNTRIES)} source country and ISO pairs",
                    "Import the complete CountryCodes worksheet (task 3)."),
        rubric_item(1, int(parameters is not None and parameters.sheet_state != "visible"), "Parameters",
                    "Hidden worksheet", parameters.sheet_state if parameters else "Missing",
                    "Keep the Parameters worksheet and hide it (task 10)."),
        rubric_item(2, int(bool(found) and not duplicate), detail.location("Country", duplicate),
                    "One transaction per country", f"{len(found)} distinct countries; {len(duplicate)} rows in duplicate groups",
                    "Remove duplicates by Country (task 4)."),
        rubric_item(3, int(bool(adjacent)), detail.sheet.title, "Order Date beside Ship Date", str(dates),
                    "Move Order Date and Ship Date next to each other (task 5)."),
    ]


def width(cell, default):
    for dim in cell.parent.column_dimensions.values():
        if dim.min is not None and dim.min <= cell.column <= dim.max:
            return 0 if dim.hidden else dim.width
        if dim.index == cell.column_letter:
            return 0 if dim.hidden else dim.width
    return default


def display_width(cell, value, name):
    if value is None:
        return 0
    if isinstance(value, date):
        fmt = cell.number_format.casefold()
        return 15 if "mmmm" in fmt else 11
    if numeric(value):
        if currency(cell.number_format):
            decimals = re.search(r"\.([0#]+)", cell.number_format)
            places = len(decimals[1]) if decimals else 0
            return len(f"${value:,.{places}f}")
        return len(f"{value:g}")
    text = str(value)
    return max((sum(.55 if ch in "ilI.,' " else 1.2 if ch in "MW@" else .93 for ch in line)
                for line in text.splitlines()), default=0)


def format_items(detail):
    sheet = detail.sheet
    narrow, mistyped = [], []
    for name, column in detail.columns.items():
        header = sheet.cell(detail.header, column)
        default = sheet.sheet_format.defaultColWidth or 8.43
        available = width(header, default)
        samples = [(header, header.value)] + [(detail.cell(row, name), detail.value(row, name)) for row in detail.rows]
        if any(not cell.alignment.wrap_text and not cell.alignment.shrink_to_fit
               and display_width(cell, value, name) > available + 1.5 for cell, value in samples):
            narrow.append(f"{header.column_letter} ({name}, width {available:g})")
        if name in FORMULAS:
            continue  # Formula formatting is scored in its own item.
        invalid = []
        for cell, value in samples[1:]:
            valid = value is not None and format_ok(cell, name, value)
            if name in MONEY | {"Units Sold", "Order ID"}:
                valid = valid and numeric(value)
            if name in {"Order Date", "Ship Date"}:
                valid = valid and day(value, sheet.parent.epoch) is not None
            if not valid:
                invalid.append(cell.coordinate)
        if invalid:
            mistyped.append(f"{name}: {', '.join(invalid[:6])}")
    tables = list(sheet.tables.values())
    table = next((t for t in tables if key(t.name) == "detailtable"), tables[0] if tables else None)
    proper = False
    if table and detail.columns and detail.rows:
        left, top, right, bottom = range_boundaries(table.ref)
        proper = (key(table.name) == "detailtable" and top == detail.header
                  and left <= min(detail.columns.values()) and right >= max(detail.columns.values())
                  and bottom - bool(table.totalsRowCount) == max(detail.rows))
    return [
        rubric_item(4, bands(narrow) if detail.rows else 0, sheet.title, "Widths that display data (task 1)",
                    "; ".join(narrow) or "No narrow columns detected",
                    "Widen these columns on " + sheet.title + ": " + "; ".join(narrow) + "." if narrow else
                    "" if detail.rows else "No transaction rows were available to check column widths.",
                    "Excel rendering depends on fonts and display settings. Widths are estimated from saved column settings."),
        rubric_item(5, bands(mistyped) if detail.rows else 0, sheet.title, "Appropriate dates, numeric types and currency (task 1)",
                    "; ".join(mistyped) or "No base-column type problems detected",
                    "Correct the value types or formats in " + "; ".join(mistyped) + "." if mistyped else
                    "" if detail.rows else "No transaction rows were available to check data types."),
        rubric_item(6, 2 if proper else 1 if table else 0, sheet.title, "DetailTable covers the transaction data (task 2)",
                    ", ".join(f"{t.name} {t.ref}" for t in tables) or "No Excel table", "Create or correct DetailTable and its data range."),
    ]


def formula_items(detail):
    expectations = ["Units Sold × Unit Price", "Units Sold × Unit Cost", "Total Revenue − Total Cost",
                    "Ship Date − Order Date", 'IF with AND: critical (C) and more than 14 days → Delayed, otherwise On time',
                    "Exact-match VLOOKUP or XLOOKUP of Country returning the three-character ISO code"]
    for offset, (name, expected_text) in enumerate(zip(FORMULAS, expectations)):
        good, uncertain, errors, fmt = [], [], [], []
        seen_required = 0
        for row in detail.rows:
            cell = detail.cell(row, name)
            value = detail.value(row, name)
            expected = detail.expected(row, name)
            formula = str(cell.value) if cell and cell.data_type == "f" else ""
            required = (name not in {"Order Status", "ISO Codes"} or
                        bool(re.search(r"\b" + ("IF" if name == "Order Status" else "(?:VLOOKUP|XLOOKUP)") + r"\s*\(", formula, re.I)))
            if formula and required:
                seen_required += 1
            matched = formula_matches(formula, name, row, detail.columns)
            cache_correct = value is not None and expected is not None and equal(value, expected)
            if formula and required and ((matched and (value is None or expected is None or cache_correct))
                                         or (cache_correct and name not in {"Order Status", "ISO Codes"})):
                good.append(row)
                if not matched or expected is None:
                    uncertain.append(row)
                if not format_ok(cell, name, value):
                    fmt.append(row)
            else:
                errors.append((row, f"{cell.coordinate if cell else 'Missing column'}: {formula or value!s} → {value!s}; expected {expected!s}"))
                if formula and required and (value is None or cache_correct) and not matched:
                    uncertain.append(row)
        position_issue = False
        col = detail.columns.get(name)
        if col and name == "Order Status" and "Days to Fulfil" in detail.columns:
            position_issue = col != detail.columns["Days to Fulfil"] + 1
        if col and name == "ISO Codes" and {"Country", "Item Type"} <= detail.columns.keys():
            position_issue = not detail.columns["Country"] + 1 == col == detail.columns["Item Type"] - 1
        full = bool(good) and not (errors or fmt or position_issue)
        score = 2 if full else 1 if good or (name in {"Order Status", "ISO Codes"} and seen_required) else 0
        detected = f"{len(good)}/{len(detail.rows)} rows supported; {len(fmt)} format issues"
        if position_issue:
            detected += "; column is out of position"
        if errors:
            detected += "\n" + "\n".join(message for _, message in errors[:20])
        issues = []
        if errors:
            labels = "Use the required formula and check its results" if name != "ISO Codes" else "Use exact-match VLOOKUP or XLOOKUP to return the three-letter ISO code"
            issues.append(f"{labels} in {detail.location(name, [row for row, _ in errors])}.")
        if fmt:
            target = {
                "Days to Fulfil": "General or Number without scientific notation",
                "Order Status": "Text or General",
                "ISO Codes": "Text or General",
            }.get(name, "Accounting or Currency")
            issues.append(f"Use {target} in {detail.location(name, fmt)}.")
        if position_issue:
            target = "immediately after Days to Fulfil" if name == "Order Status" else "between Country and Item Type"
            issues.append(f"Move {name} {target}.")
        if not detail.rows:
            issues.append("No transaction rows were available to check this formula.")
        yield rubric_item(7+offset, score, detail.location(name, [row for row, _ in errors] or fmt or uncertain),
                          expected_text, detected, " ".join(issues),
                          f"Review {len(uncertain)} row(s) whose formula equivalence or inputs could not be fully verified." if uncertain else "")


CHARTS = {
    "Bar": ("Financial Performance by Region", "Region", ("Total Revenue", "Total Cost", "Total Profit")),
    "Line": ("Annual Sales by Channel", "Order Date", ("Sales Channel",)),
    "Combo": ("Units Sold and Fulfillment Time by Priority", "Order Priority", ("Units Sold", "Days to Fulfil")),
}


def summary(detail, group):
    _, category, fields = CHARTS[group]
    buckets = defaultdict(lambda: defaultdict(list))
    for row in detail.rows:
        label = detail.value(row, category)
        if group == "Line":
            order_date = day(label, detail.sheet.parent.epoch)
            label = order_date.year if order_date else None
        if label is None:
            continue
        label = key(label)
        for field in fields:
            value = detail.value(row, field)
            if group == "Line":
                if value:
                    buckets[label][key(value)].append(1)
                continue
            if value is None:
                value = detail.expected(row, field)
            if isinstance(value, timedelta):
                value = value.days
            if numeric(value):
                buckets[label][field].append(value)
    metrics = sorted({field for bucket in buckets.values() for field in bucket}) if group == "Line" else fields
    return {label: {field: (sum(bucket[field]) / len(bucket[field]) if group == "Combo" and
                            field == "Days to Fulfil" and bucket[field] else sum(bucket[field]))
                    for field in metrics} for label, bucket in buckets.items()}


def ref_values(reference, book):
    if reference is None:
        return [], ""
    formula = getattr(reference, "f", None)
    if formula and "!" in formula:
        sheet, cells = formula.rsplit("!", 1)
        sheet = sheet.strip("'").replace("''", "'")
        if sheet in book.sheetnames:
            try:
                lo, top, hi, bottom = range_boundaries(cells)
                if None not in (lo, top, hi, bottom) and (hi-lo+1)*(bottom-top+1) <= 20000:
                    matrix = book[sheet].iter_rows(min_col=lo, max_col=hi, min_row=top, max_row=bottom, values_only=True)
                    values = [value for row in matrix for value in row]
                    if any(value is not None for value in values):
                        return values, formula
            except ValueError:
                pass
    cache = getattr(reference, "numCache", None) or getattr(reference, "strCache", None) or reference
    points = getattr(cache, "pt", [])
    if not points:
        levels = getattr(getattr(reference, "multiLvlStrCache", None), "lvl", [])
        points = levels[0].pt if levels else []
    size = max((p.idx for p in points), default=-1) + 1
    if size > 20000:
        return [], formula or ""
    result = [None] * size
    for point in points:
        result[point.idx] = point.v
    return result, formula or ""


def axis_values(source, book):
    if source is None:
        return [], ""
    return ref_values(next((getattr(source, name) for name in ("numRef", "strRef", "multiLvlStrRef", "numLit", "strLit")
                            if getattr(source, name, None) is not None), None), book)


def series_data(chart, book):
    for layer in chart._charts:
        for series in layer.series:
            title = series.title
            label = getattr(title, "v", None)
            if not label and title:
                values, _ = ref_values(title.strRef, book)
                label = next((value for value in values if value is not None), None)
            categories, category_ref = axis_values(series.cat or getattr(series, "xVal", None), book)
            numbers, number_ref = axis_values(series.val or getattr(series, "yVal", None), book)
            yield type(layer).__name__, label, categories, numbers, category_ref, number_ref


def chart_title(title):
    if title is None or title.tx is None or title.tx.rich is None:
        return ""
    return " ".join(run.t for paragraph in title.tx.rich.p for run in paragraph.r if run.t)


def metric_label(raw, metrics):
    text = key(raw).replace("fulfillment", "fulfil").replace("fulfilment", "fulfil").replace("fulfill", "fulfil")
    if text in {"fulfildays", "daystofulfil", "avgfulfildays", "averagefulfildays"} and "Days to Fulfil" in metrics:
        return "Days to Fulfil"
    candidates = [metric for metric in metrics if key(metric) in text]
    return max(candidates, key=len, default=None)


def category_key(value, group):
    if group == "Line" and isinstance(value, date):
        return str(value.year)
    if numeric(value) and value == int(value):
        return str(int(value))
    return key(value)


def compare_summary(rows, expected):
    expected_cells = {(label, metric): value for label, metrics in expected.items() for metric, value in metrics.items()}
    if not expected_cells:
        return 0, 0, []
    mismatches = []
    matches = 0
    labels = {label for label, _ in rows}
    for position, value in expected_cells.items():
        found = rows.get(position)
        if found is None and value == 0 and position[0] in labels:
            matches += 1
        elif equal(found, value):
            matches += 1
        else:
            mismatches.append(f"{position[0]} / {position[1]}: expected {format(value, ',.8f').rstrip('0').rstrip('.')}, found {found}")
    extras = set(rows) - set(expected_cells)
    mismatches.extend(f"Unexpected category / metric: {label} / {metric}" for label, metric in sorted(extras))
    return matches, len(expected_cells), mismatches


def chart_matrix(series, expected, group):
    matrix, captured, unresolved = {}, 0, 0
    metrics = {field for bucket in expected.values() for field in bucket}
    for _, label, categories, numbers, _, _ in series:
        metric = metric_label(label, metrics)
        if metric is None:
            unresolved += 1
            continue
        for category, number in zip(categories, numbers):
            if category is not None and number is not None:
                try:
                    number = float(number)
                except (TypeError, ValueError):
                    continue
                matrix[(category_key(category, group), metric)] = number
                captured += 1
    return matrix, captured, unresolved


def aggregate_tables(sheet, cached, expected, group):
    metrics = {field for bucket in expected.values() for field in bucket}
    best = (0, 0, [], "", False, [])
    if not hasattr(sheet, "_cells") or not metrics:
        return best
    pivot_bounds = [range_boundaries(p.location.ref) for p in sheet._pivots if p.location and p.location.ref]
    cells = sorted((cell.row, cell.column, cell.value) for cell in sheet._cells.values() if isinstance(cell.value, str))
    for row, col, value in cells:
        if any(l <= col <= r and t <= row <= b for l, t, r, b in pivot_bounds):
            continue
        if key(value) not in {"Bar": {"region"}, "Line": {"year", "orderyear", "orderdate", "salesyear", "years"},
                              "Combo": {"orderpriority", "priority"}}[group]:
            continue
        headings = {}
        for column in range(col+1, min(col+12, sheet.max_column)+1):
            metric = metric_label(sheet.cell(row, column).value, metrics)
            if metric:
                headings[column] = metric
        if not headings:
            continue
        matrix, formula_cells, row_names = {}, [], []
        for index in range(row+1, min(row + len(expected) + 4, sheet.max_row)+1):
            label = category_key(cached.cell(index, col).value, group)
            if label in ("", "total", "grandtotal"):
                break
            row_names.append(label)
            for column, metric in headings.items():
                cell = sheet.cell(index, column)
                detected = cached.cell(index, column).value
                if numeric(detected):
                    matrix[(label, metric)] = detected
                if cell.data_type == "f":
                    formula_cells.append(f"{cell.coordinate} {cell.value}")
        matches, total, differences = compare_summary(matrix, expected)
        if (matches, bool(formula_cells), len(matrix)) > (best[0], best[4], best[1]):
            end = get_column_letter(max(headings))
            best = (matches, len(matrix), differences, f"{sheet.title}!{get_column_letter(col)}{row}:{end}{row+len(row_names)}",
                    bool(formula_cells), formula_cells)
    return best


def pivot_source_kind(source, chart_sheet, detail):
    if source is None:
        return ""
    name = key(source.name)
    if name and any(key(table) == name for table in detail.sheet.tables):
        return "main transaction table"
    if source.ref and similar(source.sheet, detail.sheet.title):
        left, top, right, bottom = range_boundaries(source.ref)
        if None not in (left, top, right, bottom) and top == detail.header and bottom > top:
            return "main transaction range"
    if (name and any(key(table) == name for table in getattr(chart_sheet, "tables", {}))
            or source.ref and similar(source.sheet, chart_sheet.title)):
        return "transaction copy on the chart worksheet"
    return ""


def pivot_checks(sheet, cached, group, expected, detail):
    if not sheet or not getattr(sheet, "_pivots", []):
        return 0, "No PivotTable on the chart worksheet", "", "Create a PivotTable on the chart worksheet."
    best = (0, "", "", "")
    for pivot in sheet._pivots:
        cache = pivot.cache
        fields = [field.name for field in cache.cacheFields] if cache else []
        source = cache.cacheSource.worksheetSource if cache and cache.cacheSource else None
        source_text = (source.name or f"{source.sheet}!{source.ref}") if source else "Unknown"
        source_kind = pivot_source_kind(source, sheet, detail)
        values = [(fields[f.fld] if 0 <= f.fld < len(fields) else f.fld, f.subtotal, f.numFmtId)
                  for f in (pivot.dataFields or [])]
        location = pivot.location.ref
        metric_names = {field for bucket in expected.values() for field in bucket}
        requirements = ({"totalrevenue": "sum", "totalcost": "sum", "totalprofit": "sum"} if group == "Bar" else
                        {"unitssold": "sum", "daystofulfil": "average"} if group == "Combo" else {})
        settings = (all(any(similar(name, needed) and operation == op for name, operation, _ in values)
                        for needed, op in requirements.items()) if requirements else
                    any(op in {"count", "countNums"} for _, op, _ in values))
        rows = [fields[f.x] for f in (pivot.rowFields or []) if 0 <= f.x < len(fields)]
        columns = [fields[f.x] for f in (pivot.colFields or []) if 0 <= f.x < len(fields)]
        grouping = {"Bar": ("region",), "Combo": ("orderpriority",), "Line": ("saleschannel",)}[group]
        settings = settings and all(any(key(f) == name or (group == "Combo" and key(f) == "priority")
                                        for f in rows + columns) for name in grouping)
        if group == "Line":
            grouped_date = any(getattr(f, "fieldGroup", None) is not None
                               and getattr(getattr(f.fieldGroup, "rangePr", None), "groupBy", None) == "years"
                               for f in cache.cacheFields) if cache else False
            settings = settings and (grouped_date or any(key(f) in {"year", "years", "orderyear"} for f in rows + columns))
        left, top, right, bottom = range_boundaries(location)
        matrix = {}
        for column in range(left+1, min(right, left+40)+1):
            metric = next((item for row in range(top, min(top+4, bottom+1))
                           if (item := metric_label(cached.cell(row, column).value, metric_names))), None)
            if not metric:
                continue
            for row in range(top+1, min(bottom, top+len(expected)+12)+1):
                label = category_key(cached.cell(row, left).value, group)
                value = cached.cell(row, column).value
                if label in expected and numeric(value):
                    matrix[(label, metric)] = value
        if group == "Line" and columns and len(columns) >= 1:
            # Excel may put years across the columns and Sales Channel down the rows.
            for column in range(left+1, min(right, left+40)+1):
                year = next((label for row in range(top, min(top+4, bottom+1))
                             if (label := category_key(cached.cell(row, column).value, group)) in expected), None)
                if year:
                    for row in range(top+1, min(bottom, top+18)+1):
                        channel = metric_label(cached.cell(row, left).value, metric_names)
                        value = cached.cell(row, column).value
                        if channel and numeric(value):
                            matrix[(year, channel)] = value
        matches, count, differences = compare_summary(matrix, expected)
        formats = True
        if group == "Bar":
            from openpyxl.styles.numbers import BUILTIN_FORMATS
            def value_format(fmt):
                if fmt is None:
                    return ""
                if fmt < 164:
                    return BUILTIN_FORMATS.get(fmt, "")
                custom = detail.sheet.parent._number_formats
                return custom[fmt-164] if 0 <= fmt-164 < len(custom) else ""
            formats = all(currency(value_format(fmt)) for name, _, fmt in values if key(name) in requirements)
            if not formats:
                displayed = [cell for cell in sheet._cells.values()
                             if top < cell.row <= bottom and left < cell.column <= right and numeric(cell.value)]
                formats = bool(displayed) and all(currency(cell.number_format) for cell in displayed)
        score = 2 if source_kind and settings and count and matches == count and not differences and formats else 1
        evidence = f"{sheet.title}!{location} · {pivot.name}; source {source_text} ({source_kind or 'unverified source'}); rows {rows}; columns {columns}; value settings {values}; {matches}/{count} expected values matched"
        if differences:
            evidence += "\n" + "\n".join(differences[:12])
        review = "Pivot layout or displayed values could not be fully confirmed." if settings and not matrix else ""
        issues = []
        if not source_kind:
            issues.append("Use DetailTable or the copied transactions on this chart worksheet as the PivotTable source.")
        if not settings:
            issues.append("Correct the PivotTable grouping or value-field aggregation.")
        if not count or matches != count or differences:
            issues.append(f"The saved PivotTable matches {matches} of {count} expected summary values.")
        if not formats:
            issues.append("Use Currency or Accounting in the PivotTable value field settings.")
        if score > best[0]:
            best = score, evidence, review, " ".join(issues)
    return best


def chart_type_ok(chart, group):
    layers = chart._charts
    if group == "Bar":
        return len(layers) == 1 and type(chart).__name__ == "BarChart" and getattr(chart, "grouping", None) in {"clustered", None}
    if group == "Line":
        return len(layers) == 1 and type(chart).__name__ == "LineChart"
    return (len(layers) == 2 and any(type(layer).__name__ == "LineChart" for layer in layers)
            and any(type(layer).__name__ == "BarChart" and layer.type == "col"
                    and layer.grouping in {"clustered", None} for layer in layers))


def chart_candidate_items(cached, detail, group, expected, sheet, chart):
    title = CHARTS[group][0]
    target = f"{group} Chart"
    expected_count = sum(len(values) for values in expected.values())
    group_index = GROUPS.index(group)
    index = 13 + group_index * 4
    location = sheet.title if sheet is not None else target
    data = list(series_data(chart, cached)) if chart else []
    matrix, captured, unresolved = chart_matrix(data, expected, group)
    matches, count, differences = compare_summary(matrix, expected)
    chart_refs = "\n".join(f"{kind}: {label!s} · categories {categories or '(embedded)'} · values {values or '(embedded)'}"
                           for kind, label, _, _, categories, values in data)
    table = (aggregate_tables(sheet, cached[sheet.title], expected, group)
             if sheet is not None and hasattr(sheet, "_cells") else (0, 0, [], "", False, []))
    table_matches, _, table_diffs, address, formulas, examples = table
    full_table = bool(expected_count and table_matches == expected_count and not table_diffs)
    line_sorted = True
    if group == "Line":
        for _, _, categories, _, _, _ in data:
            years = [category_key(value, group) for value in categories if value is not None]
            line_sorted = line_sorted and years == sorted(years)
    aggregation = 2 if full_table and formulas and line_sorted else 1 if table_matches or matches or formulas else 0
    issues = []
    if not address:
        issues.append(f"Could not identify a separate {group.lower()} summary table on {location}.")
    elif not full_table:
        details = f" {table_diffs[0]}." if table_diffs else ""
        issues.append(f"Only {table_matches} of {expected_count} expected summary values match DetailTable.{details}")
    if address and not formulas:
        issues.append(f"No aggregation formulas were found in {address}.")
    if not line_sorted:
        issues.append("Sort the years in ascending order in the plotted source data.")
    yield rubric_item(index, aggregation, address or location, f"Formula-based {group.lower()} summary of DetailTable (task 11)",
                      f"{table_matches}/{expected_count} summary cells matched; {matches}/{expected_count} plotted cells matched; chronological order: {line_sorted}\n"
                      + "\n".join(examples[:12] + table_diffs[:12]),
                      " ".join(issues),
                      "The workbook cannot prove whether the student used filters or copied the data as instructed.")
    correct_type = bool(chart and chart_type_ok(chart, group))
    correct_data = bool(count and matches == count and not differences and not unresolved)
    # An upstream arithmetic mistake belongs to aggregation when the chart faithfully
    # plots the complete summary range with the right categories and series.
    expected_positions = {(label, metric) for label, metrics in expected.items() for metric in metrics}
    follows_summary = bool(address and expected_positions == matrix.keys() and all(
        number_ref and similar(number_ref.rsplit("!", 1)[0].strip("'"), location)
        for _, _, _, _, _, number_ref in data))
    correct_data = correct_data or follows_summary
    if correct_data and group == "Combo":
        correct_data = all((metric_label(label, {"Units Sold", "Days to Fulfil"}) == "Units Sold") ==
                           (kind == "BarChart") for kind, label, *_ in data)
    chart_score = 2 if correct_type and correct_data and chart.legend else 1 if chart and (correct_type or matches) else 0
    chart_issues = []
    if not chart:
        chart_issues.append(f"Create the {group.lower()} chart on {location}.")
    else:
        if not correct_type:
            chart_issues.append(f"Use the required {group.lower()} chart type.")
        if not correct_data:
            chart_issues.append(f"The plotted series or values differ from the required data ({matches}/{count} values match).")
        if not chart.legend:
            chart_issues.append("Add a chart legend.")
    yield rubric_item(index+1, chart_score, location, f"{group} chart with the required series, data and legend (task 11)",
                      (chart_refs or "No suitable chart") + f"\nType matches: {correct_type}; legend: {bool(chart and chart.legend)}; {captured} plotted values\n"
                      + "\n".join(differences[:12]), " ".join(chart_issues),
                      "Some chart source values could not be read from saved ranges or caches." if chart and not captured else "")
    title_issues = []
    if not sheet or not similar(sheet.title, target):
        title_issues.append(f"sheet name: expected {target}")
    if not chart or not similar(chart_title(chart.title), title):
        title_issues.append(f"chart title: {chart_title(chart.title) if chart else 'missing'}")
    for axis, label in (("x_axis", "horizontal/category axis title"), ("y_axis", "vertical/value axis title")):
        if not chart or not chart_title(getattr(getattr(chart, axis, None), "title", None)):
            title_issues.append(label)
    if group == "Bar" and chart:
        axes = [getattr(chart, axis, None) for axis in ("x_axis", "y_axis")]
        units = any(getattr(axis, "dispUnits", None) or currency(str(getattr(axis, "numFmt", "")))
                    or currency(chart_title(getattr(axis, "title", None))) for axis in axes)
        if not units:
            title_issues.append("currency or display units")
    yield rubric_item(index+2, bands(title_issues) if chart else 0, location, f'{target} · “{title}” · appropriate axis titles and units (task 11)',
                      "; ".join(title_issues) or "Required labels present",
                      "Correct the following: " + "; ".join(title_issues) + "." if chart else "Add the chart and its required labels.",
                      "Axis title wording and visual readability may need instructor review." if chart else "")
    pivot_score, pivot_evidence, review, pivot_feedback = pivot_checks(sheet, cached[sheet.title] if sheet else None, group, expected, detail)
    yield rubric_item(index+3, pivot_score, location, "PivotTable mirrors the source summary with appropriate value settings (task 12)",
                      pivot_evidence, pivot_feedback, review)



def chart_items(book, cached, detail):
    used = set()
    for group in GROUPS:
        target = f"{group} Chart"
        title = CHARTS[group][0]
        other_titles = [CHARTS[other][0] for other in GROUPS if other != group]
        other_sheets = [f"{other} Chart" for other in GROUPS if other != group]
        expected = summary(detail, group)
        metrics = {field for bucket in expected.values() for field in bucket}
        candidates = []
        for sheet in book:
            named = similar(sheet.title, target)
            other_sheet = any(similar(sheet.title, name) for name in other_sheets)
            for chart in sheet._charts:
                right_title = similar(chart_title(chart.title), title)
                other_title = any(similar(chart_title(chart.title), name) for name in other_titles)
                if id(chart) in used or other_title or (other_sheet and not right_title):
                    continue
                matching_series = chart_type_ok(chart, group) and any(
                    metric_label(label, metrics) for _, label, *_ in series_data(chart, cached))
                if named or right_title or matching_series:
                    candidates.append((sheet, chart))
            if named or (not other_sheet and not sheet._charts and sheet is not detail.sheet
                         and hasattr(sheet, "_cells")):
                candidates.append((sheet, None))

        best_rank = (-1,)
        for sheet, chart in candidates:
            items = list(chart_candidate_items(cached, detail, group, expected, sheet, chart))
            if chart is None and not similar(sheet.title, target) and not items[0]["score"]:
                continue
            rank = (sum(item["score"] for item in items), items[1]["score"], items[0]["score"],
                    bool(chart and similar(chart_title(chart.title), title)), similar(sheet.title, target))
            if rank > best_rank:
                best_rank, best_chart, best_items = rank, chart, items
        if best_rank == (-1,):
            yield from chart_candidate_items(cached, detail, group, expected, None, None)
        else:
            if best_chart is not None:
                used.add(id(best_chart))
            yield from best_items

def level(total):
    return "Level 4" if total >= 42 else "Level 3" if total >= 34 else "Level 2" if total >= 26 else "Poor"


def select_detail(book, cached):
    """Score each transaction-sheet candidate as a unit."""
    best = None
    for sheet in book.worksheets:
        detail = Detail(sheet, cached[sheet.title])
        if (len(detail.columns) < 4 or "Country" not in detail.columns
                or not {"Order ID", "Order Date", "Ship Date"} & detail.columns.keys()):
            continue
        items = setup_items(book, detail) + format_items(detail) + list(formula_items(detail))
        rank = (sum(item["score"] for item in items), items[6]["score"],
                sum(item["score"] for item in items[7:]), len(detail.columns),
                len(detail.rows), similar(sheet.title, "Detail"))
        if best is None or rank > best[0]:
            best = rank, detail, items
    if best is None:
        raise ValueError("Could not identify the assignment transaction headers.")
    return best[1], best[2]


def grade(path):
    path = Path(path).resolve()
    result = {"filename": path.name, "source_path": str(path), "version": VERSION}
    try:
        if path.suffix.lower() not in WORKBOOK_SUFFIXES:
            raise ValueError("Select an .xlsx or .xlsm workbook.")
        book, cached = load_pair(path)
        try:
            detail, items = select_detail(book, cached)
            sheet = detail.sheet
            result["detail_sheet"] = sheet.title
            items.extend(chart_items(book, cached, detail))
            total = sum(item["score"] for item in items)
            advisories = []
            if len(book.worksheets) < 5:
                advisories.append(f"Workbook has {len(book.worksheets)} regular worksheets; the instructions require at least five.")
            if (sheet.sheet_view.zoomScale or 100) != 100:
                advisories.append(f"{sheet.title}: worksheet zoom is {sheet.sheet_view.zoomScale}%, expected 100%.")
            if str(cached[sheet.title]["Z1"].value or "").strip().casefold() == "cats":
                result["marker"] = f"{sheet.title}!Z1 contains “cats”"
            if not key(path.stem).startswith("salesdata"):
                advisories.append("The instructions ask for a Sales Data filename with the student’s last name and ID.")
            result.update(items=items, total=total, level=level(total), advisories=advisories,
                          status="Review suggested" if advisories or any(item["review"] for item in items) else "Scored")
        finally:
            book.close()
            cached.close()
    except Exception as exc:
        result.update(status="Error", error=f"{type(exc).__name__}: {exc}")
    return result
