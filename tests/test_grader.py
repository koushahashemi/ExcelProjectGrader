import json
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.pivot.cache import CacheDefinition, CacheField, CacheSource, WorksheetSource
from openpyxl.pivot.table import DataField, Location, RowColField, TableDefinition
from openpyxl.worksheet.table import Table

from grader import COUNTRY_ROWS, MAXIMUM, RUBRIC, formula_matches, grade
from reports import benchmark, export_report, html_report


def submission(path):
    book = Workbook()
    detail = book.active
    detail.title = "Detail"
    detail.append(["Region", "Country", "ISO Codes", "Item Type", "Sales Channel", "Order Priority",
                   "Order Date", "Ship Date", "Order ID", "Units Sold", "Unit Price", "Unit Cost",
                   "Total Revenue", "Total Cost", "Total Profit", "Days to Fulfil", "Order Status"])
    detail.append(["Asia", "Afghanistan", "=VLOOKUP([@Country],CodesTable,3,FALSE)", "Fruit", "Online",
                   "C", 46000, 46020, 1234, 10, 25, 5, "=J2*K2", "=PRODUCT(J2,L2)", "=M2-N2",
                   "=DAYS(H2,G2)", '=IF(AND(P2>14,F2="C"),"Delayed","On time")'])
    detail["Z1"] = "cats"
    for cell in detail[1][:17]:
        detail.column_dimensions[cell.column_letter].width = 24
    for addr in ("G2", "H2"):
        detail[addr].number_format = "yyyy-mm-dd"
    for addr in ("K2", "L2", "M2", "N2", "O2"):
        detail[addr].number_format = '"$"#,##0.00'
    detail.add_table(Table(displayName="DetailTable", ref="A1:Q2"))
    book.create_sheet("Parameters").sheet_state = "hidden"
    codes = book.create_sheet("CountryCodes")
    codes.append(["COUNTRY", "COUNTRY CODE", "ISO CODES", "POPULATION", "AREA KM2", "GDP $USD"])
    for row in COUNTRY_ROWS:
        codes.append(row)
    codes.add_table(Table(displayName="CodesTable", ref=f"A1:F{len(COUNTRY_ROWS)+1}"))
    bar_sheet = book.create_sheet("Bar Chart")
    bar_sheet.append(["Region", "Total Revenue", "Total Cost", "Total Profit"])
    bar_sheet.append(["Asia", 250, 50, 200])
    bar = BarChart()
    bar.title = "Financial Performance by Region"
    bar.x_axis.title = "Region"
    bar.y_axis.title = "USD ($)"
    bar.add_data(Reference(bar_sheet, min_col=2, max_col=4, min_row=1, max_row=2), titles_from_data=True)
    bar.set_categories(Reference(bar_sheet, min_col=1, min_row=2, max_row=2))
    bar_sheet.add_chart(bar, "H3")
    book.create_sheet("Line Chart")
    book.create_sheet("Combo Chart")
    book.save(path)


class TestGrader(unittest.TestCase):
    def test_rubric_and_formula_variants(self):
        self.assertEqual((len(RUBRIC), MAXIMUM), (25, 46))
        columns = {"Units Sold": 10, "Unit Price": 11}
        self.assertTrue(formula_matches("=PRODUCT([@[Unit Price]],[@[Units Sold]])", "Total Revenue", 2, columns))
        self.assertFalse(formula_matches("=J3*K2", "Total Revenue", 2, columns))
        columns = {"Order Priority": 6, "Days to Fulfil": 16}
        self.assertFalse(formula_matches('=IF(AND(F2="C",P2>=14),"Delayed","On time")', "Order Status", 2, columns))

    def test_saved_workbook_and_chart(self):
        with TemporaryDirectory() as root:
            path = Path(root, "Sales Data Fixture 123.xlsx")
            submission(path)
            outcome = grade(path)
            self.assertNotIn("error", outcome, outcome)
            self.assertEqual([item["score"] for item in outcome["items"][:13]], [1]*4 + [2]*9)
            self.assertEqual([item["score"] for item in outcome["items"][13:17]], [1, 2, 2, 0])
            self.assertEqual(outcome["total"], 27)

    def test_line_combo_and_pivot(self):
        with TemporaryDirectory() as root:
            path = Path(root, "Sales Data Full Charts.xlsx")
            submission(path)
            book = load_workbook(BytesIO(path.read_bytes()))
            line_sheet = book["Line Chart"]
            line_sheet.append(["Year", "Online"])
            line_sheet.append([2025, 1])
            line = LineChart()
            line.title = "Annual Sales by Channel"
            line.x_axis.title = "Year"
            line.y_axis.title = "Orders"
            line.add_data(Reference(line_sheet, min_col=2, min_row=1, max_row=2), titles_from_data=True)
            line.set_categories(Reference(line_sheet, min_col=1, min_row=2, max_row=2))
            line_sheet.add_chart(line, "H2")

            combo_sheet = book["Combo Chart"]
            combo_sheet.append(["Order Priority", "Units Sold", "Average Days to Fulfill"])
            combo_sheet.append(["C", 10, 20])
            columns = BarChart()
            columns.title = "Units Sold and Fulfillment Time by Priority"
            columns.x_axis.title = "Order Priority"
            columns.y_axis.title = "Units sold"
            days = LineChart()
            days.y_axis.axId = 200
            days.y_axis.title = "Days"
            for chart, col in ((columns, 2), (days, 3)):
                chart.add_data(Reference(combo_sheet, min_col=col, min_row=1, max_row=2), titles_from_data=True)
                chart.set_categories(Reference(combo_sheet, min_col=1, min_row=2, max_row=2))
            columns += days
            combo_sheet.add_chart(columns, "H2")

            bar_sheet = book["Bar Chart"]
            for row, data in enumerate((["Region", "Sum of Total Revenue", "Sum of Total Cost", "Sum of Total Profit"],
                                        ["Asia", 250, 50, 200]), 5):
                for column, value in enumerate(data, 1):
                    bar_sheet.cell(row, column, value)
            pivot = TableDefinition(name="RegionTotals", cacheId=1, dataCaption="Values",
                                    location=Location(ref="A5:D6", firstHeaderRow=1, firstDataRow=1, firstDataCol=1),
                                    rowFields=[RowColField(x=0)],
                                    dataFields=[DataField(fld=i, numFmtId=7) for i in (1, 2, 3)])
            pivot.cache = CacheDefinition(cacheSource=CacheSource(type="worksheet", worksheetSource=WorksheetSource(
                                          name="DetailTable")), cacheFields=[CacheField(name=name) for name in
                                          ("Region", "Total Revenue", "Total Cost", "Total Profit")])
            bar_sheet.add_pivot(pivot)
            book.save(path)
            outcome = grade(path)
            self.assertNotIn("error", outcome, outcome)
            for index in (16, 18, 19, 22, 23):
                self.assertEqual(outcome["items"][index]["score"], 2, outcome["items"][index])

            book = load_workbook(BytesIO(path.read_bytes()))
            bar_sheet = book["Bar Chart"]
            for row, data in enumerate((["Region", "Total Revenue", "Total Cost", "Total Profit"],
                                        ["Asia", 250, 50, 200]), 10):
                for column, value in enumerate(data, 1):
                    bar_sheet.cell(row, column, value)
            pivot = bar_sheet._pivots[0]
            pivot.cache.cacheSource.worksheetSource = WorksheetSource(ref="A10:D11", sheet="Bar Chart")
            book.save(path)
            book.close()
            outcome = grade(path)
            local_pivot = outcome["items"][16]
            self.assertEqual(local_pivot["score"], 2, local_pivot)
            self.assertIn("transaction copy on the chart worksheet", local_pivot["detected"])

    def test_grades_more_complete_detail_and_chart(self):
        with TemporaryDirectory() as root:
            path = Path(root, "Sales Data Duplicate.xlsx")
            submission(path)
            book = load_workbook(BytesIO(path.read_bytes()))
            book["Detail"].title = "Completed transactions"
            draft = book.copy_worksheet(book["Completed transactions"])
            draft.title = "Detail"
            for col in ("M", "N", "O", "P", "Q"):
                draft[f"{col}2"] = None
            old_bar = book["Bar Chart"]
            old_bar._charts.clear()
            incomplete = BarChart()
            incomplete.title = "Financial Performance by Region"
            old_bar.add_chart(incomplete, "H1")
            finished = book.create_sheet("Question 11a")
            finished.append(["Region", "Total Revenue", "Total Cost", "Total Profit"])
            finished.append(["Asia", 250, 50, 200])
            bar = BarChart()
            bar.title = "Financial Performance by Region"
            bar.x_axis.title = "Region"
            bar.y_axis.title = "USD ($)"
            bar.add_data(Reference(finished, min_col=2, max_col=4, min_row=1, max_row=2), titles_from_data=True)
            bar.set_categories(Reference(finished, min_col=1, min_row=2, max_row=2))
            finished.add_chart(bar, "H1")
            book.save(path)
            book.close()

            result = grade(path)
            self.assertEqual(result.get("detail_sheet"), "Completed transactions", result)
            self.assertEqual(result["items"][7]["score"], 2)
            self.assertEqual(result["items"][14]["score"], 2)
            self.assertEqual(result["items"][14]["location"], "Question 11a")

    def test_html_lists_similarity_groups_after_summary(self):
        results = [{"filename": name, "source_path": name, "status": "Error", "error": "Fixture"}
                   for name in ("A&.xlsx", "B.xlsx")]
        self.assertNotIn("Similar submissions", html_report(results))
        from unittest.mock import patch

        for result, peer in zip(results, reversed(results)):
            result["similarity"] = [{"peer": peer["source_path"]}]
        with patch("reports.flag_details", return_value=""):
            page = html_report(results)
        self.assertLess(page.index("Class summary"), page.index("Similar submissions"))
        self.assertLess(page.index("Similar submissions"), page.index('class="student"'))
        group_table = page.split('<section class="similarity">')[1].split("</section>")[0]
        self.assertIn('Group 1', group_table)
        self.assertIn('<a href="#s0">A&amp;.xlsx</a><br><a href="#s1">B.xlsx</a>', group_table)

    def test_exports_and_benchmark(self):
        with TemporaryDirectory() as root:
            root = Path(root)
            valid = root / "fixture.xlsx"
            submission(valid)
            broken = root / "=broken.xlsx"
            broken.write_text("not a workbook")
            results = [grade(broken), grade(valid)]
            self.assertEqual(results[0]["status"], "Error")
            self.assertIn("total", results[1])
            for kind in ("html", "txt", "json", "csv", "feedback"):
                path = root / kind
                export_report(path, results, kind)
                self.assertGreater(path.stat().st_size, 0)
            self.assertIn("'=broken.xlsx", (root / "csv").read_text(encoding="utf-8-sig"))
            self.assertNotIn("BadZipFile", (root / "feedback").read_text(encoding="utf-8"))
            self.assertIn("<details>", (root / "html").read_text(encoding="utf-8"))
            report = (root / "html").read_text(encoding="utf-8")
            self.assertIn("<th>Added sentence</th>", report)
            self.assertIn('class="badge flag-yes">Yes</span>', report)
            self.assertIn('class="badge flag-no">No</span>', report)
            self.assertIn("Hidden-instruction sentence: <b>Yes</b>", report)
            self.assertIn("Hidden-instruction sentence: <b>No</b>", report)
            self.assertNotIn("cats", report.split('class="copy-text" hidden>')[-1].split("</pre>")[0])
            self.assertIn("submissions", json.loads((root / "json").read_text(encoding="utf-8")))
            reference = root / "benchmark.json"
            reference.write_text(json.dumps([{"workbook": valid.name, "scores": dict(RUBRIC), "notes": "Fixture"}]))
            comparison = benchmark(reference, grade)
            self.assertEqual(comparison[0]["instructor_total"], 46)
            self.assertTrue(comparison[0]["differences"])


if __name__ == "__main__":
    unittest.main()
