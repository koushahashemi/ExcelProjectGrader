from datetime import datetime
from io import BytesIO
from pathlib import Path
from shutil import copyfile
from tempfile import TemporaryDirectory
import unittest

from openpyxl import load_workbook

from batch_tools import compare_workbooks, trim_filenames
from grader import formula_matches, grade
from reports import export_report
from test_grader import submission


class TestVersion110(unittest.TestCase):
    def test_formula_references_and_lookup_modes(self):
        columns = {"Country": 2, "Order Priority": 6, "Days to Fulfil": 16}
        status = '=IF(AND(Table1[[#This Row],[Order Priority]]="C", Table1[[#This Row],[Days to Fulfil]]>14),"delayed","on time")'
        self.assertTrue(formula_matches(status, "Order Status", 2, columns))
        self.assertTrue(formula_matches('=VLOOKUP(B2,CodesTable[[COUNTRY]:[ISO CODES]],3,FALSE)', "ISO Codes", 2, columns))
        lookup = '=XLOOKUP(B2,CodesTable[COUNTRY],CodesTable[ISO CODES]'
        for arguments in (')', ',,0)', ',"missing",0,-1)'):
            self.assertTrue(formula_matches(lookup + arguments, "ISO Codes", 2, columns))
        self.assertFalse(formula_matches(lookup + ',,1)', "ISO Codes", 2, columns))

    def test_formats_and_instructor_exports(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            dodgy = root / "1_Sales Data.xlsm"
            submission(dodgy)
            book = load_workbook(BytesIO(dodgy.read_bytes()))
            detail = book["Detail"]
            detail["K2"].number_format = "0.00E+00"
            detail["Q2"].number_format = "0.00"
            detail.column_dimensions["B"].width = 3
            book.save(dodgy)
            book.close()
            result = grade(dodgy)
            self.assertNotIn("error", result, result)
            items = {item["name"]: item for item in result["items"]}
            self.assertEqual(items["Data types"]["score"], 1)
            self.assertEqual(items["Order Status"]["score"], 1)
            self.assertIn("B (Country", items["Column widths"]["feedback"])
            self.assertIn("Detail!Z1", result["marker"])
            for kind in ("csv", "txt", "json", "html", "feedback"):
                output = root / kind
                export_report(output, [result], kind)
                self.assertEqual("Detail!Z1" in output.read_text(encoding="utf-8-sig"), kind != "feedback")

    def test_comparison_and_collision_safe_trim(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            left = root / "1_Sales Data.xlsx"
            right = root / "2_Sales Data.xlsx"
            subset = root / "3_Subset.xlsx"
            submission(left)
            book = load_workbook(BytesIO(left.read_bytes()))
            detail = book["Detail"]
            for row in range(3, 13):
                for col in range(1, 18):
                    detail.cell(row, col, detail.cell(2, col).value)
                detail[f"I{row}"] = row + 1000
            book.properties.creator = "Fixture author"
            book.properties.modified = datetime(2026, 2, 1)
            detail.tables["DetailTable"].ref = "A1:Q12"
            book.save(left)
            copyfile(left, right)
            detail.tables["DetailTable"].ref = "A1:Q9"
            book.save(subset)
            book.close()
            flags, errors = compare_workbooks([left, right, subset])
            self.assertFalse(errors)
            identical, copied = flags[str(left.resolve())]
            self.assertEqual(identical["overlap"], 100)
            self.assertEqual(identical["matching_rows"], 11)
            self.assertEqual(copied["peer_overlap"], 100)
            self.assertLess(copied["overlap"], 80)
            self.assertEqual(copied["peer_row_count"], 8)
            renamed, skipped = trim_filenames([left, right, subset])
            self.assertEqual(len(renamed), 2)
            self.assertIn("already exists", skipped[0])
            self.assertTrue(right.exists())
            self.assertTrue((root / "Sales Data.xlsx").exists())
            self.assertTrue((root / "Subset.xlsx").exists())


    def test_comparison_strict_threshold_ignores_metadata(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            left, boundary, above = (root / f"{name}.xlsx" for name in ("left", "boundary", "above"))
            submission(left)
            book = load_workbook(BytesIO(left.read_bytes()))
            sheet = book["Detail"]
            for row in range(3, 12):
                for column in range(1, 18):
                    sheet.cell(row, column, sheet.cell(2, column).value)
                sheet[f"I{row}"] = row
            sheet.tables["DetailTable"].ref = "A1:Q11"
            book.properties.creator = "Shared author"
            book.save(left)
            sheet["I10"], sheet["I11"] = 2010, 2011
            book.save(boundary)
            sheet["I10"] = 10
            book.save(above)
            book.close()

            flags, errors = compare_workbooks([left, boundary])
            self.assertFalse(errors)
            self.assertFalse(any(flags.values()))
            flags, errors = compare_workbooks([left, above])
            self.assertFalse(errors)
            self.assertEqual(flags[str(left.resolve())][0]["overlap"], 90)
            self.assertEqual(flags[str(above.resolve())][0]["matching_rows"], 9)


if __name__ == "__main__":
    unittest.main()
