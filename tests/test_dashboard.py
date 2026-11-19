from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from dashboard import POINTS, read_report, rubric_stats
from grader import grade
from reports import export_filename, export_report
from test_grader import submission


class TestDashboard(unittest.TestCase):
    def test_named_exports_round_trip(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            workbook = root / "Example.xlsx"
            submission(workbook)
            result = grade(workbook)
            self.assertNotIn("error", result, result)
            entries = []
            for kind in ("csv", "json", "html", "txt", "feedback"):
                target = root / export_filename(kind, "Section A/B")
                self.assertIn("Section A-B", target.name)
                export_report(target, [result], kind, "Section A/B")
                self.assertIn("Section A/B", target.read_text(encoding="utf-8-sig"))
                if kind in ("csv", "json"):
                    entries.extend(read_report(target))
            self.assertEqual(entries[0], entries[1])
            self.assertEqual(entries[0].set_name, "Section A/B")
            self.assertEqual(entries[0].percent(POINTS), 100 * result["total"] / 46)

    def test_legacy_reports_and_missing_scores(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            report = root / "Lab 01.csv"
            report.write_text(
                "Filename,Total /46,Status,DetailTable /2,Total Revenue /2\n"
                "A.xlsx,10,Scored,2,0\n"
                "B.xlsx,20,Scored,1,2\n"
                "Error.xlsx,,Error,,\n", encoding="utf-8")
            entries = read_report(report)
            self.assertEqual(entries[0].set_name, "Lab 01")
            self.assertIsNone(entries[2].percent(("Total Revenue",)))
            rows = rubric_stats(entries, ("Excel Table Creation", "Total Revenue"))
            self.assertEqual(rows[0], ("Total Revenue", 2, 2, 1, 0, 2, 1, 0, 1, 50, 2))
            legacy = root / "Lab 02.json"
            legacy.write_text('[{"filename":"C.xlsx","total":30,"items":[]}]', encoding="utf-8")
            self.assertEqual(read_report(legacy)[0].set_name, "Lab 02")
            self.assertIsNone(read_report(legacy)[0].percent(("Total Revenue",)))
