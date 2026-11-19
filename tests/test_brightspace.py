"""Exercise the read-only importer without a Brightspace tenant."""
from io import BytesIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from brightspace import Brightspace, main
from test_grader import submission


class TestBrightspace(unittest.TestCase):
    def test_import_one_assignment_and_grade_workbook(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            fixture = root / "fixture.xlsx"
            submission(fixture)
            page = {"Items": [{"Entity": {"EntityId": 3, "EntityType": "User"},
                              "Submissions": [{"Id": 4, "Files": [
                                  {"FileId": 5, "FileName": "Sales Data Fixture.xlsx"},
                                  {"FileId": 6, "FileName": "ignored.pdf"}]}]}],
                    "PagingInfo": {"HasMoreItems": True, "Bookmark": "next"}}
            routes = []

            def fake_open(request, timeout):
                routes.append((request.full_url, request.get_method()))
                self.assertEqual(request.get_header("Authorization"), "Bearer fixture-token")
                if request.full_url.endswith("/submissions/4/files/5"):
                    return BytesIO(fixture.read_bytes())
                if request.full_url.endswith("?bookmark=next"):
                    return BytesIO(b'{"Items":[],"PagingInfo":{"HasMoreItems":false}}')
                return BytesIO(json.dumps(page).encode())

            output = root / "output"
            arguments = ["brightspace.py", "--tenant", "https://test.example", "--api-version", "1.82",
                         "--course-id", "7", "--assignment-id", "8", "--download-dir", str(output)]
            with patch("brightspace.build_opener") as opener, patch("sys.argv", arguments), patch.dict(
                    "os.environ", {"BRIGHTSPACE_ACCESS_TOKEN": "fixture-token"}):
                opener.return_value.open.side_effect = fake_open
                self.assertEqual(main(), 0)
            report = json.loads((output / "grades.json").read_text(encoding="utf-8"))
            record, = report["submissions"]
            self.assertEqual(record["total"], 27)
            self.assertEqual(record["brightspace"]["entity_id"], 3)
            self.assertEqual(len(routes), 3)
            self.assertTrue(all(method == "GET" for _, method in routes))
            self.assertTrue(all("/7/dropbox/folders/8/submissions/" in url for url, _ in routes))

    def test_rejects_untrusted_routes(self):
        for tenant in ("http://tenant.example", "https://tenant.example/path", "https://user@tenant.example"):
            with self.subTest(tenant=tenant), self.assertRaises(ValueError):
                Brightspace(tenant, "1.82", 7, 8, "fixture-token")
