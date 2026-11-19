"""Read-only pilot for importing one Brightspace assignment."""
import argparse
from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
import shutil
import sys
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

from grader import WORKBOOK_SUFFIXES, grade
from reports import export_report


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # Keep the instructor's bearer token on the explicitly configured tenant.
        return None


@dataclass(frozen=True)
class Attachment:
    submission_id: int
    file_id: int
    filename: str
    entity_id: int | None
    entity_type: str | None


def identifier(value):
    raw = str(value)
    if not re.fullmatch(r"[1-9][0-9]*", raw):
        raise ValueError("Brightspace IDs must be positive integers.")
    return int(raw)


class Brightspace:
    def __init__(self, tenant, version, course_id, assignment_id, token):
        origin = urlsplit(tenant)
        if (origin.scheme != "https" or not origin.hostname or origin.username or origin.password
                or origin.path not in ("", "/") or origin.query or origin.fragment):
            raise ValueError("Use the HTTPS Brightspace tenant origin, without a course path.")
        if not re.fullmatch(r"[0-9]+\.[0-9]+", version):
            raise ValueError("Use a Brightspace LE API version such as 1.82.")
        self.origin = f"https://{origin.netloc}"
        self.route = f"/d2l/api/le/{version}/{identifier(course_id)}/dropbox/folders/{identifier(assignment_id)}/submissions/"
        self.token = token
        self.opener = build_opener(NoRedirect())

    def request(self, route):
        request = Request(self.origin + route, headers={"Authorization": f"Bearer {self.token}"}, method="GET")
        try:
            return self.opener.open(request, timeout=30)
        except HTTPError as error:
            code = error.code
            error.close()
            raise ValueError(f"Brightspace returned HTTP {code}. Check the API version, IDs and authorization.") from None
        except URLError as error:
            raise ValueError(f"Could not reach the Brightspace tenant: {error.reason}") from None

    def files(self):
        results = {}
        bookmark = None
        seen_bookmarks = set()
        while True:
            query = "?" + urlencode({"bookmark": bookmark}) if bookmark else ""
            with self.request(self.route + query) as response:
                page = json.load(response)
            entities = page if isinstance(page, list) else page["Items"]
            for entry in entities:
                entity = entry.get("Entity") or {}
                for submission in entry.get("Submissions", []):
                    sid = identifier(submission["Id"])
                    for file in submission.get("Files", []):
                        filename = file.get("FileName") or file.get("Name") or ""
                        if Path(filename).suffix.lower() not in WORKBOOK_SUFFIXES:
                            continue
                        fid = identifier(file["FileId"] if "FileId" in file else file["Id"])
                        results[sid, fid] = Attachment(sid, fid, filename, entity.get("EntityId"), entity.get("EntityType"))
            paging = page.get("PagingInfo", {}) if isinstance(page, dict) else {}
            if not paging.get("HasMoreItems"):
                return list(results.values())
            bookmark = paging.get("Bookmark")
            if not bookmark or bookmark in seen_bookmarks:
                raise ValueError("Brightspace did not return a new continuation bookmark.")
            seen_bookmarks.add(bookmark)

    def download(self, attachment, directory):
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", attachment.filename)
        if re.fullmatch(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])", Path(safe).stem, re.I):
            safe = "_" + safe
        folder = directory / f"{attachment.submission_id}-{attachment.file_id}"
        folder.mkdir(mode=0o700, exist_ok=True)
        path = folder / safe[-150:]
        route = f"{self.route}{attachment.submission_id}/files/{attachment.file_id}"
        with self.request(route) as response, path.open("xb") as stream:
            try:
                shutil.copyfileobj(response, stream)
            except Exception:
                stream.close()
                path.unlink()
                raise
        return path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tenant", required=True)
    parser.add_argument("--api-version", required=True)
    parser.add_argument("--course-id", required=True)
    parser.add_argument("--assignment-id", required=True)
    parser.add_argument("--download-dir", type=Path, help="New local directory for downloaded workbooks and grades.")
    args = parser.parse_args()
    token = os.environ.get("BRIGHTSPACE_ACCESS_TOKEN", "").strip()
    if not token:
        parser.exit(2, "BRIGHTSPACE_ACCESS_TOKEN is missing. Use an authorized OAuth access token.\n")
    try:
        client = Brightspace(args.tenant, args.api_version, args.course_id, args.assignment_id, token)
        attachments = client.files()
        print(f"Found {len(attachments)} Excel attachment(s) across {len({item.submission_id for item in attachments})} submission(s).")
        if args.download_dir:
            args.download_dir.mkdir(mode=0o700)
            results = []
            for item in attachments:
                try:
                    result = grade(client.download(item, args.download_dir))
                    result["filename"] = item.filename
                except (OSError, ValueError) as error:
                    result = {"filename": item.filename, "status": "Error", "error": str(error)}
                result["brightspace"] = {"submission_id": item.submission_id, "file_id": item.file_id,
                                         "entity_id": item.entity_id, "entity_type": item.entity_type}
                results.append(result)
            export_report(args.download_dir / "grades.json", results, "json")
            errors = sum("error" in result for result in results)
            print(f"Wrote grades.json with {len(results) - errors} graded and {errors} error(s).")
            return 1 if errors else 0
    except (OSError, ValueError, TypeError, KeyError) as error:
        print(f"Brightspace import failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
