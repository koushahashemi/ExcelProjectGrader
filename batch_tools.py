"""On-demand workbook comparison and filename cleanup."""
from collections import defaultdict
from hashlib import file_digest
from itertools import combinations
import os
from pathlib import Path

from grader import Detail, WORKBOOK_SUFFIXES, day, detail_sheet, key, load_pair

TRANSACTION_FIELDS = ("Order ID", "Country", "Order Date", "Ship Date",
                      "Units Sold", "Unit Price", "Unit Cost")
METADATA = {"creator": "Author", "lastModifiedBy": "Last Modified By", "modified": "Modified timestamp"}


def workbook_file(path):
    return path.is_file() and path.suffix.lower() in WORKBOOK_SUFFIXES and not path.name.startswith("~$")


def snapshot(path):
    with path.open("rb") as stream:
        digest = file_digest(stream, "sha256").hexdigest()
    book, cached = load_pair(path)
    try:
        sheet = detail_sheet(book)
        detail = Detail(sheet, cached[sheet.title])
        rows = defaultdict(list)
        if all(field in detail.columns for field in TRANSACTION_FIELDS):
            for row in detail.rows:
                values = [detail.value(row, field) for field in TRANSACTION_FIELDS]
                if any(value is None or value == "" for value in values):
                    continue
                dates = [day(value, book.epoch) for value in values[2:4]]
                if all(dates):
                    signature = (values[0], key(values[1]), *(value.isoformat()[:10] for value in dates), *values[4:])
                    rows[signature].append(row)
        metadata = {label: str(value).strip() for name, label in METADATA.items()
                    if (value := getattr(book.properties, name, None)) and str(value).strip()}
        return digest, sheet.title, rows, len(detail.rows), metadata
    finally:
        book.close()
        cached.close()


def compare_workbooks(paths):
    """Flag overlaps above 90% in either direction, including reordered or copied subsets."""
    flags = {str(Path(path).resolve()): [] for path in paths}
    snapshots, errors = {}, {}
    for path in flags:
        try:
            snapshots[path] = snapshot(Path(path))
        except Exception as exc:
            errors[path] = f"{type(exc).__name__}: {exc}"
    for (left, a), (right, b) in combinations(snapshots.items(), 2):
        matches = []
        for signature in a[2].keys() & b[2].keys():
            matches.extend(zip(a[2][signature], b[2][signature]))
        matches.sort()
        count = len(matches)
        percentages = [100 * count / data[3] if data[3] else 0 for data in (a, b)]
        metadata = {name: value for name, value in a[4].items()
                    if value.casefold() == b[4].get(name, "").casefold()}
        identical = a[0] == b[0]
        if not (identical or max(percentages) > 90 or metadata):
            continue
        for index, (path, peer, own, other) in enumerate(((left, right, a, b), (right, left, b, a))):
            flags[path].append(dict(peer=peer, identical_file=identical, matching_rows=count,
                                    row_count=own[3], peer_row_count=other[3], overlap=percentages[index],
                                    peer_overlap=percentages[1-index], metadata=metadata,
                                    sheet=own[1], peer_sheet=other[1],
                                    row_pairs=[(pair[index], pair[1-index]) for pair in matches]))
    return flags, errors


def trim_filenames(paths):
    """Remove the prefix through the first underscore without overwriting files."""
    renamed, skipped = {}, []
    for path in dict.fromkeys(Path(value).absolute() for value in paths):
        if not workbook_file(path):
            continue
        if path.is_symlink():
            skipped.append(f"{path.name}: symbolic links are skipped.")
            continue
        _, separator, remainder = path.stem.partition("_")
        if not separator or not remainder.strip(" ."):
            skipped.append(f"{path.name}: no usable name after the first underscore.")
            continue
        target = path.with_name(remainder + path.suffix)
        try:
            # A hard link reserves the exact destination atomically. Windows rename already
            # rejects an existing destination and also works on drives without hard links.
            if os.name == "nt":
                path.rename(target)
            else:
                target.hardlink_to(path)
                try:
                    path.unlink()
                except OSError:
                    target.unlink()
                    raise
            renamed[str(path)] = str(target)
        except FileExistsError:
            skipped.append(f"{path.name}: {target.name} already exists.")
        except OSError as error:
            skipped.append(f"{path.name}: {error}")
    return renamed, skipped
