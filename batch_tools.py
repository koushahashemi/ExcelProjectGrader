"""On-demand workbook comparison and filename cleanup."""
from collections import defaultdict
from itertools import combinations
import os
from pathlib import Path
from typing import NamedTuple

from openpyxl.utils.cell import range_boundaries

from grader import WORKBOOK_SUFFIXES, day, key, load_pair, select_detail

TRANSACTION_FIELDS = ("Order ID", "Country", "Order Date", "Ship Date",
                      "Units Sold", "Unit Price", "Unit Cost")


class Snapshot(NamedTuple):
    sheet: str
    rows: dict[tuple, list[int]]
    row_count: int


def workbook_file(path):
    return path.is_file() and path.suffix.lower() in WORKBOOK_SUFFIXES and not path.name.startswith("~$")


def snapshot(path):
    book, cached = load_pair(path)
    try:
        detail, _ = select_detail(book, cached)
        sheet = detail.sheet
        rows = defaultdict(list)
        data_rows = detail.rows
        table = next((table for table in sheet.tables.values() if key(table.name) == "detailtable"), None)
        if table:
            left, top, right, bottom = range_boundaries(table.ref)
            last_row = bottom - (table.totalsRowCount or int(bool(table.totalsRowShown)))
            data_rows = [row for row in data_rows if top < row <= last_row]
        if all(field in detail.columns and
               (not table or left <= detail.columns[field] <= right) for field in TRANSACTION_FIELDS):
            for row in data_rows:
                values = [detail.value(row, field) for field in TRANSACTION_FIELDS]
                if any(value is None or value == "" for value in values):
                    continue
                dates = [day(value, book.epoch) for value in values[2:4]]
                if all(dates):
                    signature = (values[0], key(values[1]), *(value.isoformat()[:10] for value in dates), *values[4:])
                    rows[signature].append(row)
        return Snapshot(sheet.title, rows, len(data_rows))
    finally:
        book.close()
        cached.close()


def compare_workbooks(paths):
    """Flag overlaps above 80% in either direction, including reordered or copied subsets."""
    flags = {str(Path(path).resolve()): [] for path in paths}
    snapshots, errors = {}, {}
    for path in flags:
        try:
            snapshots[path] = snapshot(Path(path))
        except Exception as exc:
            errors[path] = f"{type(exc).__name__}: {exc}"
    for (left, a), (right, b) in combinations(snapshots.items(), 2):
        matches = []
        for signature in a.rows.keys() & b.rows.keys():
            matches.extend(zip(a.rows[signature], b.rows[signature]))
        matches.sort()
        count = len(matches)
        percentages = [100 * count / data.row_count if data.row_count else 0 for data in (a, b)]
        if max(percentages) <= 80:
            continue
        for index, (path, peer, own, other) in enumerate(((left, right, a, b), (right, left, b, a))):
            flags[path].append(dict(peer=peer, matching_rows=count,
                                    row_count=own.row_count, peer_row_count=other.row_count, overlap=percentages[index],
                                    peer_overlap=percentages[1-index],
                                    sheet=own.sheet, peer_sheet=other.sheet,
                                    row_pairs=[(pair[index], pair[1-index]) for pair in matches]))
    return flags, errors


def similarity_groups(results):
    """Assign one group to each connected set of directly matched submissions."""
    peers = {result["source_path"]: set() for result in results}
    for result in results:
        path = result["source_path"]
        for match in result.get("similarity", []):
            peer = match["peer"]
            if peer in peers and peer != path:
                peers[path].add(peer)
                peers[peer].add(path)

    groups = {}
    group = 0
    for path, matched in peers.items():
        if not matched or path in groups:
            continue
        group += 1
        pending = [path]
        groups[path] = group
        while pending:
            for peer in peers[pending.pop()]:
                if peer not in groups:
                    groups[peer] = group
                    pending.append(peer)
    return groups


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
