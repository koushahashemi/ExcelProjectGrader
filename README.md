# Excel Project Grader

Windows desktop grader for the BSYS-1000 Individual Project (Excel), Fall 2026. It evaluates 25 rubric items out of 46 and assigns the rubric’s overall level.

## Windows

Download `ExcelProjectGrader.exe` from GitHub Releases and open it. The standalone build bundles Python and its dependencies. If you also download `Install and Run.bat`, place both files together and open the batch file to launch the app.

From a source checkout, install Python 3.12 including Tcl/Tk and the Python launcher, then run `Install and Run.bat`. It creates a local virtual environment and installs dependencies. Run `Build Windows EXE.bat` on Windows to build `dist\ExcelProjectGrader.exe`. GitHub Actions builds an artifact on main and publishes release assets on version tags. The executable is unsigned.

## Grading

Drop .xlsx or .xlsm workbooks into the window, select individual or multiple files, or choose a folder. Select a submission to see rubric scores and feedback. Select a rubric item for technical evidence. Double-click the filename to open the original workbook. Grading reads the workbooks without launching Excel, executing macros, updating external links, or saving changes to submissions.

Export the class grade table as CSV. Save the student feedback separately as text. HTML, text and JSON instructor reports include technical evidence and may contain student data or local file paths.

The checks use the Fall 2026 Excel instructions and rubric. The bundled reference is the country-code data extracted from the supplied starter. The rubric’s single-point items use 0 or 1 and the other items use 0, 1 or 2. Separate advisories cover worksheet count, Detail zoom and filename prefix. A saved `cats` value in Z1 on the transaction worksheet is highlighted and recorded in instructor exports. The workbook does not establish identity, lateness, original macro execution, or whether filters and copying were used. No separate Excel style guide was supplied.

Excel’s saved data, formulas, metadata and caches are used as evidence. Common equivalent arithmetic formulas are recognized and valid alternatives with usable saved results can be compared numerically. Formula engines, visual rendering and every PivotTable or chart layout are outside the reader’s coverage. Widths are approximate. Save recalculated submissions in Excel when values are missing or stale. Check cases marked for review using the locations and evidence shown.

## Batch utilities

**Check for similar submissions** runs on demand against the current batch. It compares complete source transaction signatures independently of row order. A match above 90% of either file is flagged, with percentages in both directions and the matching row locations. Identical file contents and shared nonblank Author, Last Modified By or modification timestamp are separate signals. Metadata can come from the common starter or a shared machine. Flags are for instructor review and carry zero points. Student feedback excludes these flags and their evidence.

**Trim filenames** opens a dialog for selecting workbooks or a folder. It removes the prefix through the first underscore, preserves the extension and skips collisions. Review the dialog’s result for any skipped files.

## Development and calibration

Use the existing isolated cloud checkout. From the repository root run:

```sh
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python desktop.py --json batch.json submission1.xlsx submission2.xlsx
```

For calibration, keep instructor benchmark workbooks locally. Create a JSON array of entries with `workbook` (a path relative to the JSON file), `scores` (an object keyed by the 25 exact rubric names from the CSV headers, without the point suffixes), and optional `notes`. Run `python desktop.py --benchmark benchmarks.json --json comparison.json`. The comparison records differences by item, instructor notes, and automatic evidence to guide rule changes and regression checks. It does not apply benchmark scores to unrelated submissions.
