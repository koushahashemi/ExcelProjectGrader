# Excel Project Grader

Windows desktop grader for the BSYS-1000 Individual Project (Excel), Fall 2026. It evaluates 25 rubric items out of 46 and assigns the rubric’s overall level.

## Windows

Download `ExcelProjectGrader.exe` from GitHub Releases and open it. The standalone build bundles Python and its dependencies. If you also download `Install and Run.bat`, place both files together and open the batch file to launch the app.

From a source checkout, install Python 3.12 including Tcl/Tk and the Python launcher, then run `Install and Run.bat`. It creates a local virtual environment and installs dependencies. Run `Build Windows EXE.bat` on Windows to build `dist\ExcelProjectGrader.exe`. GitHub Actions builds an artifact on main and publishes release assets on version tags. The executable is unsigned.

## Grading

Drop .xlsx or .xlsm workbooks into the window, select individual or multiple files, or choose a folder. Select a submission to see rubric scores and feedback. Select a rubric item for technical evidence. Double-click the filename to open the original workbook. Right-click a submission for **Open file**, **Remove from list**, or **Regrade**. Regrading replaces that row’s result and clears its previous similarity matches. Grading reads the workbooks without launching Excel, executing macros, updating external links, or saving changes to submissions.

Export the class grade table as CSV. Save the student feedback separately as text. HTML, text and JSON instructor reports include technical evidence and may contain student data or local file paths.

The checks use the Fall 2026 Excel instructions and rubric. The bundled reference is the country-code data extracted from the supplied starter. The rubric’s single-point items use 0 or 1 and the other items use 0, 1 or 2. Separate advisories cover worksheet count, Detail zoom and filename prefix. A saved `cats` value in Z1 on the transaction worksheet is highlighted and recorded in instructor exports. The workbook does not establish identity, lateness, original macro execution, or whether filters and copying were used. No separate Excel style guide was supplied.

Excel’s saved data, formulas, metadata and caches are used as evidence. Common equivalent arithmetic formulas are recognized and valid alternatives with usable saved results can be compared numerically. Formula engines, visual rendering and every PivotTable or chart layout are outside the reader’s coverage. Widths are approximate. Save recalculated submissions in Excel when values are missing or stale. Check cases marked for review using the locations and evidence shown.

## Batch utilities

**Check for similar submissions** runs on demand against the current batch. It compares complete source transaction signatures independently of row order. A pair is flagged if more than 80% of the transaction rows in either detail table match. The report shows percentages in both directions and the matching row locations. Flags are for instructor review and carry zero points. Student feedback excludes these flags and their evidence.

**Trim filenames** opens a dialog for selecting workbooks or a folder. It removes the prefix through the first underscore, preserves the extension and skips collisions. Review the dialog’s result for any skipped files.

## Results dashboard

Enter a **Set name** before exporting a batch. Each export includes the set name in its suggested filename and in the report. Use the same set name when exporting CSV and JSON from the same batch.

Open **Results dashboard** from the grader or run `Open Results Dashboard.bat` next to `Install and Run.bat`. You can also launch it without opening the grader with `ExcelProjectGrader.exe --dashboard` or `python dashboard.py`.

Load several grader CSV or JSON reports together. Filter by set, category, or individual rubric item. The score cards and set comparison use percentages for the selected items. The rubric table ranks items by the share of graded submissions below full marks and shows full, partial and zero counts. The dashboard shows charts and tables side by side on one page. The score histogram, stacked rubric outcomes, set-by-rubric heatmap, box plots and points-lost chart all follow the active filters. Use the zoom controls or Ctrl+wheel to resize the page. Use the scrollbars or Shift+wheel to move horizontally, and Fit width to see both columns. Missing scores and grading errors are excluded from averages. Reports without a set name use the report filename, and duplicate submission names within a set count once.

## Development and calibration

Use the existing isolated cloud checkout. From the repository root run:

```sh
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python desktop.py --json batch.json submission1.xlsx submission2.xlsx
```

For calibration, keep instructor benchmark workbooks locally. Create a JSON array of entries with `workbook` (a path relative to the JSON file), `scores` (an object keyed by the 25 exact rubric names from the CSV headers, without the point suffixes), and optional `notes`. Run `python desktop.py --benchmark benchmarks.json --json comparison.json`. The comparison records differences by item, instructor notes, and automatic evidence to guide rule changes and regression checks. It does not apply benchmark scores to unrelated submissions.
