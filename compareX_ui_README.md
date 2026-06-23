# Excel Diff Checker — UI Version

This is the UI version of `diff checker.py`. It keeps the exact same comparison logic and adds a drag-and-drop interface plus column toggles.

## Files

- `diff_checker_ui.py` — main UI application
- `run_diff_checker_ui.bat` — double-click this to launch the UI
- `requirements_diff_checker_ui.txt` — Python dependencies

## How to run

1. Double-click `run_diff_checker_ui.bat`
2. Drag and drop your **Old/Prod** file onto the left drop zone
3. Drag and drop your **New/Dev** file onto the right drop zone
4. Use the checkboxes to include/exclude columns from comparison
5. Choose the output Excel path
6. Adjust the row match threshold if needed (default `0.65`)
7. Click **Compare & Generate Report**

## Notes

- Supported file formats: `.xlsx`, `.xls`, `.csv`
- The `created_timestamp` column is unchecked by default (matches the original script)
- The output report highlights:
  - **Yellow** — modified cells (values are shown as `old -> new`)
  - **Green** — added rows
  - **Red** — deleted rows
- After a successful comparison, click **Open output folder** to jump straight to the saved report
- If drag-and-drop does not work, make sure `tkinterdnd2` is installed (it is already included in the local `.venv`)
