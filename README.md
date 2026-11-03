# Seed Runner Builder

Small desktop utility that converts Excel workbooks from a Hebrew `תרגול` folder into a standalone `SEED_RUNNER.py`.

## Run

```powershell
py -m pip install -r requirements.txt
py app.py
```

Select:

1. The folder containing the practice Excel workbooks.
2. The destination for `SEED_RUNNER.py`.

The generated runner exposes `run_seed(conn)` for inserting the parsed questions and answers into a compatible SQLite connection.

## Folder structure

Excel files may be nested in the selected folder. The immediate parent folder becomes the subject, and the workbook name becomes the file-level topic.
