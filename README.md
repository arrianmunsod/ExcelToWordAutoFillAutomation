# Excel to Word Autofill Automation

A local Windows desktop application that creates one Word document for each checked Excel row by replacing named placeholders in a `.docx` template with values from an `.xlsx` workbook.

## Quick start

1. Install Python 3.11 or newer from https://www.python.org/downloads/windows/ and enable the Python launcher (`py`).
2. Extract this ZIP to a normal folder.
3. Double-click `run.bat`. It installs the listed dependencies and launches the app.
4. In Word, place markers such as `{{Full Name}}`, `{{Department}}`, or `{{Email Address}}` in the template.
5. Select the Word template, Excel workbook, worksheet, and output folder.
6. Review placeholder mappings, check the Excel rows you want to process, and click **Generate Word Documents**.

## Selecting Excel rows

- The app displays the worksheet data in a table.
- Check any rows you want to process, including non-consecutive rows.
- Use **Select All Rows** or **Deselect All Rows** for convenience.
- Only checked rows are used to generate Word documents.

## Placeholder behavior

- Marker names match Excel headings case-insensitively, ignoring surrounding whitespace.
- Placeholders are supported in ordinary paragraphs, table cells (including merged cells and nested tables), and section headers/footers.
- A placeholder split across Word text runs is handled. When a replacement spans differently formatted runs, the inserted value uses the first run's formatting; formatting around the placeholder is retained where possible.
- Text boxes, SmartArt, and some drawing-layer objects are not supported by `python-docx` and are not guaranteed to be processed.
- Blank cells become empty text. Dates are formatted as `YYYY-MM-DD`.
- Values stored as numeric Excel cells may not retain leading zeros. Store phone numbers and IDs as Text in Excel to preserve them.
- Source files are not modified. Existing output filenames are not overwritten; a suffix is added instead.

## Excel expectations

The first row of the selected worksheet must contain column headings. Each subsequent non-empty row is treated as one record. Rename duplicate or blank headers in Excel for best results.

## Privacy and logs

Processing is local; the app does not upload files. A `generation_log.json` file is written to the output folder. It records output filenames and errors but not the full Excel row contents.

## Tests

Run the engine tests with:

```bat
py -3 -m unittest discover -s tests -v
```

## Optional executable packaging

Install PyInstaller and run:

```bat
py -3 -m pip install pyinstaller
py -3 -m PyInstaller --noconfirm --windowed --name ExcelToWordAutofillAutomation app.py
```

The executable is created under `dist\\ExcelToWordAutofillAutomation\\` (or `dist\\ExcelToWordAutofillAutomation.exe`, depending on PyInstaller options/version). Test the packaged build on the target Windows version before distributing it.
