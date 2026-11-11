# Excel to Word Autofill Automation

A Python desktop application that automates the generation of Microsoft Word documents using data from Microsoft Excel.

The application allows users to select an Excel workbook, choose specific worksheet rows, map Excel column headers to Word placeholders, and automatically generate individual Word documents using a predefined `.docx` template.

## Features

- **Excel Integration** – Load Excel workbooks (`.xlsx`) and select worksheets.
- **Word Template Integration** – Use existing Word documents containing customizable placeholders.
- **Automatic Placeholder Detection** – Detect placeholders enclosed in double curly braces, such as `{{Full Name}}`.
- **Column Mapping** – Automatically match Word placeholders with Excel column headers.
- **Manual Row Selection** – Choose exactly which Excel records should be processed.
- **Bulk Document Generation** – Generate multiple Word documents in one operation.
- **Select All / Deselect All** – Quickly control which Excel rows are included.
- **Output Folder Selection** – Choose where generated documents are saved.
- **Duplicate Filename Protection** – Automatically generate unique filenames instead of overwriting existing documents.
- **Project Configuration** – Save and load placeholder mapping configurations.
- **Generation Logging** – Record successful document generation and errors.
- **Formatting Preservation** – Retain Word document formatting, tables, headers, and footers where supported.

## Application Preview

The application follows a simple workflow:

1. Select a Word template.
2. Select an Excel workbook.
3. Choose the worksheet.
4. Review placeholder mappings.
5. Select the Excel rows to process.
6. Choose the output folder.
7. Generate the Word documents.

## Technologies Used

| Technology | Purpose |
|---|---|
| Python | Main programming language |
| PySide6 | Desktop graphical user interface |
| openpyxl | Reading Excel workbooks |
| python-docx | Reading and generating Word documents |
| JSON | Saving application configurations and generation logs |

## Project Structure

```text
ExcelToWordAutofillAutomation/
│
├── app.py
├── engine.py
├── requirements.txt
├── README.md
├── run.bat
│
├── tests/
│   └── test_engine.py
│
├── sample_excel_data.xlsx
└── sample_word_template.docx
```

### File Descriptions

**app.py**

Contains the main application interface and GUI functionality.

Responsibilities include:
- Creating the application window and interface.
- Selecting Excel workbooks and Word templates.
- Displaying worksheet data in a selectable table.
- Managing manual Excel row selection.
- Displaying and managing placeholder mappings.
- Managing output folder selection.
- Running document generation in a background thread.
- Displaying progress and generation results.

**engine.py**

Contains the core document processing functionality.

Responsibilities include:
- Reading Excel worksheets using openpyxl.
- Extracting column headers and worksheet records.
- Detecting placeholders in Word documents.
- Matching Excel values with Word placeholders.
- Replacing placeholders in Word paragraphs and tables.
- Processing document headers and footers.
- Formatting Excel values for insertion into Word.
- Generating safe filenames.
- Preventing duplicate output filenames.

**requirements.txt**

Contains the Python libraries required to run the application.

**run.bat**

A Windows batch file that installs the required dependencies and launches the application.

**tests/**

Contains automated tests for the document generation engine.

## Requirements

- Windows 10 or Windows 11
- Python 3.11 or newer
- Microsoft Excel or another application capable of editing `.xlsx` files
- Microsoft Word or another application capable of opening `.docx` files

## Installation

### 1. Install Python

Download Python from:

https://www.python.org/downloads/windows/

During installation, enable the Python launcher (`py`).

### 2. Download the Project

Clone the repository:

```bash
git clone https://github.com/YOUR-USERNAME/ExcelToWordAutofillAutomation.git
```

Alternatively, download the repository as a ZIP file and extract it.

### 3. Install Dependencies

Open Command Prompt inside the project folder.

Run:

```bash
py -3 -m pip install -r requirements.txt
```

### 4. Launch the Application

Double-click:

```text
run.bat
```

Alternatively, launch it manually:

```bash
py -3 app.py
```

## Preparing the Word Template

Open Microsoft Word and create a document containing placeholders.

For example:

```text
EMPLOYEE INFORMATION

Full Name: {{Full Name}}

Department: {{Department}}

Position: {{Position}}

Email Address: {{Email Address}}

Contact Number: {{Contact Number}}
```

The placeholder names should correspond to the column headings in your Excel workbook.

Save the template as a `.docx` file.

## Preparing the Excel Workbook

Create an Excel worksheet with column headers in the first row.

Example:

| Full Name | Department | Position | Email Address |
|---|---|---|---|
| John Smith | IT | Technician | john@example.com |
| Jane Doe | HR | Manager | jane@example.com |
| Mark Wilson | Finance | Analyst | mark@example.com |

Each subsequent non-empty row represents one record.

**Important:** Store identification numbers, telephone numbers, and other values requiring leading zeros as Text in Excel.

## Using the Application

### Step 1 – Select the Word Template

Click **Choose Word Template** and select your `.docx` file.

### Step 2 – Select the Excel Workbook

Click **Choose Excel Workbook** and select your `.xlsx` file.

### Step 3 – Choose the Worksheet

Select the worksheet containing your data.

The application will display the worksheet records in a table.

### Step 4 – Review Placeholder Mappings

The application automatically detects placeholders in the Word template and attempts to match them with Excel column headers.

Review the mappings and make adjustments if necessary.

### Step 5 – Select Excel Rows

Use the checkboxes in the Excel preview table.

You can:
- Select individual rows.
- Select multiple rows.
- Select non-consecutive rows.
- Select all rows.
- Deselect all rows.

Only checked rows will be processed.

### Step 6 – Choose the Output Folder

Select the destination folder where the generated Word documents will be saved.

### Step 7 – Generate Word Documents

Click **Generate Word Documents**.

The application will process each selected Excel row and create an individual Word document using the selected template.

## How the Application Works

The application uses two main Python modules.

### Application Layer – app.py

The GUI is developed using PySide6.

It handles user interactions, file selection, worksheet previews, placeholder mapping, row selection, and generation progress.

Document generation runs in a separate QThread worker to prevent the interface from freezing during processing.

### Processing Layer – engine.py

The document processing engine uses openpyxl to read Excel data and python-docx to manipulate Word documents.

The engine extracts worksheet records, detects placeholders, maps Excel values, and replaces matching placeholders in the Word template.

It supports placeholders inside:
- Normal paragraphs.
- Table cells.
- Nested tables.
- Headers and footers.

Placeholders split across multiple Word text runs are also handled.

Text boxes, SmartArt, and certain drawing-layer objects are outside the supported scope of python-docx.

## Output Example

For three selected Excel rows, the application generates three individual Word documents.

```text
Generated Word Documents/
│
├── John_Smith.docx
├── Jane_Doe.docx
└── Mark_Wilson.docx
```

Existing filenames are protected. If a filename already exists, the application automatically adds a numerical suffix.

## Running Unit Tests

The project includes automated tests for the document processing engine.

Run the following command from the project directory:

```bash
py -3 -m unittest discover -s tests -v
```

## Optional – Convert to an Executable

You can package the application into a Windows executable using PyInstaller.

Install PyInstaller:

```bash
py -3 -m pip install pyinstaller
```

Build the application:

```bash
py -3 -m PyInstaller --noconfirm --windowed --name ExcelToWordAutofillAutomation app.py
```

The packaged application will be available inside the `dist` directory.

Test the packaged application on the target Windows system before distributing it.

## Notes and Limitations

- The original Excel workbook and Word template are not modified.
- The application processes files locally.
- Blank Excel cells are replaced with empty text.
- Excel dates are formatted as `YYYY-MM-DD`.
- Unmatched Word placeholders remain in the generated document and are reported as warnings.
- When a placeholder spans differently formatted Word runs, the replacement uses the formatting of the first run.
- Word text boxes and certain drawing-layer elements are not guaranteed to be processed.

## License

This project is provided for educational and personal use. You may modify and extend the source code to suit your requirements.

## Development Note 

AI-assisted development. Code reviewed and tested by the author.
<br>
*

<p align="center">
  <a href="https://github.com/arrianmunsod">
    <strong>← BACK TO ARRIAN'S GITHUB PROFILE</strong>
  </a>
  <br>
  <sub>Explore more projects, scripts, and technical documentation.</sub>
</p>
