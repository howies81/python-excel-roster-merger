# Automated Multi-File Roster Merger & Excel Generator

A one-click Python data automation pipeline (`pandas` & `openpyxl`) that ingests 30 separate team roster exports, resolves mixed file formats and character encoding issues, standardizes player names and contact emails, and outputs a presentation-ready Excel workbook.

## Key Features
* **Smart Multi-Format Ingestion:** Automatically reads folders of `.xls` files using `pathlib` and seamlessly falls back to HTML table parsing (`pd.read_html`) for web-exported spreadsheets disguised with `.xls` extensions.
* **UTF-8 & Accent Normalization:** Preserves authentic international characters (e.g., *Nikola Vučević*, *Luka Dončić*) in display columns while normalizing names to plain ASCII (`unicodedata`) to generate valid email usernames.
* **Multi-Suffix & Tag Stripping:** Iteratively strips trailing contract tags and name suffixes (`Jr.`, `Sr.`, `II`, `III`, `IV`, `(TW)`) so email addresses remain clean and consistent.
* **Automated Excel Styling:** Uses `openpyxl` to freeze the top header row, apply custom header fill and typography, center-align headers, and auto-fit all column widths with padding.
* **Execution Logging:** Tracks file ingestion counts and output status via Python's built-in `logging` module.

## Project Structure
```text
python-excel-roster-merger/
├── nba_roster_generator.py   # Main data cleaning and Excel export pipeline
├── requirements.txt          # Python dependencies
├── .gitignore                # Excluded virtual environments and raw data folders
└── README.md                 # Project documentation
```

## Quick Start

1. **Clone the repository:**
   ```bash
   git clone https://github.com<your-username>/python-excel-roster-merger.git
   cd python-excel-roster-merger
   ```

2. **Create a virtual environment and install dependencies:**
   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # macOS/Linux:
   source .venv/bin/activate

   pip install -r requirements.txt
   ```

3. **Add input files & run the pipeline:**
   Place your raw `.xls` roster exports inside an `NBA_rosters/` directory and run:
   ```bash
   python nba_roster_generator.py
   ```
   The formatted workbook will be generated at `Output_files/NBA_Roster_20262027.xlsx`.