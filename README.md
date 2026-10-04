# Abbreviation Finder and Document Analyzer

A Python-based utility suite designed to analyze Word Documents (`.docx`). It offers a Command Line Interface (CLI), a FastAPI Web API, and a standalone abbreviation finder module.

## Features
- **UK/US Word Analysis**: Compares document vocabulary against British and American spelling dictionaries.
- **Abbreviation Extraction**: Extracts abbreviations and matching full forms from the text.
- **Formatting Checks**: Detects styling issues (e.g., double spaces, double punctuation, spacing around endnote/footnote citations, incorrect serial commas).
- **Hyphenation Variant Analysis**: Analyzes spelling consistency of compound words with hyphens vs. spaces vs. combined forms.

---

## 🛠 Prerequisites

Ensure you have **Python 3.8+** installed.

Install the required dependencies:
```bash
pip install -r requirements.txt
```

---

## 🚀 How to Execute

### 1. Document Analyzer CLI (`document_analyzer_cli.py`)
The CLI supports several commands depending on the analysis you want to perform.

#### A. Run All Analyses (Combined Report)
Performs word analysis, abbreviation finding, serial comma checks, hyphenated words, double spaces, and formatting checks.
```bash
python document_analyzer_cli.py analyze-all -i input.docx -o output_report.json
```
*Note: If no arguments are passed, it defaults to: `analyze-all -i input.docx -o input_report.json`.*

#### B. UK/US Word Spelling Analysis
```bash
python document_analyzer_cli.py analyze -i input.docx -o output_report.txt --us_dict us_dict.txt --uk_dict uk_dict.txt
```

#### C. Find Abbreviations & Oxford Commas
```bash
python document_analyzer_cli.py abbreviations -i input.docx -o abbreviation_report.json
```

#### D. Check Formatting/Styling Rules
```bash
python document_analyzer_cli.py formatting -i input.docx -o formatting_report.json
```

#### E. Manage Section Words Form
Open the interactive form to view, add (single or comma-separated), and delete prohibited section terms:
```bash
python document_analyzer_cli.py open-form
```

---

### 2. Document Analyzer Web API (`UK_US_Word_Analyzer_API.py`)
Run the FastAPI backend server using Uvicorn:
```bash
uvicorn UK_US_Word_Analyzer_API:app --host 127.0.0.1 --port 8000 --reload
```
Once the server is running, visit **`http://127.0.0.1:8000/docs`** in your browser to view the interactive Swagger API documentation and try out endpoints like `/analyze` or `/abbreviation`.

---

### 3. Standalone Abbreviation Finder (`abbreviation_finder.py`)
To run the quick standalone finder on a sample document:
```bash
python abbreviation_finder.py
```

---

## 📦 Converting CLI to Executable (`.exe`)

You can compile the CLI tool into a standalone Windows executable (`.exe`) using **PyInstaller** and the included specification file (`document_analyzer_cli.spec`).

### Step-by-Step Instructions

#### **Step 1: Install PyInstaller**
Install PyInstaller in your Python environment:
```bash
pip install pyinstaller
```

#### **Step 2: Build the Executable**
Run PyInstaller using the existing configuration specification file:
```bash
pyinstaller document_analyzer_cli.spec
```

#### **Step 3: Locate the Executable**
After the compilation completes:
- The compiled `.exe` will be located in the **`dist/`** directory:  
  `dist/document_analyzer_cli.exe`
- Temporary build files are created in the `build/` directory (these can be safely ignored or deleted).

#### **Step 4: Copy Configuration & Dictionaries**
To run the executable properly, make sure the following files are placed in the **same directory** as the `document_analyzer_cli.exe` file (or set paths explicitly):
- `formatting_rules.json`
- `us_dict.txt`
- `uk_dict.txt`

#### **Step 5: Run the Executable**
Execute it directly from your command prompt:
```cmd
dist\document_analyzer_cli.exe analyze-all -i input.docx -o output_report.json
```
