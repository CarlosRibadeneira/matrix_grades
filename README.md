# Grading Matrix Generator

A flexible tool for creating student grading matrices with customizable scales, weights, and qualitative grades. Available as both a Streamlit web app and a legacy CLI script.

## Quick Start (Streamlit App)

### Prerequisites

- **Python 3.10 or higher** - [Download Python](https://www.python.org/downloads/)

Check your version:
```bash
python --version
```

### 1. Install Dependencies

```bash
# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install packages
pip install -r requirements.txt
```

### 2. Run the App

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

## Using the Streamlit App

### Configure Your Matrix (Sidebar)

1. **Grade Scale**: Set min/max values (e.g., 0-100 or 1-10) and decimal places
2. **Set A & Set B**: 
   - Name each set (e.g., "Projects", "Exams")
   - Adjust weight sliders (should sum to 100%)
   - Add/remove projects (one per line)
3. **Trimesters**: Define your grading periods
4. **Qualitative Grades**: Edit the table to define grade labels and ranges

### Add Students

- **Paste**: Enter names directly (one per line)
- **Upload**: Import a CSV or TXT file

### Enter Grades

1. Select a trimester tab
2. Either:
   - Download the CSV template, fill it in, and upload it back
   - Enter grades directly in the data editor
3. View calculated averages and qualitative grades in the preview

### Generate Excel

1. Review validation warnings/errors
2. Click "Generate Excel"
3. Download the file

### Save/Load Configuration

- **Export**: Click "Download Current Config" in sidebar to save your setup
- **Import**: Upload a previously saved JSON config to restore settings

## Legacy CLI Usage

The original command-line script is still available:

```bash
# Edit config.json and students.txt first
python generate.py
```

This generates an Excel file based on `config.json` and `students.txt`.

## Configuration File Format

```json
{
  "scale": {
    "min": 0,
    "max": 100,
    "decimal_places": 1
  },
  "set_a": {
    "name": "Set A",
    "weight": 0.70,
    "projects": ["Project 1", "Project 2"]
  },
  "set_b": {
    "name": "Set B",
    "weight": 0.30,
    "projects": ["Project 1"]
  },
  "qualitative_grades": [
    {"label": "Consistent Impact", "min": 90, "max": 100},
    {"label": "Developing Impact", "min": 80, "max": 89},
    {"label": "Emerging", "min": 70, "max": 79},
    {"label": "Needs Support", "min": 60, "max": 69},
    {"label": "Not Yet Meeting", "min": 0, "max": 59}
  ],
  "trimesters": ["Trimester 1", "Trimester 2", "Trimester 3"],
  "output_file": "grades.xlsx"
}
```

## File Structure

```
matrix_grades/
├── app.py               # Streamlit web application
├── core/                # Core logic module
│   ├── __init__.py
│   ├── config_schema.py # Config defaults and validation
│   ├── excel_generator.py
│   └── validators.py
├── generate.py          # Legacy CLI script
├── config.json          # Default configuration
├── students.txt         # Sample student list (legacy)
└── requirements.txt
```

## Requirements

- Python 3.10+
- Dependencies: `streamlit`, `pandas`, `openpyxl`
