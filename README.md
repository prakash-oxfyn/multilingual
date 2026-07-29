# Localization JSON to Excel Converter

A simple Python utility to convert nested locale JSON files into Excel (`.xlsx`) files for translation management.

Each language JSON file is converted into its own Excel file in the following format:

| Module | Translation Key | EN |
|--------|-----------------|----|
| errors | appCrashErrorHeader | Application Error |
| locale | unavailable.title | Language unavailable |

---

## Features

- Converts nested JSON files into Excel (`.xlsx`)
- Supports unlimited projects
- Automatically detects all language JSON files
- Preserves Unicode (Tamil, Telugu, Hindi, Gujarati, etc.)
- Automatically creates the output directory if it doesn't exist
- Auto-adjusts Excel column widths
- Bold header row
- Freezes the header row
- No code changes required when adding new languages

---

## Folder Structure

```
multilingual/
│
├── convert_locales.py
├── requirements.txt
├── README.md
│
├── setbet/
│   ├── locales-json/
│   │   ├── en.json
│   │   ├── gu.json
│   │   ├── hi.json
│   │   ├── ta.json
│   │   └── te.json
│   │
│   └── locales-excel/
│
├── playwin/
│   ├── locales-json/
│   └── locales-excel/
│
└── another-project/
    ├── locales-json/
    └── locales-excel/
```

---

## Installation

Install the required dependency:

```bash
pip install -r requirements.txt
```

---

## Usage

Run the converter by providing the project name.

Example:

```bash
python convert_locales.py setbet
```

For another project:

```bash
python convert_locales.py playwin
```

---

## Input

Place all locale JSON files inside:

```
<project>/locales-json/
```

Example:

```
setbet/
└── locales-json/
    ├── en.json
    ├── gu.json
    ├── hi.json
    ├── ta.json
    └── te.json
```

---

## Output

The generated Excel files will be created inside:

```
<project>/locales-excel/
```

Example:

```
setbet/
└── locales-excel/
    ├── en.xlsx
    ├── gu.xlsx
    ├── hi.xlsx
    ├── ta.xlsx
    └── te.xlsx
```

---

## Excel Format

Example:

| Module | Translation Key | EN |
|--------|-----------------|----|
| errors | appCrashErrorHeader | Application Error |
| errors | appCrashError | The application crashed unexpectedly. |
| locale | unavailable.title | Language unavailable |
| locale | unavailable.message | Language {{locale}} is not available yet. Showing default language. |

---

## Supported Languages

The script automatically converts every `.json` file found in the `locales-json` directory.

Examples:

- en.json
- ta.json
- te.json
- hi.json
- gu.json
- bn.json
- ml.json
- kn.json

No code changes are required when adding new languages.

---

## Adding a New Project

Simply create the following structure:

```
new-project/
├── locales-json/
└── locales-excel/
```

Copy your locale JSON files into `locales-json` and run:

```bash
python convert_locales.py new-project
```

---

## Requirements

- Python 3.9+
- openpyxl

---

## Notes

- The first level of the JSON becomes the **Module**.
- Nested keys become the **Translation Key** using dot notation.
- The third column header is automatically derived from the JSON filename (EN, TA, HI, etc.).
- UTF-8 encoding is used to preserve all Unicode characters.

---

## Example

Input JSON:

```json
{
  "errors": {
    "header": "Oops!",
    "generic": "Something went wrong."
  },
  "locale": {
    "unavailable": {
      "title": "Language unavailable"
    }
  }
}
```

Output:

| Module | Translation Key | EN |
|--------|-----------------|----|
| errors | header | Oops! |
| errors | generic | Something went wrong. |
| locale | unavailable.title | Language unavailable |

---

## Future Enhancements

- Excel → JSON conversion
- Translation validation
- Missing key detection
- Placeholder validation (`{{count}}`, `{{name}}`, etc.)
- Compare translations across languages
- Export reports for missing translations