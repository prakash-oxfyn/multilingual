import json
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


def flatten_json(data, module="", prefix=""):
    """
    Flattens nested JSON into:
    (module, translation_key) -> value
    """
    result = {}

    if not isinstance(data, dict):
        return result

    for key, value in data.items():
        if module == "":
            # First level becomes Module
            result.update(flatten_json(value, key, ""))
        else:
            if isinstance(value, dict):
                new_prefix = f"{prefix}.{key}" if prefix else key
                result.update(flatten_json(value, module, new_prefix))
            else:
                translation_key = f"{prefix}.{key}" if prefix else key
                result[(module, translation_key)] = value

    return result


def auto_adjust_columns(ws):
    for column in ws.columns:
        length = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[column[0].column_letter].width = min(length + 4, 80)


def workbook_matches(output_file, headers, rows):
    """
    Returns True if the existing workbook already contains exactly
    the same data as what we're about to generate.
    """
    if not output_file.exists():
        return False

    try:
        wb = load_workbook(output_file, data_only=True)
        ws = wb.active

        actual = []
        for row in ws.iter_rows(values_only=True):
            actual.append(
                tuple("" if value is None else value for value in row)
            )

        expected = [tuple(headers)]
        expected.extend(
            tuple("" if value is None else value for value in row)
            for row in rows
        )

        return actual == expected

    except Exception:
        return False


def main():

    if len(sys.argv) != 2:
        print("Usage:")
        print("    python convert_locales.py <project_name>")
        print()
        print("Example:")
        print("    python convert_locales.py setbet")
        sys.exit(1)

    project = sys.argv[1]

    base_dir = Path(__file__).parent
    project_dir = base_dir / project

    input_dir = project_dir / "locales-json"
    output_dir = project_dir / "locales-excel"

    if not input_dir.exists():
        print(f"Folder not found: {input_dir}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    json_files = sorted(input_dir.glob("*.json"))

    if not json_files:
        print("No JSON files found.")
        sys.exit(1)

    print(f"Project : {project}")
    print(f"Languages Found : {len(json_files)}")
    print()

    language_data = {}
    all_keys = set()

    # Read every language file
    for file in json_files:
        language = file.stem.upper()

        print(f"Reading {file.name}")

        with open(file, "r", encoding="utf-8") as f:
            data = json.load(f)

        flattened = flatten_json(data)

        language_data[language] = flattened
        all_keys.update(flattened.keys())

    languages = sorted(language_data.keys())

    headers = ["Module", "Translation Key"] + languages

    rows = []

    # Build all rows
    for module, translation_key in sorted(all_keys):

        row = [module, translation_key]

        for language in languages:
            row.append(
                language_data[language].get(
                    (module, translation_key),
                    ""
                )
            )

        rows.append(row)

    output_file = output_dir / f"{project}_translations.xlsx"

    # Skip rewriting if workbook content is unchanged
    if workbook_matches(output_file, headers, rows):
        print()
        print(f"✓ {output_file.name} unchanged")
        print("Done!")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Translations"

    # Header
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Data
    for row in rows:
        ws.append(row)

    ws.freeze_panes = "A2"

    auto_adjust_columns(ws)

    wb.save(output_file)

    print()
    print(f"Created: {output_file}")
    print("Done!")


if __name__ == "__main__":
    main()