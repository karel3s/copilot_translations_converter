# Copilot translations converter

## Pre-requisites

- Python3
- Pandas

## JSON to Excel Converter

### Key Features:
- ✅ Converts JSON files to Excel (.xlsx) format
- ✅ Supports both **horizontal** (multi-column) and **vertical** (Key|Value) layouts
- ✅ Flattens nested JSON structures automatically
- ✅ Handles arrays, objects, and primitive values
- ✅ Auto-sizes columns for readability
- ✅ Supports NDJSON (newline-delimited JSON) format

### Usage Examples:

**Vertical layout (recommended for translations):**
```bash
python3 json_to_excel.py 20251117_Sam_Translations_FR.json --vertical
```
Output: Creates a two-column layout with "Key" and "Value" columns

**Horizontal layout (default):**
```bash
python3 json_to_excel.py input.json
```
Output: Flattens nested JSON into columns

**Custom output file:**
```bash
python3 json_to_excel.py input.json -o output.xlsx --vertical
```

**Custom separator for nested keys:**
```bash
python3 json_to_excel.py input.json --sep "_" --vertical
```
Default separator is "." (e.g., `parent.child.key`)

**NDJSON format:**
```bash
python3 json_to_excel.py data.ndjson --ndjson --vertical
```

**Custom root sheet name:**
```bash
python3 json_to_excel.py input.json --root-sheet "Translations" --vertical
```

### Layout Options:

**Vertical Layout (`--vertical`):**
- Best for flat key-value pairs (like translations)
- Creates two columns: Key | Value
- Example:
  ```
  Key                    | Value
  ----------------------|--------
  greeting.hello        | Hello
  greeting.goodbye      | Goodbye
  ```

**Horizontal Layout (default):**
- Best for structured data with multiple records
- Creates columns for each field
- Nested objects become dot-separated column names

## Excel to JSON Converter

### Key Features:
- ✅ Converts XLSX back to **flat JSON format** (matches original structure)
- ✅ Adds timestamp to output filename
- ✅ Handles Key|Value column layout
- ✅ Preserves all data types correctly
- ✅ Round-trip tested and verified (JSON → Excel → JSON produces identical results)

### Usage Examples:

**Basic conversion with timestamp:**
```bash
python3 excel_to_json.py input.xlsx --timestamp
```
Output: `input_20251117_153424.json`

**Custom output:**
```bash
python3 excel_to_json.py input.xlsx -o output.json --timestamp
```

**Specify sheet:**
```bash
python3 excel_to_json.py input.xlsx --sheet "data" --timestamp
```

**Custom indentation:**
```bash
python3 excel_to_json.py input.xlsx --indent 4 --timestamp
```

### Complete Workflow:

1. **JSON to Excel:**
   ```bash
   python3 json_to_excel.py data.json --vertical
   ```

2. **Excel to JSON (with timestamp):**
   ```bash
   python3 excel_to_json.py data.xlsx --timestamp
   ```

The script has been tested and verified to produce identical output to the original JSON file! 🎉
