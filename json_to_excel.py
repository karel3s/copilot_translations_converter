#!/usr/bin/env python3
import argparse
import json
import sys
from collections import OrderedDict
from pathlib import Path
from typing import Any, Dict, List, Tuple

try:
    import pandas as pd
except ImportError:
    print("This script requires pandas. Install with: pip install pandas openpyxl", file=sys.stderr)
    sys.exit(1)


def load_json(path: Path, ndjson: bool) -> Any:
    with path.open("r", encoding="utf-8") as f:
        if ndjson:
            records = []
            for i, line in enumerate(f, start=1):
                line = line.strip()
                if not line:
                    continue
                try:
                    records.append(json.loads(line))
                except json.JSONDecodeError as e:
                    raise ValueError(f"Invalid JSON on line {i}: {e}")
            return records
        return json.load(f)


def is_list_of_dicts(val: Any) -> bool:
    return isinstance(val, list) and all(isinstance(x, dict) for x in val)


def make_sheet_name(name: str) -> str:
    # Excel sheet names max 31 chars and must not contain []:*?/\
    invalid = set('[]:*?/\\')
    safe = "".join(ch if ch not in invalid else "_" for ch in name)
    return (safe or "Sheet")[:31]


def uniq_sheet_name(base: str, used: set) -> str:
    name = make_sheet_name(base)
    if name not in used:
        used.add(name)
        return name
    # Append counter while respecting 31 char limit
    i = 2
    while True:
        suffix = f"_{i}"
        candidate = (name[: 31 - len(suffix)]) + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate
        i += 1


def flatten_dict(d: Dict[str, Any], parent_key: str = "", sep: str = ".") -> Dict[str, Any]:
    """Recursively flatten nested dictionaries."""
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)


def build_frames(obj: Any, root_name: str = "data", sep: str = ".", vertical: bool = False) -> "OrderedDict[str, pd.DataFrame]":
    frames: "OrderedDict[str, pd.DataFrame]" = OrderedDict()

    def normalize_list(name: str, lst: List[Any]) -> pd.DataFrame:
        if is_list_of_dicts(lst):
            return pd.json_normalize(lst, sep=sep)
        return pd.DataFrame({"value": lst})

    used_names: set = set()

    if isinstance(obj, list):
        frames[uniq_sheet_name(root_name, used_names)] = normalize_list(root_name, obj)
        return frames

    if isinstance(obj, dict):
        # Flatten the entire dictionary first
        flattened = flatten_dict(obj, sep=sep)
        
        if vertical:
            # Create vertical layout: Key | Value
            df = pd.DataFrame([
                {"Key": k, "Value": v} for k, v in flattened.items()
            ])
            frames[uniq_sheet_name(root_name, used_names)] = df
        else:
            # Original horizontal layout
            scalars: Dict[str, Any] = {}
            for k, v in obj.items():
                if isinstance(v, list):
                    frames[uniq_sheet_name(k, used_names)] = normalize_list(k, v)
                elif isinstance(v, dict):
                    df = pd.json_normalize(v, sep=sep)
                    frames[uniq_sheet_name(k, used_names)] = df
                else:
                    scalars[k] = v

            if scalars:
                frames[uniq_sheet_name(root_name, used_names)] = pd.json_normalize(scalars, sep=sep)
            if not frames:
                frames[uniq_sheet_name(root_name, used_names)] = pd.DataFrame()
        return frames

    # Primitive value
    if vertical:
        frames[uniq_sheet_name(root_name, used_names)] = pd.DataFrame({"Key": ["value"], "Value": [obj]})
    else:
        frames[uniq_sheet_name(root_name, used_names)] = pd.DataFrame({"value": [obj]})
    return frames


def autosize_columns_if_openpyxl(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    # Only works with openpyxl engine
    try:
        import openpyxl  # noqa: F401
    except Exception:
        return
    ws = writer.sheets.get(sheet_name)
    if ws is None:
        return
    # Compute simple max width for each column
    for idx, col in enumerate(df.columns, start=1):
        series = df[col].astype(str)
        max_len = max([len(str(col))] + [len(x) for x in series.tolist()]) if not df.empty else len(str(col))
        # Add a little padding
        adjusted = min(max_len + 2, 60)
        col_letter = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[col_letter].width = adjusted


def write_excel(frames: "OrderedDict[str, pd.DataFrame]", out_path: Path, engine: str) -> None:
    with pd.ExcelWriter(out_path, engine=engine) as writer:
        for sheet_name, df in frames.items():
            # Ensure at least an index-free sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            autosize_columns_if_openpyxl(writer, sheet_name, df)


def main():
    parser = argparse.ArgumentParser(description="Convert JSON to Excel (.xlsx)")
    parser.add_argument("input", type=Path, help="Path to input JSON file")
    parser.add_argument("-o", "--output", type=Path, help="Path to output .xlsx (default: input name with .xlsx)")
    parser.add_argument("--root-sheet", default="data", help="Sheet name for root list/scalars (default: data)")
    parser.add_argument("--sep", default=".", help="Separator for flattened columns (default: .)")
    parser.add_argument("--ndjson", action="store_true", help="Treat input as newline-delimited JSON (one object per line)")
    parser.add_argument("--vertical", action="store_true", help="Create vertical layout with Key|Value columns (default: horizontal)")
    parser.add_argument("--engine", choices=["openpyxl", "xlsxwriter"], default="openpyxl", help="Excel writer engine (default: openpyxl)")
    args = parser.parse_args()

    if not args.input.exists():
        print(f"Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    if args.output is None:
        args.output = args.input.with_suffix(".xlsx")

    try:
        data = load_json(args.input, ndjson=args.ndjson)
    except Exception as e:
        print(f"Failed to read JSON: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        frames = build_frames(data, root_name=args.root_sheet, sep=args.sep, vertical=args.vertical)
        write_excel(frames, args.output, engine=args.engine)
    except Exception as e:
        print(f"Failed to write Excel: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
