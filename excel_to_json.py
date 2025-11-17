#!/usr/bin/env python3
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("This script requires pandas. Install with: pip install pandas openpyxl", file=sys.stderr)
    sys.exit(1)


def read_excel_to_json(xlsx_path, sheet_name=None):
    """Read Excel file and convert back to flat JSON structure."""
    # Read the Excel file
    if sheet_name:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name, engine='openpyxl')
    else:
        # Read first sheet by default
        df = pd.read_excel(xlsx_path, sheet_name=0, engine='openpyxl')
    
    # Check if it's a vertical layout (Key | Value columns)
    if 'Key' in df.columns and 'Value' in df.columns:
        # Vertical layout - convert to flat dictionary
        result = {}
        for _, row in df.iterrows():
            key = str(row['Key'])
            value = row['Value']
            # Handle different value types
            try:
                if pd.isna(value):
                    result[key] = None
                elif isinstance(value, (np.integer, np.floating)):
                    result[key] = value.item()
                elif isinstance(value, np.ndarray):
                    result[key] = value.tolist()
                else:
                    result[key] = value
            except:
                # Fallback for any type conversion issues
                result[key] = str(value) if value is not None else None
        return result
    else:
        raise ValueError("Excel file must have 'Key' and 'Value' columns")


def main():
    parser = argparse.ArgumentParser(description="Convert Excel (.xlsx) back to JSON")
    parser.add_argument("input", type=Path, help="Path to input .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="Path to output JSON file (default: input name with .json)")
    parser.add_argument("--sheet", help="Sheet name to read (default: first sheet)")
    parser.add_argument("--timestamp", action="store_true", help="Add timestamp to output filename")
    parser.add_argument("--indent", type=int, default=2, help="JSON indentation (default: 2)")
    args = parser.parse_args()
    
    if not args.input.exists():
        print(f"Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    
    # Determine output path
    if args.output is None:
        output_path = args.input.with_suffix(".json")
    else:
        output_path = args.output
    
    # Add timestamp if requested
    if args.timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = output_path.stem
        suffix = output_path.suffix
        output_path = output_path.parent / f"{stem}_{timestamp}{suffix}"
    
    try:
        data = read_excel_to_json(args.input, sheet_name=args.sheet)
    except Exception as e:
        print(f"Failed to read Excel: {e}", file=sys.stderr)
        sys.exit(1)
    
    try:
        with output_path.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=args.indent, ensure_ascii=False)
    except Exception as e:
        print(f"Failed to write JSON: {e}", file=sys.stderr)
        sys.exit(1)
    
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
