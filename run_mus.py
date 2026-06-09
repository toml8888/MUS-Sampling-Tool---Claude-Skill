"""
MUS Sampler - Main Script
Orchestrates: Excel read -> column detection -> MUS engine -> Excel output
Usage: python run_mus.py <input_file> <pm> <risk_factor> [ref_col] [value_col]
"""

import sys
import os
import pandas as pd
from mus_engine import load_population, split_populations, run_mus
from excel_builder import build_workbook

# Common names the script will look for to suggest a value column
VALUE_COLUMN_HINTS = [
    "amount", "value", "net", "gross", "debit", "credit",
    "net amount", "gross amount", "transaction amount", "tran amount",
    "balance", "total", "sum", "gbp", "usd", "eur", "fx amount",
    "net value", "gross value", "dr", "cr", "dr amount", "cr amount",
]

REF_COLUMN_HINTS = [
    "ref", "reference", "transaction ref", "tran ref", "id",
    "transaction id", "journal ref", "line ref", "entry ref",
    "number", "no", "no.", "doc", "document", "posting ref",
]


def suggest_column(headers, hints):
    """Return the first header that fuzzy-matches a known hint."""
    for h in headers:
        if h.strip().lower() in hints:
            return h
    return None


def get_column_from_user(headers, column_type, suggestion=None):
    """Print headers and ask user to pick."""
    print(f"\n  Available columns:")
    for i, h in enumerate(headers):
        print(f"    [{i}] {h}")

    if suggestion:
        print(f"\n  Best guess for {column_type} column: '{suggestion}'")
        confirm = input(f"  Use '{suggestion}' as the {column_type} column? (y/n): ").strip().lower()
        if confirm == "y":
            return suggestion

    while True:
        choice = input(f"\n  Enter the column name or number for the {column_type} column: ").strip()
        if choice.isdigit():
            idx = int(choice)
            if 0 <= idx < len(headers):
                return headers[idx]
            else:
                print(f"  Invalid index. Enter 0-{len(headers)-1}.")
        elif choice in headers:
            return choice
        else:
            print(f"  Column '{choice}' not found. Try again.")


def main():
    print("\n" + "="*60)
    print("  MONETARY UNIT SAMPLING - AUDIT WORKPAPER TOOL")
    print("="*60)

    # ── Parse args or prompt ──────────────────────────────────────────────────
    if len(sys.argv) >= 4:
        input_file = sys.argv[1]
        pm = float(sys.argv[2])
        risk_factor = float(sys.argv[3])
        ref_col_arg = sys.argv[4] if len(sys.argv) >= 5 else None
        value_col_arg = sys.argv[5] if len(sys.argv) >= 6 else None
    else:
        print("\n  Interactive mode.\n")
        input_file = input("  Input file path (Excel or CSV): ").strip().strip('"')
        pm = float(input("  Performance materiality: ").strip())
        risk_factor = float(input("  Risk factor: ").strip())
        ref_col_arg = None
        value_col_arg = None

    print(f"\n  File:                    {input_file}")
    print(f"  Performance materiality: {pm:,.2f}")
    print(f"  Risk factor:             {risk_factor}")

    # ── Load file ─────────────────────────────────────────────────────────────
    ext = os.path.splitext(input_file)[1].lower()
    if ext in [".xlsx", ".xls", ".xlsm"]:
        df = pd.read_excel(input_file, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(input_file, dtype=str)
    else:
        print(f"\n  ERROR: Unsupported file type '{ext}'. Use .xlsx or .csv.")
        sys.exit(1)

    # Strip whitespace from headers
    df.columns = [str(c).strip() for c in df.columns]
    headers = list(df.columns)
    print(f"\n  Loaded {len(df)} rows. Columns: {headers}")

    # ── Identify ref column ───────────────────────────────────────────────────
    if ref_col_arg and ref_col_arg in headers:
        ref_col = ref_col_arg
        print(f"\n  Reference column: '{ref_col}' (from argument)")
    else:
        suggestion = suggest_column(headers, REF_COLUMN_HINTS)
        ref_col = get_column_from_user(headers, "REFERENCE", suggestion)
    print(f"  Using reference column: '{ref_col}'")

    # ── Identify value column ─────────────────────────────────────────────────
    if value_col_arg and value_col_arg in headers:
        value_col = value_col_arg
        print(f"\n  Value column: '{value_col}' (from argument)")
    else:
        suggestion = suggest_column(headers, VALUE_COLUMN_HINTS)
        value_col = get_column_from_user(headers, "VALUE", suggestion)
    print(f"  Using value column: '{value_col}'")

    # ── Load into engine ──────────────────────────────────────────────────────
    rows = df.to_dict(orient="records")
    transactions, skipped = load_population(rows, ref_col, value_col)

    if skipped:
        print(f"\n  WARNINGS ({len(skipped)} rows skipped):")
        for s in skipped:
            print(f"    {s}")

    print(f"\n  Transactions loaded: {len(transactions)}")

    positives, negatives, zeros = split_populations(transactions)
    if zeros:
        print(f"  NOTE: {len(zeros)} zero-value rows excluded from both populations.")
    print(f"  Positives: {len(positives)} | Negatives: {len(negatives)}")

    # ── Run MUS ───────────────────────────────────────────────────────────────
    pos_results = run_mus(positives, pm, risk_factor)
    neg_results = run_mus(negatives, pm, risk_factor)

    if pos_results:
        w = pos_results["workings"]
        print(f"\n  POSITIVE: IS={w['is_count']}, MUS selected={w['mus_selected_count']}, "
              f"Total selected={w['total_selected']}/{w['total_count']}")

    if neg_results:
        w = neg_results["workings"]
        print(f"  NEGATIVE: IS={w['is_count']}, MUS selected={w['mus_selected_count']}, "
              f"Total selected={w['total_selected']}/{w['total_count']}")

    # ── Build Excel ───────────────────────────────────────────────────────────
    output_path = os.path.join(os.path.dirname(input_file), "MUS_Sample.xlsx")

    build_workbook(
        output_path=output_path,
        pos_results=pos_results,
        neg_results=neg_results,
        original_headers=headers,
        ref_col=ref_col,
        value_col=value_col,
        pm=pm,
        risk_factor=risk_factor,
        filename=os.path.basename(input_file),
    )

    print(f"\n  Output: {output_path}")
    print("\n" + "="*60)
    print("  COMPLETE")
    print("="*60 + "\n")


if __name__ == "__main__":
    main()
