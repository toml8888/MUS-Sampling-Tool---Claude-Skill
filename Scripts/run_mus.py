"""
MUS Sampler - Main Script
Orchestrates: file read -> column detection -> MUS engine -> Excel output
Usage: python run_mus.py <input_file> <pm> <risk_factor> <popnum_col> <value_col> <decimal_system>
       decimal_system is 'normal' (1,234.56) or 'eur' (1.234,56)
"""

import sys
import os
import pandas as pd
from mus_engine import (
    load_population, split_populations, run_mus, find_duplicate_popnums
)
from excel_builder import build_workbook

VALUE_COLUMN_HINTS = [
    "amount", "value", "net", "gross", "debit", "credit",
    "net amount", "gross amount", "transaction amount", "tran amount",
    "balance", "total", "sum", "gbp", "usd", "eur", "fx amount",
    "net value", "gross value", "dr", "cr", "dr amount", "cr amount",
]

POPNUM_COLUMN_HINTS = [
    "population number", "pop number", "pop num", "popnum", "population no",
    "population", "item number", "item no", "item", "sample number",
    "line", "line number", "row", "number", "no", "no.", "seq", "sequence",
]


def suggest_column(headers, hints):
    for h in headers:
        if h.strip().lower() in hints:
            return h
    return None


def get_column_from_user(headers, column_type, suggestion=None):
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
        popnum_col_arg = sys.argv[4] if len(sys.argv) >= 5 else None
        value_col_arg = sys.argv[5] if len(sys.argv) >= 6 else None
        decimal_system = sys.argv[6].strip().lower() if len(sys.argv) >= 7 else "normal"
    else:
        print("\n  Interactive mode.\n")
        input_file = input("  Input file path (Excel or CSV): ").strip().strip('"')
        pm = float(input("  Performance materiality: ").strip())
        risk_factor = float(input("  Risk factor: ").strip())
        popnum_col_arg = None
        value_col_arg = None
        ds = input("  Decimal system - normal (1,234.56) or eur (1.234,56)? [normal]: ").strip().lower()
        decimal_system = ds if ds in ("normal", "eur") else "normal"

    if decimal_system not in ("normal", "eur"):
        print(f"\n  ERROR: decimal_system must be 'normal' or 'eur', got '{decimal_system}'.")
        sys.exit(1)

    print(f"\n  File:                    {input_file}")
    print(f"  Performance materiality: {pm:,.2f}")
    print(f"  Risk factor:             {risk_factor}")
    print(f"  Decimal system:          {decimal_system}")

    # ── Load file ─────────────────────────────────────────────────────────────
    ext = os.path.splitext(input_file)[1].lower()
    if ext in [".xlsx", ".xlsm"]:
        df = pd.read_excel(input_file, dtype=str, keep_default_na=False)
    elif ext == ".csv":
        df = pd.read_csv(input_file, dtype=str, keep_default_na=False)
    else:
        print(f"\n  ERROR: Unsupported file type '{ext}'. Use .xlsx, .xlsm or .csv.")
        sys.exit(1)

    df.columns = [str(c).strip() for c in df.columns]
    headers = list(df.columns)
    source_rows = len(df)
    print(f"\n  Loaded {source_rows} rows. Columns: {headers}")

    # ── Identify population number column ──────────────────────────────────────
    if popnum_col_arg and popnum_col_arg in headers:
        popnum_col = popnum_col_arg
        print(f"\n  Population number column: '{popnum_col}' (from argument)")
    else:
        suggestion = suggest_column(headers, POPNUM_COLUMN_HINTS)
        popnum_col = get_column_from_user(headers, "POPULATION NUMBER", suggestion)
    print(f"  Using population number column: '{popnum_col}'")

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
    transactions, skipped = load_population(rows, popnum_col, value_col, decimal_system)

    # ── Duplicate population number check (Issue 1) ───────────────────────────
    duplicates = find_duplicate_popnums(transactions)
    if duplicates:
        print(f"\n  WARNING: duplicate population numbers found: {duplicates}")
        print(f"  Selection is keyed by row position so flags remain correct, but every")
        print(f"  item should carry a unique population number. Consider re-numbering.")

    if skipped:
        print(f"\n  EXCLUSIONS ({len(skipped)} rows could not be parsed):")
        for s in skipped:
            print(f"    Row {s['row']} (pop {s['popnum']}): {s['reason']}")

    print(f"\n  Transactions loaded: {len(transactions)}")

    positives, negatives, zeros = split_populations(transactions)
    if zeros:
        print(f"  NOTE: {len(zeros)} zero-value rows excluded from both populations.")
    print(f"  Positives: {len(positives)} | Negatives: {len(negatives)}")

    # ── Completeness reconciliation: source = pos + neg + zero + skipped ──────
    reconcile_total = len(positives) + len(negatives) + len(zeros) + len(skipped)
    print(f"\n  Reconciliation: pos {len(positives)} + neg {len(negatives)} + "
          f"zero {len(zeros)} + excluded {len(skipped)} = {reconcile_total} "
          f"(source {source_rows}) -> {'OK' if reconcile_total == source_rows else 'CHECK'}")

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

    # ── Build reconciliation + exclusions payloads for the workpaper ──────────
    reconciliation = {
        "source_rows": source_rows,
        "loaded": len(transactions),
        "skipped": len(skipped),
        "pos_count": len(positives),
        "pos_value": round(sum(abs(t.value) for t in positives), 2),
        "neg_count": len(negatives),
        "neg_value": round(sum(abs(t.value) for t in negatives), 2),
        "zero_count": len(zeros),
        "zero_value": 0.0,
    }

    exclusions = []
    for z in zeros:
        exclusions.append({
            "row": z.row_index + 2,
            "popnum": z.popnum,
            "raw": z.original_row.get(value_col, ""),
            "type": "Zero",
            "reason": "Zero value, excluded from sampling",
        })
    for s in skipped:
        exclusions.append({
            "row": s["row"],
            "popnum": s["popnum"],
            "raw": s["raw"],
            "type": "Unparseable/Blank",
            "reason": s["reason"],
        })

    # ── Build Excel ───────────────────────────────────────────────────────────
    output_path = os.path.join(os.path.dirname(input_file) or ".", "MUS_Sample.xlsx")

    build_workbook(
        output_path=output_path,
        pos_results=pos_results,
        neg_results=neg_results,
        original_headers=headers,
        popnum_col=popnum_col,
        value_col=value_col,
        pm=pm,
        risk_factor=risk_factor,
        filename=os.path.basename(input_file),
        decimal_system=decimal_system,
        reconciliation=reconciliation,
        exclusions=exclusions,
    )

    print(f"\n  Output: {output_path}")
    print("\n" + "="*60)
    print("  COMPLETE")
    print("="*60 + "\n")


if __name__ == "__main__":
    main()
