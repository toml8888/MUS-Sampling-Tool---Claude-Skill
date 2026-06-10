"""
MUS Sampling Engine
Core sampling logic. Takes a list of dicts (population rows), population number
column name, value column name, performance materiality, and risk factor.
Returns annotated rows and workings dict for both positive and negative populations.

Selections are keyed by row index internally so duplicate population numbers can
never corrupt IS/MUS flags. clean_value is the single shared parser used by both
this engine and the Excel builder.
"""

import math
import random
from collections import namedtuple

RANDOM_SEED = 42

Transaction = namedtuple("Transaction", ["popnum", "value", "row_index", "original_row"])

CURRENCY_SYMBOLS = ["\u00a3", "$", "\u20ac"]  # £, $, €


def clean_value(raw, decimal_system="normal"):
    """
    Single shared value parser. Used by the engine and the Excel builder so the
    workpaper formulas always match the engine result.

    decimal_system:
      "normal" -> 1,234.56  (comma thousands, dot decimal)
      "eur"    -> 1.234,56  (dot thousands, comma decimal)

    Returns:
      float for a parseable value (bracketed negatives supported)
      None for blank / NaN / empty
    Raises:
      ValueError for a non-empty value that cannot be parsed
    """
    if raw is None:
        return None

    s = str(raw).strip()
    if s == "" or s.lower() in ("nan", "none", "null"):
        return None

    # Strip currency symbols and spaces
    for sym in CURRENCY_SYMBOLS:
        s = s.replace(sym, "")
    s = s.replace(" ", "").replace("\u00a0", "")  # normal and non-breaking space

    if s in ("", "-", "\u2014"):
        return None

    # Bracketed negative: (1,234.56) -> negative
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    # Leading minus
    if s.startswith("-"):
        negative = not negative
        s = s[1:].strip()

    # Decimal system normalisation
    if decimal_system == "eur":
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")

    value = float(s)  # raises ValueError if junk
    if math.isnan(value) or math.isinf(value):
        return None
    if negative:
        value = -value
    return value


def load_population(rows, popnum_col, value_col, decimal_system="normal"):
    """
    Takes list of dicts (from pandas or csv.DictReader).
    Returns (transactions, skipped) where:
      transactions is a list of Transaction namedtuples
      skipped is a list of dicts: {row, popnum, raw, reason}
    """
    transactions = []
    skipped = []

    for i, row in enumerate(rows):
        popnum = str(row.get(popnum_col, f"ROW_{i+2}")).strip()
        raw_value = row.get(value_col, "")

        try:
            value = clean_value(raw_value, decimal_system)
        except (ValueError, TypeError):
            skipped.append({
                "row": i + 2,
                "popnum": popnum,
                "raw": raw_value,
                "reason": f"Could not parse value '{raw_value}'",
            })
            continue

        if value is None:
            skipped.append({
                "row": i + 2,
                "popnum": popnum,
                "raw": raw_value,
                "reason": "Blank or empty value",
            })
            continue

        transactions.append(Transaction(popnum=popnum, value=value, row_index=i, original_row=row))

    return transactions, skipped


def find_duplicate_popnums(transactions):
    """Return a sorted list of population numbers that appear more than once."""
    seen = {}
    for t in transactions:
        seen[t.popnum] = seen.get(t.popnum, 0) + 1
    return sorted([p for p, n in seen.items() if n > 1])


def split_populations(transactions):
    positives = [t for t in transactions if t.value > 0]
    negatives = [t for t in transactions if t.value < 0]
    zeros = [t for t in transactions if t.value == 0]
    return positives, negatives, zeros


def run_mus(transactions, performance_materiality, risk_factor):
    """
    Runs full MUS on a population (all positive or all negative, passed as-is).
    Works in absolute values throughout. Selections keyed by row_index so
    duplicate population numbers cannot corrupt flags.

    Returns dict with workings and annotated list. annotated items carry:
      transaction, status ('IS' | 'MUS' | ''), running_total (None for IS)
    For the select-all branch, interval and random_start are None (no walk done).
    """
    if not transactions:
        return None

    pm = abs(performance_materiality)
    total_population = sum(abs(t.value) for t in transactions)

    # Pass 1: individually significant
    individually_significant = [t for t in transactions if abs(t.value) >= pm]
    remaining = [t for t in transactions if abs(t.value) < pm]

    is_total = sum(abs(t.value) for t in individually_significant)
    remaining_total = sum(abs(t.value) for t in remaining)
    is_indices = set(t.row_index for t in individually_significant)

    # Pass 2: MUS on remaining
    if not remaining:
        raw_n = 0.0
        sample_size = 0
        interval = None
        random_start = None
        mus_indices = set()
        hit_points = []
        full_residual = False
    else:
        raw_n = (remaining_total * risk_factor) / pm
        sample_size = math.ceil(raw_n)

        if sample_size >= len(remaining):
            # Select entire remaining population. No walk performed, so no
            # interval or random start is documented (would be fabricated).
            interval = None
            random_start = None
            mus_indices = set(t.row_index for t in remaining)
            hit_points = []
            full_residual = True
        else:
            full_residual = False
            interval = remaining_total / sample_size
            rng = random.Random(RANDOM_SEED)
            random_start = rng.uniform(0, interval)
            hit_points = [random_start + (i * interval) for i in range(sample_size)]

            mus_indices = set()
            cumulative = 0.0
            hit_idx = 0

            for t in remaining:
                cumulative += abs(t.value)
                while hit_idx < len(hit_points) and hit_points[hit_idx] <= cumulative:
                    mus_indices.add(t.row_index)
                    hit_idx += 1
                if hit_idx >= len(hit_points):
                    break

    # Build annotated list with running cumulative total for the residual population
    cumulative = 0.0
    annotated = []
    for t in transactions:
        if t.row_index in is_indices:
            annotated.append({
                "transaction": t,
                "status": "IS",
                "running_total": None,
            })
        else:
            cumulative += abs(t.value)
            status = "MUS" if t.row_index in mus_indices else ""
            annotated.append({
                "transaction": t,
                "status": status,
                "running_total": round(cumulative, 2),
            })

    workings = {
        "total_count": len(transactions),
        "total_value": round(total_population, 2),
        "performance_materiality": pm,
        "risk_factor": risk_factor,
        "is_count": len(individually_significant),
        "is_value": round(is_total, 2),
        "remaining_count": len(remaining),
        "remaining_total": round(remaining_total, 2),
        "raw_sample_size": round(raw_n, 6),
        "sample_size": sample_size,
        "interval": round(interval, 2) if interval is not None else None,
        "random_start": round(random_start, 2) if random_start is not None else None,
        "random_seed": RANDOM_SEED,
        "full_residual_selected": full_residual,
        "hit_points": [round(h, 2) for h in hit_points],
        "mus_selected_count": len(mus_indices),
        "total_selected": len(is_indices) + len(mus_indices),
    }

    return {"workings": workings, "annotated": annotated}
