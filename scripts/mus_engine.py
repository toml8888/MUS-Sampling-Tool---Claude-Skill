"""
MUS Sampling Engine
Core sampling logic. Takes a list of dicts (population rows), ref column name,
value column name, performance materiality, and risk factor.
Returns annotated rows and workings dict for both positive and negative populations.
"""

import math
import random
from collections import namedtuple

RANDOM_SEED = 42

Transaction = namedtuple("Transaction", ["ref", "value", "row_index", "original_row"])


def clean_value(raw):
    """Strip currency symbols, commas, spaces. Return float or raise."""
    cleaned = str(raw).strip().replace("£", "").replace("$", "").replace("€", "").replace(",", "").replace(" ", "")
    if cleaned in ("", "-", "—"):
        return None
    return float(cleaned)


def load_population(rows, ref_col, value_col):
    """
    Takes list of dicts (from pandas or csv.DictReader).
    Returns (transactions, skipped_refs) where transactions is a list of Transaction namedtuples.
    """
    transactions = []
    skipped = []

    for i, row in enumerate(rows):
        ref = str(row.get(ref_col, f"ROW_{i+2}")).strip()
        raw_value = row.get(value_col, "")

        try:
            value = clean_value(raw_value)
        except (ValueError, TypeError):
            skipped.append(f"Row {i+2}: could not parse value '{raw_value}' for ref '{ref}'")
            continue

        if value is None:
            skipped.append(f"Row {i+2}: empty value for ref '{ref}'")
            continue

        transactions.append(Transaction(ref=ref, value=value, row_index=i, original_row=row))

    return transactions, skipped


def split_populations(transactions):
    positives = [t for t in transactions if t.value > 0]
    negatives = [t for t in transactions if t.value < 0]
    zeros = [t for t in transactions if t.value == 0]
    return positives, negatives, zeros


def run_mus(transactions, performance_materiality, risk_factor):
    """
    Runs full MUS on a population (all positive or all negative, passed as-is).
    Works in absolute values throughout.
    Returns dict with workings and annotated list of (transaction, status, cumulative_at_selection).
    status is one of: 'IS', 'MUS', ''
    cumulative_at_selection is the running cumulative total at point of selection (for MUS items),
    None for IS items.
    """
    if not transactions:
        return None

    pm = abs(performance_materiality)
    abs_values = [abs(t.value) for t in transactions]
    total_population = sum(abs_values)

    # Pass 1: individually significant
    individually_significant = [t for t in transactions if abs(t.value) >= pm]
    remaining = [t for t in transactions if abs(t.value) < pm]

    is_total = sum(abs(t.value) for t in individually_significant)
    remaining_total = sum(abs(t.value) for t in remaining)
    is_refs = set(t.ref for t in individually_significant)

    # Pass 2: MUS on remaining
    if not remaining:
        raw_n = 0.0
        sample_size = 0
        interval = 0.0
        random_start = 0.0
        mus_refs = set()
        hit_points = []
    else:
        raw_n = (remaining_total * risk_factor) / pm
        sample_size = math.ceil(raw_n)

        if sample_size >= len(remaining):
            # Select entire remaining population
            interval = remaining_total / len(remaining) if remaining else 0
            random_start = 0.0
            mus_refs = set(t.ref for t in remaining)
            hit_points = []
        else:
            interval = remaining_total / sample_size
            rng = random.Random(RANDOM_SEED)
            random_start = rng.uniform(0, interval)
            hit_points = [random_start + (i * interval) for i in range(sample_size)]

            mus_refs = set()
            cumulative = 0.0
            hit_idx = 0

            for t in remaining:
                prev = cumulative
                cumulative += abs(t.value)
                while hit_idx < len(hit_points) and hit_points[hit_idx] <= cumulative:
                    if t.ref not in mus_refs:
                        mus_refs.add(t.ref)
                    hit_idx += 1
                if hit_idx >= len(hit_points):
                    break

    # Build annotated list with running cumulative total for the REMAINING population
    # (IS items don't participate in the MUS walk so their cumulative is N/A)
    cumulative = 0.0
    annotated = []
    for t in transactions:
        if t.ref in is_refs:
            annotated.append({
                "transaction": t,
                "status": "IS",
                "running_total": None,
            })
        else:
            cumulative += abs(t.value)
            status = "MUS" if t.ref in mus_refs else ""
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
        "interval": round(interval, 2),
        "random_start": round(random_start, 2),
        "random_seed": RANDOM_SEED,
        "hit_points": [round(h, 2) for h in hit_points],
        "mus_selected_count": len(mus_refs),
        "total_selected": len(is_refs) + len(mus_refs),
    }

    return {"workings": workings, "annotated": annotated}
