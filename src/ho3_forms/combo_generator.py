from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from typing import Dict, List, Sequence, Set

import pandas as pd


@dataclass(frozen=True)
class FormRecord:
    code: str
    category: str


def load_forms(csv_path: str | Path) -> List[FormRecord]:
    df = pd.read_csv(csv_path)

    required_columns = {"form_code", "category"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    records: List[FormRecord] = []
    for _, row in df.iterrows():
        code = str(row["form_code"]).strip()
        category = str(row["category"]).strip().upper()

        if not code:
            continue

        records.append(FormRecord(code=code, category=category))

    return records


def split_forms(records: Sequence[FormRecord]) -> tuple[list[str], list[str]]:
    mandatory = [r.code for r in records if r.category == "MANDATORY"]
    optional = [r.code for r in records if r.category == "OPTIONAL"]
    return mandatory, optional


def normalize_rules(rules: Dict[str, Sequence[str]]) -> Dict[str, Set[str]]:
    normalized: Dict[str, Set[str]] = {}
    for key, values in rules.items():
        normalized[str(key)] = {str(v) for v in values}
    return normalized


def is_valid_combo(combo: Sequence[str], exclusion_rules: Dict[str, Set[str]]) -> bool:
    combo_set = set(combo)

    for form in combo:
        blocked = exclusion_rules.get(form, set())
        if combo_set.intersection(blocked):
            return False

    return True


def build_valid_combinations(
    mandatory: Sequence[str],
    optional: Sequence[str],
    exclusion_rules: Dict[str, Set[str]],
) -> List[List[str]]:
    valid_combos: List[List[str]] = []

    for size in range(len(optional) + 1):
        for optional_subset in combinations(optional, size):
            combo = list(mandatory) + list(optional_subset)

            if is_valid_combo(combo, exclusion_rules):
                valid_combos.append(sorted(combo))

    return valid_combos


def build_combo_matrix(valid_combos: Sequence[Sequence[str]]) -> pd.DataFrame:
    all_forms = sorted({form for combo in valid_combos for form in combo})
    rows = []

    for index, combo in enumerate(valid_combos, start=1):
        combo_set = set(combo)

        row = {"combo_id": f"COMBO {index}"}

        for form in all_forms:
            row[form] = "X" if form in combo_set else ""

        rows.append(row)

    return pd.DataFrame(rows)


def save_combo_matrix(df: pd.DataFrame, output_path: str | Path) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
