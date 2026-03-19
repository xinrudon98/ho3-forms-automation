from __future__ import annotations

import itertools
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

from .config import load_json, load_yaml
from .utils import ensure_parent

REQUIRED_COLUMNS = {"FORM CODE", "CONTINGENT"}


@dataclass
class ComboArtifacts:
    forms_df: pd.DataFrame
    matrix_df: pd.DataFrame
    valid_combos: list[list[str]]



def load_forms_table(forms_file: str | Path, sheet_name: str | None = None, header_row: int = 0) -> pd.DataFrame:
    forms_file = Path(forms_file)
    suffix = forms_file.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        df = pd.read_excel(forms_file, sheet_name=sheet_name, header=header_row)
    elif suffix == ".csv":
        df = pd.read_csv(forms_file)
    else:
        raise ValueError(f"Unsupported input file type: {suffix}")

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    if "FORM NAME" not in df.columns:
        df["FORM NAME"] = df["FORM CODE"]

    return df



def normalize_code_list(values: Iterable[str]) -> list[str]:
    return [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]



def is_combo_valid(combo: set[str], rules: dict) -> bool:
    for group in rules.get("exclusion_groups", []):
        present_count = sum(1 for item in group if item in combo)
        if present_count != 1:
            return False

    for group in rules.get("require_together", []):
        present_count = sum(1 for item in group if item in combo)
        if 0 < present_count < len(group):
            return False

    for rule in rules.get("if_present_require_any", []):
        trigger = rule["if_present"]
        required_any = rule.get("require_any_of", [])
        if trigger in combo and not any(item in combo for item in required_any):
            return False

    return True



def generate_valid_combos(forms_df: pd.DataFrame, rules: dict) -> list[list[str]]:
    valid_forms_df = forms_df[forms_df["FORM CODE"].notna()].copy()
    form_order = {code: i for i, code in enumerate(normalize_code_list(valid_forms_df["FORM CODE"].tolist()))}

    mandatory_forms = normalize_code_list(
        valid_forms_df.loc[valid_forms_df["CONTINGENT"].astype(str).str.upper() == "MANDATORY", "FORM CODE"].tolist()
    )
    optional_forms = normalize_code_list(
        valid_forms_df.loc[valid_forms_df["CONTINGENT"].astype(str).str.upper() == "OPTIONAL", "FORM CODE"].tolist()
    )

    valid_combos: list[list[str]] = []
    for bits in itertools.product([0, 1], repeat=len(optional_forms)):
        selected_optional = [optional_forms[i] for i, bit in enumerate(bits) if bit == 1]
        combined = set(mandatory_forms + selected_optional)
        if is_combo_valid(combined, rules):
            ordered = sorted(combined, key=lambda x: form_order.get(x, 10**9))
            valid_combos.append(ordered)

    return valid_combos



def build_matrix(forms_df: pd.DataFrame, valid_combos: list[list[str]]) -> pd.DataFrame:
    form_list = forms_df["FORM CODE"].tolist()
    matrix_data: list[dict] = []
    previous_values = [False] * len(valid_combos)

    for form in form_list:
        display_name = form if pd.notna(form) else ""
        row = {"FORM NAME": display_name}
        current_values: list[bool] = []
        for i, combo in enumerate(valid_combos, start=1):
            value = (form in combo) if pd.notna(form) else previous_values[i - 1]
            row[f"COMBO {i}"] = value
            current_values.append(value)
        previous_values = current_values
        matrix_data.append(row)

    return pd.DataFrame(matrix_data)



def save_combo_artifacts(matrix_df: pd.DataFrame, valid_combos: list[list[str]], combo_matrix_file: str | Path, combo_json_file: str | Path) -> None:
    combo_matrix_path = ensure_parent(combo_matrix_file)
    combo_json_path = ensure_parent(combo_json_file)

    matrix_df.to_excel(combo_matrix_path, index=False)
    with open(combo_json_path, "w", encoding="utf-8") as f:
        json.dump(valid_combos, f, indent=2)



def build_combo_artifacts(config_path: str | Path) -> ComboArtifacts:
    config_path = Path(config_path)
    settings = load_yaml(config_path)
    base_dir = config_path.parent.parent if config_path.parent.name == "config" else config_path.parent
    rules_path = Path(settings["rules_file"])
    if not rules_path.is_absolute():
        rules_path = base_dir / rules_path
    rules = load_json(rules_path)
    input_cfg = settings["input"]
    output_cfg = settings["output"]

    forms_file = Path(input_cfg["forms_file"])
    if not forms_file.is_absolute():
        forms_file = base_dir / forms_file
    combo_matrix_file = Path(output_cfg["combo_matrix_file"])
    if not combo_matrix_file.is_absolute():
        combo_matrix_file = base_dir / combo_matrix_file
    combo_json_file = Path(output_cfg["combo_json_file"])
    if not combo_json_file.is_absolute():
        combo_json_file = base_dir / combo_json_file

    forms_df = load_forms_table(
        forms_file=forms_file,
        sheet_name=input_cfg.get("sheet_name"),
        header_row=input_cfg.get("header_row", 0),
    )
    valid_combos = generate_valid_combos(forms_df, rules)
    matrix_df = build_matrix(forms_df, valid_combos)
    save_combo_artifacts(matrix_df, valid_combos, combo_matrix_file, combo_json_file)
    return ComboArtifacts(forms_df=forms_df, matrix_df=matrix_df, valid_combos=valid_combos)
