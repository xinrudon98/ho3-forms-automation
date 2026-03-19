from pathlib import Path

from src.ho3_forms.combo_generator import (
    build_combo_matrix,
    build_valid_combinations,
    load_forms,
    normalize_rules,
    save_combo_matrix,
    split_forms,
)


def main() -> None:
    base_dir = Path(__file__).resolve().parents[1]

    data_file = base_dir / "data" / "sample" / "forms_public_demo.csv"
    output_file = base_dir / "output" / "combo_matrix_output.xlsx"

    # Demo mutual exclusion rules (public-safe)
    rules = {
        "FORM_A": ["FORM_B"],
        "FORM_C": ["FORM_D"],
        "FORM_E": ["FORM_F"],
    }

    records = load_forms(data_file)
    mandatory, optional = split_forms(records)

    valid_combos = build_valid_combinations(
        mandatory=mandatory,
        optional=optional,
        exclusion_rules=normalize_rules(rules),
    )

    matrix_df = build_combo_matrix(valid_combos)
    save_combo_matrix(matrix_df, output_file)

    print(f"Generated {len(valid_combos)} valid combinations.")
    print(f"Saved matrix to: {output_file}")


if __name__ == "__main__":
    main()
