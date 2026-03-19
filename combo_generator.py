from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from ho3_forms.combo_generator import build_combo_artifacts


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a forms combo matrix")
    parser.add_argument("--config", required=True, help="Path to YAML config file")
    args = parser.parse_args()

    artifacts = build_combo_artifacts(args.config)
    print(f"Generated {len(artifacts.valid_combos)} valid combinations.")
    print(f"Matrix rows: {len(artifacts.matrix_df)}")
    print("Saved combo matrix and combo JSON.")


if __name__ == "__main__":
    main()
