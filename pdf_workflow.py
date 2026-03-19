from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from ho3_forms.pdf_workflow import (
    copy_demo_templates_message,
    export_schedule_pdfs_with_excel,
    merge_supplemental_pdfs,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate schedule PDFs and merge supplemental PDFs")
    parser.add_argument("--config", required=True, help="Path to YAML config file")
    parser.add_argument("--skip-export", action="store_true", help="Skip Excel schedule PDF export")
    parser.add_argument("--skip-merge", action="store_true", help="Skip Acrobat PDF merge")
    args = parser.parse_args()

    if not Path(args.config).exists():
        raise FileNotFoundError(args.config)

    if not args.skip_export:
        try:
            export_schedule_pdfs_with_excel(args.config)
            print("Schedule PDF export finished.")
        except Exception as exc:
            print(f"Schedule export not completed: {exc}")
            print(copy_demo_templates_message())

    if not args.skip_merge:
        try:
            merge_supplemental_pdfs(args.config)
            print("Supplemental PDF merge finished.")
        except Exception as exc:
            print(f"PDF merge not completed: {exc}")
            print(copy_demo_templates_message())


if __name__ == "__main__":
    main()
