# ho3-forms-automation

Public-safe demo repo for form-combination generation and PDF package assembly.

This version is designed for a **public GitHub repo**:
- no internal network paths
- no company templates or signed PDFs
- configurable with JSON/YAML
- Windows Excel / Acrobat automation kept optional

## What it does

1. Read a forms table from Excel or CSV.
2. Build valid combinations from mandatory + optional forms.
3. Apply inclusion rules and exclusion groups from config.
4. Export a combo matrix to Excel.
5. Optionally generate schedule PDFs from an Excel template.
6. Optionally merge supplemental PDFs.

## Project structure

```text
src/ho3_forms/
  combo_generator.py
  pdf_workflow.py
  config.py
  utils.py
scripts/
  generate_combo_matrix.py
  generate_schedule_pdfs.py
config/
  public_demo_settings.yaml
  rules.public_demo.json
data/sample/
  forms_public_demo.csv
```

## Install

```bash
pip install -r requirements.txt
```

## Run combo generation

```bash
python scripts/generate_combo_matrix.py --config config/public_demo_settings.yaml
```

Output goes to `output/combo_matrix_output.xlsx` by default.

## Run PDF workflow

```bash
python scripts/generate_schedule_pdfs.py --config config/public_demo_settings.yaml
```

This step is **Windows-only** and requires:
- Microsoft Excel desktop
- Python package `xlwings`
- Adobe Acrobat desktop if you want PDF merging with COM

## Public repo safety notes

Keep these out of the public repo:
- real templates
- real signed PDFs
- real internal paths
- real business-only form code mappings

Instead, put them in ignored local folders and point to them through your local config file.
