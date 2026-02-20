import pandas as pd
import itertools

# Settings
EXCEL_PATH = r"K:\\Kmis_Public\\~Program Department\\1 - Programs\\02 - Active\\Knight - Direct Property Program\\8 - Knight Templates & Documents\\Dec Page Issuance\\2025_07 GENERATE DEC PAGE.xlsm"
SHEET_NAME = "HO3 FORMS"

# Mutually exclusive sets
mutual_exclusions = [
    {"FBR1HO (05 25)", "HO 04 27 05 11"}, # Fungi Exclusion vs Sublimit
    {"SPLS1HO (05 25)", "SPLE1HO (05 25)"}, # Swimming Pool Exclusion vs Sublimit
    {"LCEHO 05 24", "TE1HO (05 25)"}, # Liability vs Trampoline Exclusion
    {"LCEHO 05 24", "FRO1HO (05 25)"}, # Liability vs Farm And Ranch Exclusion
    {"LCEHO 05 24", "AIDE1HO (05 25)"}, # Liability vs Animal-Caused Injury Exclusion
    {"LCEHO 05 24", "EPL1HO (05 25)"}, # Liability vs Employment Practices Exclusion
    {"LCEHO 05 24", "PD1HO (05 25)"}, # Liability vs Punitive Damages Exclusion
    {"LCEHO 05 24", "SA1HO (05 25)"}, # Liability vs Sexual Abuse Exclusion
]

# Load form data from Excel
forms_df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=9)
form_list = forms_df['FORM CODE'].tolist()
form_order = {code: i for i, code in enumerate(forms_df['FORM CODE'].tolist())}
valid_forms_df = forms_df[forms_df['FORM CODE'].notna()]

# Get mandatory and optional forms
mandatory_forms = forms_df[(forms_df['CONTINGENT'] == "MANDATORY")]['FORM CODE'].tolist()
optional_forms = forms_df[(forms_df['CONTINGENT'] == "OPTIONAL")]['FORM CODE'].tolist()

# Generate all valid combinations
valid_combos = []
for bits in itertools.product([0, 1], repeat=len(optional_forms)):
    selected_optional = [optional_forms[i] for i in range(len(bits)) if bits[i] == 1]
    combined = set(mandatory_forms + selected_optional)

    # Check mutual exclusions
    valid = True
    for ex_set in mutual_exclusions:
        present = [item for item in ex_set if item in combined]
        if len(present) != 1:
            valid = False
            break

    if valid:
        ordered = sorted(combined, key=lambda x: form_order.get(x, float('inf')))
        valid_combos.append(ordered)

# Check the output
print(f"Generated {len(valid_combos)} valid combinations.\n")

# Create DataFrame
combo_cols = [f"COMBO {i+1}" for i in range(len(valid_combos))]
matrix_data = []

# Fill each combo column with True/False values
previous_values = [False] * len(valid_combos)

for form in form_list:
    row = {"FORM NAME": form if pd.notna(form) else ""}
    current_values = []
    for i, combo in enumerate(valid_combos):
        if pd.notna(form):
            value = form in combo
        else:
            value = previous_values[i]  # inherit from previous row
        row[f"COMBO {i+1}"] = value
        current_values.append(value)
    previous_values = current_values
    matrix_data.append(row)

# Final Output
matrix_df = pd.DataFrame(matrix_data)
print(matrix_df)
matrix_df.to_excel(r"K:\Kmis_Public\Clare\HO3 PDF MERGER\combo_matrix_output.xlsx", index=False)

