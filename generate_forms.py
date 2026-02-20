import os
import shutil

base_path = r"K:\Kmis_Public\~Program Department\1 - Programs\02 - Active\Knight - Direct Property Program\15 - Systems\2025 07 15"

for i in range(1, 2049):
    folder_name = '_' + str(i).zfill(4)
    folder_path = os.path.join(base_path, folder_name)
    
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    else:
        print(f"Already exists: {folder_path}")

from openpyxl import load_workbook
import pandas as pd
import numpy as np

file_path = r"K:\Kmis_Public\~Program Department\1 - Programs\02 - Active\Knight - Direct Property Program\15 - Systems\2025 07 15\_TEMPLATE.xlsx"

wb = load_workbook(file_path)
ws = wb['MASTER']

df = pd.read_excel(file_path, sheet_name='MASTER', usecols='E:E', skiprows=9)

from generate_combos_matrix import matrix_df
df = matrix_df.copy().astype(str)
matrix_df.to_excel(r"K:\Kmis_Public\Clare\combo_matrix_output.xlsx", index=False)

df.head(10)

# CREATE SCHEDULE OF FORMS
import xlwings as xw
import os
import json

base_path = r"K:\Kmis_Public\~Program Department\1 - Programs\02 - Active\Knight - Direct Property Program\15 - Systems\2025 07 15"
excel_path = os.path.join(base_path, "_TEMPLATE.xlsx")
supp_forms_path = r"K:\Kmis_Public\~Program Department\1 - Programs\02 - Active\Knight - Direct Property Program\15 - Systems\2025 07 15\_FORMS_SIGNATURES"

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

# TODO CHANGE 6 to 2048
for i in range(1, 2049):
    folder_name = '_' + str(i).zfill(4)
    
    wb = app.books.open(excel_path)
    ws = wb.sheets['MASTER']

    combo_col = f"COMBO {i}"
    if combo_col not in df.columns:
        raise KeyError(f"{combo_col} not found in combo matrix")

    values_to_write = [[j] for j in df[combo_col].tolist()[:56]]
    ws.range('D11:D66').value = values_to_write
    
    print_ws = wb.sheets['SCHEDULE OF FORMS']
    pdf_path = os.path.join(base_path, folder_name, '_Q_SOF.pdf')

    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    print_ws.api.ExportAsFixedFormat(0, pdf_path)

    wb.close()


    if i % 10 == 0:
        print(f"{i}th iteration")

app.quit()

# ADDITIONAL FORMS THAT HAVE TO BE GENERATED
import os
import win32com.client

base_path = r"K:\Kmis_Public\~Program Department\1 - Programs\02 - Active\Knight - Direct Property Program\15 - Systems\2025 07 15"
supp_forms_path = r"K:\Kmis_Public\~Program Department\1 - Programs\02 - Active\Knight - Direct Property Program\15 - Systems\2025 07 15\_FORMS_SIGNATURES"

# TODO CHANGE 6 to 2048
for i in range(1, 2049):
    combo_col = f"COMBO {i}"
    folder_name = '_' + str(i).zfill(4)

    supp_forms = {
        "IACITV 06 25 - SIG.pdf": "True",
        "Supplemental Application.pdf": str(df.loc[34, combo_col] == "True"),
        "Supp App w Liability.pdf": str(df.loc[34, combo_col] == "False"),
        "BCD1HO (05 25) - SIG.pdf": "True",
        "WDSE1HO (05 25) - SIG.pdf": "True",
        "SPLS1HO (05 25) - SIG.pdf": df.loc[42, combo_col],
        "SPLE1HO (05 25) - SIG.pdf": df.loc[43, combo_col],
        "ATSPEHO 05 25 - SIG.pdf":  df.loc[52, combo_col],
        "IL 12 01 11 85 - SIG.pdf": df.loc[53, combo_col],
    }

    acrobat_app = win32com.client.Dispatch("AcroExch.App")
    av_doc = win32com.client.Dispatch("AcroExch.AVDoc")

    first_pdf = None
    for x, y in supp_forms.items():
        if y == "True":
            first_pdf = os.path.join(supp_forms_path, x)
            break

    if not first_pdf:
        raise Exception("No PDFs to merge")

    # Open first PDF in AVDoc
    if not av_doc.Open(first_pdf, ""):
        raise Exception(f"Failed to open {first_pdf}")

    base_pd_doc = av_doc.GetPDDoc()

    # Append pages from remaining PDFs
    for filename, include in supp_forms.items():
        if include == "True" and os.path.join(supp_forms_path, filename) != first_pdf:
            pdf_path = os.path.join(supp_forms_path, filename)

            append_doc = win32com.client.Dispatch("AcroExch.PDDoc")
            if not append_doc.Open(pdf_path):
                print(f"Failed to open {pdf_path}")
                continue

            num_pages = append_doc.GetNumPages()
            base_pd_doc.InsertPages(base_pd_doc.GetNumPages() - 1, append_doc, 0, num_pages, True)
            append_doc.Close()

    output_path = os.path.join(base_path, folder_name, "_Q_AF.pdf")
    if not base_pd_doc.Save(1, output_path):
        raise Exception(f"Failed to save merged PDF to {output_path}")

    av_doc.Close(True)
    acrobat_app.Exit()

    if i % 10 == 0:
        print(f"{i}th iteration")
