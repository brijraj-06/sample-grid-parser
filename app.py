
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from io import BytesIO

# Style definitions
bold_font = Font(bold=True)
italic_font = Font(italic=True)
wrap_alignment = Alignment(wrap_text=True, vertical="top")
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

def generate_composition_table(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "COMPOSITION TABLE FORMAT"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 50

    # Headers
    ws.merge_cells('A1:E1')
    ws['A1'] = "1. COMPOSITION TABLE FORMAT"
    ws['A1'].font = bold_font
    ws.merge_cells('A2:E2')
    ws['A2'] = "Each 50g contains:"

    ws.append(["#", "English Transliterated Name (Botanical Name)/ हिंदी नाम", "Part Used Full Form", "Quantity", "Proof Of Concept"])
    for col in range(1, 6):
        c = ws.cell(row=3, column=col)
        c.font = bold_font
        c.border = thin_border
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    index = 1
    row_idx = 4
    for _, row in df.iterrows():
        if pd.isna(row["#"]) and str(row["English Name"]).strip().endswith(":"):
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
            c = ws.cell(row=row_idx, column=1, value=row["English Name"])
            c.font = bold_font
            c.alignment = wrap_alignment
            c.border = thin_border
        else:
            ws.cell(row=row_idx, column=1, value=index).border = thin_border
            ws.cell(row=row_idx, column=2, value=row["English Name"]).border = thin_border
            ws.cell(row=row_idx, column=3, value=row["Part Used Full Form"]).border = thin_border
            ws.cell(row=row_idx, column=4, value=row["Quantity"]).border = thin_border
            ws.cell(row=row_idx, column=5, value=row["Proof Of Concept"]).border = thin_border
            for col in range(1, 6):
                ws.cell(row=row_idx, column=col).alignment = wrap_alignment
            index += 1
        row_idx += 1

    ws.merge_cells(start_row=row_idx + 1, start_column=1, end_row=row_idx + 1, end_column=5)
    ws.cell(row=row_idx + 1, column=1, value="*Official Substitute")

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def generate_paragraph_excel(title, subtitle, lines):
    wb = Workbook()
    ws = wb.active
    ws.title = title.replace(" ", "_")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 150

    row = 1
    ws.cell(row=row, column=1, value=title).font = bold_font
    row += 1
    ws.cell(row=row, column=1, value=subtitle).font = italic_font
    row += 2

    for line in lines:
        ws.cell(row=row, column=1, value=line).alignment = wrap_alignment
        row += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit app
st.title("Master Grid → 3 Output Generator")

uploaded_file = st.file_uploader("Upload the Master Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if not all(col in df.columns for col in ["#", "English Name", "Part Used Full Form", "Quantity", "Proof Of Concept"]):
        st.error("Uploaded file must contain the columns: #, English Name, Part Used Full Form, Quantity, Proof Of Concept")
    else:
        st.success("File uploaded and validated.")

        # Generate all three outputs
        comp_file = generate_composition_table(df)

        hindi_lines = [
            "Each 50g contains:",
            "Amla Pishti:",
            "Abhrak Bhasma (RT)/ अभ्रक भस्म 1g; Bel (Aegle marmelos)/ बेल (St. Bk.), Choti Elaichi (Elettaria cardamomum)/ छोटी इलायची (Sd.) each 1.75mg.",
            "Kwath Dravya (Coarse Powders Of):",
            "Giloy (Tinospora cordifolia)/ गिलोय (St.) 20mg; Kakoli (Withania somnifera*)/ अश्वगंधा (Rt.) 10mg; Munnaka (Vitis vinifera)/ मुनक्का (Fr.), Badi Kateri (Solanum indicum)/ बड़ी कटेरी (Pl.) each 5mg. Permitted Additives QS.",
            "*Official Substitute"
        ]
        hindi_file = generate_paragraph_excel("PARAGRAPH FORMAT (ENGLISH-HINDI MIX)",
                                              "English Transliterated Name (Botanical Name)/ हिंदी नाम (Part Used Short Form) Qty.",
                                              hindi_lines)

        eng_lines = [
            "Each 50g contains:",
            "Amla Pishti:",
            "Abhrak Bhasma (RT) 1g; Bel (Aegle marmelos) (St. Bk.), Choti Elaichi (Elettaria cardamomum) (Sd.) each 1.75mg.",
            "Kwath Dravya (Coarse Powders Of):",
            "Giloy (Tinospora cordifolia) (St.) 20mg; Kakoli (Withania somnifera*) (Rt.) 10mg; Munnaka (Vitis vinifera) (Fr.), Badi Kateri (Solanum indicum) (Pl.) each 5mg. Permitted Additives QS.",
            "*Official Substitute"
        ]
        eng_file = generate_paragraph_excel("PARAGRAPH FORMAT (ENGLISH ONLY)",
                                            "English Transliterated Name (Botanical Name) (Part Used Short Form) Qty.",
                                            eng_lines)

        st.download_button("Download Composition Table Format", comp_file, file_name="COMPOSITION_TABLE_FORMAT.xlsx")
        st.download_button("Download Paragraph Format (English-Hindi Mix)", hindi_file, file_name="PARAGRAPH_FORMAT_ENGLISH_HINDI_MIX.xlsx")
        st.download_button("Download Paragraph Format (English Only)", eng_file, file_name="PARAGRAPH_FORMAT_ENGLISH_ONLY.xlsx")
