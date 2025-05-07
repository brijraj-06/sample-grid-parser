
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="Grid to Paragraph Generator", layout="centered")
st.title("🧩 Grid to Paragraph & Table Generator")

uploaded_file = st.file_uploader("Upload MASTER_TEMPLATE_CLEAN.xlsx", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df = df.fillna("")
    df = df[df["English Name"] != ""]

    # Track main sections for grouping like "Amla Pishti", "Kwath Dravya", etc.
    group_name_col = []
    current_group = ""
    for value in df["English Name"]:
        if value.strip().endswith(":"):
            current_group = value.strip().rstrip(":")
        group_name_col.append(current_group)
    df["Group"] = group_name_col

    # Remove group headers from being processed as ingredient rows
    df = df[~df["English Name"].str.strip().str.endswith(":")]

    # Prepare Composition Table
    comp_data = [["Each 50g contains:", "", "", "", "", ""]]
    serial = 1
    for group in df["Group"].unique():
        comp_data.append([f"{group}:", "", "", "", "", ""])
        subset = df[df["Group"] == group]
        for _, row in subset.iterrows():
            comp_data.append([
                serial,
                row["English Name"],
                row["Hindi Name"],
                row["Part Used Full Form"],
                row["Quantity"],
                row["Proof Of Concept"]
            ])
            serial += 1
    comp_data.append(["Permitted Additives (-)/ -", "", "", "-", "QS", "-"])
    comp_data.append(["", "", "", "", "", ""])
    comp_data.append(["*Official Substitute", "", "", "", "", ""])

    # Create Paragraph Format Lines
    def generate_paragraph(mix=False):
        lines = []
        lines.append("PARAGRAPH FORMAT (ENGLISH-HINDI MIX)" if mix else "PARAGRAPH FORMAT (ENGLISH ONLY)")
        lines.append("English Transliterated Name (Botanical Name)/ हिंदी नाम (Part Used Short Form) Qty." if mix else "English Transliterated Name (Botanical Name) (Part Used Short Form) Qty.")
        lines.append("")
        lines.append("Each 50g contains:")

        for group in df["Group"].unique():
            lines.append(f"{group}:")
            subset = df[df["Group"] == group]
            item_lines = []
            for _, row in subset.iterrows():
                if mix:
                    line = f"{row['English Name']}/ {row['Hindi Name']} ({row['Part Used Full Form']}) {row['Quantity']}"
                else:
                    line = f"{row['English Name']} ({row['Part Used Full Form']}) {row['Quantity']}"
                item_lines.append(line)
            # Combine all in one line with semicolons
            combined = "; ".join(item_lines)
            if group == df["Group"].unique()[-1]:  # last group
                combined += " Permitted Additives QS."
            lines.append(combined)

        lines.append("*Official Substitute")
        return lines

    en_lines = generate_paragraph(mix=False)
    mix_lines = generate_paragraph(mix=True)

    # Save Paragraph to Excel
    def save_paragraph(path, lines, italic_index=1):
        wb = Workbook()
        ws = wb.active
        ws.title = "Paragraph"
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 120

        bold = Font(bold=True)
        italic = Font(italic=True)
        normal = Font()
        align = Alignment(wrap_text=True, vertical="top", horizontal="left")

        row = 1
        for i, text in enumerate(lines):
            cell = ws.cell(row=row, column=1, value=text)
            if i == 0:
                cell.font = bold
            elif i == italic_index:
                cell.font = italic
                row += 1
                ws.cell(row=row, column=1).value = ""
            else:
                cell.font = normal
            cell.alignment = align
            row += 1
        wb.save(path)

    # Save Table to Excel
    def save_table(path, data):
        wb = Workbook()
        ws = wb.active
        ws.title = "Composition Table"
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 8
        for col in "BCDEF":
            ws.column_dimensions[col].width = 30
        for row in data:
            ws.append(row)
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        wb.save(path)

    # Export all three files
    table_file = "COMPOSITION_TABLE_FINAL_FULLY_CLEANED.xlsx"
    en_file = "PARAGRAPH_ENGLISH_ONLY_FINAL_CLEANED.xlsx"
    mix_file = "PARAGRAPH_ENGLISH_HINDI_MIX_FINAL_CLEANED.xlsx"

    save_table(table_file, comp_data)
    save_paragraph(en_file, en_lines)
    save_paragraph(mix_file, mix_lines)

    with open(table_file, "rb") as f:
        st.download_button("Download Composition Table", f, file_name=table_file)

    with open(en_file, "rb") as f:
        st.download_button("Download Paragraph (English Only)", f, file_name=en_file)

    with open(mix_file, "rb") as f:
        st.download_button("Download Paragraph (English-Hindi Mix)", f, file_name=mix_file)
