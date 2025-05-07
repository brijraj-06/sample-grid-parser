
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="Grid to Paragraph & Table Generator", layout="centered")
st.title("🧩 Visual Output Generator")

uploaded_file = st.file_uploader("Upload MASTER_TEMPLATE_CLEAN.xlsx", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df = df.fillna("")
    df = df[df["English Name"] != ""]

    group_col = []
    current_group = ""
    for val in df["English Name"]:
        if val.strip().endswith(":"):
            current_group = val.strip().rstrip(":")
        group_col.append(current_group)
    df["Group"] = group_col
    df = df[~df["English Name"].str.strip().str.endswith(":")]

    # Final composition table
    table_rows = [["Each 50g contains:", "", "", "", "", ""]]
    serial = 1
    for group in df["Group"].unique():
        table_rows.append([f"{group}:", "", "", "", "", ""])
        subset = df[df["Group"] == group]
        for _, row in subset.iterrows():
            table_rows.append([
                serial,
                row["English Name"],
                row["Hindi Name"],
                row["Part Used Full Form"],
                row["Quantity"],
                row["Proof Of Concept"]
            ])
            serial += 1
    table_rows.append(["", "Permitted Additives (-)/ -", "", "-", "QS", "-"])
    table_rows.append(["", "", "", "", "", ""])
    table_rows.append(["*Official Substitute", "", "", "", "", ""])

    # Paragraphs
    def get_paragraph_lines(mix=False):
        lines = []
        lines.append("PARAGRAPH FORMAT (ENGLISH-HINDI MIX)" if mix else "PARAGRAPH FORMAT (ENGLISH ONLY)")
        lines.append("English Transliterated Name (Botanical Name)/ हिंदी नाम (Part Used Short Form) Qty." if mix else "English Transliterated Name (Botanical Name) (Part Used Short Form) Qty.")
        lines.append("")
        lines.append("Each 50g contains:")
        for group in df["Group"].unique():
            lines.append(f"{group}:")
            subset = df[df["Group"] == group]
            entries = []
            for _, row in subset.iterrows():
                if mix:
                    entry = f"{row['English Name']}/ {row['Hindi Name']} ({row['Part Used Full Form']}) {row['Quantity']}".strip()
                else:
                    entry = f"{row['English Name']} ({row['Part Used Full Form']}) {row['Quantity']}".strip()
                entry = entry.replace(" ()", "")  # remove empty brackets
                entries.append(entry)
            combined = "; ".join(entries)
            if group == df["Group"].unique()[-1]:
                combined += " Permitted Additives QS."
            lines.append(combined)
        lines.append("*Official Substitute")
        return lines

    lines_en = get_paragraph_lines(mix=False)
    lines_mix = get_paragraph_lines(mix=True)

    def save_paragraph(path, lines, italic_line=1):
        wb = Workbook()
        ws = wb.active
        ws.title = "Paragraph"
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 120

        bold = Font(bold=True)
        italic = Font(italic=True)
        normal = Font()
        align = Alignment(horizontal="left", vertical="top", wrap_text=False)

        row = 1
        for i, line in enumerate(lines):
            cell = ws.cell(row=row, column=1, value=line)
            if i == 0:
                cell.font = bold
            elif i == italic_line:
                cell.font = italic
                row += 1
                ws.cell(row=row, column=1, value="")  # blank line
            else:
                cell.font = normal
            cell.alignment = align
            row += 1

        wb.save(path)

    def save_composition_table(path, rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Composition Table"
        ws.sheet_view.showGridLines = False
        widths = [8, 35, 30, 20, 10, 25]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[chr(64+i)].width = width

        # Apply border
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for row in rows:
            ws.append(row)

        for r in ws.iter_rows():
            for cell in r:
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                cell.border = border

        wb.save(path)

    # Save all 3 files
    table_xlsx = "COMPOSITION_TABLE_FINAL_FULLY_CLEANED.xlsx"
    eng_xlsx = "PARAGRAPH_ENGLISH_ONLY_FINAL_CLEANED.xlsx"
    mix_xlsx = "PARAGRAPH_ENGLISH_HINDI_MIX_FINAL_CLEANED.xlsx"

    save_composition_table(table_xlsx, table_rows)
    save_paragraph(eng_xlsx, lines_en)
    save_paragraph(mix_xlsx, lines_mix)

    with open(table_xlsx, "rb") as f:
        st.download_button("Download Composition Table", f, file_name=table_xlsx)

    with open(eng_xlsx, "rb") as f:
        st.download_button("Download Paragraph (English Only)", f, file_name=eng_xlsx)

    with open(mix_xlsx, "rb") as f:
        st.download_button("Download Paragraph (English-Hindi Mix)", f, file_name=mix_xlsx)
