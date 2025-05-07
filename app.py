
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="Grid to Paragraph Generator", layout="centered")

st.title("🔠 Master Grid to Composition Table & Paragraph Format Converter")

uploaded_file = st.file_uploader("Upload MASTER_TEMPLATE_CLEAN.xlsx", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Clean and prepare data
    df = df.dropna(subset=["English Name"])
    df = df.fillna("")
    df["Group"] = df["English Name"].where(df["Part Used Full Form"] == "", None)
    df["Group"] = df["Group"].fillna(method="ffill")

    # Initialize outputs
    table_data = []
    paragraph_en = []
    paragraph_mix = []

    # Build Composition Table Format
    table_data.append(["Each 50g contains:", "", "", "", "", ""])
    grouped = df.groupby("Group")

    serial = 1
    for group, items in grouped:
        table_data.append([group + ":", "", "", "", "", ""])
        for _, row in items.iterrows():
            if row["Part Used Full Form"] == "":
                continue
            entry = [
                serial,
                row["English Name"],
                row["Hindi Name"],
                row["Part Used Full Form"],
                row["Quantity"],
                row["Proof Of Concept"]
            ]
            table_data.append(entry)
            serial += 1

    table_data.append(["Additives:", "", "", "", "", ""])
    table_data.append([serial, "Permitted Additives (-)/ -", "", "-", "QS", "-"])
    table_data.append(["", "", "", "", "", ""])
    table_data.append(["*Official Substitute", "", "", "", "", ""])

    # Paragraphs
    paragraph_en.append("PARAGRAPH FORMAT (ENGLISH ONLY)")
    paragraph_en.append("English Transliterated Name (Botanical Name) (Part Used Short Form) Qty.")
    paragraph_en.append("")
    paragraph_en.append("Each 50g contains:")
    paragraph_mix.append("PARAGRAPH FORMAT (ENGLISH-HINDI MIX)")
    paragraph_mix.append("English Transliterated Name (Botanical Name)/ हिंदी नाम (Part Used Short Form) Qty.")
    paragraph_mix.append("")
    paragraph_mix.append("Each 50g contains:")

    def collect_lines(group_rows, is_mix=False):
        lines = []
        group_title = group_rows[0]["Group"]
        lines.append(f"{group_title}:")
        for r in group_rows:
            if r["Part Used Full Form"] == "":
                continue
            if is_mix:
                line = f"{r['English Name']}/ {r['Hindi Name']} ({r['Part Used Full Form']}) {r['Quantity']}"
            else:
                line = f"{r['English Name']} ({r['Part Used Full Form']}) {r['Quantity']}"
            lines.append(line)
        return lines

    for group, rows in grouped:
        group_rows = rows.to_dict(orient="records")
        lines_en = collect_lines(group_rows, is_mix=False)
        lines_mix = collect_lines(group_rows, is_mix=True)
        if "permitted additives" not in lines_en[-1].lower():
            lines_en[-1] += " Permitted Additives QS."
            lines_mix[-1] += " Permitted Additives QS."
        paragraph_en.extend(lines_en)
        paragraph_mix.extend(lines_mix)

    paragraph_en.append("*Official Substitute")
    paragraph_mix.append("*Official Substitute")

    # Write Excel outputs
    def save_excel(path, lines, sheetname, bold_idx=0, italic_idx=1):
        wb = Workbook()
        ws = wb.active
        ws.title = sheetname
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 120
        bold = Font(bold=True)
        italic = Font(italic=True)
        normal = Font()
        align = Alignment(horizontal="left", vertical="top", wrap_text=True)
        row = 1
        for i, line in enumerate(lines):
            cell = ws.cell(row=row, column=1, value=line)
            if i == bold_idx:
                cell.font = bold
            elif i == italic_idx:
                cell.font = italic
                row += 1
                ws.cell(row=row, column=1, value="")  # spacing
            else:
                cell.font = normal
            cell.alignment = align
            row += 1
        wb.save(path)

    def save_table(path, rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Composition Table Format"
        ws.sheet_view.showGridLines = False
        headers = ["S. No.", "English Transliterated Name (Botanical Name)/ हिंदी नाम", "Part Used Full Form", "Quantity", "Proof Of Concept"]
        ws.append(["Each 50g contains:"] + [""] * 5)
        font_bold = Font(bold=True)
        for row in rows:
            ws.append(row)
        ws.append(["*Official Substitute"] + [""] * 5)
        for col in "ABCDEF":
            ws.column_dimensions[col].width = 28
        for cell in ws["A"]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        wb.save(path)

    path_en = "PARAGRAPH_ENGLISH_ONLY_FINAL_CLEANED.xlsx"
    path_mix = "PARAGRAPH_ENGLISH_HINDI_MIX_FINAL_CLEANED.xlsx"
    path_table = "COMPOSITION_TABLE_FINAL_FULLY_CLEANED.xlsx"

    save_excel(path_en, paragraph_en, "ENGLISH ONLY")
    save_excel(path_mix, paragraph_mix, "ENGLISH-HINDI MIX")
    save_table(path_table, table_data)

    with open(path_en, "rb") as f:
        st.download_button("Download Paragraph (English Only)", f, file_name=path_en)

    with open(path_mix, "rb") as f:
        st.download_button("Download Paragraph (English-Hindi Mix)", f, file_name=path_mix)

    with open(path_table, "rb") as f:
        st.download_button("Download Composition Table Format", f, file_name=path_table)
