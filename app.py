
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Pravek Composition Formatter", layout="centered")

st.title("🌿 Pravek Composition Formatter")

uploaded_file = st.file_uploader("Upload the COMPOSITION ELEMENTS Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, skiprows=5)
    df = df.dropna(axis=1, how='all')
    df.columns = [
        "#", "Group", "English Name", "Botanical Name", "Hindi Name",
        "Part Used Full Form", "Part Used Short Form", "Quantity",
        "Unit", "Book", "Page No.", "Proof Of Concept", "Extra 1", "Extra 2"
    ]
    df = df.drop(columns=["#", "Extra 1", "Extra 2", "Book", "Page No."])
    df = df[df["English Name"].notna()].reset_index(drop=True)

    # 1. COMPOSITION TABLE FORMAT
    df_table = df.copy()
    df_table["Name Combined"] = df_table["English Name"] + " (" + df_table["Botanical Name"] + ")/ " + df_table["Hindi Name"]
    df_table_output = df_table[[
        "Name Combined", "Part Used Full Form", "Quantity", "Unit", "Proof Of Concept"
    ]]
    df_table_output.columns = [
        "English Transliterated Name (Botanical Name)/ हिंदी नाम",
        "Part Used Full Form", "Quantity", "Unit", "Proof Of Concept"
    ]

    st.subheader("📥 Download: COMPOSITION TABLE FORMAT")
    table_buffer = io.BytesIO()
    df_table_output.to_excel(table_buffer, index=False)
    st.download_button("Download Table Format", table_buffer.getvalue(), file_name="COMPOSITION_TABLE_FORMAT.xlsx")

    # 2. PARAGRAPH FORMATS
    from collections import defaultdict

    para_hindi = "Each 50g contains:

"
    para_eng = "Each 50g contains:

"

    grouped = df.groupby("Group")

    for group, group_df in grouped:
        para_hindi += f"{group}:
"
        para_eng += f"{group}:
"

        quantity_map = defaultdict(list)
        for _, row in group_df.iterrows():
            qty = f"{row['Quantity']}{row['Unit']}"
            quantity_map[qty].append(row)

        hindi_parts = []
        eng_parts = []

        for qty, rows in quantity_map.items():
            if len(rows) == 1:
                row = rows[0]
                hindi_parts.append(f"{row['English Name']} ({row['Botanical Name']})/ {row['Hindi Name']} ({row['Part Used Short Form']}) {qty}")
                eng_parts.append(f"{row['English Name']} ({row['Botanical Name']}) ({row['Part Used Short Form']}) {qty}")
            else:
                hp = ", ".join([f"{r['English Name']} ({r['Botanical Name']})/ {r['Hindi Name']} ({r['Part Used Short Form']})" for r in rows])
                ep = ", ".join([f"{r['English Name']} ({r['Botanical Name']}) ({r['Part Used Short Form']})" for r in rows])
                hindi_parts.append(hp + f" each {qty}")
                eng_parts.append(ep + f" each {qty}")

        para_hindi += "; ".join(hindi_parts) + ".

"
        para_eng += "; ".join(eng_parts) + ".

"

    st.subheader("📥 Download: PARAGRAPH FORMAT (ENGLISH-HINDI MIX)")
    df_hindi = pd.DataFrame({"PARAGRAPH FORMAT (ENGLISH-HINDI MIX)": para_hindi.strip().split("\n")})
    hindi_buffer = io.BytesIO()
    df_hindi.to_excel(hindi_buffer, index=False)
    st.download_button("Download Hindi Mix Paragraph", hindi_buffer.getvalue(), file_name="PARAGRAPH_FORMAT_ENGLISH_HINDI_MIX.xlsx")

    st.subheader("📥 Download: PARAGRAPH FORMAT (ENGLISH ONLY)")
    df_eng = pd.DataFrame({"PARAGRAPH FORMAT (ENGLISH ONLY)": para_eng.strip().split("\n")})
    eng_buffer = io.BytesIO()
    df_eng.to_excel(eng_buffer, index=False)
    st.download_button("Download English Only Paragraph", eng_buffer.getvalue(), file_name="PARAGRAPH_FORMAT_ENGLISH_ONLY.xlsx")

    st.subheader("🔍 Preview Paragraphs")
    st.text_area("ENGLISH-HINDI MIX", para_hindi.strip(), height=200)
    st.text_area("ENGLISH ONLY", para_eng.strip(), height=200)
