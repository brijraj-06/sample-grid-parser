
import streamlit as st
import pandas as pd

st.title("Grid to Paragraph Converter (English-Hindi Mix & English Only)")

uploaded_file = st.file_uploader("Upload your Excel/CSV file", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Uploaded Data")
    st.write(df)

    # Group and generate paragraphs
    grouped = df.groupby("Group")
    english_hindi_output = []
    english_only_output = []

    for group_name, group_df in grouped:
        para_hindi = group_name + ":
"
        para_eng = group_name + ":
"

        parts_hindi = []
        parts_eng = []

        for _, row in group_df.iterrows():
            name = row["English Name"]
            botanical = row["Botanical Name"]
            hindi = row["Hindi Name"]
            part = row["Part Used Short Form"]
            qty = f"{row['Quantity']} {row['Unit']}"

            parts_hindi.append(f"{name} ({botanical})/ {hindi} ({part}) {qty}")
            parts_eng.append(f"{name} ({botanical}) ({part}) {qty}")

        para_hindi += "; ".join(parts_hindi) + "."
        para_eng += "; ".join(parts_eng) + "."

        english_hindi_output.append(para_hindi)
        english_only_output.append(para_eng)

    st.subheader("English-Hindi Mix Paragraph")
    st.text("\n\n".join(english_hindi_output))

    st.subheader("English Only Paragraph")
    st.text("\n\n".join(english_only_output))
