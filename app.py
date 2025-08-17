import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document
import io

def extract_tables_from_pdf(file):
    tables = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_tables()
            for tbl in extracted:
                df = pd.DataFrame(tbl[1:], columns=tbl[0])
                tables.append(df)

    # Auto-merge if all tables have same columns
    if len(tables) > 1:
        first_cols = tables[0].columns
        can_join = all(list(t.columns) == list(first_cols) for t in tables)
        if can_join:
            merged_df = pd.concat(tables, ignore_index=True)
            return [merged_df]
    return tables

def table_to_docx(df, filename="table.docx"):
    doc = Document()
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
    # Header
    for j, column in enumerate(df.columns):
        table.cell(0, j).text = column
    # Data
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            table.cell(i+1, j).text = str(df.iloc[i, j])
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ---------------- Streamlit UI ------------------

st.title("Jamabandi Table Extractor")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    tables = extract_tables_from_pdf(uploaded_file)

    if not tables:
        st.error("No tables found in PDF. Check if PDF has proper grid lines.")
    else:
        st.success(f"Extracted {len(tables)} table(s). (Merged into one if possible)")
        selected_df = tables[0]

        st.subheader("Preview of Table")
        st.dataframe(selected_df)

        # Download buttons
        csv_data = selected_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", csv_data,
                           file_name="table.csv", mime="text/csv")

        excel_buf = io.BytesIO()
        selected_df.to_excel(excel_buf, index=False)
        excel_buf.seek(0)
        st.download_button("Download Excel", excel_buf.getvalue(),
                           file_name="table.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        docx_buf = table_to_docx(selected_df, filename="table.docx")
        st.download_button("Download DOCX", docx_buf.getvalue(),
                           file_name="table.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
