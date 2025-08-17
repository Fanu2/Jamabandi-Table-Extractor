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

    # Auto-merge if multiple segments detected
    if len(tables) > 1:
        base_cols = tables[0].columns
        merged_frames = [tables[0]]

        for t in tables[1:]:
            # Fix missing or extra columns
            if len(t.columns) < len(base_cols):
                for i in range(len(base_cols) - len(t.columns)):
                    t[f"_padding_{i}"] = ""
            elif len(t.columns) > len(base_cols):
                t = t.iloc[:, :len(base_cols)]
            t.columns = base_cols
            merged_frames.append(t)

        merged_df = pd.concat(merged_frames, ignore_index=True)
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
            table.cell(i + 1, j).text = str(df.iloc[i, j])
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ---------------- Streamlit UI ------------------

st.title("Jamabandi Table Extractor (Auto Merge Tables)")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    tables = extract_tables_from_pdf(uploaded_file)

    if not tables:
        st.error("No tables found in PDF. Check if PDF has proper grid lines.")
    else:
        st.success(f"Extracted {len(tables)} table(s) — combined automatically")

        df = tables[0]
        st.subheader("Preview of Combined Table")
        st.dataframe(df)

        # Download CSV
        csv_data = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", csv_data, file_name="table.csv", mime="text/csv")

        # Download Excel
        excel_buf = io.BytesIO()
        df.to_excel(excel_buf, index=False)
        excel_buf.seek(0)
        st.download_button("Download Excel", excel_buf.getvalue(),
                           file_name="table.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Download DOCX
        docx_buf = table_to_docx(df, filename="table.docx")
        st.download_button("Download DOCX", docx_buf.getvalue(),
                           file_name="table.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
