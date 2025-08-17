import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document
import io

# ---- HELPER: Fix duplicate columns ----
def fix_duplicate_columns(df):
    new_cols = []
    counts = {}
    for col in df.columns:
        if col in counts:
            counts[col] += 1
            new_cols.append(f"{col}_{counts[col]}")
        else:
            counts[col] = 0
            new_cols.append(col)
    df.columns = new_cols
    return df

# ---- Extract tables from PDF ----
def extract_tables_from_pdf(file):
    tables = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_tables()
            for tbl in extracted:
                df = pd.DataFrame(tbl[1:], columns=tbl[0])
                df = fix_duplicate_columns(df)
                tables.append(df)
    return tables

# Convert to DOCX
def table_to_docx(df):
    doc = Document()
    table = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
    # headers
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = col

    # data
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            table.cell(i+1, j).text = str(df.iloc[i, j])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ---- Streamlit UI ----
st.title("Jamabandi Table Extractor")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])
if uploaded_file:
    tables = extract_tables_from_pdf(uploaded_file)

    if not tables:
        st.error("No tables found in PDF. Make sure it has grid lines.")
    else:
        st.success(f"Found {len(tables)} table(s)")
        index = st.selectbox("Select Table to Preview / Export",
                             list(range(1, len(tables)+1)),
                             format_func=lambda x: f"Table {x}")

        selected_df = tables[index - 1]
        st.subheader(f"Table {index}")
        st.dataframe(selected_df)

        # Download as CSV
        csv_data = selected_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", csv_data,
                           file_name=f"table_{index}.csv", mime="text/csv")

        # Download as Excel
        excel_buf = io.BytesIO()
        selected_df.to_excel(excel_buf, index=False)
        excel_buf.seek(0)
        st.download_button("Download Excel", excel_buf.getvalue(),
                           file_name=f"table_{index}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Download as DOCX
        docx_buf = table_to_docx(selected_df)
        st.download_button("Download DOCX", docx_buf.getvalue(),
                           file_name=f"table_{index}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
