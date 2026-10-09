import streamlit as st
import pdfplumber
import pandas as pd
import re

# Page Configuration
st.set_page_config(
    page_title="Store Signage Schedule",
    page_icon="📋",
    layout="wide"
)

st.title("Store Signage Schedule Extractor")
st.markdown("Drag and drop your store PDF package below to instantly get a clean list of signage codes, positions, and quantities.")

# File uploader with built-in drag-and-drop support
uploaded_file = st.file_uploader(
    "Upload or Drag & Drop Store PDF Document", 
    type=["pdf"],
    help="Drag and drop your PDF file here, or click Browse files."
)

if uploaded_file is not None:
    with st.spinner("Extracting signage schedule..."):
        signage_tables = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        df_table = pd.DataFrame(table)
                        header_text = " ".join([str(cell) for row in df_table.head(2).values for cell in row if cell])
                        if any(keyword in header_text.upper() for keyword in ["TK CODE", "SIGNAGE", "FAMILY", "SCHEDULE"]):
                            signage_tables.append((page_num + 1, df_table))

        if signage_tables:
            master_df_list = []
            for page_num, df in signage_tables:
                if len(df) > 1:
                    raw_headers = [str(c).strip() if c is not None and str(c).strip() != "" else f"Col_{i}" for i, c in enumerate(df.iloc[0])]
                    cols = pd.Series(raw_headers)
                    for dup in cols[cols.duplicated()].unique():
                        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
                    
                    df_clean = df[1:].copy()
                    df_clean.columns = cols
                    df_clean = df_clean.reset_index(drop=True)
                    master_df_list.append(df_clean)
            
            if master_df_list:
                master_df = pd.concat(master_df_list, ignore_index=True, join='outer')
                
                # Identify key columns dynamically
                cols_lower = {c.lower(): c for c in master_df.columns}
                code_col = next((cols_lower[c] for c in cols_lower if 'code' in c), master_df.columns[2] if len(master_df.columns) > 2 else None)
                level_col = next((cols_lower[c] for c in cols_lower if 'level' in c), master_df.columns[0] if len(master_df.columns) > 0 else None)
                count_col = next((cols_lower[c] for c in cols_lower if 'count' in c or 'qty' in c or 'quantity' in c), master_df.columns[3] if len(master_df.columns) > 3 else None)
                desc_col = next((cols_lower[c] for c in cols_lower if 'family' in c or 'type' in c or 'description' in c), master_df.columns[1] if len(master_df.columns) > 1 else None)

                simplified_data = []
                for _, row in master_df.iterrows():
                    code = str(row[code_col]).strip() if code_col and pd.notna(row[code_col]) else ""
                    if code and code.upper() not in ["TK CODE", "SIGNAGE CODE", "COL_2", ""]:
                        level = str(row[level_col]).strip() if level_col and pd.notna(row[level_col]) else "L00"
                        count = str(row[count_col]).strip() if count_col and pd.notna(row[count_col]) else "1"
                        desc = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else ""
                        
                        try:
                            qty_val = int(float(count))
                        except:
                            qty_val = 1

                        simplified_data.append({
                            "Signage Code": code,
                            "Description / Size": desc,
                            "Position / Level": level,
                            "Quantity": qty_val
                        })
                
                if simplified_data:
                    df_simple = pd.DataFrame(simplified_data).drop_duplicates(subset=["Signage Code", "Position / Level"])
                    
                    st.success(f"Successfully extracted {len(df_simple)} signage items!")
                    
                    # Totals Metric
                    total_qty = df_simple["Quantity"].sum()
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total Signage Quantity", total_qty)
                    with col2:
                        st.metric("Unique Signage Codes", len(df_simple))
                    
                    st.markdown("---")
                    st.subheader("📋 Simplified Signage List")
                    st.dataframe(df_simple, use_container_width=True)
                    
                    # CSV Download
                    csv_data = df_simple.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Simplified Schedule as CSV",
                        data=csv_data,
                        file_name="simplified_signage_schedule.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("Could not isolate signage fields cleanly. Showing raw table:")
                    st.dataframe(master_df, use_container_width=True)
        else:
            st.warning("No signage schedule tables found in this PDF.")
