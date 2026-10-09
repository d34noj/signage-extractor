import streamlit as st
import pdfplumber
import pandas as pd
import re

st.set_page_config(
    page_title="Store Signage Schedule",
    page_icon="📋",
    layout="wide"
)

st.title("Store Signage Schedule Extractor")
st.markdown("Drag and drop your store PDF package below to instantly get a clean list of signage codes, descriptions, positions, and quantities.")

uploaded_file = st.file_uploader(
    "Upload or Drag & Drop Store PDF Document", 
    type=["pdf"],
    help="Drag and drop your PDF file here, or click Browse files."
)

if uploaded_file is not None:
    with st.spinner("Extracting master signage schedule..."):
        signage_tables = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        df_table = pd.DataFrame(table)
                        header_snippet = " ".join([str(cell) for row in df_table.head(3).values for cell in row if cell]).upper()
                        
                        # Specifically target the main TK_Signage table, ignoring B1-B6 finish schedules
                        if "FAMILY" in header_snippet and "TK CODE" in header_snippet and "COUNT" in header_snippet:
                            if not any(sched in header_snippet for sched in ["B1", "B2", "B3", "B4", "B5", "B6", "DUVKS"]):
                                signage_tables.append((page_num + 1, df_table))

        if signage_tables:
            master_df_list = []
            for page_num, df in signage_tables:
                if len(df) > 1:
                    # Find the actual header row dynamically
                    header_row_idx = 0
                    for i, row in df.iterrows():
                        row_str = " ".join([str(c) for c in row if c is not None]).upper()
                        if "TK CODE" in row_str or ("FAMILY" in row_str and "COUNT" in row_str):
                            header_row_idx = i
                            break
                    
                    raw_headers = [str(c).strip() if c is not None and str(c).strip() != "" else f"Col_{i}" for i, c in enumerate(df.iloc[header_row_idx])]
                    cols = pd.Series(raw_headers)
                    for dup in cols[cols.duplicated()].unique():
                        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
                    
                    df_clean = df.iloc[header_row_idx + 1:].copy()
                    df_clean.columns = cols
                    df_clean = df_clean.reset_index(drop=True)
                    master_df_list.append(df_clean)
            
            if master_df_list:
                master_df = pd.concat(master_df_list, ignore_index=True, join='outer')
                
                cols_lower = {c.lower(): c for c in master_df.columns}
                code_col = next((cols_lower[c] for c in cols_lower if 'tk code' in c or 'code' in c), None)
                level_col = next((cols_lower[c] for c in cols_lower if 'level' in c), None)
                count_col = next((cols_lower[c] for c in cols_lower if 'count' in c or 'qty' in c or 'quantity' in c), None)
                desc_col = next((cols_lower[c] for c in cols_lower if 'family' in c or 'type' in c or 'description' in c), None)

                simplified_data = []
                for _, row in master_df.iterrows():
                    code = str(row[code_col]).strip() if code_col and pd.notna(row[code_col]) else ""
                    
                    if code and code.upper() not in ["TK CODE", "SIGNAGE CODE", "NAN", "NONE", ""] and not code.lower().startswith("nan"):
                        level = str(row[level_col]).strip() if level_col and pd.notna(row[level_col]) else "L00"
                        count = str(row[count_col]).strip() if count_col and pd.notna(row[count_col]) else "1"
                        desc = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else ""
                        
                        if "nan" in desc.lower():
                            desc = desc.replace("nan", "").strip("-_ ")

                        try:
                            qty_val = int(float(count))
                        except:
                            qty_val = 1

                        simplified_data.append({
                            "Signage Code": code,
                            "Description": desc,
                            "Position / Level": level,
                            "Quantity": qty_val
                        })
                
                if simplified_data:
                    df_simple = pd.DataFrame(simplified_data).drop_duplicates(subset=["Signage Code", "Position / Level"])
                    
                    st.success(f"Successfully extracted {len(df_simple)} clean signage items!")
                    
                    total_qty = df_simple["Quantity"].sum()
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total Signage Quantity", total_qty)
                    with col2:
                        st.metric("Unique Signage Codes", len(df_simple))
                    
                    st.markdown("---")
                    st.subheader("📋 Clean Signage Schedule")
                    st.dataframe(df_simple, use_container_width=True)
                    
                    csv_data = df_simple.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Clean Schedule as CSV",
                        data=csv_data,
                        file_name="clean_signage_schedule.csv",
                        mime="text/csv"
                    )
                else:
                    st.warning("Found schedule table but could not extract rows cleanly:")
                    st.dataframe(master_df, use_container_width=True)
        else:
            st.warning("No main signage schedule tables found in this PDF.")
