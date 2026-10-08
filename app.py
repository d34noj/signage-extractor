import streamlit as st
import pdfplumber
import pandas as pd
import re

# Page Configuration
st.set_page_config(
    page_title="Store Signage Quantity & Level Analyzer",
    page_icon="📊",
    layout="wide"
)

st.title("Store Design Signage Schedule & Level Analyzer")
st.markdown("Upload your store PDF drawing package to automatically extract signage schedules, sort by level, and auto-calculate total quantities.")

# File uploader
uploaded_file = st.file_uploader("Upload Store PDF Document", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Analyzing PDF pages, parsing levels, and computing totals..."):
        signage_tables = []
        raw_text_matches = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # 1. Extract structured tables
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        df_table = pd.DataFrame(table)
                        header_text = " ".join([str(cell) for row in df_table.head(2).values for cell in row if cell])
                        # Filter for signage-related schedule tables
                        if any(keyword in header_text.upper() for keyword in ["TK CODE", "SIGNAGE", "FAMILY", "SCHEDULE", "SIGNAGE REFERENCE"]):
                            signage_tables.append((page_num + 1, df_table))
                
                # 2. Backup text code search
                text = page.extract_text()
                if text:
                    found_codes = re.findall(r'\b(?:TK\d{2}-\d{2}[A-Z]?|HANGNAV\d+[A-Z]?|TOV\d+[A-Z]?|HS\d+[A-Z]*|NAV\d+)\b', text)
                    for code in found_codes:
                        raw_text_matches.append({"Page": page_num + 1, "Signage / Reference Code": code})

        if signage_tables:
            st.success(f"Successfully extracted schedule table(s) from PDF!")
            
            # Combine and clean tables safely
            master_df_list = []
            for page_num, df in signage_tables:
                if len(df) > 1:
                    # Clean and ensure unique column headers to prevent InvalidIndexError
                    raw_headers = [str(c).strip() if c is not None and str(c).strip() != "" else f"Col_{i}" for i, c in enumerate(df.iloc[0])]
                    
                    # Make headers unique
                    cols = pd.Series(raw_headers)
                    for dup in cols[cols.duplicated()].unique():
                        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
                    
                    df_clean = df[1:].copy()
                    df_clean.columns = cols
                    df_clean = df_clean.reset_index(drop=True)
                    df_clean["Source Page"] = page_num
                    master_df_list.append(df_clean)
            
            if master_df_list:
                # Concatenate safely filling missing columns with NaN
                master_df = pd.concat(master_df_list, ignore_index=True, join='outer')
                
                # Find Level and Count columns dynamically
                cols = master_df.columns
                level_col = next((c for c in cols if 'level' in c.lower()), None)
                count_col = next((c for c in cols if 'count' in c.lower() or 'qty' in c.lower() or 'quantity' in c.lower()), None)
                
                if count_col:
                    master_df[count_col] = pd.to_numeric(master_df[count_col], errors='coerce').fillna(0)

                st.markdown("---")
                st.header("📈 Project Summary & Totals")
                
                # Metric Cards
                col1, col2, col3 = st.columns(3)
                with col1:
                    total_items = int(master_df[count_col].sum()) if count_col else len(master_df)
                    st.metric("Total Signage Quantity", total_items)
                with col2:
                    unique_codes = master_df.iloc[:, 0].nunique() if len(master_df.columns) > 0 else len(master_df)
                    st.metric("Unique Signage Entries", unique_codes)
                with col3:
                    levels_count = master_df[level_col].nunique() if level_col else 1
                    st.metric("Levels Covered", levels_count)

                st.markdown("---")
                st.header("🗂️ Filter & View by Level")
                
                if level_col:
                    levels = ["All Levels"] + list(master_df[level_col].dropna().unique())
                    selected_level = st.selectbox("Select Store Level", levels)
                    
                    if selected_level != "All Levels":
                        filtered_df = master_df[master_df[level_col] == selected_level]
                    else:
                        filtered_df = master_df
                else:
                    filtered_df = master_df
                    selected_level = "All"

                st.dataframe(filtered_df, use_container_width=True)

                # Export CSV
                csv_data = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label=f"📥 Download Schedule for [{selected_level}] as CSV",
                    data=csv_data,
                    file_name=f"signage_schedule_{selected_level.lower().replace(' ', '_')}.csv",
                    mime="text/csv"
                )
                
                # Quantity Breakdown by Level
                if level_col and count_col:
                    st.markdown("---")
                    st.header("📊 Quantity Breakdown by Level")
                    level_summary = master_df.groupby(level_col)[count_col].sum().reset_index()
                    level_summary.columns = ["Level", "Total Quantity"]
                    st.dataframe(level_summary, use_container_width=True)

        else:
            st.warning("No structured signage schedule tables detected. Showing reference codes found in drawings:")
            if raw_text_matches:
                df_codes = pd.DataFrame(raw_text_matches).drop_duplicates().reset_index(drop=True)
                st.dataframe(df_codes, use_container_width=True)
                st.download_button(
                    label="📥 Download Reference Codes CSV",
                    data=df_codes.to_csv(index=False).encode('utf-8'),
                    file_name="signage_reference_codes.csv",
                    mime="text/csv"
                )
