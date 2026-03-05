import streamlit as st
import pandas as pd
from io import BytesIO

# ---------- UTILITIES ----------

def rename_columns(df):
    return df.rename(columns={
        'Column1': 'SiteCode',
        'Column2': 'LocationCode',
        'Column3': 'Product',
        'Column4': 'Description',
        'Column5': 'Date',
        'Column6': 'Qty'
    })

def revert_column_names(df):
    return df.rename(columns={
        'SiteCode': 'Column1',
        'LocationCode': 'Column2',
        'Product': 'Column3',
        'Description': 'Column4',
        'Date': 'Column5',
        'Qty': 'Column6'
    })

# ---------- STREAMLIT APP ----------

st.set_page_config(page_title="Sales Forecast Adjustment", layout="wide")

st.markdown("""
<style>
/* ---- TURQUOISE THEME ---- */
:root {
  --turq: #1abc9c;
  --turq-dark: #148f77;
  --bg: #e8fbf7;
  --bg2: #f6fffd;
  --text: #0f2f2a;
}

/* Main app background */
.stApp {
  background: linear-gradient(180deg, var(--bg) 0%, var(--bg2) 100%) !important;
  color: var(--text) !important;
}

/* Sidebar */
section[data-testid="stSidebar"] > div {
  background: linear-gradient(180deg, #d7faf3 0%, #eafffa 100%) !important;
}

/* Titles */
h1, h2, h3, h4 {
  color: var(--turq-dark) !important;
}

/* Buttons (download + normal) */
div.stButton > button,
div.stDownloadButton > button {
  background: var(--turq) !important;
  color: #fff !important;
  border: 1px solid var(--turq-dark) !important;
  border-radius: 10px !important;
}
div.stButton > button:hover,
div.stDownloadButton > button:hover {
  background: var(--turq-dark) !important;
}

/* File uploader box */
div[data-testid="stFileUploader"] section {
  border: 2px dashed rgba(26,188,156,0.55) !important;
  border-radius: 12px !important;
  background: rgba(26,188,156,0.06) !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Sales Forecast Adjustment")


step = st.radio("Step", ["Upload Sage X3 File", "Generate Pivot", "Upload Modified Pivot", "Download Final Sage X3 Format"])

if step == "Upload Sage X3 File":
    uploaded_file = st.file_uploader("Upload Original CSV from Sage X3", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file, encoding='latin1', header=0, usecols=lambda col: col != 'Column7')
        df = rename_columns(df)
        st.session_state['df_original'] = df
        st.write("Preview:", df.head())

elif step == "Generate Pivot":
    if 'df_original' in st.session_state:
        df = st.session_state['df_original'].copy()
        df['Date'] = df['Date'].astype(str).str.zfill(8)
        df['Date2'] = pd.to_datetime(df['Date'], format='%Y%m%d', errors='coerce')
        df['Month'] = df['Date2'].dt.to_period('M').dt.to_timestamp()
        df['Site'] = df['SiteCode'] + '-' + df['LocationCode']
        df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)

        desc_lookup = df[['Site', 'Product', 'Description']].drop_duplicates(subset=['Site', 'Product'])

        pivot = pd.pivot_table(
            df,
            index=['Site', 'Product'],
            columns='Month',
            values='Qty',
            aggfunc='sum'
        ).reset_index()

        pivot = pd.merge(pivot, desc_lookup, how='left', on=['Site', 'Product'])
        month_cols = [col for col in pivot.columns if isinstance(col, pd.Timestamp)]
        pivot = pivot[['Site', 'Product', 'Description'] + month_cols]
        pivot.columns = [col.strftime('%Y-%m') if isinstance(col, pd.Timestamp) else col for col in pivot.columns]

        st.session_state['pivot'] = pivot
        st.dataframe(pivot.fillna('').head(), use_container_width=True)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            pivot.to_excel(writer, index=False)
        st.download_button("Download Pivot for Sales Team", data=buffer.getvalue(), file_name="pivot_table.xlsx")
    else:
        st.warning("Please upload the original Sage X3 file first.")

elif step == "Upload Modified Pivot":
    uploaded_pivot = st.file_uploader("Upload Modified Pivot Table", type=["xlsx"])
    if uploaded_pivot:
        df = pd.read_excel(uploaded_pivot)
        st.session_state['pivot_modified'] = df
        st.success("Modified pivot uploaded.")
        st.dataframe(df.head(), use_container_width=True)

elif step == "Download Final Sage X3 Format":
    if 'pivot_modified' in st.session_state and 'df_original' in st.session_state:
        original_df = st.session_state['df_original'].copy()
        pivot_df = st.session_state['pivot_modified'].copy()

        # Preprocess original
        original_df['Date'] = original_df['Date'].astype(str).str.zfill(8)
        original_df['ParsedDate'] = pd.to_datetime(original_df['Date'], format='%Y%m%d', errors='coerce')
        original_df['MonthStart'] = original_df['ParsedDate'].dt.to_period('M').dt.start_time
        original_df['Site'] = original_df['SiteCode'] + '-' + original_df['LocationCode']
        original_df['Qty'] = pd.to_numeric(original_df['Qty'], errors='coerce')
        original_df['key'] = original_df[['Site', 'Product', 'MonthStart']].astype(str).agg('-'.join, axis=1)

        # Build pivot_long
        pivot_long = pivot_df.melt(id_vars=['Site', 'Product', 'Description'], var_name='Month', value_name='Qty')
        pivot_long['Month'] = pd.to_datetime(pivot_long['Month'], errors='coerce')
        pivot_long = pivot_long.dropna(subset=['Month'])
        pivot_long['MonthStart'] = pivot_long['Month'].dt.to_period('M').dt.start_time
        pivot_long['key'] = pivot_long[['Site', 'Product', 'MonthStart']].astype(str).agg('-'.join, axis=1)
        pivot_long['SiteCode'] = pivot_long['Site'].str.split('-').str[0]
        pivot_long['LocationCode'] = pivot_long['Site'].str.split('-').str[1]

        # Reallocate or create new rows
        reallocated_rows = []
        new_rows = []

        for _, row in pivot_long.iterrows():
            key = row['key']
            updated_qty = row['Qty']

            if key in original_df['key'].values:
                subset = original_df[original_df['key'] == key].copy()
                total_original = subset['Qty'].sum()
                if total_original > 0:
                    weights = subset['Qty'] / total_original
                    subset['Qty'] = weights * updated_qty
                    reallocated_rows.append(subset)
                else:
                    new_rows.append(pd.DataFrame([{
                        'SiteCode': row['SiteCode'],
                        'LocationCode': row['LocationCode'],
                        'Product': row['Product'],
                        'Description': row['Description'],
                        'Date': row['MonthStart'].strftime('%Y%m%d'),
                        'Qty': updated_qty
                    }]))
            else:
                new_rows.append(pd.DataFrame([{
                    'SiteCode': row['SiteCode'],
                    'LocationCode': row['LocationCode'],
                    'Product': row['Product'],
                    'Description': row['Description'],
                    'Date': row['MonthStart'].strftime('%Y%m%d'),
                    'Qty': updated_qty
                }]))

        # Assemble final
        reallocated_df = pd.concat(reallocated_rows, ignore_index=True) if reallocated_rows else pd.DataFrame()
        new_df = pd.concat(new_rows, ignore_index=True) if new_rows else pd.DataFrame()
        touched_keys = set(pivot_long['key'])
        untouched_df = original_df[~original_df['key'].isin(touched_keys)].copy()

        final_df = pd.concat([untouched_df, reallocated_df, new_df], ignore_index=True)
        final_df = final_df[['SiteCode', 'LocationCode', 'Product', 'Description', 'Date', 'Qty']]
        final_df = final_df[final_df[['SiteCode', 'LocationCode', 'Product', 'Date', 'Qty']].notna().all(axis=1)]
        final_df = final_df[final_df['Product'].astype(str).str.strip() != '']
        final_df['Site'] = final_df['SiteCode'].astype(str) + '-' + final_df['LocationCode'].astype(str)

        final_df['Date_sort'] = pd.to_datetime(
            final_df['Date'].astype(str).str.zfill(8),
            format='%Y%m%d',
            errors='coerce'
        )

        final_df = (
            final_df
            .sort_values(['Site', 'Product', 'Date_sort'], na_position='last')
            .drop(columns=['Site', 'Date_sort'])
        )

        df_export = revert_column_names(final_df)
        df_export = df_export[['Column1', 'Column2', 'Column3', 'Column4', 'Column5', 'Column6']]
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df_export.to_excel(writer, index=False, header=False, sheet_name="Sage_Changes")
        st.download_button(
            "Download Final Excel for Sage X3",
            data=buffer.getvalue(),
            file_name="final_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")    
