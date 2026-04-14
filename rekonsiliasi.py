import streamlit as st
import pandas as pd
import io

# --- CONFIG & STYLING ---
st.set_page_config(page_title="Audit RFC Mitra vs SAP", layout="wide")
st.title("🛡️ Tool Rekonsiliasi Material & RFC")
st.markdown("Versi 4.0: Dynamic Sheet & Column Selection")

# --- HELPER FUNCTIONS ---
def clean_key(series):
    return (series.astype(str)
            .str.strip()
            .str.upper()
            .str.replace(r'^0+', '', regex=True)
            .str.replace(r'\s+', ' ', regex=True))

# --- SIDEBAR: UPLOAD & DYNAMIC SHEET ---
st.sidebar.header("1. Upload & Pilih Sheet")

# --- FILE 1 ---
file1 = st.sidebar.file_uploader("Data 1 (Klaim Mitra)", type=['xlsx'])
sheet_opt1 = []
if file1:
    # Ambil daftar nama sheet secara otomatis
    sheet_opt1 = pd.ExcelFile(file1).sheet_names
selected_sheet1 = st.sidebar.selectbox("Pilih Sheet Data 1", sheet_opt1) if sheet_opt1 else None

st.sidebar.markdown("---")

# --- FILE 2 ---
file2 = st.sidebar.file_uploader("Data 2 (SAP Gudang)", type=['xlsx'])
sheet_opt2 = []
if file2:
    sheet_opt2 = pd.ExcelFile(file2).sheet_names
selected_sheet2 = st.sidebar.selectbox("Pilih Sheet Data 2", sheet_opt2) if sheet_opt2 else None

# --- MAPPING KOLOM (Hanya muncul jika sheet sudah dipilih) ---
if file1 and file2 and selected_sheet1 and selected_sheet2:
    st.sidebar.header("2. Mapping Kolom")
    
    # Load sample data berdasarkan sheet yang dipilih
    df1_sample = pd.read_excel(file1, sheet_name=selected_sheet1, nrows=0)
    df2_sample = pd.read_excel(file2, sheet_name=selected_sheet2, nrows=0)

    st.sidebar.subheader("Kolom Data 1 (Mitra)")
    col_rfc1 = st.sidebar.selectbox("Kolom RFC (D1)", df1_sample.columns, index=df1_sample.columns.get_loc("NO RFC") if "NO RFC" in df1_sample.columns else 0)
    col_mat1 = st.sidebar.selectbox("Kolom Material (D1)", df1_sample.columns, index=df1_sample.columns.get_loc("MATERIAL") if "MATERIAL" in df1_sample.columns else 0)
    col_qty1 = st.sidebar.selectbox("Kolom Quantity (D1)", df1_sample.columns, index=df1_sample.columns.get_loc("QUANTITY") if "QUANTITY" in df1_sample.columns else 0)

    st.sidebar.subheader("Kolom Data 2 (SAP)")
    col_rfc2 = st.sidebar.selectbox("Kolom RFC (D2)", df2_sample.columns, index=df2_sample.columns.get_loc("RFC") if "RFC" in df2_sample.columns else 0)
    col_mat2 = st.sidebar.selectbox("Kolom Material (D2)", df2_sample.columns, index=df2_sample.columns.get_loc("Material") if "Material" in df2_sample.columns else 0)
    col_qty2 = st.sidebar.selectbox("Kolom Quantity (D2)", df2_sample.columns, index=df2_sample.columns.get_loc("Total Quantity") if "Total Quantity" in df2_sample.columns else 0)

    if st.sidebar.button("🚀 Mulai Rekonsiliasi"):
        # Load Full Data berdasarkan sheet terpilih
        df1 = pd.read_excel(file1, sheet_name=selected_sheet1)
        df2 = pd.read_excel(file2, sheet_name=selected_sheet2)

        # Normalisasi
        df1['RFC_KEY'] = clean_key(df1[col_rfc1])
        df1['MAT_KEY'] = clean_key(df1[col_mat1])
        df2['RFC_KEY'] = clean_key(df2[col_rfc2])
        df2['MAT_KEY'] = clean_key(df2[col_mat2])

        # Aggregasi
        g1 = df1.groupby(['RFC_KEY', 'MAT_KEY'])[col_qty1].sum().reset_index()
        g2 = df2.groupby(['RFC_KEY', 'MAT_KEY'])[col_qty2].sum().reset_index()

        all_rfc_mitra = g1['RFC_KEY'].unique()
        results = []

        # LOGIKA AUDIT
        for rfc in all_rfc_mitra:
            items_mitra = g1[g1['RFC_KEY'] == rfc]
            items_sap = g2[g2['RFC_KEY'] == rfc]

            if items_sap.empty:
                for _, row in items_mitra.iterrows():
                    results.append({'RFC_KEY': rfc, 'MAT_KEY': row['MAT_KEY'], 'Qty_Mitra': row[col_qty1], 'Qty_SAP': 0, 'Selisih': row[col_qty1], 'Status': 'RFC Tidak Ada di SAP'})
                continue

            merged_items = pd.merge(items_mitra, items_sap, on='MAT_KEY', how='outer', suffixes=('_Mtr', '_SAP'))
            for _, row in merged_items.iterrows():
                q1 = row[col_qty1] if not pd.isna(row[col_qty1]) else 0
                q2 = row[col_qty2] if not pd.isna(row[col_qty2]) else 0
                status = "Lurus"
                if pd.isna(row['RFC_KEY_SAP']): status = "Perbedaan Material (Mitra Ada, SAP Tidak)"
                elif pd.isna(row['RFC_KEY_Mtr']): status = "Perbedaan Material (SAP Ada, Mitra Tidak)"
                elif q1 != q2: status = "Selisih Quantity"
                results.append({'RFC_KEY': rfc, 'MAT_KEY': row['MAT_KEY'], 'Qty_Mitra': q1, 'Qty_SAP': q2, 'Selisih': q1 - q2, 'Status': status})

        df_results = pd.DataFrame(results)

        # VALIDASI MANUAL (Filtered Scope)
        validasi_mat1 = g1.groupby('MAT_KEY')[col_qty1].sum().reset_index().rename(columns={col_qty1: 'Total_Qty_Mitra'})
        g2_filtered = g2[g2['RFC_KEY'].isin(all_rfc_mitra)]
        validasi_mat2 = g2_filtered.groupby('MAT_KEY')[col_qty2].sum().reset_index().rename(columns={col_qty2: 'Total_Qty_SAP'})

        # SUMMARY METRICS
        total_rfc_count = len(all_rfc_mitra)
        rfc_bermasalah_list = df_results[df_results['Status'] != 'Lurus']['RFC_KEY'].unique()
        rfc_masalah_count = len(rfc_bermasalah_list)
        rfc_lurus_count = total_rfc_count - rfc_masalah_count

        # KATEGORISASI
        data_lurus = df_results[~df_results['RFC_KEY'].isin(rfc_bermasalah_list)]
        data_missing_rfc = df_results[df_results['Status'] == 'RFC Tidak Ada di SAP']
        data_diff_qty = df_results[df_results['Status'] == 'Selisih Quantity']
        data_diff_mat = df_results[df_results['Status'].str.contains("Perbedaan Material")]

        # --- TAMPILAN DASHBOARD ---
        st.header("📈 Summary Rekonsiliasi")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total RFC Mitra", total_rfc_count)
        c2.metric("RFC Match (100% Lurus)", rfc_lurus_count)
        c3.metric("RFC Bermasalah", rfc_masalah_count)

        st.subheader("🔍 Validasi Manual per Material (Scope: RFC Mitra)")
        v_col1, v_col2 = st.columns(2)
        with v_col1:
            st.write("Total Qty per Material di Data 1 (Mitra)")
            st.dataframe(validasi_mat1, use_container_width=True)
        with v_col2:
            st.write("Total Qty per Material di Data 2 (SAP - Filtered)")
            st.dataframe(validasi_mat2, use_container_width=True)

        st.subheader("📋 Detail Hasil Audit")
        t1, t2, t3, t4, t5 = st.tabs(["All Data", "Data Lurus", "Missing RFC", "Selisih Qty", "Beda Material"])
        with t1: st.dataframe(df_results)
        with t2: st.dataframe(data_lurus)
        with t3: st.dataframe(data_missing_rfc)
        with t4: st.dataframe(data_diff_qty)
        with t5: st.dataframe(data_diff_mat)

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_results.to_excel(writer, sheet_name='All_Data', index=False)
            validasi_mat1.to_excel(writer, sheet_name='Summary_Mat_Mitra', index=False)
            validasi_mat2.to_excel(writer, sheet_name='Summary_Mat_SAP_Filtered', index=False)
            data_lurus.to_excel(writer, sheet_name='Data_Lurus', index=False)
            data_missing_rfc.to_excel(writer, sheet_name='Missing_RFC', index=False)
            data_diff_qty.to_excel(writer, sheet_name='Selisih_Quantity', index=False)
            data_diff_mat.to_excel(writer, sheet_name='Beda_Material', index=False)
        
        st.download_button(label="📥 Download Hasil Audit v4", data=output.getvalue(), file_name=f"Audit_RFC_Final_{pd.Timestamp.now().strftime('%d%m%Y')}.xlsx")
else:
    st.info("💡 Silakan upload kedua file Excel untuk memilih sheet dan kolom yang akan di-audit.")