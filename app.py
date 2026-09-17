import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import os
import plotly.express as px
from datetime import datetime

# ==========================================
# 0. CONFIG & ENTERPRISE THEME SETUP
# ==========================================
st.set_page_config(
    page_title="PT Geoservices - Supply Chain ERP",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Enterprise Look
st.markdown("""
    <style>
        .main-header {font-size: 2rem; font-weight: 700; color: #1e3d59; margin-bottom: 0px;}
        .sub-header {font-size: 1rem; color: #5f6c7b; margin-bottom: 20px;}
        .erp-card {padding: 20px; border-radius: 8px; background-color: #f8f9fa; border: 1px solid #e9ecef; margin-bottom: 15px;}
        .badge-new {background-color: #2b9348; color: white; padding: 2px 8px; border-radius: 12px; font-size: 0.75rem; font-weight: bold;}
    </style>
""", unsafe_allow_html=True)

OUTPUT_FOLDER = "saved_reports"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Top Bar / Enterprise Header
col_title, col_date = st.columns([3, 1])
with col_title:
    st.markdown('<div class="main-header">🏭 PT GEOSERVICES — WMS & PROCUREMENT ERP</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Module: Purchase Request to Inbound Tracking Engine</div>', unsafe_allow_html=True)
with col_date:
    today_str = datetime.now().strftime("%A, %d %B %Y | %H:%M WIB")
    st.markdown(f"<div style='text-align: right; background: #e9ecef; padding: 10px; border-radius: 6px; font-size: 0.85rem; font-weight: 600;'>📅 Hari Ini:<br>{today_str}</div>", unsafe_allow_html=True)

st.divider()

# ==========================================
# 1. SIDEBAR NAVIGATION & UPLOADERS
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/warehouse.png", width=60)
    st.title("Navigation")
    menu = st.radio("Pilih Menu:", ["🚀 Processing & Dashboard", "📁 Arsip Dokumen (Repository)"])
    
    st.divider()
    st.header("📂 Data Source Input")
    st.info("Unggah dokumen sumber berformat .xlsx")
    
    # Nama disesuaikan tapi variabel backend aman
    file_pr_2026 = st.file_uploader("PR Data", type=['xlsx'], key="pr_base")
    file_pr_lama = st.file_uploader("PRN Data", type=['xlsx'], key="pr_closed")
    file_po_lokal = st.file_uploader("PO Data - Lokal", type=['xlsx'], key="po_lok")
    file_po_impor = st.file_uploader("PO Data - Impor", type=['xlsx'], key="po_imp")
    file_inbound = st.file_uploader("Inbound Data", type=['xlsx'], key="inb")
    
    process_btn = st.button("⚙️ Execute Processing", use_container_width=True, type="primary")

# ==========================================
# 2. CORE PROCESSING ENGINE (SAFE & ROBUST)
# ==========================================
@st.cache_data
def process_tracking_data(pr_new, pr_old, po_lok, po_imp, inb):
    # Proses PR Base
    pr_df = pd.read_excel(pr_new, header=1)
    pr_data = pr_df[['Date', 'PR Number \n(Manual)', 'Item \nCode', 'Item Description', 'Qty']].copy()
    pr_data.columns = ['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty']
    pr_data['Item_Code'] = pr_data['Item_Code'].astype(str).str.strip()
    pr_data['PR_Manual_No_Clean'] = pr_data['PR_Manual_No'].astype(str).str.replace(" ", "")
    pr_data['PR_Date'] = pd.to_datetime(pr_data['PR_Date'], errors='coerce')
    pr_data = pr_data[pr_data['PR_Manual_No'].notna() & (pr_data['PR_Manual_No'] != '') & (pr_data['PR_Manual_No'].astype(str) != 'nan')]

    # Ambil Status 'Closed' dari PRN Data
    pr_2426_df = pd.read_excel(pr_old, sheet_name=0, header=None)
    closed_info = pr_2426_df.iloc[7:, [2, 6, 7]].copy()
    closed_info.columns = ['PR_Manual_No', 'RequestClosed', 'Item_Code']
    closed_info['Item_Code'] = closed_info['Item_Code'].astype(str).str.strip()
    closed_info['PR_Manual_No_Clean'] = closed_info['PR_Manual_No'].astype(str).str.replace(" ", "")
    closed_info = closed_info.drop_duplicates(subset=['PR_Manual_No_Clean', 'Item_Code'], keep='last')
    
    pr_data = pd.merge(pr_data, closed_info[['PR_Manual_No_Clean', 'Item_Code', 'RequestClosed']], on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    pr_data['RequestClosed'] = pr_data['RequestClosed'].fillna('No')

    # Proses PO (Ffill & Expand)
    def clean_po(file, tipe):
        po = pd.read_excel(file, header=13)
        po = po[['Purchase Order Number', 'PO Date', 'PR Manual No.', 'Item Code', 'Qty', 'Unnamed: 5']].copy()
        po.columns = ['PO_No', 'PO_Date', 'PR_Manual_No_Orig', 'Item_Code', 'PO_Qty', 'Vendor']
        po['PO_No'] = po['PO_No'].ffill()
        po['PO_Date'] = po['PO_Date'].ffill()
        po['Vendor'] = po['Vendor'].ffill()
        po = po.dropna(subset=['Item_Code'])
        po = po[po['Item_Code'].astype(str).str.strip() != 'nan']
        po['Tipe_PO'] = tipe
        return po

    po_df = pd.concat([clean_po(po_lok, "Lokal"), clean_po(po_imp, "Impor")], ignore_index=True)
    po_df['Item_Code'] = po_df['Item_Code'].astype(str).str.strip()

    expanded_rows = []
    for _, row in po_df.iterrows():
        pr_str = str(row['PR_Manual_No_Orig']).strip()
        if pd.isna(pr_str) or pr_str == 'nan': continue
        if ',' in pr_str or '/' in pr_str:
            parts = re.split(r'[,/]', pr_str)
            base_prefix = ""
            for p in [x.strip() for x in parts if x.strip()]:
                if not p.isdigit():
                    match = re.match(r'([A-Za-z.\-]+)(\d+)', p)
                    if match: base_prefix = match.group(1)
                new_row = row.to_dict()
                new_row['PR_Manual_No_Clean'] = (base_prefix + p).replace(" ", "") if p.isdigit() and base_prefix else p.replace(" ", "")
                expanded_rows.append(new_row)
        else:
            new_row = row.to_dict()
            new_row['PR_Manual_No_Clean'] = pr_str.replace(" ", "")
            expanded_rows.append(new_row)

    po_exp = pd.DataFrame(expanded_rows)
    po_exp['PO_Date'] = pd.to_datetime(po_exp['PO_Date'], errors='coerce')
    po_exp['PO_Qty'] = pd.to_numeric(po_exp['PO_Qty'], errors='coerce').fillna(0)
    
    po_agg = po_exp.groupby(['PR_Manual_No_Clean', 'Item_Code']).agg(
        PO_No=('PO_No', lambda x: ', '.join(x.dropna().unique().astype(str))),
        PO_Date=('PO_Date', 'max'),
        PO_Qty=('PO_Qty', 'sum'),
        Vendor=('Vendor', lambda x: ', '.join(x.dropna().unique().astype(str))),
        Tipe_PO=('Tipe_PO', lambda x: ', '.join(x.dropna().unique().astype(str)))
    ).reset_index()

    # Gabung PR & PO
    merged = pd.merge(pr_data, po_agg, on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    merged['PO_No'] = merged['PO_No'].fillna('')
    merged = merged[~((merged['PO_No'] == '') & (merged['RequestClosed'] == 'Yes'))]

    # Proses Inbound Data
    inb_xls = pd.ExcelFile(inb)
    inb_df = pd.concat([pd.read_excel(inb_xls, sheet_name=s) for s in inb_xls.sheet_names])
    inb_df['ItemCode'] = inb_df['ItemCode'].astype(str).str.strip()
    inb_df['POnumber'] = inb_df['POnumber'].astype(str).str.strip()
    inb_agg = inb_df.groupby(['POnumber', 'ItemCode']).agg(Rcv_Date=('ReceivedDate', 'max'), Rcv_Qty=('RcvQty', 'sum')).reset_index()

    def get_inb(po_str, item_code):
        if pd.isna(po_str) or po_str == '': return pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})
        pos = [p.strip() for p in po_str.split(',')]
        subset = inb_agg[(inb_agg['POnumber'].isin(pos)) & (inb_agg['ItemCode'] == item_code)]
        if subset.empty: return pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})
        return pd.Series({'Rcv_Date': subset['Rcv_Date'].max(), 'Rcv_Qty': subset['Rcv_Qty'].sum()})

    merged[['Rcv_Date', 'Rcv_Qty']] = merged.apply(lambda row: get_inb(row['PO_No'], row['Item_Code']), axis=1)

    def get_status(row):
        if row['PO_No'] == '': return 'Routing Approval'
        elif row['Rcv_Qty'] == 0: return 'Menunggu Pengiriman'
        elif row['Rcv_Qty'] < row['PO_Qty']: return 'Diterima Sebagian'
        else: return 'Sudah Diterima'

    merged['Status'] = merged.apply(get_status, axis=1)
    
    merged['PR_Date'] = merged['PR_Date'].dt.strftime('%m/%d/%Y').fillna('-')
    merged['PO_Date'] = merged['PO_Date'].dt.strftime('%m/%d/%Y').fillna('-')
    merged['Rcv_Date'] = pd.to_datetime(merged['Rcv_Date'], errors='coerce').dt.strftime('%m/%d/%Y').fillna('-')
    merged['Vendor'] = merged['Vendor'].fillna('-')
    merged['Tipe_PO'] = merged['Tipe_PO'].fillna('-')
    
    final_cols = ['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty', 'PO_Date', 'PO_No', 'Vendor', 'Tipe_PO', 'PO_Qty', 'Rcv_Date', 'Rcv_Qty', 'Status']
    return merged[final_cols]

# Helper function untuk format arsip dengan Tanggal/Jam & Badge Terbaru
def get_formatted_archive_list():
    files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.xlsx')]
    if not files:
        return []
    # Sort by modification time descending (latest first)
    files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
    
    formatted_list = []
    for idx, f in enumerate(files):
        mtime = os.path.getmtime(os.path.join(OUTPUT_FOLDER, f))
        dt_str = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M:%S")
        label = f"[{dt_str}] {f}"
        if idx == 0:
            label += " ⭐ [TERBARU]"
        formatted_list.append((label, f))
    return formatted_list

# ==========================================
# 3. ROUTING MENU
# ==========================================
if menu == "🚀 Processing & Dashboard":
    if process_btn:
        if all([file_pr_2026, file_pr_lama, file_po_lokal, file_po_impor, file_inbound]):
            with st.spinner('Menjalankan pipeline ERP... Normalisasi data sedang berjalan...'):
                df_final = process_tracking_data(file_pr_2026, file_pr_lama, file_po_lokal, file_po_impor, file_inbound)
                
                # Simpan otomatis ke server (arsip) dengan timestamp human-friendly
                now_dt = datetime.now()
                timestamp_file = now_dt.strftime("%Y%m%d_%H%M%S")
                saved_filename = f"Tracking_Final_{timestamp_file}.xlsx"
                saved_filepath = os.path.join(OUTPUT_FOLDER, saved_filename)
                df_final.to_excel(saved_filepath, index=False)
                
                st.success(f"✅ Eksekusi selesai! File terarsip sebagai `{saved_filename}`")
                
                # Simpan ke session state buat tampilkan dashboard
                st.session_state['df_final'] = df_final
                st.session_state['last_saved'] = saved_filename
        else:
            st.error("⚠️ Mohon lengkapi kelima file sumber di panel sebelah kiri.")
            
    # Tampilkan Dashboard jika data tersedia di session atau baru diproses
    if 'df_final' in st.session_state:
        df_final = st.session_state['df_final']
        
        st.subheader("📊 Executive Summary Metrics")
        col1, col2, col3, col4 = st.columns(4)
        status_counts = df_final['Status'].value_counts()
        col1.metric("Routing Approval", status_counts.get("Routing Approval", 0))
        col2.metric("Menunggu Pengiriman", status_counts.get("Menunggu Pengiriman", 0))
        col3.metric("Diterima Sebagian", status_counts.get("Diterima Sebagian", 0))
        col4.metric("Sudah Diterima", status_counts.get("Sudah Diterima", 0))

        st.divider()

        c1, c2 = st.columns(2)
        with c1:
            fig_pie = px.pie(df_final, names='Status', title='Proporsi Status Supply Chain', color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig_pie, use_container_width=True)
        with c2:
            top_vendors = df_final[df_final['Vendor'] != '-']['Vendor'].value_counts().head(10).reset_index()
            top_vendors.columns = ['Vendor', 'Jumlah']
            fig_bar = px.bar(top_vendors, x='Jumlah', y='Vendor', orientation='h', title='Top 10 Active Vendors', color='Jumlah', color_continuous_scale='viridis')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("📋 Master Tracking Dataset")
        st.dataframe(df_final, use_container_width=True, height=450)
        
        # Download Quick Action
        output = io.BytesIO()
        df_final.to_excel(output, index=False, sheet_name='Tracking Data')
        st.download_button(
            label="📥 Download Hasil Proses Sesi Ini",
            data=output.getvalue(),
            file_name=st.session_state.get('last_saved', 'Tracking_Final.xlsx'),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("👈 Silakan unggah dokumen di *sidebar* dan klik **Execute Processing** untuk melihat *dashboard*.")

elif menu == "📁 Arsip Dokumen (Repository)":
        st.subheader("🗄️ Enterprise Repository & Audit Trail")
        st.markdown("Daftar seluruh laporan hasil *processing* masa lalu yang tersimpan di server.")
        
        archive_items = get_formatted_archive_list()
        
        if archive_items:
            labels = [item[0] for item in archive_items]
            filenames = [item[1] for item in archive_items]
            
            selected_label = st.selectbox("Pilih Dokumen Arsip:", labels)
            selected_filename = dict(zip(labels, filenames))[selected_label]
            
            file_path_dl = os.path.join(OUTPUT_FOLDER, selected_filename)
            if os.path.exists(file_path_dl):
                # Preview mini
                df_preview = pd.read_excel(file_path_dl)
                st.write(f"**Preview file:** `{selected_filename}` ({len(df_preview)} baris)")
                st.dataframe(df_preview.head(50), use_container_width=True)
                
                with open(file_path_dl, "rb") as f:
                    st.download_button(
                        label=f"📥 Download Arsip ({selected_filename})",
                        data=f.read(),
                        file_name=selected_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
        else:
            st.warning("Belum ada arsip tersimpan. Jalankan *Processing* terlebih dahulu di menu sebelah.")
