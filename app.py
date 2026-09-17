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
    page_title="Purchase Data Tracking - PT. Geoservices",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        .main-title {font-size: 1.5rem; font-weight: 700; color: #1e3d59; margin: 0;}
        .sub-meta {font-size: 0.8rem; color: #5f6c7b;}
        
        .erp-panel {
            background: #ffffff;
            border: 1px solid #c0c0c0;
            border-radius: 3px;
            padding: 15px;
            margin-bottom: 15px;
        }
    </style>
""", unsafe_allow_html=True)

OUTPUT_FOLDER = "saved_reports"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ==========================================
# AUTO-LOAD LATEST REPORT ON STARTUP (NON-EMPTY)
# ==========================================
def auto_load_latest_dashboard():
    if 'df_final' not in st.session_state:
        files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.xlsx')]
        if files:
            files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
            latest_file = os.path.join(OUTPUT_FOLDER, files[0])
            try:
                st.session_state['df_final'] = pd.read_excel(latest_file)
                st.session_state['last_saved'] = files[0]
            except Exception:
                pass

auto_load_latest_dashboard()

# ==========================================
# CLEAN CORPORATE HEADER WITH LOGO
# ==========================================
c_left, c_right = st.columns([3, 1])
with c_left:
    c_img, c_txt = st.columns([1, 6])
    with c_img:
        logo_path = "Logo_PT_Geoservices_4K_Transparent.jpg"
        if os.path.exists(logo_path):
            st.image(logo_path, width=48)
        else:
            st.markdown("🌐", unsafe_allow_html=True)
    with c_txt:
        st.markdown('<div class="main-title">Purchase Data Tracking</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-meta">PT. Geoservices — Warehouse & Procurement Analytics</div>', unsafe_allow_html=True)

with c_right:
    today_str = datetime.now().strftime("%A, %d %b %Y | %H:%M WIB")
    st.markdown(f"""
        <div style='text-align: right; background: #f8f9fa; padding: 8px 12px; border: 1px solid #e9ecef; border-radius: 4px; font-size: 0.78rem; color: #333;'>
            <b>Business Unit:</b> PT. Geoservices (7001)<br>
            <span style='color: #666;'>📅 {today_str}</span>
        </div>
    """, unsafe_allow_html=True)

st.write("")

# ==========================================
# CORE PROCESSING ENGINE
# ==========================================
@st.cache_data
def process_tracking_data(pr_new, pr_old, po_lok, po_imp, inb):
    pr_df = pd.read_excel(pr_new, header=1)
    pr_data = pr_df[['Date', 'PR Number \n(Manual)', 'Item \nCode', 'Item Description', 'Qty']].copy()
    pr_data.columns = ['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty']
    pr_data['Item_Code'] = pr_data['Item_Code'].astype(str).str.strip()
    pr_data['PR_Manual_No_Clean'] = pr_data['PR_Manual_No'].astype(str).str.replace(" ", "")
    pr_data['PR_Date'] = pd.to_datetime(pr_data['PR_Date'], errors='coerce')
    pr_data = pr_data[pr_data['PR_Manual_No'].notna() & (pr_data['PR_Manual_No'] != '') & (pr_data['PR_Manual_No'].astype(str) != 'nan')]

    pr_2426_df = pd.read_excel(pr_old, sheet_name=0, header=None)
    closed_info = pr_2426_df.iloc[7:, [2, 6, 7]].copy()
    closed_info.columns = ['PR_Manual_No', 'RequestClosed', 'Item_Code']
    closed_info['Item_Code'] = closed_info['Item_Code'].astype(str).str.strip()
    closed_info['PR_Manual_No_Clean'] = closed_info['PR_Manual_No'].astype(str).str.replace(" ", "")
    closed_info = closed_info.drop_duplicates(subset=['PR_Manual_No_Clean', 'Item_Code'], keep='last')
    
    pr_data = pd.merge(pr_data, closed_info[['PR_Manual_No_Clean', 'Item_Code', 'RequestClosed']], on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    pr_data['RequestClosed'] = pr_data['RequestClosed'].fillna('No')

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

    merged = pd.merge(pr_data, po_agg, on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    merged['PO_No'] = merged['PO_No'].fillna('')
    merged = merged[~((merged['PO_No'] == '') & (merged['RequestClosed'] == 'Yes'))]

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

def get_formatted_archive_list():
    files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.xlsx')]
    if not files: return []
    files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
    formatted_list = []
    for idx, f in enumerate(files):
        mtime = os.path.getmtime(os.path.join(OUTPUT_FOLDER, f))
        dt_str = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M:%S")
        label = f"[{dt_str}] {f}"
        if idx == 0: label += " ⭐ [TERBARU]"
        formatted_list.append((label, f))
    return formatted_list

# ==========================================
# 3 TABS MENU NAVIGATION
# ==========================================
selected_tab = st.radio(
    "Navigation Menu",
    ["📊 Dashboard", "📤 Upload Data", "📥 Download / Arsip"],
    horizontal=True,
    label_visibility="collapsed"
)

st.markdown("""<hr style="margin: 5px 0 15px 0; border: none; border-top: 1px solid #1e3d59;">""", unsafe_allow_html=True)

# ==========================================
# TAB 1: DASHBOARD (AUTO-LOADS LATEST)
# ==========================================
if selected_tab == "📊 Dashboard":
    st.subheader("Purchase Data Tracking | Executive Dashboard")
    
    if 'df_final' in st.session_state:
        df_final = st.session_state['df_final']
        if 'last_saved' in st.session_state:
            st.caption(f"📁 Active Dataset: `{st.session_state['last_saved']}`")
        
        c1, c2, c3, c4 = st.columns(4)
        status_counts = df_final['Status'].value_counts()
        with c1: st.metric("Routing Approval", status_counts.get("Routing Approval", 0))
        with c2: st.metric("Menunggu Pengiriman", status_counts.get("Menunggu Pengiriman", 0))
        with c3: st.metric("Diterima Sebagian", status_counts.get("Diterima Sebagian", 0))
        with c4: st.metric("Sudah Diterima", status_counts.get("Sudah Diterima", 0))

        st.write("")
        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            fig_pie = px.pie(df_final, names='Status', title='Proporsi Status Supply Chain', color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_pie, use_container_width=True)
        with col_chart2:
            top_v = df_final[df_final['Vendor'] != '-']['Vendor'].value_counts().head(10).reset_index()
            top_v.columns = ['Vendor', 'Jumlah']
            fig_b = px.bar(top_v, x='Jumlah', y='Vendor', orientation='h', title='Top 10 Active Vendors', color='Jumlah')
            fig_b.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_b, use_container_width=True)

        st.subheader("Master Data Grid Preview")
        st.dataframe(df_final.head(100), use_container_width=True, height=380)
    else:
        st.warning("⚠️ Belum ada file arsip ditemukan. Silakan unggah dokumen di menu **📤 Upload Data**.")

# ==========================================
# TAB 2: UPLOAD DATA
# ==========================================
elif selected_tab == "📤 Upload Data":
    st.subheader("Purchase Data Tracking | Document Source Hub")
    st.markdown('<div class="erp-panel">Unggah dokumen sumber (.xlsx). Sistem akan menormalisasi duplikasi spasi, konsolidasi multi-PR dalam satu PO, dan pengecekan status <i>RequestClosed</i> secara otomatis.</div>', unsafe_allow_html=True)
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        file_pr_2026 = st.file_uploader("PR Data (Base 2026)", type=['xlsx'], key="u_pr_base")
        file_pr_lama = st.file_uploader("PRN Data (Closed Status Audit)", type=['xlsx'], key="u_pr_closed")
        file_po_lokal = st.file_uploader("PO Data - Lokal", type=['xlsx'], key="u_po_lok")
    with col_u2:
        file_po_impor = st.file_uploader("PO Data - Impor", type=['xlsx'], key="u_po_imp")
        file_inbound = st.file_uploader("Inbound Data (Warehouse Receipts)", type=['xlsx'], key="u_inb")
        
    st.write("")
    run_process = st.button("⚙️ Execute ETL & Sync Engine", type="primary")
    
    if run_process:
        if all([file_pr_2026, file_pr_lama, file_po_lokal, file_po_impor, file_inbound]):
            with st.spinner('Menjalankan Data Transformation Pipeline...'):
                df_final = process_tracking_data(file_pr_2026, file_pr_lama, file_po_lokal, file_po_impor, file_inbound)
                now_dt = datetime.now()
                timestamp_file = now_dt.strftime("%Y%m%d_%H%M%S")
                saved_filename = f"Tracking_Final_{timestamp_file}.xlsx"
                saved_filepath = os.path.join(OUTPUT_FOLDER, saved_filename)
                df_final.to_excel(saved_filepath, index=False)
                
                st.session_state['df_final'] = df_final
                st.session_state['last_saved'] = saved_filename
                st.success(f"✅ Data berhasil diproses & diarsipkan ke server sebagai `{saved_filename}`! Silakan cek menu **📊 Dashboard**.")
        else:
            st.error("⚠️ Error: Kelima file sumber wajib diunggah lengkap!")

# ==========================================
# TAB 3: DOWNLOAD / ARSIP
# ==========================================
elif selected_tab == "📥 Download / Arsip":
    st.subheader("Purchase Data Tracking | Audit & Document Repository")
    st.markdown('<div class="erp-panel">Arsip historis hasil eksekusi pengolahan data tersimpan di server. File terbaru diberi penanda khusus.</div>', unsafe_allow_html=True)
    
    archive_items = get_formatted_archive_list()
    
    if archive_items:
        labels = [item[0] for item in archive_items]
        filenames = [item[1] for item in archive_items]
        
        selected_label = st.selectbox("Select Archived Document:", labels)
        selected_filename = dict(zip(labels, filenames))[selected_label]
        
        file_path_dl = os.path.join(OUTPUT_FOLDER, selected_filename)
        if os.path.exists(file_path_dl):
            df_preview = pd.read_excel(file_path_dl)
            st.markdown(f"**Selected File Metadata:** `📁 {selected_filename}` | Total Records: `{len(df_preview)} rows`")
            st.dataframe(df_preview.head(50), use_container_width=True, height=380)
            
            with open(file_path_dl, "rb") as f:
                st.download_button(
                    label=f"📥 Download Repository File ({selected_filename})",
                    data=f.read(),
                    file_name=selected_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
    else:
        st.warning("⚠️ Belum ada file arsip ditemukan pada server direktori `saved_reports/`.")
