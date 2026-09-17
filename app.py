import gradio as gr
import pandas as pd
import numpy as np
import re
import tempfile

def process_data(pr_new, pr_old, po_lok, po_imp, inb):
    # Proses PR 2026
    pr_df = pd.read_excel(pr_new.name, header=1)
    pr_data = pr_df[['Date', 'PR Number \n(Manual)', 'Item \nCode', 'Item Description', 'Qty']].copy()
    pr_data.columns = ['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty']
    pr_data['Item_Code'] = pr_data['Item_Code'].astype(str).str.strip()
    pr_data['PR_Manual_No_Clean'] = pr_data['PR_Manual_No'].astype(str).str.replace(" ", "")
    pr_data = pr_data[pr_data['PR_Manual_No'].notna() & (pr_data['PR_Manual_No'] != '') & (pr_data['PR_Manual_No'].astype(str) != 'nan')]

    # Proses Status Closed
    pr_2426_df = pd.read_excel(pr_old.name, sheet_name=0, header=None)
    closed_info = pr_2426_df.iloc[7:, [2, 6, 7]].copy()
    closed_info.columns = ['PR_Manual_No', 'RequestClosed', 'Item_Code']
    closed_info['Item_Code'] = closed_info['Item_Code'].astype(str).str.strip()
    closed_info['PR_Manual_No_Clean'] = closed_info['PR_Manual_No'].astype(str).str.replace(" ", "")
    closed_info = closed_info.drop_duplicates(subset=['PR_Manual_No_Clean', 'Item_Code'], keep='last')
    
    pr_data = pd.merge(pr_data, closed_info[['PR_Manual_No_Clean', 'Item_Code', 'RequestClosed']], on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    pr_data['RequestClosed'] = pr_data['RequestClosed'].fillna('No')

    # Proses PO
    def clean_po(filepath, tipe):
        po = pd.read_excel(filepath, header=13)
        po = po[['Purchase Order Number', 'PO Date', 'PR Manual No.', 'Item Code', 'Qty', 'Unnamed: 5']].copy()
        po.columns = ['PO_No', 'PO_Date', 'PR_Manual_No_Orig', 'Item_Code', 'PO_Qty', 'Vendor']
        po['PO_No'] = po['PO_No'].ffill()
        po['PO_Date'] = po['PO_Date'].ffill()
        po['Vendor'] = po['Vendor'].ffill()
        po = po.dropna(subset=['Item_Code'])
        po = po[po['Item_Code'].astype(str).str.strip() != 'nan']
        po['Tipe_PO'] = tipe
        return po

    po_df = pd.concat([clean_po(po_lok.name, "Lokal"), clean_po(po_imp.name, "Impor")], ignore_index=True)
    po_df['Item_Code'] = po_df['Item_Code'].astype(str).str.strip()

    exp_rows = []
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
                exp_rows.append(new_row)
        else:
            new_row = row.to_dict()
            new_row['PR_Manual_No_Clean'] = pr_str.replace(" ", "")
            exp_rows.append(new_row)

    po_exp = pd.DataFrame(exp_rows)
    po_exp['PO_Qty'] = pd.to_numeric(po_exp['PO_Qty'], errors='coerce').fillna(0)
    
    po_agg = po_exp.groupby(['PR_Manual_No_Clean', 'Item_Code']).agg(
        PO_No=('PO_No', lambda x: ', '.join(x.dropna().unique().astype(str))),
        PO_Qty=('PO_Qty', 'sum'),
        PO_Date=('PO_Date', 'max'),
        Vendor=('Vendor', lambda x: ', '.join(x.dropna().unique().astype(str))),
        Tipe_PO=('Tipe_PO', lambda x: ', '.join(x.dropna().unique().astype(str)))
    ).reset_index()

    merged = pd.merge(pr_data, po_agg, on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    merged['PO_No'] = merged['PO_No'].fillna('')
    merged = merged[~((merged['PO_No'] == '') & (merged['RequestClosed'] == 'Yes'))]

    # Proses Inbound
    inb_xls = pd.ExcelFile(inb.name)
    inb_df = pd.concat([pd.read_excel(inb_xls, sheet_name=s) for s in inb_xls.sheet_names])
    inb_df['ItemCode'] = inb_df['ItemCode'].astype(str).str.strip()
    inb_df['POnumber'] = inb_df['POnumber'].astype(str).str.strip()
    inb_agg = inb_df.groupby(['POnumber', 'ItemCode']).agg(Rcv_Date=('ReceivedDate', 'max'), Rcv_Qty=('RcvQty', 'sum')).reset_index()

    def get_inb(po_str, item_code):
        if not po_str: return pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})
        subset = inb_agg[(inb_agg['POnumber'].isin([p.strip() for p in po_str.split(',')])) & (inb_agg['ItemCode'] == item_code)]
        return pd.Series({'Rcv_Date': subset['Rcv_Date'].max(), 'Rcv_Qty': subset['Rcv_Qty'].sum()}) if not subset.empty else pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})

    merged[['Rcv_Date', 'Rcv_Qty']] = merged.apply(lambda row: get_inb(row['PO_No'], row['Item_Code']), axis=1)

    def get_status(row):
        if row['PO_No'] == '': return 'Routing Approval'
        elif row['Rcv_Qty'] == 0: return 'Menunggu Pengiriman'
        elif row['Rcv_Qty'] < row['PO_Qty']: return 'Diterima Sebagian'
        return 'Sudah Diterima'

    merged['Status'] = merged.apply(get_status, axis=1)
    
    merged['PR_Date'] = pd.to_datetime(merged['PR_Date']).dt.strftime('%m/%d/%Y').fillna('-')
    merged['PO_Date'] = pd.to_datetime(merged['PO_Date']).dt.strftime('%m/%d/%Y').fillna('-')
    merged['Rcv_Date'] = pd.to_datetime(merged['Rcv_Date']).dt.strftime('%m/%d/%Y').fillna('-')
    
    out_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    merged[['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty', 'PO_Date', 'PO_No', 'Vendor', 'Tipe_PO', 'PO_Qty', 'Rcv_Date', 'Rcv_Qty', 'Status']].to_excel(out_file, index=False)
    return out_file

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 📦 Tracker Pengadaan Barang")
    with gr.Row():
        with gr.Column():
            f1 = gr.File(label="1. PR 2026 (Base)")
            f2 = gr.File(label="2. PR 2024-2026 (Closed Status)")
            f3 = gr.File(label="3. PO 2026 Lokal")
            f4 = gr.File(label="4. PO 2026 Impor")
            f5 = gr.File(label="5. Inbound 2025-2026")
            btn = gr.Button("🚀 Proses Data", variant="primary")
        with gr.Column():
            out = gr.File(label="📥 Hasil Tracking (Excel)")
    
    btn.click(fn=process_data, inputs=[f1, f2, f3, f4, f5], outputs=out)

demo.launch()
