import streamlit as st
import pandas as pd
import io
import os

# --- 設定 ---
# 確保 'Product Info_All Countries.xlsx' 與此 webapp.py 位於同一個資料夾內
MASTER_FILE = 'Product Info_All Countries.xlsx'

def process_order_data(uploaded_file):
    """處理上傳的 Excel 檔案並回傳處理後的 DataFrame"""
    # 1. 讀取上傳的檔案
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row = None
    for idx, row in raw_df.iterrows():
        if row.astype(str).str.contains('XITEM|ITEM#|ITEM #', case=False, na=False).any():
            header_row = idx
            break
    
    if header_row is None:
        raise ValueError("無法找到標題列 (Header Row)")

    po_df = pd.read_excel(uploaded_file, header=header_row)
    po_df.columns = po_df.columns.astype(str).str.strip()
    po_df['original_index'] = po_df.index
    
    # 2. 識別關鍵欄位
    item_col = next((c for c in po_df.columns if any(x in c.upper() for x in ['XITEM', 'ITEM#', 'ITEM #'])), None)
    country_cols = [c for c in po_df.columns if 'country' in c.lower()]
    
    country_col = None
    known_codes = ['HK', 'TW', 'ER', 'ZENZO', 'ZENSO', 'MKT', 'TWHR', 'JP']
    for c in country_cols:
        sample_values = po_df[c].fillna('').astype(str).str.upper()
        if any(code in ' '.join(sample_values) for code in known_codes):
            country_col = c
            break
    
    if not item_col or not country_col:
        raise ValueError(f"無法識別欄位。找到的欄位: {list(po_df.columns)}")
        
    po_df.rename(columns={item_col: 'XITEM'}, inplace=True)
    po_df = po_df[po_df['XITEM'].notna() & (po_df['XITEM'] != '')].copy()
    
    # 3. 處理尺寸 (Matrix to Long)
    possible_sizes = ['0', '1', '2', '3', '4', '24', '26', '28', '30', '32', '34', '36', '38', '40', '42', '44', 'M', 'L', 'UNI', 'S', 'XL']
    existing_size_cols = [c for c in po_df.columns if str(c) in possible_sizes]
    
    if existing_size_cols:
        id_vars = [c for c in po_df.columns if c not in existing_size_cols]
        po_long = po_df.melt(id_vars=id_vars, value_vars=existing_size_cols, var_name='SIZENAME', value_name='qty')
    else:
        po_long = po_df.copy()
        po_long['SIZENAME'] = 'UNIQUE'
        qty_cols = [c for c in po_df.columns if 'qty' in c.lower() or 'quantity' in c.lower()]
        po_long['qty'] = po_df[qty_cols[0]] if qty_cols else 1
        
    po_long = po_long[po_long['qty'].notna() & (po_long['qty'] != 0)]
    po_long['SIZENAME'] = po_long['SIZENAME'].astype(str).str.strip().replace({'UNI': 'UNIQUE'})
    
    # 4. 國家映射
    mapping = {
        'HK': 'HKD', 'MKT': 'HKD', 'TW': 'TWD', 'TWHR': 'TWD', 'JP': 'JPY', 
        'Zenzo (SG)': 'Zenzo_SGD', 'Zenso (SG)': 'Zenzo_SGD', 'Zenzo': 'Zenzo_SGD', 'ER': 'ER'
    }
    po_long['Target_Sheet'] = po_long[country_col].apply(lambda x: next((v for k, v in mapping.items() if k.lower() in str(x).lower()), None))
    
    # 5. 合併與最終處理
    results = []
    for sheet, group in po_long.groupby('Target_Sheet'):
        if not sheet or not os.path.exists(MASTER_FILE): continue
        
        master_df = pd.read_excel(MASTER_FILE, sheet_name=sheet)
        master_df.rename(columns={'SIZENAME': 'SIZENAME'}, inplace=True)
        master_df['XITEM'] = master_df['XITEM'].astype(str).str.strip()
        master_df['Color'] = master_df['Color'].astype(str).str.zfill(3)
        master_df['SIZENAME'] = master_df['SIZENAME'].astype(str).str.strip()
        
        group = group.copy()
        group['XITEM'] = group['XITEM'].astype(str).str.strip()
        color_col = next((c for c in group.columns if 'color' in c.lower()), None)
        group['Color'] = group[color_col].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(3)
        
        merged = pd.merge(group, master_df, on=['XITEM', 'Color', 'SIZENAME'], how='left')
        results.append(merged)
    
    final_df = pd.concat(results, ignore_index=True)
    final_df = final_df.sort_values(['original_index', 'SIZENAME'])
    
    # 應用更新後的邏輯
    final_df['MPO'] = final_df['PO#'] if 'PO#' in final_df.columns else ""
    final_df['UPC'] = final_df['BCODE13'] if 'BCODE13' in final_df.columns else ""
    final_df['qty'] = final_df['qty'].astype(int)
    
    if 'Currency' in final_df.columns:
        final_df['Currency'] = final_df['Currency'].astype(str).str.strip().replace({'TWD': 'NT$', 'twd': 'NT$', 'Twd': 'NT$'})
        
    for col in ['Code', 'Color']:
        if col in final_df.columns:
            final_df[col] = final_df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(3)
    
    final_cols = ['MPO', 'UPC', 'XITEM', 'BCODE13', 'PL1C1C2', 'Designation', 'Color', 'Color desc', 'SIZENAME', 'Price', 'Currency', 'LORDS', 'Code', 'qty']
    return final_df[[c for c in final_cols if c in final_df.columns]]

# --- Streamlit 網頁介面 ---
st.title("Order Form Generator")
uploaded_file = st.file_uploader("請上傳您的 PO Excel 檔案", type=["xlsx"])

if uploaded_file is not None:
    if st.button("開始處理"):
        try:
            with st.spinner('處理中...'):
                df_result = process_order_data(uploaded_file)
                
                # 轉換為 CSV 格式供下載
                csv_buffer = io.BytesIO()
                df_result.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                
                st.success("成功！訂單表已生成。")
                st.download_button(
                    label="下載 Order_Form_Output.csv",
                    data=csv_buffer.getvalue(),
                    file_name="Order_Form_Output.csv",
                    mime="text/csv"
                )
        except Exception as e:
            st.error(f"發生錯誤: {e}")
