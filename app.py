import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from datetime import datetime
import openpyxl

# Opt-in to future pandas behavior to avoid FutureWarning during replace
pd.set_option('future.no_silent_downcasting', True)

st.set_page_config(page_title="Excel Processor", layout="wide")

st.title("📊 Xử lí dữ liệu raw cho NVL")

st.markdown("""
Cập nhật các cột (Sản phẩm, LỆNH SX) trong Sheet library của file Material report với thông tin của lô hàng mới trước khi upload file

Upload:
1. **Báo cáo đóng lô hàng** (Chứa các sheets "Báo cáo tổng hợp" + "Mẫu đóng lô")
2. **Báo cáo NVL (Material report)** (Chứa các sheets "RawData" + "Library")

Click **Process** để tạo kết quả đầu ra.
""")

# Upload files
main_file = st.file_uploader("Upload Báo cáo đóng lô hàng", type=["xlsx"])
raw_file = st.file_uploader("Upload Báo cáo NVL", type=["xlsx"])

def get_next_non_empty(df, search_text):
    for _, row in df.iterrows():
        for col_idx, val in enumerate(row):
            if isinstance(val, str):
                clean_val = " ".join(val.replace("\n", " ").split())
                if search_text.lower() in clean_val.lower():
                    for next_val in row.iloc[col_idx+1:]:
                        if pd.notna(next_val) and str(next_val).strip() != '':
                            try:
                                return int(float(next_val))
                            except:
                                continue
    return None

if st.button("🚀 Process"):

    if not main_file or not raw_file:
        st.warning("Please upload both files")
        st.stop()

    gif_placeholder = st.empty()
    gif_path = "1711970569877.gif"
    if os.path.exists(gif_path):
        with gif_placeholder.container():
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.image(gif_path)

    try:
        # =============================
        # READ MAIN FILE
        # =============================
        # Use openpyxl to get ONLY visible sheets
        wb_check = openpyxl.load_workbook(main_file, read_only=True)
        sheet_names = [ws.title for ws in wb_check.worksheets if ws.sheet_state == 'visible']
        wb_check.close()

        # Reset main_file pointer after reading with openpyxl
        main_file.seek(0)

        def find_sheet(partial_list):
            if isinstance(partial_list, str):
                partial_list = [partial_list]
            for partial in partial_list:
                for s in sheet_names:
                    if partial.lower() in s.lower():
                        return s
            raise Exception(f"Sheet containing any of {partial_list} not found among visible sheets")

        data_sheet = find_sheet(["Báo cáo tổng hợp", "BÁO CÁO ĐÓNG LỆNH"])
        mau_sheet = find_sheet(["Mẫu đóng lô"])

        df_data = pd.read_excel(main_file, sheet_name=data_sheet, engine="openpyxl")
        main_file.seek(0)

        # Read raw to get metadata from top rows
        df_mau_raw = pd.read_excel(main_file, sheet_name=mau_sheet, engine="openpyxl")
        sl_lo_sx = get_next_non_empty(df_mau_raw, "Số lượng lô sx")
        sl_sx_pcba = get_next_non_empty(df_mau_raw, "Số lượng sx PCBA")
        sl_sx_pkg = get_next_non_empty(df_mau_raw, "Số lượng sx PKG")
        if sl_lo_sx is None:
            sl_lo_sx = get_next_non_empty(df_mau_raw, "Số lượng lệnh sx")
        main_file.seek(0)

        # Read actual data table without header to find 'Level'
        temp_df = pd.read_excel(main_file, sheet_name=mau_sheet, header=None, engine="openpyxl")
        main_file.seek(0)

        header_idx = 12 # default fallback
        for idx, row in temp_df.iterrows():
            if any(str(cell).strip().lower() == 'level' for cell in row):
                header_idx = idx
                break

        df_mau = pd.read_excel(main_file, sheet_name=mau_sheet, header=header_idx, engine="openpyxl")

        # Fix unnamed columns & clean up whitespace/newlines
        new_cols = []
        for col in df_mau.columns:
            if str(col).startswith("Unnamed"):
                idx = df_mau[col].first_valid_index()
                val = str(df_mau.loc[idx, col]) if idx is not None else str(col)
            else:
                val = str(col)

            # Clean up newlines and extra spaces
            val = " ".join(val.replace("\n", " ").replace("\r", "").split())
            new_cols.append(val)

        df_mau.columns = new_cols
        df_mau = df_mau.iloc[5:].reset_index(drop=True)

        # End the table where 'Level' (case-insensitive) is empty or has value 'Tổng'
        level_col = next((c for c in df_mau.columns if str(c).strip().lower() == 'level'), None)

        if level_col:
            level_str = df_mau[level_col].astype(str).str.strip().str.lower()
            end_condition = df_mau[level_col].isna() | (level_str == '') | (level_str == 'tổng')

            if end_condition.any():
                first_empty_idx = end_condition.idxmax()
                df_mau = df_mau.iloc[:first_empty_idx]

        # Helper to robustly find column names safely
        def get_col_data(df, possible_names):
            if isinstance(possible_names, str):
                possible_names = [possible_names]
            for p_name in possible_names:
                for c in df.columns:
                    if p_name.lower() in str(c).lower():
                        return df[c]
                # Try removing spaces to match 'tỉlệ' vs 'tỉ lệ'
                compact_p = p_name.lower().replace(" ", "")
                for c in df.columns:
                    if compact_p in str(c).lower().replace(" ", ""):
                        return df[c]
            return pd.Series([np.nan] * len(df))

        # Find prod_name by searching for "Tên sản phẩm" and getting value 3, 2, or 1 rows down
        prod_name = None
        for r_idx in range(len(df_data)):
            for c_idx in range(len(df_data.columns)):
                val = df_data.iloc[r_idx, c_idx]
                if isinstance(val, str) and "Tên sản phẩm" in val:
                    for offset in [3, 2, 1]:
                        if r_idx + offset < len(df_data):
                            prod_val = df_data.iloc[r_idx + offset, c_idx]
                            if pd.notna(prod_val) and str(prod_val).strip() != '':
                                prod_name = str(prod_val).strip()
                                break
                    break
            if prod_name:
                break

        # Find lenh_sx by searching for "Số lệnh:"
        lenh_sx = None
        for row_idx, row in df_data.iterrows():
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and "Số lệnh:" in val:
                    parts = val.split("Số lệnh:")
                    if len(parts) > 1 and parts[1].strip() != "":
                        lenh_sx = parts[1].strip()
                    elif col_idx + 1 < len(row):
                        next_val = row.iloc[col_idx + 1]
                        if pd.notna(next_val):
                            lenh_sx = str(next_val).strip()
                    break
            if lenh_sx:
                break

        # =============================
        # READ RAW FILE
        # =============================
        df_rawdata = pd.read_excel(raw_file, sheet_name="RawData",
                                  usecols=["MÃ VT", "GIÁ TRỊ VẬT TƯ/SP", "BoM TYPE"], engine="openpyxl")
        raw_file.seek(0)

        df_library = pd.read_excel(raw_file, sheet_name="Library", header=1, engine="openpyxl")
        df_library.columns = [str(c).strip() for c in df_library.columns]

        # =============================
        # BUILD OUTPUT
        # =============================
        columns = [
            "PRODUCT","LỆNH SX","SL CỦA LỆNH","KQSX","MÃ VT 14","MÃ VT",
            "TÊN VT","CĐ","ĐM VẬT TƯ","TỔNG VT SD","SL TIÊU HAO THỰC TẾ",
            "SL TIÊU HAO ĐM","TỶ LỆ TH THỰC TẾ","TỶ LỆ TH ĐM",
            "CHI PHÍ TIÊU HAO THỰC TẾ","CHI PHÍ TIÊU HAO ĐM","CHÊNH LỆCH CHI PHÍ",
            "GIÁ TRỊ VẬT TƯ/SP","GHI CHÚ","BoM TYPE","PRODUCT SERIES","Months",
            "Sản phẩm mất đồng bộ","Giá trị vật tư lãng phí/ Sản phẩm",
            "Giá trị vật tư lãng phí/ đơn vị vật tư","Date","Year"
        ]

        raw_data = pd.DataFrame(columns=columns)

        # Safely extract columns without KeyError
        raw_data['MÃ VT 14'] = get_col_data(df_mau, 'Mã vật tư 14 ký tự')
        raw_data['MÃ VT'] = get_col_data(df_mau, 'Mã vật tư sử dụng 16 ký tự')
        raw_data['TÊN VT'] = get_col_data(df_mau, 'Tên vật tư')
        raw_data['CĐ'] = get_col_data(df_mau, 'Công đoạn').replace(['TOP', 'BOT'], 'SMT')
        raw_data['ĐM VẬT TƯ'] = pd.to_numeric(get_col_data(df_mau, ['Tổng ĐM', 'ĐM', 'TỔNG ĐM']), errors='coerce')

        raw_data['TỔNG VT SD'] = pd.to_numeric(get_col_data(df_mau, ['Tổng vật tư sử dụng theo BOM', 'Vật tư sử dụng trong LSX']), errors='coerce')
        raw_data['SL TIÊU HAO THỰC TẾ'] = pd.to_numeric(get_col_data(df_mau, 'Tổng vật tư tiêu hao'), errors='coerce')

        # Also handle potential 'tỉ lệ' vs 'tỷ lệ'
        raw_data['TỶ LỆ TH ĐM'] = pd.to_numeric(get_col_data(df_mau, ['Tỉ lệ tiêu hao định mức', 'Tỷ lệ tiêu hao định mức', 'lệ tiêu hao định mức']), errors='coerce')

        raw_data['PRODUCT'] = prod_name
        raw_data['LỆNH SX'] = lenh_sx
        raw_data.loc[0, 'SL CỦA LỆNH'] = sl_lo_sx
        raw_data.loc[0, 'KQSX'] = sl_sx_pkg if sl_sx_pkg is not None else sl_sx_pcba

        raw_data['SL TIÊU HAO ĐM'] = raw_data['TỔNG VT SD'] * raw_data['TỶ LỆ TH ĐM']
        raw_data['TỶ LỆ TH THỰC TẾ'] = (
            raw_data['SL TIÊU HAO THỰC TẾ'] / raw_data['TỔNG VT SD']
        ).replace([np.inf, -np.inf], 0).fillna(0)

        # Mapping
        map_price = df_rawdata.set_index('MÃ VT')['GIÁ TRỊ VẬT TƯ/SP'].to_dict()
        raw_data['GIÁ TRỊ VẬT TƯ/SP'] = raw_data['MÃ VT'].map(map_price)
        bom_mapping = df_rawdata.set_index('MÃ VT')['BoM TYPE'].to_dict()
        raw_data['BoM TYPE'] = raw_data['MÃ VT'].map(bom_mapping)

        raw_data['CHI PHÍ TIÊU HAO THỰC TẾ'] = raw_data['SL TIÊU HAO THỰC TẾ'] * raw_data['GIÁ TRỊ VẬT TƯ/SP']
        raw_data['CHI PHÍ TIÊU HAO ĐM'] = raw_data['SL TIÊU HAO ĐM'] * raw_data['GIÁ TRỊ VẬT TƯ/SP']
        raw_data['CHÊNH LỆCH CHI PHÍ'] = raw_data['CHI PHÍ TIÊU HAO THỰC TẾ'] - raw_data['CHI PHÍ TIÊU HAO ĐM']

        raw_sp_mat = ((raw_data['SL TIÊU HAO THỰC TẾ'] - raw_data['SL TIÊU HAO ĐM']) / raw_data['ĐM VẬT TƯ']).replace([np.inf, -np.inf], 0).fillna(0)
        raw_data['Sản phẩm mất đồng bộ'] = np.where(raw_sp_mat < 0, np.floor(raw_sp_mat), np.ceil(raw_sp_mat))
        raw_data['Sản phẩm mất đồng bộ'] = pd.Series(raw_data['Sản phẩm mất đồng bộ']).fillna(0)

        raw_data['Giá trị vật tư lãng phí/ đơn vị vật tư'] = (
            raw_data['CHI PHÍ TIÊU HAO THỰC TẾ'] / raw_data['TỔNG VT SD']
        ).replace([np.inf, -np.inf], 0).fillna(0)

        # Date mapping
        lenh_col = next((c for c in df_library.columns if 'lệnh' in c.lower()), None)

        if lenh_col and 'Date' in df_library.columns and 'Year' in df_library.columns:
            date_map = df_library.set_index(lenh_col)['Date'].to_dict()
            year_map = df_library.set_index(lenh_col)['Year'].to_dict()

            raw_data['Months'] = raw_data['LỆNH SX'].map(date_map)
            raw_data['Year'] = raw_data['LỆNH SX'].map(year_map)

        def make_date(row):
            try:
                dt = datetime.strptime(
                    f"{int(row['Year'])}-{str(row['Months'])[:3]}-17",
                    "%Y-%b-%d"
                )
                return dt.strftime("%d-%b-%Y")
            except:
                return pd.NaT

        raw_data['Date'] = raw_data.apply(make_date, axis=1)

        # =============================
        # DOWNLOAD
        # =============================
        output = io.BytesIO()
        raw_data.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        gif_placeholder.empty()

        st.success("✅ Done! Ấn download để tải file excel về. Copy kết quả sang Báo cáo NVL (Material report) và xử lí nốt các cột.")

        st.download_button(
            label="📥 Download Result",
            data=output,
            file_name="RawData.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(raw_data.head())

    except Exception as e:
        gif_placeholder.empty()
        st.error(f"❌ Error: {e}")
