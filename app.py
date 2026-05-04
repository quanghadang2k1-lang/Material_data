import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

st.set_page_config(page_title="Excel Processor", layout="wide")

st.title("📊 Excel Processing Tool")

st.markdown("""
Upload:
1. **Main Excel file** (contains "Data" + "Mẫu đóng lô")
2. **RawData file** (contains "RawData" + "Library")

Then click **Process** to generate output.
""")

# Upload files
main_file = st.file_uploader("Upload Main Excel file", type=["xlsx"])
raw_file = st.file_uploader("Upload RawData Excel file", type=["xlsx"])

def get_next_non_empty(df, search_text):
    for _, row in df.iterrows():
        for col_idx, val in enumerate(row):
            if isinstance(val, str) and search_text.lower() in val.lower():
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

    try:
        # =============================
        # READ MAIN FILE
        # =============================
        xls = pd.ExcelFile(main_file)
        sheet_names = xls.sheet_names

        def find_sheet(partial):
            for s in sheet_names:
                if partial.lower() in s.lower():
                    return s
            raise Exception(f"Sheet '{partial}' not found")

        data_sheet = find_sheet("Data")
        mau_sheet = find_sheet("Mẫu đóng lô")

        df_data = pd.read_excel(main_file, sheet_name=data_sheet)
        df_mau = pd.read_excel(main_file, sheet_name=mau_sheet, header=12)

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

        # Extract values
        sl_lo_sx = get_next_non_empty(df_mau, "Số lượng lô sx")
        sl_sx_pkg = get_next_non_empty(df_mau, "Số lượng sx PKG")

        prod_name = df_data['SẢN PHẨM'].iloc[0]
        lenh_sx = df_data['Sản phẩm - lệnh sản xuất'].iloc[0]

        # =============================
        # READ RAW FILE
        # =============================
        df_rawdata = pd.read_excel(raw_file, sheet_name="RawData",
                                  usecols=["MÃ VT", "GIÁ TRỊ VẬT TƯ/SP"])

        df_library = pd.read_excel(raw_file, sheet_name="Library", header=1)
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
        raw_data['CĐ'] = get_col_data(df_mau, 'Công đoạn')
        raw_data['ĐM VẬT TƯ'] = get_col_data(df_mau, 'ĐM')
        
        raw_data['TỔNG VT SD'] = pd.to_numeric(get_col_data(df_mau, 'Tổng vật tư sử dụng theo BOM'), errors='coerce')
        raw_data['SL TIÊU HAO THỰC TẾ'] = pd.to_numeric(get_col_data(df_mau, 'Tổng vật tư tiêu hao'), errors='coerce')
        
        # Also handle potential 'tỉ lệ' vs 'tỷ lệ'
        raw_data['TỶ LỆ TH ĐM'] = pd.to_numeric(get_col_data(df_mau, ['Tỉ lệ tiêu hao định mức', 'Tỷ lệ tiêu hao định mức', 'lệ tiêu hao định mức']), errors='coerce')

        raw_data['PRODUCT'] = prod_name
        raw_data['LỆNH SX'] = lenh_sx
        raw_data.loc[0, 'SL CỦA LỆNH'] = sl_lo_sx
        raw_data.loc[0, 'KQSX'] = sl_sx_pkg

        raw_data['SL TIÊU HAO ĐM'] = raw_data['TỔNG VT SD'] * raw_data['TỶ LỆ TH ĐM']
        raw_data['TỶ LỆ TH THỰC TẾ'] = (
            raw_data['SL TIÊU HAO THỰC TẾ'] / raw_data['TỔNG VT SD']
        ).replace([np.inf, -np.inf], 0).fillna(0)

        # Mapping
        map_price = df_rawdata.set_index('MÃ VT')['GIÁ TRỊ VẬT TƯ/SP'].to_dict()
        raw_data['GIÁ TRỊ VẬT TƯ/SP'] = raw_data['MÃ VT'].map(map_price)

        raw_data['CHI PHÍ TIÊU HAO THỰC TẾ'] = raw_data['SL TIÊU HAO THỰC TẾ'] * raw_data['GIÁ TRỊ VẬT TƯ/SP']
        raw_data['CHI PHÍ TIÊU HAO ĐM'] = raw_data['SL TIÊU HAO ĐM'] * raw_data['GIÁ TRỊ VẬT TƯ/SP']
        raw_data['CHÊNH LỆCH CHI PHÍ'] = raw_data['CHI PHÍ TIÊU HAO THỰC TẾ'] - raw_data['CHI PHÍ TIÊU HAO ĐM']

        raw_data['Sản phẩm mất đồng bộ'] = np.ceil(
            ((raw_data['SL TIÊU HAO THỰC TẾ'] - raw_data['SL TIÊU HAO ĐM']) / raw_data['ĐM VẬT TƯ'])
        ).replace([np.inf, -np.inf], 0).fillna(0)

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
                return datetime.strptime(
                    f"{int(row['Year'])}-{str(row['Months'])[:3]}-17",
                    "%Y-%b-%d"
                )
            except:
                return pd.NaT

        raw_data['Date'] = raw_data.apply(make_date, axis=1)

        # =============================
        # DOWNLOAD
        # =============================
        output = io.BytesIO()
        raw_data.to_excel(output, index=False)
        output.seek(0)

        st.success("✅ Done!")

        st.download_button(
            label="📥 Download Result",
            data=output,
            file_name="processed_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(raw_data.head())

    except Exception as e:
        st.error(f"❌ Error: {e}")
