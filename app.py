import streamlit as st
import pandas as pd

st.title("Gộp nhiều sheet Excel")

uploaded_file = st.file_uploader("Tải lên file Excel (.xlsx)", type="xlsx")

if uploaded_file:
    col = ['STT', 'MÃ VẬT TƯ', 'TÊN HÀNG', 'ĐVT', 'SL.NHẬP ', 'CÔNG TY CUNG CẤP HÀNG', 'SL.XUẤT (PXGV)', 'GHI CHÚ']

    try:
        xls = pd.ExcelFile(uploaded_file)
        dfs = []

        for sheet in xls.sheet_names:
            df = xls.parse(sheet, header=5)

            if set(col).issubset(df.columns):
                df = df[col]
                df = df[pd.to_numeric(df['STT'], errors='coerce').notnull()]
                df['STT'] = df['STT'].astype(int)
                dfs.append(df)
            else:
                st.warning(f"Sheet '{sheet}' thiếu cột cần thiết. Bỏ qua.")

        if dfs:
            df_all = pd.concat(dfs, ignore_index=True)
            st.success("Gộp thành công!")

            st.dataframe(df_all)

            # Xuất ra file Excel
            output_file = "merged_sheets.xlsx"
            df_all.to_excel(output_file, index=False)

            with open(output_file, "rb") as f:
                st.download_button("📥 Tải file Excel đã gộp", f, file_name="merged_sheets.xlsx")

        else:
            st.error("Không có dữ liệu hợp lệ để gộp.")

    except Exception as e:
        st.error(f"Lỗi: {e}")
