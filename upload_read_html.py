from io import StringIO, BytesIO
import streamlit as st
from bs4 import BeautifulSoup
import pandas as pd
import zipfile

st.set_page_config(layout="wide")
st.title("HTML/ZIP Table Viewer & Excel Exporter")

uploaded_file = st.file_uploader("Upload a ZIP file or a single HTML file", type=["zip", "html"])

def extract_tables_from_html(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    tables = soup.find_all("table")
    if tables:
        try:
            df = pd.read_html(StringIO(str(tables[0])))[0]
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [' '.join(map(str, col)).strip() for col in df.columns]
            df.iloc[0, 0] = str(df.iloc[0, 0]).replace('.', '').replace(')', '').replace('(', '-')
            for col in df.columns[1:]:
                df[col] = df[col].astype(str).str.replace('.', '', regex=False)
                df[col] = df[col].astype(str).str.replace(')', '', regex=False)
                df[col] = df[col].astype(str).str.replace('(', '-', regex=False)
                df[col] = pd.to_numeric(df[col], errors="coerce")
            return df
        except Exception as e:
            return f"Lỗi khi đọc bảng: {e}"
    return "Không tìm thấy bảng nào trong file HTML."

html_tables = {}

if uploaded_file is not None:
    if uploaded_file.name.lower().endswith(".zip"):
        with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
            html_files = [f for f in zip_ref.namelist() if f.lower().endswith(".html")]
            for html_file in html_files:
                with zip_ref.open(html_file) as file:
                    html_content = file.read().decode("utf-8")
                    result = extract_tables_from_html(html_content)
                    html_tables[html_file] = result
    elif uploaded_file.name.lower().endswith(".html"):
        html_content = uploaded_file.read().decode("utf-8")
        result = extract_tables_from_html(html_content)
        html_tables[uploaded_file.name] = result

if html_tables:
    # Tạo file Excel chứa tất cả bảng
    output_all = BytesIO()
    with pd.ExcelWriter(output_all, engine='xlsxwriter') as writer:
        for name, df in html_tables.items():
            if isinstance(df, pd.DataFrame):
                sheet_name = name[:31]
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                for i, col in enumerate(df.columns):
                    column_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
                    align_format = workbook.add_format({'align': 'left' if i == 0 else 'right'})
                    worksheet.set_column(i, i, column_len, align_format)
    output_all.seek(0)

    # Nút tải tất cả bảng ở góc phải phía trên
    top_col1, top_col2 = st.columns([5, 1])
    with top_col2:
        st.download_button(
            label="📥 Tải tất cả bảng",
            data=output_all.getvalue(),
            file_name="all_tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Hiển thị bảng trong các tab
    tabs = st.tabs(list(html_tables.keys()))
    for tab, name in zip(tabs, html_tables.keys()):
        with tab:
            table = html_tables[name]
            if isinstance(table, pd.DataFrame):
                df_formatted = table.copy()
                for col in df_formatted.select_dtypes(include='number').columns:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else "")
                df_formatted.columns = df_formatted.columns.map(str)
                styles = [{'selector': 'th', 'props': [('text-align', 'center')]}]
                styles.append({'selector': 'td.col0', 'props': [('text-align', 'left')]} )
                for i in range(1, len(df_formatted.columns)):
                    styles.append({'selector': f'td.col{i}', 'props': [('text-align', 'right')]} )
                styled_table = df_formatted.style.set_table_styles(styles).set_table_attributes('style="font-size:12px;"')
                st.markdown(styled_table.to_html(), unsafe_allow_html=True)
            else:
                st.error(table)
else:
    st.info("Vui lòng tải lên file HTML hoặc ZIP chứa các file HTML.")
