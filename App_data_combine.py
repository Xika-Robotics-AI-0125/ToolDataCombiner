import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. Cấu hình trang Dashboard giao diện rộng (Wide Mode)
st.set_page_config(
    page_title="Excel Combiner",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS giúp giao diện chuyên nghiệp
st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #64748B;
        margin-bottom: 2rem;
    }
    .stMetric {
        background-color: #F8FAFC;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Thanh bên (Sidebar)
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/732/732220.png", width=50)
    st.title("⚙️ Advanced Configuration")
    st.caption("Customize data processing rules")
    
    st.divider()
    
    st.subheader("1. Data Processing")
    remove_dup = st.checkbox("Remove duplicates", value=True)
    drop_na_rows = st.checkbox("Remove completely empty rows", value=True)
    add_timestamp = st.checkbox("Add Consolidate Timestamp Column", value=True)
    add_source_file = st.checkbox("Add Source Filename Column", value=True)
    
    st.divider()
    
    st.subheader("2. Excel Sheet Options")
    # Đã sửa lỗi thừa dấu ngoặc kép
    sheet_option = st.radio(
        "Select sheet to extract data from:",
        ["First sheet", "Combine all sheets"]
    )
    
    st.divider()
    st.info("💡 **Tip:** Double-check header column names in individual files to ensure accurate data matching.")

# 3. Màn hình chính (Main Content)
st.markdown('<div class="main-header">⚡ Excel Data Combiner Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automated system for consolidating and standardizing Excel data from multiple sources</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📥 1. Source Files")
    uploaded_files = st.file_uploader(
        "Drag or select multiple individual files:",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="You can select dozens of Excel files at once."
    )

with col2:
    st.markdown("### 📗 2. Total File (Optional)")
    total_file = st.file_uploader(
        "Drop existing Total file to append new data:",
        type=["xlsx", "xls", "csv"],
        help="New data will be appended below existing data."
    )

st.divider()

# 4. Thao tác xử lý dữ liệu
if st.button("🚀 START CONSOLIDATING DATA", type="primary", use_container_width=True):
    if not uploaded_files and total_file is None:
        st.error("⚠️ Please upload at least one individual file or a Total file!")
    else:
        with st.spinner("⏳ Analyzing, cleaning, and merging data..."):
            all_dfs = []
            file_count = 0
            
            def read_excel_file(file, filename):
                dfs = []
                if filename.endswith('.csv'):
                    df = pd.read_csv(file)
                    if add_source_file:
                        df['Source_File'] = filename
                    dfs.append(df)
                else:
                    xl = pd.ExcelFile(file)
                    # Đã sửa lỗi so sánh điều kiện chuỗi chữ "First sheet"
                    sheets_to_read = [xl.sheet_names[0]] if sheet_option == "First sheet" else xl.sheet_names
                    for s in sheets_to_read:
                        df = pd.read_excel(file, sheet_name=s)
                        if add_source_file:
                            df['Source_File'] = f"{filename} ({s})" if len(sheets_to_read) > 1 else filename
                        dfs.append(df)
                return pd.concat(dfs, ignore_index=True)

            # Đọc file Total cũ nếu có
            if total_file is not None:
                df_total_exist = read_excel_file(total_file, total_file.name)
                all_dfs.append(df_total_exist)

            # Đọc các file nhỏ
            if uploaded_files:
                for file in uploaded_files:
                    df_small = read_excel_file(file, file.name)
                    all_dfs.append(df_small)
                    file_count += 1

            # Tiến hành Gộp Dữ Liệu
            final_df = pd.concat(all_dfs, ignore_index=True)
            
            if drop_na_rows:
                final_df.dropna(how='all', inplace=True)
                
            if remove_dup:
                cols_to_check = [c for c in final_df.columns if c not in ['Source_File', 'Consolidated_At']]
                final_df.drop_duplicates(subset=cols_to_check, inplace=True)

            if add_timestamp:
                final_df['Consolidated_At'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            st.balloons()
            st.success("🎉 **Data processing completed successfully!**")
            
            # Dashboard Thống kê (Đã chuẩn hóa Tiếng Anh)
            st.markdown("### 📊 Processing Summary")
            m1, m2, m3 = st.columns(3)
            m1.metric(label="Total Child Files Merged", value=f"{file_count} File(s)")
            m2.metric(label="Total Data Rows", value=f"{len(final_df):,} rows")
            m3.metric(label="Total Data Columns", value=f"{len(final_df.columns)} columns")

            st.divider()

            # Xem trước dữ liệu
            st.markdown("### 👀 Preview Aggregated Data")
            st.dataframe(final_df, use_container_width=True, height=350)

            st.divider()

            # Khu vực Tải Xuất Dữ Liệu
            st.markdown("### 📥 Download Results")
            d_col1, d_col2 = st.columns(2)
            
            # Xuất Excel
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            d_col1.download_button(
                label="📗 Download EXCEL File (.xlsx)",
                data=buffer_excel.getvalue(),
                file_name=f"Total_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # Xuất CSV
            csv_data = final_df.to_csv(index=False).encode('utf-8-sig')
            d_col2.download_button(
                label="📄 Download CSV File (.csv)",
                data=csv_data,
                file_name=f"Total_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
