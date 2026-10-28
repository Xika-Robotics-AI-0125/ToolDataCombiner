import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 1. Cấu hình trang Dashboard giao diện rộng (Wide Mode)
st.set_page_config(
    page_title="Excel Combiner Pro",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Thêm Custom CSS giúp giao diện chuyên nghiệp, hiện đại
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

# 2. Thanh bên (Sidebar) - Các thiết lập & Tùy chọn nâng cao
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/732/732220.png", width=50)
    st.title("⚙️ Cấu Hình Nâng Cao")
    st.caption("Tùy chỉnh quy tắc xử lý dữ liệu")
    
    st.divider()
    
    st.subheader("1. Xử lý dữ liệu")
    remove_dup = st.checkbox("Xóa trùng lặp (Remove Duplicates)", value=True)
    drop_na_rows = st.checkbox("Bỏ các dòng trống toàn bộ", value=True)
    add_timestamp = st.checkbox("Thêm cột Thời Gian Gộp", value=True)
    add_source_file = st.checkbox("Thêm cột Tên File Nguồn", value=True)
    
    st.divider()
    
    st.subheader("2. Tùy chọn Sheet Excel")
    sheet_option = st.radio(
        "Chọn Sheet cần lấy dữ liệu:",
        ["Sheet đầu tiên", "Gộp tất cả các Sheet trong file"]
    )
    
    st.divider()
    st.info("💡 **Mẹo:** Kiểm tra kỹ tên các cột tiêu đề ở các file lẻ để đảm bảo khớp dữ liệu chính xác nhất.")

# 3. Màn hình chính (Main Content)
st.markdown('<div class="main-header">⚡ Excel Data Combiner Pro</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Hệ thống tự động hợp nhất và chuẩn hóa dữ liệu Excel đa nguồn</div>', unsafe_allow_html=True)

# Chia layout thành 2 cột tải file
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📥 1. Các file Excel nhỏ")
    uploaded_files = st.file_uploader(
        "Kéo thả hoặc chọn nhiều file lẻ từ máy tính:",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="Có thể chọn cùng lúc hàng chục file Excel"
    )

with col2:
    st.markdown("### 📗 2. File Total gốc (Tùy chọn)")
    total_file = st.file_uploader(
        "Thả file Total đã có để nối tiếp dữ liệu mới:",
        type=["xlsx", "xls", "csv"],
        help="Dữ liệu mới sẽ được nối tiếp xuống dưới dữ liệu cũ"
    )

st.divider()

# 4. Thao tác xử lý dữ liệu khi bấm nút
if st.button("🚀 BẮT ĐẦU XỬ LÝ & GỘP DỮ LIỆU", type="primary", use_container_width=True):
    if not uploaded_files and total_file is None:
        st.error("⚠️ Vui lòng tải lên ít nhất 1 file Excel lẻ hoặc 1 file Total gốc!")
    else:
        with st.spinner("⏳ Đang phân tích, làm sạch và gộp dữ liệu..."):
            all_dfs = []
            file_count = 0
            
            # Hàm đọc dữ liệu linh hoạt
            def read_excel_file(file, filename):
                dfs = []
                if filename.endswith('.csv'):
                    df = pd.read_csv(file)
                    if add_source_file:
                        df['Nguồn_File'] = filename
                    dfs.append(df)
                else:
                    xl = pd.ExcelFile(file)
                    sheets_to_read = [xl.sheet_names[0]] if sheet_option == "Sheet đầu tiên" else xl.sheet_names
                    for s in sheets_to_read:
                        df = pd.read_excel(file, sheet_name=s)
                        if add_source_file:
                            df['Nguồn_File'] = f"{filename} ({s})" if len(sheets_to_read) > 1 else filename
                        dfs.append(df)
                return pd.concat(dfs, ignore_index=True)

            # Đọc file Total cũ nếu có
            if total_file is not None:
                df_total_exist = read_excel_file(total_file, total_file.name)
                all_dfs.append(df_total_exist)

            # Đọc danh sách các file nhỏ
            if uploaded_files:
                for file in uploaded_files:
                    df_small = read_excel_file(file, file.name)
                    all_dfs.append(df_small)
                    file_count += 1

            # Tiến hành Gộp Dữ Liệu
            final_df = pd.concat(all_dfs, ignore_index=True)
            
            # Áp dụng các cấu hình từ Sidebar
            if drop_na_rows:
                final_df.dropna(how='all', inplace=True)
                
            if remove_dup:
                # Loại trừ cột Nguồn_File & Thời Gian khi xét trùng lặp
                cols_to_check = [c for c in final_df.columns if c not in ['Nguồn_File', 'Thời_Gian_Gộp']]
                final_df.drop_duplicates(subset=cols_to_check, inplace=True)

            if add_timestamp:
                final_df['Thời_Gian_Gộp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            st.balloons() # Hiệu ứng bóng bay chúc mừng
            st.success("🎉 **Hoàn tất xử lý dữ liệu thành công!**")
            
            # Hiển thị Dashboard Thống Kê (KPIs)
            st.markdown("### 📊 Thống Kê Kết Quả")
            m1, m2, m3 = st.columns(3)
            m1.metric(label="Tổng số File đã gộp", value=f"{file_count} File")
            m2.metric(label="Tổng số Dòng dữ liệu", value=f"{len(final_df):,} dòng")
            m3.metric(label="Tổng số Cột", value=f"{len(final_df.columns)} cột")

            st.divider()

            # Xem trước dữ liệu trực quan
            st.markdown("### 👀 Xem Trước Dữ Liệu Tổng Hợp")
            st.dataframe(final_df, use_container_width=True, height=350)

            st.divider()

            # Khu vực Tải Xuất Dữ Liệu Đa Dạng
            st.markdown("### 📥 Tải Kết Quả Về Máy")
            d_col1, d_col2 = st.columns(2)
            
            # Xuất Excel
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            d_col1.download_button(
                label="📗 Tải về file EXCEL (.xlsx)",
                data=buffer_excel.getvalue(),
                file_name=f"Total_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # Xuất CSV
            csv_data = final_df.to_csv(index=False).encode('utf-8-sig')
            d_col2.download_button(
                label="📄 Tải về file CSV (.csv)",
                data=csv_data,
                file_name=f"Total_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
