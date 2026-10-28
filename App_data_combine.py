import streamlit as st
import pandas as pd
import io

# Cấu hình giao diện Web
st.set_page_config(page_title="Tool Tổng Hợp File Excel", page_icon="📊", layout="centered")

st.title("📊 Tool Tổng Hợp File Excel Auto")
st.markdown("Tải lên các file Excel nhỏ và file Total (nếu có) để tự động gộp dữ liệu.")

st.divider()

# 1. Khu vực Kéo-Thả các file nhỏ
st.subheader("1. Chọn hoặc Kéo-Thả các file Excel nhỏ")
uploaded_files = st.file_uploader(
    "Nhấp 'Browse files' để chọn file từ This PC hoặc Kéo-Thả file vào đây:", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

# 2. Khu vực Kéo-Thả file Total (Không bắt buộc)
st.subheader("2. File Excel Total hiện có (Tùy chọn)")
total_file = st.file_uploader(
    "Nếu đã có file Total từ trước, hãy thả vào đây để nối tiếp dữ liệu mới:", 
    type=["xlsx", "xls"]
)

st.divider()

# 3. Nút Bắt đầu gộp
if st.button("🚀 BẮT ĐẦU GỘP DỮ LIỆU", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Vui lòng tải lên ít nhất 1 file Excel nhỏ!")
    else:
        with st.spinner("Đang xử lý và hợp nhất dữ liệu..."):
            all_dfs = []
            
            # Đọc dữ liệu từ file Total cũ (nếu người dùng tải lên)
            if total_file is not None:
                df_total_exist = pd.read_excel(total_file)
                all_dfs.append(df_total_exist)
                
            # Đọc từng file nhỏ
            for file in uploaded_files:
                df = pd.read_excel(file)
                df['Nguồn_File'] = file.name  # Thêm cột ghi nhận tên file gốc
                all_dfs.append(df)
                
            # Gộp tất cả
            final_df = pd.concat(all_dfs, ignore_index=True)
            
            st.success(f"🎉 Đã gộp thành công {len(uploaded_files)} file lẻ! Tổng cộng có {len(final_df)} dòng dữ liệu.")
            
            # Xem trước dữ liệu
            with st.expander("👀 Bấm để xem trước dữ liệu vừa gộp"):
                st.dataframe(final_df, use_container_width=True)
            
            # Đưa file ra bộ nhớ tạm để tải về
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            # Nút Tải file kết quả
            st.download_button(
                label="📥 TẢI FILE EXCEL TOTAL MỚI VỀ MÁY",
                data=buffer.getvalue(),
                file_name="Excel_Total_CapNhat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True
            )