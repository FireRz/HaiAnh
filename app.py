import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN TRANG WEB
# ==========================================
st.set_page_config(page_title="Hệ thống Điều chuyển Nội bộ", page_icon="📦", layout="wide")
st.title("📦 Hệ thống Tự động Điều chuyển Hàng hóa (Tích hợp SAP)")
st.markdown("""
Ứng dụng hỗ trợ luân chuyển hàng bán chậm từ cửa hàng **Có tồn nhưng không bán** sang cửa hàng **Có bán**.
*Thuật toán tự động ưu tiên rót hàng theo: (1) Sức bán của mã hàng -> (2) Tổng doanh thu của cửa hàng.*
""")
st.divider()

# ==========================================
# 2. HÀM XỬ LÝ LÕI (MATCHING ENGINE)
# ==========================================
def run_allocation_engine(df_stock, df_master):
    # Mapping cột Doanh thu vào bảng Tồn kho dựa trên mã Plant
    df = pd.merge(df_stock, df_master, on='Plant', how='left')
    # Điền 0 cho các Plant không có doanh thu để tránh lỗi logic
    df['Doanh thu'] = df['Doanh thu'].fillna(0)

    transfer_list = []
    dead_stock_list = []

    # Lấy danh sách các Mã 13 có Tồn > 0
    unique_items = df[df['Số lượng tồn'] > 0]['Mã 13'].unique()

    for item in unique_items:
        df_item = df[df['Mã 13'] == item].copy()
        
        # SENDER (Cung): Có tồn, Không bán
        senders = df_item[(df_item['Số lượng tồn'] > 0) & (df_item['Số lượng bán'] == 0)].copy()
        
        # RECEIVER (Cầu): Có bán
        # PRIORITY RULE TÍCH HỢP: Ưu tiên Số lượng bán trước, nếu bằng nhau xét Doanh thu
        receivers = df_item[df_item['Số lượng bán'] > 0].copy()
        receivers = receivers.sort_values(by=['Số lượng bán', 'Doanh thu'], ascending=[False, False])
        
        if receivers.empty and not senders.empty:
            dead_stock_list.extend(senders.to_dict('records'))
            continue
        if senders.empty:
            continue
            
        sender_list = senders.to_dict('records')
        receiver_list = receivers.to_dict('records')
        s_idx, r_idx = 0, 0
        
        # Thuật toán Rót hàng
        while s_idx < len(sender_list) and r_idx < len(receiver_list):
            sender = sender_list[s_idx]
            receiver = receiver_list[r_idx]
            
            qty_available = sender['Số lượng tồn']
            qty_needed = receiver['Số lượng bán']
            
            transfer_qty = min(qty_available, qty_needed)
            
            if transfer_qty > 0:
                transfer_list.append({
                    'Plant Giao': sender['Plant'],
                    'Plant Nhận': receiver['Plant'],
                    'Mã 13': item,
                    'Số lượng': transfer_qty
                })
                sender_list[s_idx]['Số lượng tồn'] -= transfer_qty
                receiver_list[r_idx]['Số lượng bán'] -= transfer_qty
            
            if sender_list[s_idx]['Số lượng tồn'] == 0: s_idx += 1
            if receiver_list[r_idx]['Số lượng bán'] == 0: r_idx += 1

        for s in sender_list[s_idx:]:
            if s['Số lượng tồn'] > 0:
                dead_stock_list.append(s)

    df_transfer = pd.DataFrame(transfer_list)
    df_dead_stock = pd.DataFrame(dead_stock_list)
    
    return df_transfer, df_dead_stock

# ==========================================
# 3. THIẾT KẾ UI: UPLOAD & HIỂN THỊ
# ==========================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. File Dữ liệu Tồn/Bán")
    st.caption("Cột bắt buộc: Plant, Mã 13, Số lượng tồn, Số lượng bán")
    file_stock = st.file_uploader("Tải lên file Excel Tồn kho", type=['xlsx', 'xls'])

with col2:
    st.subheader("2. File Master Data (Ranking)")
    st.caption("Cột bắt buộc: Plant, Doanh thu")
    file_master = st.file_uploader("Tải lên file Master Data", type=['xlsx', 'xls'])

if file_stock and file_master:
    # Đọc dữ liệu khi user upload xong
    df_stock = pd.read_excel(file_stock)
    df_master = pd.read_excel(file_master)
    
    st.success("Tải dữ liệu thành công! Nhấn nút bên dưới để bắt đầu phân tích.")
    
    if st.button("🚀 Chạy Thuật Toán Điều Chuyển", use_container_width=True, type="primary"):
        with st.spinner('Đang tính toán ma trận phân bổ...'):
            df_transfer, df_dead_stock = run_allocation_engine(df_stock, df_master)
            
            # --- HIỂN THỊ KẾT QUẢ TRỰC QUAN ---
            st.divider()
            st.subheader("📊 Xem trước kết quả")
            
            tab1, tab2 = st.tabs(["📝 File SAP_Upload (Thành công)", "⚠️ File Dead_Stock (Cần xử lý)"])
            
            with tab1:
                st.dataframe(df_transfer, use_container_width=True)
                st.metric(label="Tổng số dòng lệnh điều chuyển tạo ra", value=len(df_transfer))
                
            with tab2:
                if not df_dead_stock.empty:
                    df_dead_show = df_dead_stock[['Plant', 'Mã 13', 'Số lượng tồn']].rename(columns={'Số lượng tồn': 'Tồn đọng'})
                    st.dataframe(df_dead_show, use_container_width=True)
                else:
                    st.info("Tuyệt vời! Không có hàng tồn đọng nào không thể điều chuyển.")

            # --- TẠO FILE EXCEL ẢO ĐỂ DOWNLOAD ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_transfer.empty:
                    df_transfer.to_excel(writer, sheet_name='SAP_Upload', index=False)
                if not df_dead_stock.empty:
                    df_dead_show.to_excel(writer, sheet_name='Dead_Stock', index=False)
            
            processed_data = output.getvalue()
            
            st.divider()
            st.download_button(
                label="📥 TẢI XUỐNG FILE EXCEL KẾT QUẢ",
                data=processed_data,
                file_name="Ket_Qua_Dieu_Chuyen_SAP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )