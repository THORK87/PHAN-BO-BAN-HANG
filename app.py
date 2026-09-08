import streamlit as st
import pandas as pd
import numpy as np
import io
import math
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Phân Bổ Bán Hàng Chuẩn Excel", layout="wide")
import hmac
import hashlib
from datetime import datetime, date

# =========================================================================
# HỆ THỐNG BẢN QUYỀN & TẠO KEY TRỰC TIẾP TRÊN WEB
# =========================================================================
SECRET_KEY = "KHUONG_BERUBCO_PRIVATE_2026"
ADMIN_KEY = "Duykhuong@2026"  # Mật khẩu bí mật của riêng bạn

def kiem_tra_license(key_nhap):
    try:
        parts = key_nhap.strip().upper().rsplit("-", 3)
        if len(parts) != 4:
            return False, "Định dạng License Key không đúng!"
        
        client_id, exp_str, s1, s2 = parts
        sig_nhap = s1 + s2
        
        payload = f"{client_id}|{exp_str}"
        sig_chuan = hmac.new(SECRET_KEY.encode(), payload.encode(), hashlib.sha256).hexdigest()[:8].upper()
        
        if sig_nhap != sig_chuan:
            return False, "License Key không hợp lệ hoặc đã bị chỉnh sửa!"
        
        exp_date = datetime.strptime(exp_str, "%Y%m%d")
        if datetime.now() > exp_date:
            return False, f"Bản quyền của [{client_id}] đã hết hạn vào ngày {exp_date.strftime('%d/%m/%Y')}!"
            
        return True, f"Kích hoạt thành công: {client_id} (Hạn dùng: {exp_date.strftime('%d/%m/%Y')})"
    except Exception:
        return False, "Mã kích hoạt không hợp lệ!"

if "is_licensed" not in st.session_state:
    st.session_state.is_licensed = False
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

# GIAO DIỆN KHÓA PHẦN MỀM
if not st.session_state.is_licensed and not st.session_state.is_admin:
    st.sidebar.title("🔐 KÍCH HOẠT BẢN QUYỀN")
    user_key = st.sidebar.text_input("Nhập mã License Key:", type="password")
    
    if st.sidebar.button("Kích hoạt"):
        # Nhập mã Admin -> Mở giao diện tạo key
        if user_key.strip() == ADMIN_KEY:
            st.session_state.is_admin = True
            st.rerun()
        # Khách nhập key -> Kiểm tra bản quyền
        else:
            hop_le, thong_bao = kiem_tra_license(user_key)
            if hop_le:
                st.session_state.is_licensed = True
                st.sidebar.success(thong_bao)
                st.rerun()
            else:
                st.sidebar.error(thong_bao)

    st.warning("⚠️ Vui lòng nhập License Key ở thanh menu bên trái để mở khóa phần mềm.")
    st.stop()

# GIAO DIỆN TẠO KEY (CHỈ HIỆN KHI BẠN NHẬP MÃ ADMIN)
if st.session_state.is_admin:
    st.title("🔑 BẢNG TẠO LICENSE KEY (ADMIN)")
    st.info("Chế độ quản trị viên đang bật. Bạn tạo mã xong thì copy gửi cho khách.")
    
    c_admin1, c_admin2 = st.columns(2)
    with c_admin1:
        ten_khach = st.text_input("Tên khách hàng (viết liền không dấu):", value="CONGTY_ABC")
    with c_admin2:
        ngay_het_han = st.date_input("Hạn sử dụng đến ngày:", value=date(2026, 12, 31))
        
    if st.button("🚀 BẤM ĐỂ TẠO KEY", type="primary"):
        client = ten_khach.strip().upper().replace(" ", "")
        exp_str = ngay_het_han.strftime("%Y%m%d")
        payload = f"{client}|{exp_str}"
        sig = hmac.new(SECRET_KEY.encode(), payload.encode(), hashlib.sha256).hexdigest()[:8].upper()
        
        license_tao = f"{client}-{exp_str}-{sig[:4]}-{sig[4:8]}"
        st.success(f"Mã của khách **{client}** (Hạn: {ngay_het_han.strftime('%d/%m/%Y')}):")
        st.code(license_tao, language="text")
        
    if st.button("Thoát chế độ Admin"):
        st.session_state.is_admin = False
        st.rerun()
    st.stop()

st.title("HỆ THỐNG PHÂN BỔ BÁN HÀNG THEO NGÀY")

# Hàm làm tròn chuẩn Excel (Half Up: >= 0.5 làm tròn lên)
def round_excel(val):
    if val >= 0:
        return math.floor(val + 0.5)
    else:
        return math.ceil(val - 0.5)

# =========================================================================
# 1. NẠP DỮ LIỆU TỪ SHEET T4.2026 (ĐƠN GIÁ LÀM TRÒN SỐ NGUYÊN)
# =========================================================================
st.subheader("1. DANH MỤC HÀNG HÓA TỪ SHEET T4.2026")
file_upload = st.file_uploader("Kéo thả file THORK V6 PHAN BO BAN HANG.xlsx vào đây:", type=["xlsx"])

if file_upload is not None:
    try:
        xls = pd.ExcelFile(file_upload)
        sheet_target = "T4.2026" if "T4.2026" in xls.sheet_names else xls.sheet_names[0]
        df_source = pd.read_excel(xls, sheet_target)
        df_source = df_source.dropna(subset=[df_source.columns[0]]).copy()

        df_init = pd.DataFrame({
            "MÃ VTHH": df_source["Mã VTHH (*)"],
            "TÊN VTHH": df_source["Tên VTHH (*)"],
            "ĐVT": df_source["ĐVT chính"],
            "ĐƠN GIÁ": pd.to_numeric(df_source["Giá bán cố định"], errors="coerce").fillna(0).round(0).astype(int),
            "TỒN ĐẦU": pd.to_numeric(df_source["Tồn gốc (ẩn)"], errors="coerce").fillna(0).astype(int)
        })
        st.success(f"Đã nạp chính xác {len(df_init)} mặt hàng từ sheet '{sheet_target}'.")
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        df_init = pd.DataFrame(columns=["MÃ VTHH", "TÊN VTHH", "ĐVT", "ĐƠN GIÁ", "TỒN ĐẦU"])
else:
    df_init = pd.DataFrame([
        {"MÃ VTHH": "Bánh Khoai Tây xá 4kg", "TÊN VTHH": "Bánh Khoai Tây xá 4kg", "ĐVT": "kg", "ĐƠN GIÁ": 58850, "TỒN ĐẦU": 4},
        {"MÃ VTHH": "VP26_01", "TÊN VTHH": "Nước T-REXX Oishi Túi 180ml (5 bịch/T)", "ĐVT": "Bịch", "ĐƠN GIÁ": 28160, "TỒN ĐẦU": 25},
        {"MÃ VTHH": "VP26_02", "TÊN VTHH": "Bánh snack Oishi 2K (10 bịch/bao)", "ĐVT": "Bịch", "ĐƠN GIÁ": 15248, "TỒN ĐẦU": 280},
        {"MÃ VTHH": "VP26_10", "TÊN VTHH": "Bánh snack Oishi 5K (8 bịch/bao)", "ĐVT": "Bịch", "ĐƠN GIÁ": 38500, "TỒN ĐẦU": 8},
        {"MÃ VTHH": "VP26_13", "TÊN VTHH": "Bánh Chấm 330g PC (12 cái)", "ĐVT": "Cái", "ĐƠN GIÁ": 2383, "TỒN ĐẦU": 120},
        {"MÃ VTHH": "VP26_32", "TÊN VTHH": "Thạch miếng nhỏ Suso LƯỚI 800g (62 cái)", "ĐVT": "Cái", "ĐƠN GIÁ": 550, "TỒN ĐẦU": 372},
        {"MÃ VTHH": "VP26_40", "TÊN VTHH": "Coke 300ml - Coca Chai Nhí 5K", "ĐVT": "Chai", "ĐƠN GIÁ": 4285, "TỒN ĐẦU": 24},
        {"MÃ VTHH": "VP26_41", "TÊN VTHH": "Đậu phộng Oishi 1K (20 dây/T)", "ĐVT": "Dây", "ĐƠN GIÁ": 7637, "TỒN ĐẦU": 20},
        {"MÃ VTHH": "VP26_43", "TÊN VTHH": "Thạch SUSO Cây Dài (24 cái)", "ĐVT": "Cái", "ĐƠN GIÁ": 1375, "TỒN ĐẦU": 432},
        {"MÃ VTHH": "VP26_45", "TÊN VTHH": "Mì Vàng Enaak 30g (24 gói/Hộp)", "ĐVT": "Gói", "ĐƠN GIÁ": 5500, "TỒN ĐẦU": 144},
        {"MÃ VTHH": "VP26_46", "TÊN VTHH": "Nước ÉP THẠCH Suso (48 ly)", "ĐVT": "Ly", "ĐƠN GIÁ": 3942, "TỒN ĐẦU": 480},
        {"MÃ VTHH": "VP26_47", "TÊN VTHH": "Kẹo Cao Su Dinos 500g (69 viên)", "ĐVT": "Viên", "ĐƠN GIÁ": 829, "TỒN ĐẦU": 759},
        {"MÃ VTHH": "VP26_49", "TÊN VTHH": "Bánh Mứt Dâu Xá 4kg PC", "ĐVT": "kg", "ĐƠN GIÁ": 61875, "TỒN ĐẦU": 4},
        {"MÃ VTHH": "VP26_50", "TÊN VTHH": "Nước Túi Oishi (10 túi x 5 bịch / T)", "ĐVT": "Túi", "ĐƠN GIÁ": 2560, "TỒN ĐẦU": 500},
    ])

edited_df = st.data_editor(
    df_init,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "ĐƠN GIÁ": st.column_config.NumberColumn("ĐƠN GIÁ (VNĐ)", format="%d"),
        "TỒN ĐẦU": st.column_config.NumberColumn("TỒN ĐẦU", format="%d"),
    }
)

# =========================================================================
# 2. THIẾT LẬP THÔNG SỐ PHÂN BỔ
# =========================================================================
st.subheader("2. THIẾT LẬP THÔNG SỐ PHÂN BỔ")
c1, c2, c3 = st.columns(3)

with c1:
    so_ngay = st.number_input("SỐ NGÀY BÁN (Nhập số nguyên):", min_value=1, max_value=31, value=5, step=1)
with c2:
    bien_do = st.number_input("ĐỘ LỆCH DOANH THU CÁC NGÀY (mặc định 0.3):", min_value=0.3, max_value=3.0, value=0.3, step=0.1, format="%.3f")
with c3:
    tong_tien_muc_tieu = st.number_input("DOANH THU TỔNG (VNĐ):", min_value=0.0, value=10000000.0, step=500000.0, format="%.0f")
    st.caption(f"👉 Số tiền: **{int(tong_tien_muc_tieu):,} đ**")

btn_run = st.button("🚀 CHẠY PHÂN BỔ ", type="primary")

# =========================================================================
# 3. THUẬT TOÁN TÁI TẠO NGUYÊN BẢN CÔNG THỨC EXCEL
# =========================================================================
if btn_run:
    df_items = edited_df.dropna(subset=["MÃ VTHH"]).copy().reset_index(drop=True)
    df_items["ĐƠN GIÁ"] = pd.to_numeric(df_items["ĐƠN GIÁ"], errors="coerce").fillna(0).round(0).astype(int)
    df_items["TỒN ĐẦU"] = pd.to_numeric(df_items["TỒN ĐẦU"], errors="coerce").fillna(0).round(0).astype(int)
    
    n = len(df_items)
    if n == 0 or tong_tien_muc_tieu <= 0:
        st.warning("Vui lòng kiểm tra lại danh sách hàng và số tiền phân bổ!")
        st.stop()

    so_ngay_int = int(so_ngay)
    danh_sach_ngay = [f"Ngày {d}" for d in range(1, so_ngay_int + 1)]

    ton_arr = df_items["TỒN ĐẦU"].values
    gia_arr = df_items["ĐƠN GIÁ"].values

    # --- Bước 3.1: Tính BS, BT, BU theo đúng công thức Excel ---
    tong_gia_tri_ton = np.sum(ton_arr * gia_arr)
    ti_le_chung = min(1.0, tong_tien_muc_tieu / tong_gia_tri_ton) if tong_gia_tri_ton > 0 else 0

    # BS: ROUNDDOWN(BQ * ti_le, 0)
    bs_arr = np.floor(ton_arr * ti_le_chung)

    # BT: Công thức lũy kế chính xác trong Excel
    bt_arr = np.zeros(n)
    sumprod_bs = np.sum(bs_arr * gia_arr)
    tien_bt_cum = 0.0

    for i in range(n):
        if gia_arr[i] > 0:
            con_lai = tong_tien_muc_tieu - sumprod_bs - tien_bt_cum
            sl_bu_du_kien = np.floor(con_lai / gia_arr[i]) if con_lai > 0 else 0
            sl_bu = min(ton_arr[i] - bs_arr[i], max(0, sl_bu_du_kien))
            bt_arr[i] = sl_bu
            tien_bt_cum += sl_bu * gia_arr[i]

    # BU = BS + BT
    bu_arr = (bs_arr + bt_arr).astype(int)
    df_items["TỔNG SL BÁN"] = bu_arr
    df_items["THÀNH TIỀN"] = bu_arr * gia_arr
    df_items["TỒN CUỐI"] = ton_arr - bu_arr

    # --- Bước 3.2: Tính ma trận số lượng từng ngày ---
    matrix_ngay = np.zeros((n, so_ngay_int), dtype=int)

    for i in range(n):
        bu_val = bu_arr[i]
        if bu_val <= 0:
            continue

        row_excel = i + 9
        bv_val = (row_excel - 5) % 30

        he_so = np.zeros(so_ngay_int)
        for c_idx in range(so_ngay_int):
            k = c_idx + 1
            val_in = k * 12.9898 + tong_tien_muc_tieu * 0.0001
            sin_val = math.sin(val_in)
            hash_sin = (abs(sin_val) * 43758.5453) % 1.0
            song_tuan = math.sin((bv_val + c_idx) * 2.0 * math.pi / 7.0)
            
            val = 1.0 + bien_do * (0.6 * (2.0 * hash_sin - 1.0) + 0.4 * song_tuan)
            he_so[c_idx] = max(0.1, val)

        sum_all = np.sum(he_so)
        if sum_all == 0:
            continue

        cum_he_so = np.cumsum(he_so)
        prev_round = 0
        for d_idx, cum in enumerate(cum_he_so):
            curr_round = round_excel(bu_val * cum / sum_all)
            matrix_ngay[i, d_idx] = curr_round - prev_round
            prev_round = curr_round

    df_daily = pd.DataFrame(matrix_ngay, columns=danh_sach_ngay)

    # ĐƯA CÁC CỘT NGÀY VÀO GIỮA, CÁC CỘT TỔNG HỢP CHO RA SAU CÙNG
    df_ket_qua_chi_tiet = pd.concat([
        df_items[["MÃ VTHH", "TÊN VTHH", "ĐVT", "ĐƠN GIÁ", "TỒN ĐẦU"]],
        df_daily,
        df_items[["TỔNG SL BÁN", "THÀNH TIỀN", "TỒN CUỐI"]]
    ], axis=1)

    # =========================================================================
    # 4. HIỂN THỊ KẾT QUẢ & XUẤT FILE EXCEL
    # =========================================================================
    st.divider()
    st.subheader(f"3. KẾT QUẢ PHÂN BỔ (TỔNG CỘNG {so_ngay_int} NGÀY)")

    tong_tien_thuc_te = int(np.sum(df_items["THÀNH TIỀN"]))
    chenh_lech = tong_tien_thuc_te - int(tong_tien_muc_tieu)

    # 3 thẻ chỉ số tổng quan trên cùng
    m1, m2, m3 = st.columns(3)
    m1.metric("DOANH THU MỤC TIÊU", f"{tong_tien_muc_tieu:,.0f} đ")
    m2.metric("ĐÃ PHÂN BỔ THỰC TẾ", f"{tong_tien_thuc_te:,.0f} đ")
    m3.metric("CHÊNH LỆCH", f"{chenh_lech:,.0f} đ")

    # -------------------------------------------------------------------------
    # BẢNG 1: DOANH THU TỪNG NGÀY (HIỆN NGAY TRÊN ĐẦU, ĐẬP VÀO MẮT KHÔNG CẦN CUỘN)
    # -------------------------------------------------------------------------
    st.markdown("#### 💰 TỔNG CỘNG DOANH THU VÀ CHỈ TIÊU TỪNG NGÀY")
    
    doanh_thu_tung_ngay = np.dot(gia_arr, matrix_ngay)  # Tổng tiền từng ngày
    tong_sl_tung_ngay = np.sum(matrix_ngay, axis=0)     # Tổng sản lượng từng ngày
    so_mon_tung_ngay = np.sum(matrix_ngay > 0, axis=0)  # Số món bán từng ngày

    df_tong_hop_ngay = pd.DataFrame({
        "CHỈ TIÊU": [
            "DOANH THU TỪNG NGÀY (VNĐ)",
            "Tổng số lượng bán (ĐVT)",
            "Số chủng loại mặt hàng bán"
        ]
    })
    for idx, d_col in enumerate(danh_sach_ngay):
        df_tong_hop_ngay[d_col] = [
            f"{doanh_thu_tung_ngay[idx]:,.0f} đ",
            f"{tong_sl_tung_ngay[idx]:,d}",
            f"{so_mon_tung_ngay[idx]} món"
        ]
    df_tong_hop_ngay["TỔNG CỘNG"] = [
        f"{tong_tien_thuc_te:,.0f} đ",
        f"{np.sum(bu_arr):,d}",
        f"{np.sum(bu_arr > 0)} món"
    ]

    st.dataframe(df_tong_hop_ngay, use_container_width=True, hide_index=True)

    # -------------------------------------------------------------------------
    # BẢNG 2: CHI TIẾT SẢN LƯỢNG BÁN CÁC NGÀY (CÁC CỘT TỔNG HỢP Ở SAU CÙNG)
    # -------------------------------------------------------------------------
    st.markdown("#### 📋 CHI TIẾT SẢN LƯỢNG BÁN THEO MẶT HÀNG")
    st.dataframe(
    df_ket_qua_chi_tiet,
    use_container_width=True,
    hide_index=True,
    column_config={
        "ĐƠN GIÁ": st.column_config.NumberColumn("ĐƠN GIÁ", format="%,d"),
        "THÀNH TIỀN": st.column_config.NumberColumn("THÀNH TIỀN", format="%,d")
    }
)
    # Xuất file Excel chuẩn hóa
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_ket_qua_chi_tiet.to_excel(writer, sheet_name="SL_BAN_THEO_NGAY", index=False)

        ws = writer.sheets["SL_BAN_THEO_NGAY"]
        
        last_row = ws.max_row
        row_sl_idx = last_row + 1
        row_tien_idx = last_row + 2

        # Dòng 1: Tổng số lượng
        ws.cell(row_sl_idx, 1, "TỔNG CỘNG")
        ws.cell(row_sl_idx, 2, "Tổng số lượng bán")
        ws.cell(row_sl_idx, 5, int(np.sum(ton_arr)))
        for idx in range(so_ngay_int):
            ws.cell(row_sl_idx, 6 + idx, int(tong_sl_tung_ngay[idx]))
        ws.cell(row_sl_idx, 6 + so_ngay_int, int(np.sum(bu_arr)))
        ws.cell(row_sl_idx, 7 + so_ngay_int, int(tong_tien_thuc_te))
        ws.cell(row_sl_idx, 8 + so_ngay_int, int(np.sum(ton_arr - bu_arr)))

        # Dòng 2: Tổng doanh thu từng ngày
        ws.cell(row_tien_idx, 1, "DOANH THU")
        ws.cell(row_tien_idx, 2, "Tổng tiền bán ngày")
        for idx in range(so_ngay_int):
            ws.cell(row_tien_idx, 6 + idx, int(doanh_thu_tung_ngay[idx]))
        ws.cell(row_tien_idx, 7 + so_ngay_int, int(tong_tien_thuc_te))

        # Định dạng Header và Dòng Tổng
        header_font = Font(name="Arial", size=11, bold=True, color="000000")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        sum_font = Font(name="Arial", size=11, bold=True, color="002060")
        sum_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = header_alignment
            cell.fill = header_fill
        ws.row_dimensions[1].height = 28

        for r in [row_sl_idx, row_tien_idx]:
            for cell in ws[r]:
                cell.font = sum_font
                cell.fill = sum_fill
            ws.row_dimensions[r].height = 24

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    st.download_button(
        label="📥 TẢI FILE EXCEL KẾT QUẢ (KHỚP CHUẨN FILE GỐC)",
        data=output.getvalue(),
        file_name=f"PHAN_BO_{so_ngay_int}_NGAY.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
