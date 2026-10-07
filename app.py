import base64
import streamlit.components.v1 as components
from datetime import date, datetime
import io
import json
import math
import secrets
import numpy as np
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Phân Bổ Bán Hàng Chuẩn Excel", layout="wide")

# =========================================================================
# CẤU HÌNH GITHUB DATABASE BẢN QUYỀN & NHẬT KÝ HOẠT ĐỘNG
# =========================================================================
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "THORK87/PHAN-BO-BAN-HANG")
GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")
LICENSES_FILE = "licenses.json"
ACTIVITY_LOG_FILE = "activity_log.json"
ADMIN_KEY = "Duykhuong@2026"  # Mật khẩu quản trị của bạn

# Lấy token từ Secrets của Streamlit Cloud
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")

HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
}
LICENSES_API_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO}/contents/{LICENSES_FILE}"
)
ACTIVITY_LOG_API_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ACTIVITY_LOG_FILE}"
)


# =========================================================================
# HÀM QUẢN LÝ BẢN QUYỀN
# =========================================================================
def get_remote_licenses():
    try:
        res = requests.get(
            f"{LICENSES_API_URL}?ref={GITHUB_BRANCH}", headers=HEADERS
        )
        if res.status_code == 200:
            data = res.json()
            content = base64.b64decode(data["content"]).decode("utf-8")
            return json.loads(content), data["sha"]
    except Exception:
        pass
    return {}, None


def update_remote_licenses(new_data, sha=None):
    try:
        content_str = json.dumps(new_data, ensure_ascii=False, indent=4)
        content_b64 = base64.b64encode(content_str.encode("utf-8")).decode(
            "utf-8"
        )
        payload = {
            "message": f"Update licenses.json via Admin - {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            "content": content_b64,
            "branch": GITHUB_BRANCH,
        }
        if sha:
            payload["sha"] = sha
        res = requests.put(LICENSES_API_URL, headers=HEADERS, json=payload)
        return res.status_code in [200, 201]
    except Exception:
        return False


# =========================================================================
# HÀM QUẢN LÝ NHẬT KÝ HOẠT ĐỘNG
# =========================================================================
def get_remote_activity_log():
    try:
        res = requests.get(
            f"{ACTIVITY_LOG_API_URL}?ref={GITHUB_BRANCH}", headers=HEADERS
        )
        if res.status_code == 200:
            data = res.json()
            content = base64.b64decode(data["content"]).decode("utf-8")
            return json.loads(content), data["sha"]
    except Exception:
        pass
    return [], None


def update_remote_activity_log(new_data, sha=None):
    try:
        content_str = json.dumps(new_data, ensure_ascii=False, indent=2)
        content_b64 = base64.b64encode(content_str.encode("utf-8")).decode(
            "utf-8"
        )
        payload = {
            "message": f"Update activity log - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            "content": content_b64,
            "branch": GITHUB_BRANCH,
        }
        if sha:
            payload["sha"] = sha
        res = requests.put(ACTIVITY_LOG_API_URL, headers=HEADERS, json=payload)
        return res.status_code in [200, 201]
    except Exception:
        return False


def log_activity(activity_type, details):
    activity_log, log_sha = get_remote_activity_log()
    log_entry = {
        "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "type": activity_type,
        "details": details,
    }
    activity_log.append(log_entry)
    if len(activity_log) > 500:
        activity_log = activity_log[-500:]

    if update_remote_activity_log(activity_log, log_sha):
        return True
    return False


# Đọc dữ liệu từ GitHub
db_licenses, file_sha = get_remote_licenses()

# =========================================================================
# 1. KIỂM TRA BẢN QUYỀN KHÁCH HÀNG (GIỮ NGUYÊN KEY)
# =========================================================================
if "is_licensed" not in st.session_state:
    st.session_state.is_licensed = False
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "current_client_name" not in st.session_state:
    st.session_state.current_client_name = ""

if not st.session_state.is_licensed and not st.session_state.is_admin:
    st.sidebar.title("🔐 KÍCH HOẠT BẢN QUYỀN")
    user_key = st.sidebar.text_input("Nhập mã License Key:", type="password")

    if st.sidebar.button("Kích hoạt"):
        key_input = user_key.strip()
        if key_input == ADMIN_KEY:
            st.session_state.is_admin = True
            log_activity(
                "ADMIN_LOGIN",
                {"timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
            )
            st.rerun()

        client_key = key_input.upper()
        if client_key in db_licenses:
            client_info = db_licenses[client_key]
            client_name = client_info.get("client_name", "Unknown")

            if client_info.get("status") == "blocked":
                st.sidebar.error(
                    "❌ Mã bản quyền này đã bị thu hồi hoặc khóa truy cập!"
                )
                log_activity(
                    "LOGIN_FAILED",
                    {
                        "reason": "KEY_BLOCKED",
                        "key": client_key,
                        "client_name": client_name,
                        "timestamp": datetime.now().strftime(
                            "%d/%m/%Y %H:%M:%S"
                        ),
                    },
                )
            else:
                exp_date = datetime.strptime(
                    client_info["expiry"], "%Y-%m-%d"
                ).date()
                if datetime.now().date() > exp_date:
                    st.sidebar.error(
                        f"❌ Bản quyền đã hết hạn vào ngày {exp_date.strftime('%d/%m/%Y')}!"
                    )
                    log_activity(
                        "LOGIN_FAILED",
                        {
                            "reason": "KEY_EXPIRED",
                            "key": client_key,
                            "client_name": client_name,
                            "expiry_date": exp_date.strftime("%d/%m/%Y"),
                            "timestamp": datetime.now().strftime(
                                "%d/%m/%Y %H:%M:%S"
                            ),
                        },
                    )
                else:
                    st.session_state.is_licensed = True
                    st.session_state.current_client_name = client_name
                    st.sidebar.success(
                        f"Hợp lệ! Hạn sử dụng: {exp_date.strftime('%d/%m/%Y')}"
                    )
                    log_activity(
                        "LOGIN_SUCCESS",
                        {
                            "key": client_key,
                            "client_name": client_name,
                            "expiry_date": exp_date.strftime("%d/%m/%Y"),
                            "timestamp": datetime.now().strftime(
                                "%d/%m/%Y %H:%M:%S"
                            ),
                        },
                    )
                    st.rerun()
        else:
            st.sidebar.error("Mã kích hoạt không tồn tại trên hệ thống!")
            log_activity(
                "LOGIN_FAILED",
                {
                    "reason": "KEY_NOT_FOUND",
                    "key": client_key,
                    "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                },
            )

    st.warning(
        "⚠️ Vui lòng nhập License Key ở thanh menu bên trái để mở khóa phần mềm."
    )
    st.stop()

# =========================================================================
# 2. BẢNG ĐIỀU KHIỂN QUẢN TRỊ ADMIN (TẠO - GIA HẠN - HỦY)
# =========================================================================
if st.session_state.is_admin:
    st.title("🔑 BẢNG QUẢN TRỊ BẢN QUYỀN (ĐỒNG BỘ GITHUB)")
    st.info(
        "Dữ liệu được lưu vĩnh viễn vào file licenses.json trên GitHub Repository."
    )

    tab_tao, tab_ql, tab_log = st.tabs(
        [
            "➕ Tạo Key Mới (Không ngày tháng)",
            "📋 Danh Sách & Gia Hạn / Hủy",
            "📊 Nhật Ký Hoạt Động",
        ]
    )

    # TAB 1: TẠO KEY MỚI
    with tab_tao:
        c1, c2 = st.columns(2)
        with c1:
            ten_khach = st.text_input(
                "Ghi chú tên khách hàng (quản lý nội bộ):", value="CONGTY_ABC"
            )
        with c2:
            ngay_het = st.date_input(
                "Hạn sử dụng ban đầu:", value=date(2027, 1, 1)
            )

        if st.button("🚀 Tạo License Key Quốc Tế", type="primary"):
            c_name = ten_khach.strip() if ten_khach.strip() else "KHACH_HANG"
            parts = [secrets.token_hex(2).upper() for _ in range(4)]
            generated_key = "-".join(parts)

            db_licenses[generated_key] = {
                "client_name": c_name,
                "expiry": ngay_het.strftime("%Y-%m-%d"),
                "status": "active",
            }
            if update_remote_licenses(db_licenses, file_sha):
                log_activity(
                    "KEY_CREATED",
                    {
                        "key": generated_key,
                        "client_name": c_name,
                        "expiry_date": ngay_het.strftime("%d/%m/%Y"),
                        "timestamp": datetime.now().strftime(
                            "%d/%m/%Y %H:%M:%S"
                        ),
                    },
                )
                st.success(f"Đã tạo Key thành công cho {c_name}!")
                st.code(generated_key, language="text")
                st.info(
                    "Gửi mã trên cho khách hàng. Khách sẽ dùng cố định mã này vĩnh viễn."
                )
                st.rerun()
            else:
                st.error(
                    "Lỗi khi lưu lên GitHub. Vui lòng kiểm tra lại GITHUB_TOKEN trong Secrets."
                )

    # TAB 2: QUẢN LÝ
    with tab_ql:
        if not db_licenses:
            st.info("Chưa có mã bản quyền nào trên GitHub.")
        else:
            for k, v in list(db_licenses.items()):
                ten_hien_thi = v.get("client_name", k)
                with st.expander(
                    f"Khách hàng: {ten_hien_thi} | Key: {k}", expanded=True
                ):
                    col_info, col_han, col_action = st.columns([2.5, 2, 2])
                    with col_info:
                        st.write(f"**Khách hàng:** `{ten_hien_thi}`")
                        st.code(k, language="text")
                        st.write(
                            f"Trạng thái: {'🟢 Hoạt động' if v['status'] == 'active' else '🔴 ĐÃ KHÓA'}"
                        )
                        st.write(
                            f"Hạn: **{datetime.strptime(v['expiry'], '%Y-%m-%d').strftime('%d/%m/%Y')}**"
                        )

                    with col_han:
                        cur_date = datetime.strptime(
                            v["expiry"], "%Y-%m-%d"
                        ).date()
                        new_date = st.date_input(
                            "Chọn hạn mới:", value=cur_date, key=f"date_{k}"
                        )
                        if st.button("Cập nhật hạn", key=f"btn_date_{k}"):
                            db_licenses[k]["expiry"] = new_date.strftime(
                                "%Y-%m-%d"
                            )
                            if update_remote_licenses(db_licenses, file_sha):
                                log_activity(
                                    "KEY_EXTENDED",
                                    {
                                        "key": k,
                                        "client_name": ten_hien_thi,
                                        "new_expiry_date": new_date.strftime(
                                            "%d/%m/%Y"
                                        ),
                                        "timestamp": datetime.now().strftime(
                                            "%d/%m/%Y %H:%M:%S"
                                        ),
                                    },
                                )
                                st.success("Đã gia hạn thành công!")
                                st.rerun()
                            else:
                                st.error("Lỗi khi cập nhật lên GitHub!")

                    with col_action:
                        st.write("Thao tác bản quyền:")
                        if v["status"] == "active":
                            if st.button(
                                "🚫 Khóa Key",
                                key=f"block_{k}",
                                type="secondary",
                            ):
                                db_licenses[k]["status"] = "blocked"
                                update_remote_licenses(db_licenses, file_sha)
                                log_activity(
                                    "KEY_BLOCKED",
                                    {
                                        "key": k,
                                        "client_name": ten_hien_thi,
                                        "timestamp": datetime.now().strftime(
                                            "%d/%m/%Y %H:%M:%S"
                                        ),
                                    },
                                )
                                st.warning("Đã khóa bản quyền!")
                                st.rerun()
                        else:
                            if st.button("✅ Mở khóa", key=f"unblock_{k}"):
                                db_licenses[k]["status"] = "active"
                                update_remote_licenses(db_licenses, file_sha)
                                log_activity(
                                    "KEY_UNBLOCKED",
                                    {
                                        "key": k,
                                        "client_name": ten_hien_thi,
                                        "timestamp": datetime.now().strftime(
                                            "%d/%m/%Y %H:%M:%S"
                                        ),
                                    },
                                )
                                st.success("Đã mở khóa lại!")
                                st.rerun()

                        st.write("")
                        if st.button(
                            "🗑️ Xóa vĩnh viễn", key=f"del_{k}", type="primary"
                        ):
                            del db_licenses[k]
                            if update_remote_licenses(db_licenses, file_sha):
                                log_activity(
                                    "KEY_DELETED",
                                    {
                                        "key": k,
                                        "client_name": ten_hien_thi,
                                        "timestamp": datetime.now().strftime(
                                            "%d/%m/%Y %H:%M:%S"
                                        ),
                                    },
                                )
                                st.success(f"Đã xóa vĩnh viễn key `{k}`!")
                                st.rerun()
                            else:
                                st.error("Lỗi khi xóa key trên GitHub!")

    # TAB 3: XEM NHẬT KÝ
    with tab_log:
        st.subheader("📊 Nhật Ký Hoạt Động Đăng Nhập")
        activity_log, _ = get_remote_activity_log()
        if not activity_log:
            st.info("Chưa có hoạt động nào được ghi nhận.")
        else:
            activity_log_reversed = list(reversed(activity_log))
            filter_type = st.selectbox(
                "Lọc theo loại hoạt động:",
                [
                    "Tất cả",
                    "LOGIN_SUCCESS",
                    "LOGIN_FAILED",
                    "ADMIN_LOGIN",
                    "KEY_CREATED",
                    "KEY_EXTENDED",
                    "KEY_BLOCKED",
                    "KEY_UNBLOCKED",
                    "KEY_DELETED",
                ],
            )
            if filter_type != "Tất cả":
                activity_log_reversed = [
                    log
                    for log in activity_log_reversed
                    if log.get("type") == filter_type
                ]

            df_log = pd.DataFrame(
                [
                    {
                        "⏰ Thời gian": log.get("timestamp", "N/A"),
                        "📌 Loại": log.get("type", "N/A"),
                        "👤 Khách hàng": log.get("details", {}).get(
                            "client_name", "N/A"
                        ),
                        "🔑 Key": log.get("details", {}).get("key", "N/A"),
                        "📝 Chi tiết": str(log.get("details", {})),
                    }
                    for log in activity_log_reversed[:100]
                ]
            )
            st.dataframe(df_log, use_container_width=True, hide_index=True)
            csv_log = df_log.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                label="📥 Tải xuống CSV",
                data=csv_log,
                file_name=f"activity_log_{datetime.now().strftime('%d%m%Y_%H%M%S')}.csv",
                mime="text/csv",
            )

# =========================================================================
# 1. NẠP DỮ LIỆU TỪ SHEET T4.2026
# =========================================================================
st.subheader("1. DANH MỤC HÀNG HÓA")

df_mau = pd.DataFrame(
    {
        "Mã VTHH (*)": ["SP01", "SP02", "SP03"],
        "Tên VTHH (*)": [
            "Sản phẩm mẫu A",
            "Sản phẩm mẫu B",
            "Sản phẩm mẫu C",
        ],
        "ĐVT chính": ["Cái", "Hộp", "Gói"],
        "Giá bán cố định": [50000, 120000, 25000],
        "Tồn gốc (ẩn)": [100, 50, 200],
    }
)
buffer_mau = io.BytesIO()
with pd.ExcelWriter(buffer_mau, engine="openpyxl") as writer:
    df_mau.to_excel(writer, sheet_name="T4.2026", index=False)

col_up, col_btn = st.columns([3, 1])
with col_up:
    file_upload = st.file_uploader(
        "Kéo thả file số liệu vào đây:", type=["xlsx"]
    )
with col_btn:
    st.write("")
    st.write("")
    st.download_button(
        label="📥 Tải file mẫu Excel",
        data=buffer_mau.getvalue(),
        file_name="MAU_PHAN_BO_BAN_HANG.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Tải file mẫu Excel có sẵn các cột chuẩn để nhập liệu",
    )

if file_upload is not None:
    try:
        xls = pd.ExcelFile(file_upload)
        sheet_target = (
            "T4.2026" if "T4.2026" in xls.sheet_names else xls.sheet_names[0]
        )
        df_source = pd.read_excel(xls, sheet_target)
        df_source = df_source.dropna(subset=[df_source.columns[0]]).copy()

        df_init = pd.DataFrame(
            {
                "MÃ VTHH": df_source["Mã VTHH (*)"],
                "TÊN VTHH": df_source["Tên VTHH (*)"],
                "ĐVT": df_source["ĐVT chính"],
                "ĐƠN GIÁ": pd.to_numeric(
                    df_source["Giá bán cố định"], errors="coerce"
                )
                .fillna(0)
                .round(0)
                .astype(int),
                "TỒN ĐẦU": pd.to_numeric(
                    df_source["Tồn gốc (ẩn)"], errors="coerce"
                )
                .fillna(0)
                .round(0)
                .astype(int),
            }
        )
        st.success(
            f"Đã nạp chính xác {len(df_init)} mặt hàng từ sheet '{sheet_target}'."
        )
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        df_init = pd.DataFrame(
            columns=["MÃ VTHH", "TÊN VTHH", "ĐVT", "ĐƠN GIÁ", "TỒN ĐẦU"]
        )
else:
    df_init = pd.DataFrame(
        [
            {
                "MÃ VTHH": "Bánh Khoai Tây xá 4kg",
                "TÊN VTHH": "Bánh Khoai Tây xá 4kg",
                "ĐVT": "kg",
                "ĐƠN GIÁ": 58850,
                "TỒN ĐẦU": 4,
            },
            {
                "MÃ VTHH": "VP26_01",
                "TÊN VTHH": "Nước T-REXX Oishi Túi 180ml (5 bịch/T)",
                "ĐVT": "Bịch",
                "ĐƠN GIÁ": 28160,
                "TỒN ĐẦU": 25,
            },
            {
                "MÃ VTHH": "VP26_02",
                "TÊN VTHH": "Bánh snack Oishi 2K (10 bịch/bao)",
                "ĐVT": "Bịch",
                "ĐƠN GIÁ": 15248,
                "TỒN ĐẦU": 280,
            },
            {
                "MÃ VTHH": "VP26_10",
                "TÊN VTHH": "Bánh snack Oishi 5K (8 bịch/bao)",
                "ĐVT": "Bịch",
                "ĐƠN GIÁ": 38500,
                "TỒN ĐẦU": 8,
            },
            {
                "MÃ VTHH": "VP26_13",
                "TÊN VTHH": "Bánh Chấm 330g PC (12 cái)",
                "ĐVT": "Cái",
                "ĐƠN GIÁ": 2383,
                "TỒN ĐẦU": 120,
            },
            {
                "MÃ VTHH": "VP26_32",
                "TÊN VTHH": "Thạch miếng nhỏ Suso LƯỚI 800g (62 cái)",
                "ĐVT": "Cái",
                "ĐƠN GIÁ": 550,
                "TỒN ĐẦU": 372,
            },
            {
                "MÃ VTHH": "VP26_40",
                "TÊN VTHH": "Coke 300ml - Coca Chai Nhí 5K",
                "ĐVT": "Chai",
                "ĐƠN GIÁ": 4285,
                "TỒN ĐẦU": 24,
            },
            {
                "MÃ VTHH": "VP26_41",
                "TÊN VTHH": "Đậu phộng Oishi 1K (20 dây/T)",
                "ĐVT": "Dây",
                "ĐƠN GIÁ": 7637,
                "TỒN ĐẦU": 20,
            },
            {
                "MÃ VTHH": "VP26_43",
                "TÊN VTHH": "Thạch SUSO Cây Dài (24 cái)",
                "ĐVT": "Cái",
                "ĐƠN GIÁ": 1375,
                "TỒN ĐẦU": 432,
            },
            {
                "MÃ VTHH": "VP26_45",
                "TÊN VTHH": "Mì Vàng Enaak 30g (24 gói/Hộp)",
                "ĐVT": "Gói",
                "ĐƠN GIÁ": 5500,
                "TỒN ĐẦU": 144,
            },
            {
                "MÃ VTHH": "VP26_46",
                "TÊN VTHH": "Nước ÉP THẠCH Suso (48 ly)",
                "ĐVT": "Ly",
                "ĐƠN GIÁ": 3942,
                "TỒN ĐẦU": 480,
            },
            {
                "MÃ VTHH": "VP26_47",
                "TÊN VTHH": "Kẹo Cao Su Dinos 500g (69 viên)",
                "ĐVT": "Viên",
                "ĐƠN GIÁ": 829,
                "TỒN ĐẦU": 759,
            },
            {
                "MÃ VTHH": "VP26_49",
                "TÊN VTHH": "Bánh Mứt Dâu Xá 4kg PC",
                "ĐVT": "kg",
                "ĐƠN GIÁ": 61875,
                "TỒN ĐẦU": 4,
            },
            {
                "MÃ VTHH": "VP26_50",
                "TÊN VTHH": "Nước Túi Oishi (10 túi x 5 bịch / T)",
                "ĐVT": "Túi",
                "ĐƠN GIÁ": 2560,
                "TỒN ĐẦU": 500,
            },
        ]
    )

# Cập nhật lại session_state khi người dùng upload file Excel mới
if file_upload is not None:
    if "last_uploaded_file" not in st.session_state or st.session_state.last_uploaded_file != file_upload.name:
        st.session_state.last_uploaded_file = file_upload.name
        st.session_state.df_data = df_init.copy()

if "df_data" not in st.session_state:
    st.session_state.df_data = df_init.copy()

# Định thứ tự cột: đưa TỒN CUỐI vào ngay sau TỒN ĐẦU nếu đã chạy phân bổ
cot_hien_thi = ["MÃ VTHH", "TÊN VTHH", "ĐVT", "ĐƠN GIÁ", "TỒN ĐẦU"]
if "TỒN CUỐI" in st.session_state.df_data.columns:
    cot_hien_thi.append("TỒN CUỐI")

col_config = {
    "ĐƠN GIÁ": st.column_config.NumberColumn("ĐƠN GIÁ (VNĐ)", format="%d"),
    "TỒN ĐẦU": st.column_config.NumberColumn("TỒN ĐẦU", format="%d"),
}
if "TỒN CUỐI" in st.session_state.df_data.columns:
    col_config["TỒN CUỐI"] = st.column_config.NumberColumn(
        "TỒN CUỐI (Sau phân bổ)", format="%d", disabled=True
    )

edited_df = st.data_editor(
    st.session_state.df_data[cot_hien_thi],
    num_rows="dynamic",
    use_container_width=True,
    column_config=col_config,
    key="data_editor_main",
)

# =========================================================================
# 2. THIẾT LẬP THÔNG SỐ PHÂN BỔ (BỔ SUNG CHỌN NGÀY BẮT ĐẦU BÁN)
# =========================================================================
st.subheader("2. THIẾT LẬP THÔNG SỐ PHÂN BỔ")
c_ngay_bd, c1, c2, c3 = st.columns(4)

with c_ngay_bd:
    ngay_bat_dau = st.date_input("NGÀY BẮT ĐẦU BÁN:", value=date.today())
with c1:
    so_ngay = st.number_input(
        "SỐ NGÀY BÁN (Nhập số nguyên):",
        min_value=1,
        max_value=31,
        value=5,
        step=1,
    )
with c2:
    bien_do = st.number_input(
        "ĐỘ LỆCH DOANH THU CÁC NGÀY (0.0 = san phẳng tuyệt đối):",
        min_value=0.0,
        max_value=3.0,
        value=0.0,
        step=0.05,
        format="%.2f",
    )
with c3:
    tong_tien_muc_tieu = st.number_input(
        "DOANH THU TỔNG (VNĐ):",
        min_value=0.0,
        value=10000000.0,
        step=500000.0,
        format="%.0f",
    )
    st.caption(f"👉 Số tiền: **{int(tong_tien_muc_tieu):,} đ**")

btn_run = st.button("🚀 CHẠY PHÂN BỔ ", type="primary")

# =========================================================================
# 3. THUẬT TOÁN TỐI ƯU 2 CHIỀU (KNAPSACK GREEDY)
# =========================================================================
if btn_run:
    df_items = edited_df.dropna(subset=["MÃ VTHH"]).copy().reset_index(drop=True)
    df_items["ĐƠN GIÁ"] = (
        pd.to_numeric(df_items["ĐƠN GIÁ"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )
    df_items["TỒN ĐẦU"] = (
        pd.to_numeric(df_items["TỒN ĐẦU"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )

    n = len(df_items)
    if n == 0 or tong_tien_muc_tieu <= 0:
        st.warning("Vui lòng kiểm tra lại danh sách hàng và số tiền phân bổ!")
        st.stop()

    so_ngay_int = int(so_ngay)
    danh_sach_ngay = [f"Ngày {d}" for d in range(1, so_ngay_int + 1)]

    ton_arr = df_items["TỒN ĐẦU"].values
    gia_arr = df_items["ĐƠN GIÁ"].values

    # --- TÍNH BU: LẤY ĐA DẠNG HÀNG HÓA KHI SỐ TIỀN NHỎ ---
    tong_gia_tri_ton = np.sum(ton_arr * gia_arr)
    bu_arr = np.zeros(n, dtype=int)
    tien_hien_tai = 0.0

    idx_sorted_by_price = np.argsort(gia_arr)
    for i in idx_sorted_by_price:
        if gia_arr[i] > 0 and ton_arr[i] > 0:
            if tien_hien_tai + gia_arr[i] <= tong_tien_muc_tieu:
                bu_arr[i] = 1
                tien_hien_tai += gia_arr[i]

    tien_con_lai = tong_tien_muc_tieu - tien_hien_tai
    ton_con_lai = ton_arr - bu_arr
    tong_ton_con_lai_vnd = np.sum(ton_con_lai * gia_arr)

    if tien_con_lai > 0 and tong_ton_con_lai_vnd > 0:
        ti_le_phu = min(1.0, tien_con_lai / tong_ton_con_lai_vnd)
        sl_them = np.floor(ton_con_lai * ti_le_phu).astype(int)
        sl_them = np.minimum(sl_them, ton_con_lai)
        bu_arr += sl_them
        tien_hien_tai = np.sum(bu_arr * gia_arr)

    tien_con_lai = tong_tien_muc_tieu - tien_hien_tai
    if tien_con_lai > 0:
        for i in range(n):
            if gia_arr[i] > 0 and bu_arr[i] < ton_arr[i]:
                sl_bu_max = ton_arr[i] - bu_arr[i]
                sl_can_bu = int(tien_con_lai // gia_arr[i])
                sl_thuc_bu = min(sl_bu_max, sl_can_bu)
                if sl_thuc_bu > 0:
                    bu_arr[i] += sl_thuc_bu
                    tien_con_lai -= sl_thuc_bu * gia_arr[i]
                    if tien_con_lai < np.min(gia_arr[gia_arr > 0]):
                        break

    df_items["TỔNG SL BÁN"] = bu_arr
    df_items["THÀNH TIỀN"] = bu_arr * gia_arr
    df_items["TỒN CUỐI"] = ton_arr - bu_arr

    # Cập nhật chuẩn TỒN CUỐI vào bảng nhập liệu theo MÃ VTHH
    if "df_data" in st.session_state:
        map_ton_cuoi = dict(zip(df_items["MÃ VTHH"], df_items["TỒN CUỐI"]))
        st.session_state.df_data["TỒN CUỐI"] = st.session_state.df_data["MÃ VTHH"].map(map_ton_cuoi).fillna(st.session_state.df_data["TỒN ĐẦU"]).astype(int)
        ti_le_phu = min(1.0, tien_con_lai / tong_ton_con_lai_vnd)
        sl_them = np.floor(ton_con_lai * ti_le_phu).astype(int)
        sl_them = np.minimum(sl_them, ton_con_lai)
        bu_arr += sl_them
        tien_hien_tai = np.sum(bu_arr * gia_arr)

    tien_con_lai = tong_tien_muc_tieu - tien_hien_tai
    if tien_con_lai > 0:
        for i in range(n):
            if gia_arr[i] > 0 and bu_arr[i] < ton_arr[i]:
                sl_bu_max = ton_arr[i] - bu_arr[i]
                sl_can_bu = int(tien_con_lai // gia_arr[i])
                sl_thuc_bu = min(sl_bu_max, sl_can_bu)
                if sl_thuc_bu > 0:
                    bu_arr[i] += sl_thuc_bu
                    tien_con_lai -= sl_thuc_bu * gia_arr[i]
                    if tien_con_lai < np.min(gia_arr[gia_arr > 0]):
                        break

    df_items["TỔNG SL BÁN"] = bu_arr
    df_items["THÀNH TIỀN"] = bu_arr * gia_arr
    df_items["TỒN CUỐI"] = ton_arr - bu_arr

    # 1. Xác định doanh thu mục tiêu từng ngày (theo bien_do)
    avg_day = tong_tien_muc_tieu / so_ngay_int
    day_targets = np.zeros(so_ngay_int, dtype=float)
    for d in range(so_ngay_int):
        wave = (
            np.sin((d + 1) * 2 * np.pi / so_ngay_int) if so_ngay_int > 1 else 0
        )
        day_targets[d] = avg_day * (1.0 + bien_do * wave)

    day_targets = day_targets * (tong_tien_muc_tieu / np.sum(day_targets))
    day_ratios = day_targets / np.sum(day_targets)

    # 2. Phân bổ ma trận đồng đều đa món
    matrix_ngay = np.zeros((n, so_ngay_int), dtype=int)

    for i in range(n):
        total_qty = bu_arr[i]
        if total_qty == 0:
            continue

        if total_qty <= 3:
            indices = [
                (
                    int(round(k * (so_ngay_int - 1) / (total_qty - 1)))
                    if total_qty > 1
                    else so_ngay_int // 2
                )
                for k in range(total_qty)
            ]
            for d_idx in indices:
                matrix_ngay[i, d_idx] += 1
        else:
            raw_dist = total_qty * day_ratios
            base_dist = np.floor(raw_dist).astype(int)
            matrix_ngay[i, :] += base_dist

            rem = total_qty - np.sum(base_dist)
            if rem > 0:
                fractional_parts = raw_dist - base_dist
                top_days = np.argsort(-fractional_parts)[:rem]
                for d_idx in top_days:
                    matrix_ngay[i, d_idx] += 1

    # =========================================================================
    # 4. HIỂN THỊ KẾT QUẢ DẠNG SỔ CHI TIẾT TỪ TRÊN XUỐNG & XUẤT EXCEL
    # =========================================================================
    st.divider()
    st.subheader(f"3. KẾT QUẢ PHÂN BỔ (TỔNG CỘNG {so_ngay_int} NGÀY)")

    tong_tien_thuc_te = int(np.sum(df_items["THÀNH TIỀN"]))
    chenh_lech = tong_tien_thuc_te - int(tong_tien_muc_tieu)

    m1, m2, m3 = st.columns(3)
    m1.metric("DOANH THU MỤC TIÊU", f"{tong_tien_muc_tieu:,.0f} đ")
    m2.metric("ĐÃ PHÂN BỔ THỰC TẾ", f"{tong_tien_thuc_te:,.0f} đ")
    m3.metric("CHÊNH LỆCH", f"{chenh_lech:,.0f} đ")

    # Gom dữ liệu dạng sổ dọc theo từng ngày bán
    records_hien_thi = []
    tong_sl_ban_tat_ca = 0

    for d_idx in range(so_ngay_int):
        ngay_ban_str = (ngay_bat_dau + pd.Timedelta(days=d_idx)).strftime(
            "%d/%m/%Y"
        )
        is_first = True

        for i in range(n):
            sl = matrix_ngay[i, d_idx]
            if sl > 0:
                gia = int(gia_arr[i])
                thanh_tien = int(sl * gia)
                tong_sl_ban_tat_ca += sl

                # Dòng đầu tiên của ngày in Ngày HĐ, các dòng món tiếp theo để trống
                records_hien_thi.append(
                    {
                        "ngay_in": ngay_ban_str if is_first else "",
                        "ma_hang": df_items.loc[i, "MÃ VTHH"],
                        "ten_hang": df_items.loc[i, "TÊN VTHH"],
                        "dvt": df_items.loc[i, "ĐVT"],
                        "so_luong": sl,
                        "don_gia": gia,
                        "thanh_tien": thanh_tien,
                    }
                )

    st.markdown("#### 📋 BẢNG KÊ CHI TIẾT BÁN HÀNG THEO NGÀY")

    import streamlit.components.v1 as components

    # 1. Tạo chuỗi HTML các dòng (không dùng khoảng trắng thụt lề đầu dòng để tránh lỗi Markdown)
    body_rows = []
    for r in records_hien_thi:
        sl_fmt = f"{r['so_luong']:,d}".replace(",", ".")
        gia_fmt = f"{r['don_gia']:,d}".replace(",", ".")
        tien_fmt = f"{r['thanh_tien']:,d}".replace(",", ".")
        body_rows.append(
            f'<tr>'
            f'<td class="td-c" style="font-weight:bold; color:#1b5e20;">{r["ngay_in"]}</td>'
            f'<td class="td-l">{r["ma_hang"]}</td>'
            f'<td class="td-l">{r["ten_hang"]}</td>'
            f'<td class="td-c">{r["dvt"]}</td>'
            f'<td class="td-r">{sl_fmt}</td>'
            f'<td class="td-r">{gia_fmt}</td>'
            f'<td class="td-boxed">{tien_fmt}</td>'
            f'</tr>'
        )

    sum_sl_fmt = f"{tong_sl_ban_tat_ca:,d}".replace(",", ".")
    sum_tien_fmt = f"{tong_tien_thuc_te:,d}".replace(",", ".")

    raw_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{
            margin: 0;
            padding: 0;
            font-family: Arial, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background-color: transparent;
        }}
        .table-container {{
            width: 100%;
            max-height: 580px;
            overflow-y: auto;
            border: 1px solid #d5d5d5;
            border-radius: 4px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            color: #111;
            background-color: #fff;
        }}
        thead th {{
            position: sticky;
            top: 0;
            z-index: 10;
            padding: 8px 10px;
            font-weight: bold;
            color: #ffffff;
            text-align: center;
        }}
        th.col-green {{
            background-color: #4F7942;
            width: 13%;
        }}
        th.col-orange {{
            background-color: #E26B00;
        }}
        td {{
            padding: 6px 10px;
            border-bottom: 1px solid #eeeeee;
            vertical-align: middle;
        }}
        tr:hover td {{
            background-color: #f8f9fa;
        }}
        .td-c {{ text-align: center; }}
        .td-l {{ text-align: left; }}
        .td-r {{ text-align: right; }}
        .td-boxed {{
            text-align: right;
            border: 1.5px solid #222222 !important;
            font-weight: 600;
        }}
        tfoot tr {{
            position: sticky;
            bottom: 0;
            background-color: #FFF2CC;
            font-weight: bold;
            border-top: 2px solid #888;
        }}
    </style>
    </head>
    <body>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th class="col-green">Ngày HĐ</th>
                        <th class="col-orange">Mã hàng (*)</th>
                        <th class="col-orange">Tên hàng</th>
                        <th class="col-orange">ĐVT</th>
                        <th class="col-orange">Số lượng</th>
                        <th class="col-orange">Đơn giá</th>
                        <th class="col-orange">Thành tiền</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(body_rows)}
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="4" class="td-c" style="padding: 9px;">TỔNG CỘNG PHÂN BỔ</td>
                        <td class="td-r" style="padding: 9px;">{sum_sl_fmt}</td>
                        <td></td>
                        <td class="td-boxed" style="padding: 9px; color: #b71c1c; font-size: 14px;">{sum_tien_fmt}</td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </body>
    </html>
    """

    components.html(raw_html, height=600, scrolling=True)

    # --- TẠO FILE EXCEL ĐỊNH DẠNG MÀU SẮC CHUẨN MẪU ---
    from openpyxl.styles import Border, Side

    df_excel_export = pd.DataFrame(
        [
            {
                "Ngày HĐ": r["ngay_in"],
                "Mã hàng (*)": r["ma_hang"],
                "Tên hàng": r["ten_hang"],
                "ĐVT": r["dvt"],
                "Số lượng": r["so_luong"],
                "Đơn giá": r["don_gia"],
                "Thành tiền": r["thanh_tien"],
            }
            for r in records_hien_thi
        ]
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_excel_export.to_excel(
            writer, sheet_name="SO_CHI_TIET_BAN_HANG", index=False
        )
        ws = writer.sheets["SO_CHI_TIET_BAN_HANG"]

        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        align_header = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        fill_green = PatternFill(
            start_color="4F7942", end_color="4F7942", fill_type="solid"
        )
        fill_orange = PatternFill(
            start_color="E26B00", end_color="E26B00", fill_type="solid"
        )

        ws.row_dimensions[1].height = 26
        for col_idx, cell in enumerate(ws[1], start=1):
            cell.font = font_header
            cell.alignment = align_header
            cell.fill = fill_green if col_idx == 1 else fill_orange

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        data_font = Font(name="Arial", size=10)
        box_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            ws.row_dimensions[row[0].row].height = 20
            row[0].alignment = align_center
            row[0].font = Font(name="Arial", size=10, bold=True)
            row[1].alignment = align_left
            row[1].font = data_font
            row[2].alignment = align_left
            row[2].font = data_font
            row[3].alignment = align_center
            row[3].font = data_font
            row[4].alignment = align_right
            row[4].number_format = "#,##0"
            row[4].font = data_font
            row[5].alignment = align_right
            row[5].number_format = "#,##0"
            row[5].font = data_font
            row[6].alignment = align_right
            row[6].number_format = "#,##0"
            row[6].font = data_font
            row[6].border = box_border

        # Dòng tổng cộng
        last_r = ws.max_row + 1
        ws.row_dimensions[last_r].height = 24
        ws.cell(last_r, 1, "TỔNG CỘNG").font = Font(
            name="Arial", size=11, bold=True
        )
        ws.cell(last_r, 1).alignment = align_center
        ws.cell(last_r, 5, tong_sl_ban_tat_ca).font = Font(
            name="Arial", size=11, bold=True
        )
        ws.cell(last_r, 5).number_format = "#,##0"
        ws.cell(last_r, 7, tong_tien_thuc_te).font = Font(
            name="Arial", size=11, bold=True, color="990000"
        )
        ws.cell(last_r, 7).number_format = "#,##0"
        ws.cell(last_r, 7).border = box_border

        sum_fill = PatternFill(
            start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
        )
        for c in range(1, 8):
            ws.cell(last_r, c).fill = sum_fill

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    st.download_button(
        label="📥 TẢI BẢNG KÊ CHI TIẾT (EXCEL)",
        data=output.getvalue(),
        file_name=f"BANG_KE_CHI_TIET_{so_ngay_int}_NGAY.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
