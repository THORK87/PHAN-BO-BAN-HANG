import base64
from datetime import date, datetime
import io
import json
import math
import secrets
import textwrap
import numpy as np
from openpyxl.styles import Alignment, Font, PatternFill
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
ADMIN_KEY = "Duykhuong@2026"

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


db_licenses, file_sha = get_remote_licenses()

# =========================================================================
# 1. KIỂM TRA BẢN QUYỀN KHÁCH HÀNG
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
# 2. BẢNG ĐIỀU KHIỂN QUẢN TRỊ ADMIN
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
                            if update_remote
