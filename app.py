import base64
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
