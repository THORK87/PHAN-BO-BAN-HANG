import base64
from datetime import date, datetime
import io
import json
import math
import secrets
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
ADMIN_KEY = "Duykhuong@2026"  # Mật khẩu quản trị

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


# =========================================================================
# HÀM QUẢN LÝ NHẬT KÝ HOẠT ĐỘNG
# =========================================================================
def get_remote_activity_log():
    try:
        res = requests.get(f"{ACTIVITY_LOG_API_URL}?ref={GITHUB_BRANCH}", headers=HEADERS)
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
        content_b64 = base64.b64encode(content_str.encode("utf-8")).decode("utf-8")
        payload = {
            "message": f"Update activity log - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            "content": content_b64,
            "branch": GITHUB_BRANCH
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
        "details": details
    }
    activity_log.append(log_entry)
    if len(activity_log) > 500:
        activity_log = activity_log[-500:]

    if update_remote_activity_log(activity_log, log_sha):
        return True
    return False
