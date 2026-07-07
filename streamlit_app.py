# -*- coding: utf-8 -*-
"""
Donation Box Collection System - Full Mobile Version
ระบบติดตามและจัดเก็บกล่องรับบริจาคผ่านมือถือ

คุณสมบัติหลัก:
- ใช้ Excel .xlsx ได้ทันที และเชื่อม Google Sheets/Drive ได้เมื่อใส่ Streamlit Secrets
- ระบบล็อกอินหลายผู้ใช้ พร้อมบันทึกชื่อเจ้าหน้าที่อัตโนมัติ
- Admin เพิ่ม/ลบ/แก้ไข รายการกล่อง และ User ส่งคำขอวางกล่องใหม่รออนุมัติ
- อัปโหลดข้อมูลกล่องทั้งหมดหรือบางรายการด้วยไฟล์ Excel .xlsx ตามแพทเทิร์น
- บันทึกการเก็บโดยไม่ต้องใส่จำนวนเงิน
- ถ่ายภาพ/อัปโหลดภาพไป Google Drive
- แผนที่/ลิงก์ Google Maps พร้อมแสดงรูปที่เคยบันทึก
- ดาวน์โหลดข้อมูลเป็นไฟล์ Excel .xlsx ได้
"""

from __future__ import annotations

import io
import json
import os
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import folium
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
except Exception:
    gspread = None
    Credentials = None
    build = None
    MediaIoBaseUpload = None

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
PHOTO_DIR = DATA_DIR / "photos"
BOXES_XLSX = DATA_DIR / "boxes_master.xlsx"
COLLECTIONS_XLSX = DATA_DIR / "collections.xlsx"
MISSING_XLSX = DATA_DIR / "missing_reports.xlsx"
USERS_XLSX = DATA_DIR / "users.xlsx"
PHOTO_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="ระบบเก็บกล่องบริจาค",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
:root{
  --main:#0f766e; --main2:#2563eb; --soft:#ecfeff; --warn:#fef3c7; --danger:#fee2e2;
}
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 1400px;}
html, body, [class*="css"] {font-size: 20px !important;}
h1 {font-size: 2.2rem !important; font-weight: 900 !important; color:#0f172a;}
h2, h3 {font-weight: 850 !important; color:#0f172a;}
[data-testid="stMetric"] {background: linear-gradient(135deg,#ecfeff,#ffffff); border:2px solid #99f6e4; border-radius:18px; padding:16px; box-shadow:0 4px 12px rgba(15,118,110,.12);}
[data-testid="stMetricLabel"] {font-size:1.05rem !important; font-weight:800 !important; color:#0f172a !important;}
[data-testid="stMetricValue"] {font-size:2.25rem !important; font-weight:950 !important; color:#0f766e !important;}
.stButton > button, .stDownloadButton > button {font-size: 20px !important; font-weight: 900 !important; border-radius: 14px !important; min-height: 3.2rem; width: 100%; border:2px solid #0f766e !important;}
.stLinkButton a {font-size: 20px !important; font-weight: 900 !important; border-radius: 14px !important; min-height: 3rem;}
.stTextInput input, .stNumberInput input, .stDateInput input, textarea {font-size: 20px !important; border-radius: 12px !important;}
div[data-baseweb="select"] * {font-size: 20px !important;}
.card {padding: 1.05rem; border: 2px solid #bae6fd; border-radius: 18px; background: #ffffff; box-shadow: 0 4px 12px rgba(0,0,0,.07); margin-bottom: .85rem;}
.card-title {font-size:1.25rem; font-weight:950; color:#0f172a;}
.badge {padding: .3rem .7rem; border-radius: 999px; font-weight: 900; font-size: 1rem; display:inline-block; margin:.15rem;}
.green {background:#dcfce7; color:#166534; border:1px solid #86efac;} 
.red {background:#fee2e2; color:#991b1b; border:1px solid #fca5a5;} 
.yellow {background:#fef3c7; color:#92400e; border:1px solid #fcd34d;} 
.blue {background:#dbeafe; color:#1e40af; border:1px solid #93c5fd;}
.gray {background:#f1f5f9; color:#334155; border:1px solid #cbd5e1;}
.small {font-size: 1rem; color:#475569; line-height:1.6;}
.success-box {background:#dcfce7; border:3px solid #22c55e; color:#14532d; padding:18px; border-radius:18px; font-size:1.35rem; font-weight:950; text-align:center;}
.warning-box {background:#fef3c7; border:3px solid #f59e0b; color:#78350f; padding:16px; border-radius:18px; font-size:1.15rem; font-weight:850;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

BOX_COLUMNS = [
    "box_id", "original_no", "province_zone", "subzone", "store_name", "floor", "contact",
    "landmark", "google_maps_url", "status", "cycle_months", "last_collect_month",
    "last_collect_day", "last_collect_date", "next_due_date", "box_photo_url", "note",
    "created_by", "updated_by", "updated_at"
]
COLLECTION_COLUMNS = [
    "record_id", "box_id", "store_name", "collect_date", "collector", "box_condition",
    "result_status", "photo_url", "google_maps_url", "note", "created_at"
]
MISSING_COLUMNS = [
    "report_id", "box_id", "store_name", "found_date", "contact_date", "contact_person",
    "detail", "photo_url", "google_maps_url", "case_status", "created_by", "created_at"
]
USER_COLUMNS = ["username", "password", "display_name", "role", "status", "created_at"]

REQUIRED_IMPORT_COLUMNS = [
    "box_id", "store_name", "province_zone", "subzone", "contact", "landmark",
    "google_maps_url", "cycle_months", "status", "last_collect_date", "next_due_date", "note"
]


def safe_secret(key: str, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def has_google_config() -> bool:
    return bool(safe_secret("GOOGLE_SHEET_ID", "")) and bool(safe_secret("gcp_service_account", {})) and gspread is not None


def has_drive_config() -> bool:
    return bool(safe_secret("GOOGLE_DRIVE_FOLDER_ID", "")) and bool(safe_secret("gcp_service_account", {})) and build is not None


def get_credentials(scopes: List[str]):
    info = dict(safe_secret("gcp_service_account", {}))
    if not info or Credentials is None:
        return None
    if "private_key" in info:
        info["private_key"] = info["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(info, scopes=scopes)


@st.cache_resource(show_spinner=False)
def get_gspread_client():
    creds = get_credentials(["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    if not creds:
        return None
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_drive_service():
    creds = get_credentials(["https://www.googleapis.com/auth/drive.file"])
    if not creds:
        return None
    return build("drive", "v3", credentials=creds)


def normalize_records(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    if df is None:
        df = pd.DataFrame(columns=cols)
    # รองรับไฟล์รุ่นเก่าที่มี gps_lat/gps_lng แต่เวอร์ชันนี้ใช้ลิงก์แทน
    if "google_maps_url" not in df.columns:
        df["google_maps_url"] = ""
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    out = df[cols].copy()
    return out.fillna("")


def normalize_boxes(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_records(df, BOX_COLUMNS)
    df["cycle_months"] = pd.to_numeric(df["cycle_months"], errors="coerce").fillna(3).astype(int)
    df["status"] = df["status"].replace("", "ใช้งาน").fillna("ใช้งาน")
    df["store_name"] = df["store_name"].fillna("")
    return df


def ensure_excel(path: Path, cols: List[str]):
    if not path.exists():
        pd.DataFrame(columns=cols).to_excel(path, index=False, engine="openpyxl")


def read_excel_df(path: Path, cols: List[str]) -> pd.DataFrame:
    ensure_excel(path, cols)
    df = pd.read_excel(path, dtype=str, engine="openpyxl").fillna("")
    return normalize_records(df, cols)


def write_excel_df(path: Path, df: pd.DataFrame):
    df.to_excel(path, index=False, engine="openpyxl")


def excel_bytes(df: pd.DataFrame, sheet_name: str = "data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.fillna("").to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return buf.getvalue()


def upload_bytes_to_drive(data: bytes, filename: str, mimetype: str) -> str:
    folder_id = safe_secret("GOOGLE_DRIVE_FOLDER_ID", "")
    if not folder_id or not has_drive_config():
        return ""
    try:
        service = get_drive_service()
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mimetype, resumable=False)
        meta = {"name": filename, "parents": [folder_id]}
        uploaded = service.files().create(body=meta, media_body=media, fields="id, webViewLink").execute()
        return uploaded.get("webViewLink", "")
    except Exception as e:
        st.warning(f"อัปโหลดไป Google Drive ไม่สำเร็จ: {e}")
        return ""


def backup_df_to_drive(sheet_name: str, df: pd.DataFrame):
    # สำรองข้อมูลเป็น Excel .xlsx ใน Google Drive เพื่อให้ข้อมูลไม่อยู่เฉพาะเครื่อง/มือถือ
    if not has_drive_config():
        return
    data = excel_bytes(df, sheet_name)
    filename = f"backup_{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    upload_bytes_to_drive(data, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def get_sheet_df(sheet_name: str, fallback_path: Path, cols: List[str]) -> pd.DataFrame:
    if has_google_config():
        try:
            gc = get_gspread_client()
            sh = gc.open_by_key(safe_secret("GOOGLE_SHEET_ID"))
            try:
                ws = sh.worksheet(sheet_name)
            except Exception:
                ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=max(len(cols), 20))
                ws.append_row(cols)
                if fallback_path.exists():
                    seed = pd.read_excel(fallback_path, dtype=str, engine="openpyxl").fillna("")
                    seed = normalize_records(seed, cols)
                    if len(seed):
                        ws.update([cols] + seed.astype(str).values.tolist())
            values = ws.get_all_records()
            return normalize_records(pd.DataFrame(values), cols) if values else pd.DataFrame(columns=cols)
        except Exception as e:
            st.warning(f"อ่าน Google Sheets ไม่สำเร็จ จึงใช้ไฟล์ Excel .xlsx แทน: {e}")
    return read_excel_df(fallback_path, cols)


def save_sheet_df(sheet_name: str, path: Path, df: pd.DataFrame, cols: List[str], backup=True):
    df = normalize_records(df, cols).fillna("")
    saved_google = False
    if has_google_config():
        try:
            gc = get_gspread_client()
            sh = gc.open_by_key(safe_secret("GOOGLE_SHEET_ID"))
            try:
                ws = sh.worksheet(sheet_name)
            except Exception:
                ws = sh.add_worksheet(title=sheet_name, rows=max(len(df)+20, 1000), cols=max(len(cols), 20))
            ws.clear()
            ws.update([cols] + df.astype(str).values.tolist())
            saved_google = True
        except Exception as e:
            st.warning(f"บันทึก Google Sheets ไม่สำเร็จ จึงบันทึกลง Excel .xlsx แทน: {e}")
    write_excel_df(path, df)
    if backup:
        backup_df_to_drive(sheet_name, df)
    return saved_google


@st.cache_data(ttl=20, show_spinner=False)
def load_boxes_cached() -> pd.DataFrame:
    return normalize_boxes(get_sheet_df("boxes", BOXES_XLSX, BOX_COLUMNS))


@st.cache_data(ttl=20, show_spinner=False)
def load_collections_cached() -> pd.DataFrame:
    return get_sheet_df("collections", COLLECTIONS_XLSX, COLLECTION_COLUMNS)


@st.cache_data(ttl=20, show_spinner=False)
def load_missing_cached() -> pd.DataFrame:
    return get_sheet_df("missing_reports", MISSING_XLSX, MISSING_COLUMNS)


@st.cache_data(ttl=20, show_spinner=False)
def load_users_cached() -> pd.DataFrame:
    users = get_sheet_df("users", USERS_XLSX, USER_COLUMNS)
    if len(users) == 0:
        users = pd.DataFrame([{
            "username": "admin", "password": "1234", "display_name": "ผู้ดูแลระบบ",
            "role": "admin", "status": "ใช้งาน", "created_at": datetime.now().isoformat(timespec="seconds")
        }])
        save_sheet_df("users", USERS_XLSX, users, USER_COLUMNS, backup=False)
    return users


def clear_and_save():
    st.cache_data.clear()


def save_boxes(df: pd.DataFrame):
    clear_and_save(); save_sheet_df("boxes", BOXES_XLSX, normalize_boxes(df), BOX_COLUMNS)


def save_collections(df: pd.DataFrame):
    clear_and_save(); save_sheet_df("collections", COLLECTIONS_XLSX, df, COLLECTION_COLUMNS)


def save_missing(df: pd.DataFrame):
    clear_and_save(); save_sheet_df("missing_reports", MISSING_XLSX, df, MISSING_COLUMNS)


def save_users(df: pd.DataFrame):
    clear_and_save(); save_sheet_df("users", USERS_XLSX, df, USER_COLUMNS)


def upload_photo(file_obj, prefix: str) -> str:
    if file_obj is None:
        return ""
    data = file_obj.getvalue()
    ext = "jpg"
    mime = "image/jpeg"
    filename = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}.{ext}"
    drive_link = upload_bytes_to_drive(data, filename, mime)
    if drive_link:
        return drive_link
    local_path = PHOTO_DIR / filename
    local_path.write_bytes(data)
    return str(local_path.as_posix())


def local_image_exists(url: str) -> bool:
    return bool(url) and url.startswith(str(PHOTO_DIR.as_posix())) and Path(url).exists()


def show_saved_photo(photo_url: str, caption="รูปที่บันทึกไว้"):
    if not photo_url:
        st.caption("ยังไม่มีรูป")
    elif local_image_exists(photo_url):
        st.image(photo_url, caption=caption, use_container_width=True)
    else:
        st.link_button("🖼️ เปิดรูปที่บันทึกไว้", photo_url)


def login():
    app_name = safe_secret("APP_NAME", "ระบบเก็บกล่องบริจาค")
    if "auth" not in st.session_state:
        st.session_state.auth = False
        st.session_state.username = ""
        st.session_state.display_name = ""
        st.session_state.role = "user"
    if st.session_state.auth:
        return True
    st.title(f"📦 {app_name}")
    st.markdown("<div class='warning-box'>กรุณาเข้าสู่ระบบ เพื่อให้ระบบบันทึกได้ว่าเจ้าหน้าที่คนใดเป็นผู้ทำรายการ</div>", unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("ชื่อผู้ใช้", placeholder="เช่น admin หรือชื่อเจ้าหน้าที่")
        password = st.text_input("รหัสผ่าน", type="password")
        ok = st.form_submit_button("เข้าสู่ระบบ")
    if ok:
        users = load_users_cached()
        # รองรับ secret USERS เดิมด้วย
        secret_users = dict(safe_secret("USERS", {}))
        if secret_users and username in secret_users and str(secret_users[username]) == password:
            st.session_state.auth = True
            st.session_state.username = username
            st.session_state.display_name = username
            st.session_state.role = "admin"
            st.rerun()
        matched = users[(users["username"].astype(str) == username) & (users["password"].astype(str) == password) & (users["status"].astype(str) == "ใช้งาน")]
        if len(matched):
            u = matched.iloc[0]
            st.session_state.auth = True
            st.session_state.username = username
            st.session_state.display_name = u.get("display_name", username) or username
            st.session_state.role = u.get("role", "user") or "user"
            st.rerun()
        else:
            st.error("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง หรือบัญชีถูกปิดใช้งาน")
    st.info("ค่าเริ่มต้น: admin / 1234  ควรเปลี่ยนทันทีเมื่อใช้งานจริง")
    return False


def green_success(text: str):
    st.markdown(f"<div class='success-box'>✅ {text}</div>", unsafe_allow_html=True)


def due_status(row) -> Tuple[str, str, Optional[int]]:
    today = pd.Timestamp(date.today())
    d = pd.to_datetime(row.get("next_due_date", ""), errors="coerce")
    if pd.isna(d):
        lcd = pd.to_datetime(row.get("last_collect_date", ""), errors="coerce")
        if pd.isna(lcd):
            return "ยังไม่มีวันกำหนด", "gray", None
        d = lcd + pd.DateOffset(months=int(row.get("cycle_months", 3) or 3))
    diff = (d.normalize() - today).days
    if diff < 0:
        return f"เกินกำหนด {abs(diff)} วัน", "red", diff
    if diff <= 14:
        return f"ใกล้ครบกำหนด {diff} วัน", "yellow", diff
    return f"อีก {diff} วัน", "green", diff


def google_maps_link(row_or_url) -> str:
    if isinstance(row_or_url, str):
        url = row_or_url
        return url if url else "https://www.google.com/maps"
    url = str(row_or_url.get("google_maps_url", "") or "")
    if url and url.lower() != "nan":
        return url
    q = f"{row_or_url.get('store_name','')} {row_or_url.get('subzone','')} {row_or_url.get('province_zone','')}"
    return "https://www.google.com/maps/search/?api=1&query=" + q.replace(" ", "+")


def parse_lat_lng_from_url(url: str):
    # รองรับลิงก์แบบมี @lat,lng หรือ query=lat,lng แบบคร่าว ๆ เพื่อปักหมุดได้บางกรณี
    import re
    if not url:
        return None, None
    m = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", url)
    if not m:
        m = re.search(r"(?:query=|destination=)(-?\d+\.\d+),(-?\d+\.\d+)", url)
    if m:
        return float(m.group(1)), float(m.group(2))
    return None, None


def filters(df: pd.DataFrame, key_prefix=""):
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        zones = ["ทั้งหมด"] + sorted([z for z in df["province_zone"].dropna().unique().tolist() if str(z)])
        zone = st.selectbox("จังหวัด/โซน", zones, key=f"{key_prefix}_zone")
    with c2:
        statuses = ["ทั้งหมด"] + sorted([z for z in df["status"].dropna().unique().tolist() if str(z)])
        status = st.selectbox("สถานะ", statuses, key=f"{key_prefix}_status")
    with c3:
        q = st.text_input("ค้นหา เลขกล่อง / ชื่อร้าน / เบอร์ / จุดสังเกต", key=f"{key_prefix}_q")
    out = df.copy()
    if zone != "ทั้งหมด":
        out = out[out["province_zone"].astype(str) == zone]
    if status != "ทั้งหมด":
        out = out[out["status"].astype(str) == status]
    if q:
        mask = out.astype(str).apply(lambda col: col.str.contains(q, case=False, na=False)).any(axis=1)
        out = out[mask]
    return out


def select_box_widget(boxes, label="เลือกกล่อง", key="box_select"):
    if len(boxes) == 0:
        return None
    display = boxes.apply(lambda r: f"{r['box_id']} | {r['store_name']} | {r['province_zone']}", axis=1).tolist()
    selected = st.selectbox(label, display, key=key)
    box_id = selected.split(" | ")[0]
    return boxes[boxes["box_id"] == box_id].iloc[0]


def page_dashboard(boxes, collections, missing):
    st.title("📊 Dashboard ภาพรวม")
    active = boxes[boxes["status"].astype(str).str.contains("ใช้งาน", na=False)]
    boxes = boxes.copy()
    due_labels = boxes.apply(due_status, axis=1, result_type="expand")
    boxes["due_text"] = due_labels[0]
    overdue = boxes[boxes["due_text"].astype(str).str.contains("เกินกำหนด")]
    near = boxes[boxes["due_text"].astype(str).str.contains("ใกล้ครบ")]
    today_str = date.today().isoformat()
    today_col = collections[collections["collect_date"].astype(str) == today_str] if len(collections) else collections

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("กล่องทั้งหมด", len(boxes))
    c2.metric("กล่องใช้งาน", len(active))
    c3.metric("เก็บวันนี้", len(today_col))
    c4.metric("ใกล้ครบกำหนด", len(near))
    c5.metric("เกินกำหนด", len(overdue))

    c6, c7, c8, c9 = st.columns(4)
    c6.metric("แจ้งกล่องหาย", len(missing))
    c7.metric("รออนุมัติ", int((boxes["status"].astype(str) == "รออนุมัติ").sum()))
    c8.metric("รอบ 3 เดือน", int((boxes["cycle_months"] == 3).sum()))
    c9.metric("รอบ 6 เดือน", int((boxes["cycle_months"] == 6).sum()))

    st.subheader("จำนวนกล่องตามจังหวัด/โซน")
    if len(boxes):
        st.bar_chart(boxes.groupby("province_zone")["box_id"].count().sort_values(ascending=False))
    st.subheader("กล่องที่ควรติดตาม")
    show = pd.concat([overdue, near]).head(50)
    if len(show):
        st.dataframe(show[["box_id", "store_name", "province_zone", "subzone", "status", "cycle_months", "next_due_date", "due_text"]], use_container_width=True, hide_index=True)
    else:
        st.success("ยังไม่มีกล่องเกินกำหนดหรือใกล้ครบกำหนด")


def page_boxes(boxes, collections):
    st.title("📦 รายการกล่องบริจาค")

    # Admin จัดการข้อมูลได้จากหน้านี้โดยตรง
    if st.session_state.get("role") == "admin":
        tab_list, tab_add, tab_edit, tab_delete = st.tabs(["📋 รายการและประวัติ", "➕ เพิ่ม", "✏️ แก้ไข", "🗑️ ลบ"])
    else:
        tab_list, = st.tabs(["📋 รายการและประวัติ"])

    with tab_list:
        df = filters(boxes, "boxes")
        st.caption(f"พบ {len(df):,} รายการ")
        for _, row in df.head(120).iterrows():
            text, color, _ = due_status(row)
            history = collections[collections["box_id"].astype(str) == str(row["box_id"])].copy() if len(collections) else pd.DataFrame(columns=COLLECTION_COLUMNS)
            last_text = "ยังไม่มีประวัติการเก็บ"
            if len(history):
                history = history.sort_values("collect_date", ascending=False)
                last_date = str(history.iloc[0].get("collect_date", ""))
                try:
                    days_ago = (pd.Timestamp(date.today()) - pd.to_datetime(last_date)).days
                    last_text = f"เก็บล่าสุด: {last_date} ({days_ago} วันที่ผ่านมา)"
                except Exception:
                    last_text = f"เก็บล่าสุด: {last_date}"
            st.markdown(f"""
            <div class='card'>
            <div class='card-title'>{row['box_id']} — {row['store_name']}</div>
            <span class='small'>โซน: {row.get('province_zone','')} | พื้นที่: {row.get('subzone','')} | ติดต่อ: {row.get('contact','')}</span><br>
            <span class='badge {color}'>{text}</span> <span class='badge gray'>{row.get('status','')}</span> <span class='badge blue'>{last_text}</span>
            </div>
            """, unsafe_allow_html=True)
            c1, c2 = st.columns([1, 1])
            with c1:
                st.link_button("🧭 เปิดลิงก์ Google Maps", google_maps_link(row))
            with c2:
                if row.get("box_photo_url", ""):
                    show_saved_photo(row.get("box_photo_url", ""), "รูปจุดวางกล่อง")
            with st.expander("ดูประวัติการเก็บกล่องนี้"):
                if len(history):
                    show_cols = [c for c in ["collect_date", "collector", "box_condition", "result_status", "google_maps_url", "note", "photo_url", "created_at"] if c in history.columns]
                    st.dataframe(history[show_cols], use_container_width=True, hide_index=True)
                else:
                    st.info("ยังไม่มีประวัติการเก็บกล่องนี้")
        st.download_button("⬇️ ดาวน์โหลดรายการที่กรองแล้ว Excel", excel_bytes(df, "filtered_boxes"), "filtered_boxes.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if st.session_state.get("role") == "admin":
        with tab_add:
            new = box_form(default={"box_id": f"BOX-{len(boxes)+1:04d}", "status": "ใช้งาน"}, form_key="boxes_add_form")
            if new:
                if str(new["box_id"]) in boxes["box_id"].astype(str).tolist():
                    st.error("เลขกล่องนี้มีอยู่แล้ว กรุณาใช้เลขกล่องอื่น")
                else:
                    save_boxes(pd.concat([boxes, pd.DataFrame([new])], ignore_index=True))
                    green_success("เพิ่มข้อมูลกล่องเรียบร้อยแล้ว")
                    st.rerun()
        with tab_edit:
            row = select_box_widget(boxes, "เลือกกล่องที่ต้องการแก้ไข", key="boxes_edit_select")
            if row is not None:
                edited = box_form(default=row.to_dict(), form_key="boxes_edit_form")
                if edited:
                    b = boxes.copy()
                    idx = b.index[b["box_id"].astype(str) == str(row["box_id"])]
                    if len(idx):
                        for col, val in edited.items():
                            b.loc[idx[0], col] = val
                        save_boxes(b)
                        green_success("แก้ไขข้อมูลกล่องเรียบร้อยแล้ว")
                        st.rerun()
        with tab_delete:
            row = select_box_widget(boxes, "เลือกกล่องที่ต้องการลบ", key="boxes_delete_select")
            if row is not None:
                st.warning(f"กำลังจะลบ: {row['box_id']} — {row['store_name']}")
                confirm = st.checkbox("ยืนยันว่าต้องการลบรายการนี้", key="boxes_delete_confirm")
                if st.button("🗑️ ลบกล่องนี้", disabled=not confirm, key="boxes_delete_btn"):
                    b = boxes[boxes["box_id"].astype(str) != str(row["box_id"])].copy()
                    save_boxes(b)
                    green_success("ลบข้อมูลกล่องเรียบร้อยแล้ว")
                    st.rerun()


def page_map(boxes, collections):
    st.title("🗺️ แผนที่ GPS / ลิงก์นำทาง พร้อมรูปที่บันทึก")
    st.markdown("<div class='warning-box'>เวอร์ชันนี้ใช้วิธีวางลิงก์ Google Maps แทนการบันทึกพิกัดโดยตรง</div>", unsafe_allow_html=True)
    df = filters(boxes, "map")
    latest_photo = {}
    if len(collections):
        temp = collections.sort_values("created_at").dropna(subset=["box_id"])
        for _, r in temp.iterrows():
            if r.get("photo_url", ""):
                latest_photo[str(r["box_id"])] = r.get("photo_url", "")

    # พยายามปักหมุดเฉพาะรายการที่ลิงก์มีพิกัด
    points = []
    for _, row in df.iterrows():
        lat, lng = parse_lat_lng_from_url(row.get("google_maps_url", ""))
        if lat and lng:
            points.append((lat, lng, row))
    if points:
        center = [sum(p[0] for p in points) / len(points), sum(p[1] for p in points) / len(points)]
        m = folium.Map(location=center, zoom_start=11, control_scale=True)
        for lat, lng, row in points:
            purl = latest_photo.get(str(row["box_id"]), row.get("box_photo_url", ""))
            photo_html = f"<br><a href='{purl}' target='_blank'>ดูรูปที่บันทึก</a>" if purl else ""
            popup = f"<b>{row['box_id']}</b><br>{row['store_name']}<br>{row['province_zone']}<br><a href='{google_maps_link(row)}' target='_blank'>นำทาง</a>{photo_html}"
            folium.Marker([lat, lng], popup=popup, tooltip=f"{row['box_id']} {row['store_name']}").add_to(m)
        st_folium(m, height=520, use_container_width=True)
    else:
        st.info("ยังไม่มีลิงก์ Google Maps ที่มีพิกัดแบบปักหมุด ระบบจะแสดงเป็นรายการนำทางแทน")

    st.subheader("รายการนำทางและรูปที่เคยบันทึก")
    for _, row in df.head(80).iterrows():
        purl = latest_photo.get(str(row["box_id"]), row.get("box_photo_url", ""))
        with st.container(border=True):
            st.markdown(f"**{row['box_id']} — {row['store_name']}**  \nโซน: {row.get('province_zone','')} | พื้นที่: {row.get('subzone','')}")
            c1, c2 = st.columns([1, 1])
            with c1:
                st.link_button("🧭 เปิด Google Maps", google_maps_link(row))
            with c2:
                show_saved_photo(purl, "รูปที่บันทึกไว้")


def page_collect(boxes, collections):
    st.title("✅ บันทึกการเก็บกล่อง")
    df = filters(boxes, "collect")
    row = select_box_widget(df, key="collect_select")
    if row is None:
        st.warning("ไม่พบข้อมูลกล่อง")
        return
    st.info(f"กล่อง {row['box_id']} | {row['store_name']} | {row['province_zone']}")
    st.link_button("🧭 นำทางด้วย Google Maps", google_maps_link(row))
    with st.form("collect_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            collect_date = st.date_input("วันที่เก็บ", value=date.today())
            collector = st.text_input("ผู้บันทึก/เจ้าหน้าที่", value=st.session_state.get("display_name") or st.session_state.get("username", ""), disabled=True)
            google_maps_url = st.text_input("ลิงก์ Google Maps จุดที่ไปเก็บ", value=row.get("google_maps_url", ""))
        with c2:
            condition = st.selectbox("สภาพกล่อง", ["สมบูรณ์", "ชำรุด", "กล่องหาย", "ร้านปิด", "อื่น ๆ"])
            result_status = st.selectbox("ผลการดำเนินงาน", ["เก็บแล้ว", "ไม่ได้เก็บ", "ต้องติดตาม", "ยกเลิกจุดวาง"])
            note = st.text_area("หมายเหตุ")
        photo = st.camera_input("ถ่ายภาพกล่อง/หน้าร้าน")
        submit = st.form_submit_button("💾 บันทึกการเก็บ")
    if submit:
        photo_url = upload_photo(photo, row["box_id"])
        new = {
            "record_id": uuid.uuid4().hex[:12], "box_id": row["box_id"], "store_name": row["store_name"],
            "collect_date": collect_date.isoformat(), "collector": st.session_state.get("display_name") or st.session_state.get("username", ""),
            "box_condition": condition, "result_status": result_status, "photo_url": photo_url,
            "google_maps_url": google_maps_url, "note": note, "created_at": datetime.now().isoformat(timespec="seconds")
        }
        collections = pd.concat([collections, pd.DataFrame([new])], ignore_index=True)
        save_collections(collections)
        b = boxes.copy()
        idx = b.index[b["box_id"].astype(str) == str(row["box_id"])]
        if len(idx):
            i = idx[0]
            b.loc[i, "last_collect_date"] = collect_date.isoformat()
            b.loc[i, "next_due_date"] = (pd.Timestamp(collect_date) + pd.DateOffset(months=int(row.get("cycle_months", 3)))).date().isoformat()
            b.loc[i, "google_maps_url"] = google_maps_url
            b.loc[i, "updated_by"] = st.session_state.get("username", "")
            b.loc[i, "updated_at"] = datetime.now().isoformat(timespec="seconds")
            if photo_url:
                b.loc[i, "box_photo_url"] = photo_url
            if condition == "กล่องหาย":
                b.loc[i, "status"] = "กล่องหาย"
        save_boxes(b)
        green_success("บันทึกเรียบร้อยแล้ว")
        st.balloons()


def page_missing(boxes, missing):
    st.title("🚨 กล่องหาย / กล่องคืน")
    tab_missing, tab_return, tab_list = st.tabs(["แจ้งกล่องหาย", "บันทึกกล่องคืน", "รายการติดตาม"])
    with tab_missing:
        row = select_box_widget(boxes, "เลือกกล่องที่หาย", key="missing_select")
        if row is None:
            st.warning("ไม่พบข้อมูลกล่อง")
            return
        with st.form("missing_form"):
            found_date = st.date_input("วันที่ตรวจพบ", value=date.today())
            contact_date = st.date_input("วันที่ติดต่อร้าน", value=date.today())
            contact_person = st.text_input("ผู้ให้ข้อมูล/ผู้ติดต่อ")
            google_maps_url = st.text_input("ลิงก์ Google Maps จุดที่ตรวจพบ", value=row.get("google_maps_url", ""))
            detail = st.text_area("รายละเอียดเหตุการณ์")
            case_status = st.selectbox("สถานะเคส", ["รอตรวจสอบ", "กำลังติดตาม", "ปิดเคส", "ต้องวางใหม่"])
            photo = st.camera_input("ถ่ายภาพจุดวางเดิม")
            submit = st.form_submit_button("💾 บันทึกแจ้งกล่องหาย")
        if submit:
            photo_url = upload_photo(photo, f"missing_{row['box_id']}")
            new = {
                "report_id": uuid.uuid4().hex[:12], "box_id": row["box_id"], "store_name": row["store_name"],
                "found_date": found_date.isoformat(), "contact_date": contact_date.isoformat(), "contact_person": contact_person,
                "detail": detail, "photo_url": photo_url, "google_maps_url": google_maps_url, "case_status": case_status,
                "created_by": st.session_state.get("username", ""), "created_at": datetime.now().isoformat(timespec="seconds")
            }
            save_missing(pd.concat([missing, pd.DataFrame([new])], ignore_index=True))
            b = boxes.copy(); b.loc[b["box_id"].astype(str) == str(row["box_id"]), "status"] = "กล่องหาย"; save_boxes(b)
            green_success("บันทึกแจ้งกล่องหายเรียบร้อยแล้ว")
    with tab_return:
        missing_boxes = boxes[boxes["status"].astype(str).str.contains("หาย", na=False)]
        if len(missing_boxes) == 0:
            st.success("ยังไม่มีกล่องสถานะหาย")
        else:
            row2 = select_box_widget(missing_boxes, "เลือกกล่องที่พบคืน", key="return_select")
            with st.form("return_form"):
                return_date = st.date_input("วันที่พบ/รับคืน", value=date.today())
                return_detail = st.text_area("รายละเอียดการคืนกล่อง", placeholder="เช่น ร้านพบกล่องแล้ว / นำกลับมาตรวจสอบ / วางคืนที่เดิม")
                return_photo = st.camera_input("ถ่ายภาพกล่องที่พบคืน")
                ok_return = st.form_submit_button("💾 บันทึกกล่องคืน")
            if ok_return and row2 is not None:
                photo_url = upload_photo(return_photo, f"return_{row2['box_id']}")
                new = {
                    "report_id": uuid.uuid4().hex[:12], "box_id": row2["box_id"], "store_name": row2["store_name"],
                    "found_date": return_date.isoformat(), "contact_date": return_date.isoformat(), "contact_person": st.session_state.get("display_name", ""),
                    "detail": return_detail, "photo_url": photo_url, "google_maps_url": row2.get("google_maps_url", ""), "case_status": "คืนกล่องแล้ว",
                    "created_by": st.session_state.get("username", ""), "created_at": datetime.now().isoformat(timespec="seconds")
                }
                save_missing(pd.concat([missing, pd.DataFrame([new])], ignore_index=True))
                b = boxes.copy()
                idx = b.index[b["box_id"].astype(str) == str(row2["box_id"])]
                if len(idx):
                    b.loc[idx[0], "status"] = "ใช้งาน"
                    b.loc[idx[0], "updated_by"] = st.session_state.get("username", "")
                    b.loc[idx[0], "updated_at"] = datetime.now().isoformat(timespec="seconds")
                    if photo_url:
                        b.loc[idx[0], "box_photo_url"] = photo_url
                save_boxes(b)
                green_success("บันทึกการคืนกล่องเรียบร้อยแล้ว")
    with tab_list:
        st.subheader("รายการแจ้งกล่องหาย/คืน")
        st.dataframe(missing.sort_values("created_at", ascending=False) if len(missing) else missing, use_container_width=True, hide_index=True)
        st.download_button("⬇️ ดาวน์โหลดรายการกล่องหาย/คืน", excel_bytes(missing, "missing_return_reports"), "missing_return_reports.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def box_form(default=None, form_key="box_form"):
    default = default or {}
    with st.form(form_key):
        c1, c2 = st.columns(2)
        with c1:
            box_id = st.text_input("เลขกล่อง", value=str(default.get("box_id", "")))
            store_name = st.text_input("ชื่อร้าน/สถานที่", value=str(default.get("store_name", "")))
            province_zone = st.text_input("จังหวัด/โซน", value=str(default.get("province_zone", "")))
            subzone = st.text_input("พื้นที่ย่อย/จุดสังเกต", value=str(default.get("subzone", "")))
            contact = st.text_input("เบอร์ติดต่อ/ผู้ติดต่อ", value=str(default.get("contact", "")))
        with c2:
            cycle_val = int(default.get("cycle_months", 3) or 3)
            cycle_months = st.selectbox("รอบเก็บ", [3, 6], index=0 if cycle_val == 3 else 1)
            status_options = ["ใช้งาน", "รอตรวจสอบ", "ย้ายจุด", "กล่องหาย", "ยกเลิก"]
            cur_status = default.get("status", "ใช้งาน") or "ใช้งาน"
            status = st.selectbox("สถานะ", status_options, index=status_options.index(cur_status) if cur_status in status_options else 0)
            last_collect_date = st.date_input("วันที่เริ่มวาง/วันที่เก็บล่าสุด", value=pd.to_datetime(default.get("last_collect_date", date.today()), errors="coerce").date() if str(default.get("last_collect_date", "")) else date.today())
            google_maps_url = st.text_input("วางลิงก์ Google Maps", value=str(default.get("google_maps_url", "")))
        landmark = st.text_area("รายละเอียดจุดวาง/หมายเหตุ", value=str(default.get("landmark", "")))
        note = st.text_area("หมายเหตุเพิ่มเติม", value=str(default.get("note", "")))
        photo = st.camera_input("ถ่ายภาพ/อัปเดตรูปจุดวาง")
        submit = st.form_submit_button("💾 บันทึกข้อมูลกล่อง")
    if not submit:
        return None
    if not box_id or not store_name:
        st.error("กรุณากรอกเลขกล่องและชื่อร้าน")
        return None
    photo_url = upload_photo(photo, f"box_{box_id}")
    old_photo = default.get("box_photo_url", "") if default else ""
    out = {col: "" for col in BOX_COLUMNS}
    out.update({
        "box_id": box_id, "province_zone": province_zone, "subzone": subzone, "store_name": store_name,
        "contact": contact, "landmark": landmark, "google_maps_url": google_maps_url,
        "status": status, "cycle_months": cycle_months, "last_collect_date": last_collect_date.isoformat(),
        "next_due_date": (pd.Timestamp(last_collect_date) + pd.DateOffset(months=int(cycle_months))).date().isoformat(),
        "box_photo_url": photo_url or old_photo, "note": note,
        "created_by": default.get("created_by", st.session_state.get("username", "")) if default else st.session_state.get("username", ""),
        "updated_by": st.session_state.get("username", ""),
        "updated_at": datetime.now().isoformat(timespec="seconds")
    })
    return out


def page_manage_boxes(boxes):
    st.title("📍 วางกล่องใหม่ / จัดการกล่อง")
    is_admin = st.session_state.get("role") == "admin"

    if not is_admin:
        st.markdown("<div class='warning-box'>เจ้าหน้าที่สามารถเพิ่มจุดวางกล่องใหม่ได้ ระบบจะบันทึกเป็นสถานะ ‘รออนุมัติ’ เพื่อให้ผู้ดูแลระบบตรวจสอบก่อนใช้งานจริง</div>", unsafe_allow_html=True)
        new = box_form(default={"box_id": f"BOX-{len(boxes)+1:04d}", "status": "รออนุมัติ"}, form_key="user_add_box_form")
        if new:
            if str(new["box_id"]) in boxes["box_id"].astype(str).tolist():
                st.error("เลขกล่องนี้มีอยู่แล้ว กรุณาใช้เลขกล่องอื่น")
            else:
                new["status"] = "รออนุมัติ"
                new["created_by"] = st.session_state.get("username", "")
                new["updated_by"] = st.session_state.get("username", "")
                save_boxes(pd.concat([boxes, pd.DataFrame([new])], ignore_index=True))
                green_success("ส่งคำขอวางกล่องใหม่เรียบร้อยแล้ว รอผู้ดูแลระบบอนุมัติ")
        st.subheader("รายการที่คุณส่งคำขอ")
        mine = boxes[(boxes["created_by"].astype(str) == st.session_state.get("username", "")) & (boxes["status"].astype(str) == "รออนุมัติ")]
        if len(mine):
            st.dataframe(mine[["box_id", "store_name", "province_zone", "subzone", "contact", "google_maps_url", "status", "updated_at"]], use_container_width=True, hide_index=True)
        else:
            st.info("ยังไม่มีรายการรออนุมัติของคุณ")
        return

    st.markdown("<div class='warning-box'>ผู้ดูแลระบบใช้หน้านี้สำหรับวางกล่องใหม่ อนุมัติคำขอ และจัดการ เพิ่ม / แก้ไข / ลบ รายการกล่อง</div>", unsafe_allow_html=True)
    tab_add, tab_pending, tab_edit, tab_delete = st.tabs(["➕ วางกล่องใหม่", "✅ อนุมัติคำขอ", "✏️ แก้ไขข้อมูลกล่อง", "🗑️ ลบกล่อง"])
    with tab_add:
        new = box_form(default={"box_id": f"BOX-{len(boxes)+1:04d}", "status": "ใช้งาน"}, form_key="add_box_form")
        if new:
            if str(new["box_id"]) in boxes["box_id"].astype(str).tolist():
                st.error("เลขกล่องนี้มีอยู่แล้ว กรุณาใช้เลขกล่องอื่น")
            else:
                save_boxes(pd.concat([boxes, pd.DataFrame([new])], ignore_index=True))
                green_success("เพิ่มกล่องใหม่เรียบร้อยแล้ว")
    with tab_pending:
        pending = boxes[boxes["status"].astype(str) == "รออนุมัติ"].copy()
        st.subheader(f"คำขอวางกล่องใหม่รออนุมัติ {len(pending):,} รายการ")
        if len(pending) == 0:
            st.success("ไม่มีคำขอรออนุมัติ")
        else:
            for _, row in pending.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{row['box_id']} — {row['store_name']}**  \nโซน: {row.get('province_zone','')} | พื้นที่: {row.get('subzone','')} | ผู้ส่งคำขอ: {row.get('created_by','')}")
                    st.write(f"ติดต่อ: {row.get('contact','')} | หมายเหตุ: {row.get('note','')}")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.link_button("🧭 เปิด Google Maps", google_maps_link(row))
                    with c2:
                        if st.button("✅ อนุมัติ", key=f"approve_{row['box_id']}"):
                            b = boxes.copy()
                            idx = b.index[b["box_id"].astype(str) == str(row["box_id"])]
                            if len(idx):
                                b.loc[idx[0], "status"] = "ใช้งาน"
                                b.loc[idx[0], "updated_by"] = st.session_state.get("username", "")
                                b.loc[idx[0], "updated_at"] = datetime.now().isoformat(timespec="seconds")
                                save_boxes(b)
                                green_success("อนุมัติคำขอเรียบร้อยแล้ว")
                                st.rerun()
                    with c3:
                        if st.button("↩️ ส่งกลับแก้ไข", key=f"return_{row['box_id']}"):
                            b = boxes.copy()
                            idx = b.index[b["box_id"].astype(str) == str(row["box_id"])]
                            if len(idx):
                                b.loc[idx[0], "status"] = "รอแก้ไข"
                                b.loc[idx[0], "updated_by"] = st.session_state.get("username", "")
                                b.loc[idx[0], "updated_at"] = datetime.now().isoformat(timespec="seconds")
                                save_boxes(b)
                                green_success("ส่งกลับให้แก้ไขเรียบร้อยแล้ว")
                                st.rerun()
                    show_saved_photo(row.get("box_photo_url", ""), "รูปจุดวางที่ส่งคำขอ")
    with tab_edit:
        row = select_box_widget(boxes, "เลือกกล่องที่ต้องการแก้ไข", key="edit_select")
        if row is not None:
            edited = box_form(default=row.to_dict(), form_key="edit_box_form")
            if edited:
                b = boxes.copy()
                idx = b.index[b["box_id"].astype(str) == str(row["box_id"])]
                if len(idx):
                    for col, val in edited.items():
                        b.loc[idx[0], col] = val
                    save_boxes(b)
                    green_success("แก้ไขข้อมูลกล่องเรียบร้อยแล้ว")
    with tab_delete:
        row = select_box_widget(boxes, "เลือกกล่องที่ต้องการลบ", key="delete_select")
        if row is not None:
            st.warning(f"กำลังจะลบ: {row['box_id']} — {row['store_name']}")
            confirm = st.checkbox("ยืนยันว่าต้องการลบรายการนี้")
            if st.button("🗑️ ลบกล่องนี้", disabled=not confirm):
                b = boxes[boxes["box_id"].astype(str) != str(row["box_id"])].copy()
                save_boxes(b)
                green_success("ลบข้อมูลกล่องเรียบร้อยแล้ว")


def page_import(boxes):
    st.title("📥 อัปโหลดรายการกล่องตามแพทเทิร์น Excel")
    st.markdown("""
    ใช้สำหรับอัปโหลด **รายการทั้งหมด** หรือ **บางรายการ** ได้ โดยต้องมีหัวคอลัมน์ตามแพทเทิร์น
    """)
    template = pd.DataFrame(columns=REQUIRED_IMPORT_COLUMNS)
    sample = pd.DataFrame([{
        "box_id": "BOX-001", "store_name": "ชื่อร้านตัวอย่าง", "province_zone": "นครราชสีมา",
        "subzone": "ในเมือง", "contact": "0812345678", "landmark": "ใกล้ตลาด/หน้าร้าน",
        "google_maps_url": "https://maps.google.com/?q=14.97,102.10", "cycle_months": 3,
        "status": "ใช้งาน", "last_collect_date": date.today().isoformat(),
        "next_due_date": (pd.Timestamp(date.today()) + pd.DateOffset(months=3)).date().isoformat(), "note": ""
    }])
    st.download_button("⬇️ ดาวน์โหลดแพทเทิร์น Excel เปล่า", excel_bytes(template, "template"), "box_upload_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ ดาวน์โหลดตัวอย่าง Excel", excel_bytes(sample, "sample"), "box_upload_sample.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded = st.file_uploader("เลือกไฟล์ Excel .xlsx รายการกล่อง", type=["xlsx"])
    mode = st.radio("รูปแบบการอัปโหลด", ["เพิ่ม/อัปเดตบางรายการ", "แทนที่รายการทั้งหมด"], horizontal=True)
    if uploaded is not None:
        try:
            df = pd.read_excel(uploaded, dtype=str, engine="openpyxl").fillna("")
        except Exception as e:
            st.error(f"อ่านไฟล์ Excel ไม่สำเร็จ: {e}")
            return
        missing_cols = [c for c in REQUIRED_IMPORT_COLUMNS if c not in df.columns]
        if missing_cols:
            st.error("ไฟล์ยังขาดคอลัมน์: " + ", ".join(missing_cols))
            return
        preview = df[REQUIRED_IMPORT_COLUMNS].copy()
        st.subheader("ตัวอย่างข้อมูลที่จะนำเข้า")
        st.dataframe(preview.head(50), use_container_width=True, hide_index=True)
        if st.button("✅ ยืนยันนำเข้าข้อมูล"):
            now = datetime.now().isoformat(timespec="seconds")
            import_rows = []
            for _, r in preview.iterrows():
                new = {col: "" for col in BOX_COLUMNS}
                for c in REQUIRED_IMPORT_COLUMNS:
                    new[c] = r.get(c, "")
                new["cycle_months"] = int(pd.to_numeric(new.get("cycle_months", 3), errors="coerce") if str(new.get("cycle_months", "")).strip() else 3)
                if not new["status"]:
                    new["status"] = "ใช้งาน"
                new["created_by"] = st.session_state.get("username", "")
                new["updated_by"] = st.session_state.get("username", "")
                new["updated_at"] = now
                import_rows.append(new)
            incoming = normalize_boxes(pd.DataFrame(import_rows))
            if mode == "แทนที่รายการทั้งหมด":
                save_boxes(incoming)
                green_success(f"นำเข้าและแทนที่รายการทั้งหมดแล้ว จำนวน {len(incoming):,} รายการ")
            else:
                b = boxes.copy()
                for _, r in incoming.iterrows():
                    idx = b.index[b["box_id"].astype(str) == str(r["box_id"])]
                    if len(idx):
                        for col in BOX_COLUMNS:
                            if str(r.get(col, "")) != "":
                                b.loc[idx[0], col] = r[col]
                    else:
                        b = pd.concat([b, pd.DataFrame([r])], ignore_index=True)
                save_boxes(b)
                green_success(f"เพิ่ม/อัปเดตรายการแล้ว จำนวน {len(incoming):,} รายการ")


def page_route(boxes):
    st.title("🚗 จัดโซนการเดินทางวันนี้")
    st.caption("แสดงรายการตามจังหวัด/โซน และเปิดลิงก์ Google Maps ทีละจุด เหมาะกับการวางแผนเดินทางผ่านมือถือ")
    df = filters(boxes, "route")
    due_labels = df.apply(due_status, axis=1, result_type="expand") if len(df) else pd.DataFrame()
    if len(df):
        df = df.copy(); df["สถานะครบกำหนด"] = due_labels[0]
    sort_by = st.selectbox("เรียงลำดับ", ["จังหวัด/โซน", "วันครบกำหนด", "เลขกล่อง"])
    if sort_by == "วันครบกำหนด":
        df = df.sort_values("next_due_date")
    elif sort_by == "เลขกล่อง":
        df = df.sort_values("box_id")
    else:
        df = df.sort_values(["province_zone", "subzone", "box_id"])
    st.dataframe(df[["box_id", "store_name", "province_zone", "subzone", "google_maps_url", "next_due_date", "สถานะครบกำหนด", "status"]], use_container_width=True, hide_index=True)
    st.subheader("เปิดนำทาง")
    for _, row in df.head(50).iterrows():
        with st.container(border=True):
            st.markdown(f"**{row['box_id']} — {row['store_name']}**  \n{row.get('province_zone','')} / {row.get('subzone','')}")
            st.link_button("🧭 เปิด Google Maps", google_maps_link(row))


def page_reports(boxes, collections, missing, users):
    st.title("📑 รายงานและดาวน์โหลดข้อมูล")
    tab1, tab2, tab3, tab4 = st.tabs(["ประวัติการเก็บ", "แจ้งกล่องหาย", "บัญชีผู้ใช้", "ดาวน์โหลดทั้งหมด"])
    with tab1:
        st.dataframe(collections.sort_values("created_at", ascending=False) if len(collections) else collections, use_container_width=True, hide_index=True)
        st.download_button("⬇️ ดาวน์โหลดประวัติการเก็บ", excel_bytes(collections, "collections"), "collections.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab2:
        st.dataframe(missing.sort_values("created_at", ascending=False) if len(missing) else missing, use_container_width=True, hide_index=True)
        st.download_button("⬇️ ดาวน์โหลดแจ้งกล่องหาย", excel_bytes(missing, "missing_reports"), "missing_reports.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab3:
        view_users = users.drop(columns=["password"], errors="ignore")
        st.dataframe(view_users, use_container_width=True, hide_index=True)
    with tab4:
        st.download_button("⬇️ ดาวน์โหลดรายการกล่องทั้งหมด", excel_bytes(boxes, "boxes"), "boxes.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("⬇️ ดาวน์โหลดประวัติการเก็บทั้งหมด", excel_bytes(collections, "collections"), "collections.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("⬇️ ดาวน์โหลดรายงานกล่องหายทั้งหมด", excel_bytes(missing, "missing_reports"), "missing_reports.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_users(users):
    st.title("👥 บัญชีผู้ใช้ / เจ้าหน้าที่")
    st.markdown("<div class='warning-box'>หน้านี้ใช้สำหรับผู้ดูแลระบบในการเพิ่ม แก้ไข ลบ หรือปิดใช้งานบัญชีเจ้าหน้าที่ เพื่อให้ตรวจสอบย้อนหลังได้ว่าใครเป็นผู้บันทึกข้อมูล</div>", unsafe_allow_html=True)
    if st.session_state.get("role") != "admin":
        st.warning("เมนูนี้สำหรับผู้ดูแลระบบเท่านั้น")
        return

    tab_add, tab_edit, tab_delete, tab_download = st.tabs(["➕ เพิ่มบัญชี", "✏️ แก้ไขบัญชี", "🗑️ ลบ/ปิดบัญชี", "⬇️ ดาวน์โหลดบัญชี"])

    with tab_add:
        st.subheader("เพิ่มบัญชีเจ้าหน้าที่")
        with st.form("add_user"):
            username = st.text_input("ชื่อผู้ใช้", placeholder="เช่น somchai")
            password = st.text_input("รหัสผ่าน", type="password")
            display_name = st.text_input("ชื่อ-สกุลที่แสดงในระบบ", placeholder="เช่น นายสมชาย ใจดี")
            role = st.selectbox("สิทธิ์การใช้งาน", ["user", "admin"], help="user = เจ้าหน้าที่บันทึกข้อมูล, admin = ผู้ดูแลระบบ")
            status = st.selectbox("สถานะบัญชี", ["ใช้งาน", "ปิดใช้งาน"])
            submit = st.form_submit_button("💾 สร้างบัญชี")
        if submit:
            if not username or not password:
                st.error("กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
            elif username in users["username"].astype(str).tolist():
                st.error("ชื่อผู้ใช้นี้มีอยู่แล้ว")
            else:
                new = {
                    "username": username.strip(), "password": password, "display_name": display_name.strip() or username.strip(),
                    "role": role, "status": status, "created_at": datetime.now().isoformat(timespec="seconds")
                }
                save_users(pd.concat([users, pd.DataFrame([new])], ignore_index=True))
                green_success("สร้างบัญชีผู้ใช้เรียบร้อยแล้ว")
                st.rerun()

    with tab_edit:
        st.subheader("แก้ไขบัญชีเจ้าหน้าที่")
        if len(users) == 0:
            st.info("ยังไม่มีบัญชีผู้ใช้")
        else:
            st.dataframe(users.drop(columns=["password"], errors="ignore"), use_container_width=True, hide_index=True)
            target = st.selectbox("เลือกบัญชีที่ต้องการแก้ไข", users["username"].astype(str).tolist(), key="edit_user_target")
            row = users[users["username"].astype(str) == target].iloc[0]
            with st.form("edit_user"):
                new_username = st.text_input("ชื่อผู้ใช้", value=str(row.get("username", "")), help="ถ้าเปลี่ยนชื่อผู้ใช้ ระบบจะอัปเดตบัญชีเดิมเป็นชื่อใหม่")
                password = st.text_input("รหัสผ่าน", value=str(row.get("password", "")), type="password")
                display_name = st.text_input("ชื่อ-สกุลที่แสดงในระบบ", value=str(row.get("display_name", "")))
                role_options = ["user", "admin"]
                role = st.selectbox("สิทธิ์การใช้งาน", role_options, index=role_options.index(row.get("role", "user")) if row.get("role", "user") in role_options else 0)
                status_options = ["ใช้งาน", "ปิดใช้งาน"]
                status = st.selectbox("สถานะบัญชี", status_options, index=status_options.index(row.get("status", "ใช้งาน")) if row.get("status", "ใช้งาน") in status_options else 0)
                submit = st.form_submit_button("💾 บันทึกการแก้ไขบัญชี")
            if submit:
                if not new_username or not password:
                    st.error("กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
                elif new_username != target and new_username in users["username"].astype(str).tolist():
                    st.error("ชื่อผู้ใช้นี้มีอยู่แล้ว")
                else:
                    u = users.copy()
                    idx = u.index[u["username"].astype(str) == target][0]
                    u.loc[idx, "username"] = new_username.strip()
                    u.loc[idx, "password"] = password
                    u.loc[idx, "display_name"] = display_name.strip() or new_username.strip()
                    u.loc[idx, "role"] = role
                    u.loc[idx, "status"] = status
                    if target == st.session_state.get("username"):
                        st.session_state.username = new_username.strip()
                        st.session_state.display_name = display_name.strip() or new_username.strip()
                        st.session_state.role = role
                    save_users(u)
                    green_success("แก้ไขบัญชีเรียบร้อยแล้ว")
                    st.rerun()

    with tab_delete:
        st.subheader("ลบหรือปิดใช้งานบัญชี")
        st.markdown("<div class='warning-box'>แนะนำให้ใช้ ‘ปิดใช้งาน’ แทนการลบ เพื่อเก็บประวัติว่าเจ้าหน้าที่คนใดเคยบันทึกข้อมูล</div>", unsafe_allow_html=True)
        st.dataframe(users.drop(columns=["password"], errors="ignore"), use_container_width=True, hide_index=True)
        if len(users) == 0:
            return
        target = st.selectbox("เลือกบัญชี", users["username"].astype(str).tolist(), key="delete_user_target")
        action = st.radio("เลือกวิธีจัดการ", ["ปิดใช้งานบัญชี", "ลบบัญชีถาวร"], horizontal=True)
        confirm = st.checkbox(f"ยืนยันการดำเนินการกับบัญชี {target}")
        if st.button("ดำเนินการ", type="primary"):
            if not confirm:
                st.error("กรุณาติ๊กยืนยันก่อนดำเนินการ")
            elif target == st.session_state.get("username") and action == "ลบบัญชีถาวร":
                st.error("ไม่สามารถลบบัญชีที่กำลังเข้าสู่ระบบอยู่ได้")
            else:
                u = users.copy()
                if action == "ปิดใช้งานบัญชี":
                    idx = u.index[u["username"].astype(str) == target][0]
                    u.loc[idx, "status"] = "ปิดใช้งาน"
                    save_users(u)
                    green_success("ปิดใช้งานบัญชีเรียบร้อยแล้ว")
                    st.rerun()
                else:
                    # ป้องกันการลบ admin คนสุดท้าย
                    remaining = u[u["username"].astype(str) != target]
                    active_admins = remaining[(remaining["role"].astype(str) == "admin") & (remaining["status"].astype(str) == "ใช้งาน")]
                    target_role = u.loc[u["username"].astype(str) == target, "role"].iloc[0]
                    if target_role == "admin" and len(active_admins) == 0:
                        st.error("ไม่สามารถลบผู้ดูแลระบบคนสุดท้ายได้")
                    else:
                        save_users(remaining)
                        green_success("ลบบัญชีเรียบร้อยแล้ว")
                        st.rerun()

    with tab_download:
        st.subheader("ดาวน์โหลดข้อมูลบัญชี")
        safe_users = users.drop(columns=["password"], errors="ignore")
        st.dataframe(safe_users, use_container_width=True, hide_index=True)
        st.download_button("⬇️ ดาวน์โหลดบัญชีผู้ใช้แบบไม่แสดงรหัสผ่าน", excel_bytes(safe_users, "users"), "users_no_password.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    if not login():
        return
    boxes = load_boxes_cached()
    collections = load_collections_cached()
    missing = load_missing_cached()
    users = load_users_cached()
    with st.sidebar:
        st.title("📦 Donation Box")
        st.caption(f"ผู้ใช้: {st.session_state.get('display_name') or st.session_state.username}")
        st.caption(f"สิทธิ์: {st.session_state.get('role', 'user')}")
        if st.button("ออกจากระบบ"):
            st.session_state.auth = False
            st.session_state.username = ""
            st.session_state.display_name = ""
            st.session_state.role = "user"
            st.rerun()
        st.divider()
        if st.session_state.get("role") == "admin":
            menu_items = [
                "Dashboard", "รายการกล่อง", "แผนที่ GPS", "บันทึกการเก็บ", "กล่องหาย/คืน",
                "วางกล่องใหม่", "อัปโหลดรายการกล่อง", "จัดโซนการเดินทาง", "รายงาน/ดาวน์โหลด", "บัญชีผู้ใช้"
            ]
        else:
            menu_items = [
                "Dashboard", "แผนที่ GPS", "บันทึกการเก็บ", "วางกล่องใหม่", "กล่องหาย/คืน", "จัดโซนการเดินทาง"
            ]
        menu = st.radio("เมนู", menu_items)
        st.divider()
        st.caption("โหมดข้อมูล")
        st.write("Google Sheets + Drive" if has_google_config() else "Excel .xlsx Local")
        st.caption("เก็บรูป/สำรอง Drive")
        st.write("เปิดใช้งาน" if has_drive_config() else "ยังไม่เชื่อม Drive")

    if menu == "Dashboard":
        page_dashboard(boxes, collections, missing)
    elif menu == "รายการกล่อง":
        page_boxes(boxes, collections)
    elif menu == "แผนที่ GPS":
        page_map(boxes, collections)
    elif menu == "บันทึกการเก็บ":
        page_collect(boxes, collections)
    elif menu == "กล่องหาย/คืน":
        page_missing(boxes, missing)
    elif menu == "วางกล่องใหม่":
        page_manage_boxes(boxes)
    elif menu == "อัปโหลดรายการกล่อง":
        page_import(boxes)
    elif menu == "จัดโซนการเดินทาง":
        page_route(boxes)
    elif menu == "รายงาน/ดาวน์โหลด":
        page_reports(boxes, collections, missing, users)
    elif menu == "บัญชีผู้ใช้":
        page_users(users)


if __name__ == "__main__":
    main()
