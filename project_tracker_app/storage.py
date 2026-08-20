"""
資料儲存層。

- 本機執行：存到 data/tasks.json
- 部署到 Streamlit Cloud：只要設定好 secrets 就改存 Google Sheets

Streamlit Community Cloud 沒有永久磁碟，容器每次重建（睡眠喚醒、重新部署、
維護、Reboot）都會從 repo 重拉一份，所以寫在硬碟上的檔案一定會不見。
要讓雲端版的資料留得住，就必須存到 App 外面，這裡用 Google Sheets。
"""

import json
import time
from datetime import datetime
from pathlib import Path

import streamlit as st

LOCAL_FILE = Path(__file__).parent / "data" / "tasks.json"
TASK_SHEET = "任務"
CONFIG_SHEET = "設定"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ------------------------------------------------------------
# 後端判斷
# ------------------------------------------------------------
def _session_get(key, default=None):
    # 在 Streamlit runtime 外（測試 / 匯入）讀 session_state 會丟例外
    try:
        return st.session_state.get(key, default)
    except Exception:
        return default


def _session_set(key, value):
    try:
        st.session_state[key] = value
    except Exception:
        pass


def _has_secret(key):
    # 沒有 secrets.toml 時讀 st.secrets 會直接丟例外，包起來
    try:
        return key in st.secrets
    except Exception:
        return False


def backend():
    """目前實際使用的儲存後端。"""
    if _has_secret("gcp_service_account") and _has_secret("gsheets"):
        return "gsheets"
    return "local"


def describe():
    """給側邊欄顯示用的說明文字。"""
    if backend() == "gsheets":
        return "Google Sheets", "資料存在雲端試算表，重新部署或睡眠都不會不見"
    return "本機檔案", str(LOCAL_FILE)


# ------------------------------------------------------------
# Google Sheets
# ------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _spreadsheet():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    client = gspread.authorize(creds)

    ref = str(st.secrets["gsheets"]["spreadsheet"]).strip()
    if ref.startswith("http"):
        return client.open_by_url(ref)
    if " " not in ref and len(ref) >= 30:
        return client.open_by_key(ref)
    return client.open(ref)


def _worksheet(sh, title, rows=200, cols=20):
    try:
        return sh.worksheet(title)
    except Exception:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def _retry(func, attempts=3):
    """Google API 偶爾會回 429（每分鐘寫入上限），退避後重試。"""
    delay = 1.5
    last = None
    for _ in range(attempts):
        try:
            return func()
        except Exception as exc:
            last = exc
            text = str(exc).lower()
            if "429" in text or "quota" in text or "rate limit" in text:
                time.sleep(delay)
                delay *= 2
                continue
            raise
    raise last


def _load_gsheets(columns):
    sh = _spreadsheet()
    ws = _worksheet(sh, TASK_SHEET)
    # numericise_ignore=all：不要把 WBS "1.10" 這種字串當成數字 1.1
    records = _retry(lambda: ws.get_all_records(numericise_ignore=["all"]))

    project_name = None
    try:
        rows = _retry(lambda: _worksheet(sh, CONFIG_SHEET, rows=10, cols=2).get_all_values())
        for row in rows:
            if len(row) >= 2 and row[0] == "project_name":
                project_name = row[1]
                break
    except Exception:
        pass

    return records, project_name


def _save_gsheets(records, project_name, columns):
    sh = _spreadsheet()
    ws = _worksheet(sh, TASK_SHEET, rows=max(len(records) + 20, 200), cols=len(columns))

    values = [list(columns)] + [
        [str(rec.get(col, "")) for col in columns] for rec in records
    ]
    _retry(ws.clear)
    _retry(lambda: ws.update(values=values, range_name="A1"))

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 專案名稱沒變就不用多打一次 API（每分鐘寫入次數有上限）
    if _session_get("_saved_project_name") != project_name:
        cfg = _worksheet(sh, CONFIG_SHEET, rows=10, cols=2)
        _retry(cfg.clear)
        _retry(
            lambda: cfg.update(
                values=[
                    ["key", "value"],
                    ["project_name", project_name],
                    ["saved_at", stamp],
                ],
                range_name="A1",
            )
        )
        _session_set("_saved_project_name", project_name)
    return stamp


# ------------------------------------------------------------
# 本機檔案
# ------------------------------------------------------------
def _load_local():
    if not LOCAL_FILE.exists():
        return [], None
    payload = json.loads(LOCAL_FILE.read_text(encoding="utf-8"))
    return payload.get("tasks", []), payload.get("project_name")


def _save_local(records, project_name):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    LOCAL_FILE.parent.mkdir(parents=True, exist_ok=True)
    payload = {"project_name": project_name, "saved_at": stamp, "tasks": records}
    LOCAL_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return stamp


# ------------------------------------------------------------
# 對外介面
# ------------------------------------------------------------
def load(columns):
    """回傳 (records, project_name)。沒有資料時 records 是空 list。"""
    if backend() == "gsheets":
        return _load_gsheets(columns)
    return _load_local()


def save(records, project_name, columns):
    """存檔，失敗時直接丟例外讓上層顯示錯誤。回傳存檔時間字串。"""
    if backend() == "gsheets":
        return _save_gsheets(records, project_name, columns)
    return _save_local(records, project_name)


def clear_cache():
    """換過 secrets 或想重新連線時用。"""
    _spreadsheet.clear()
