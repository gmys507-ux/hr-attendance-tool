"""
명단/설정 저장소 — 로컬 파일과 구글시트를 자동 전환.

- 클라우드(Streamlit Cloud): 파일이 영속되지 않으므로 구글시트에 저장
  · 설정 시트 ID가 지정되고(st.secrets['config_sheet_id'] 또는 app_config.json)
    서비스계정 키가 있으면 → 구글시트의 'config' 탭 A1 셀에 JSON으로 저장/로드
- 로컬/사내망: data/people.json 사용 (기존 동작)

⚠️ 명단(개인정보)은 이 코드에 두지 않는다.
   초기 명단은 구글시트(운영) 또는 data/people.json(로컬)에 있고,
   둘 다 비어 있을 때만 빈 DEFAULT_CONFIG로 시작한다.
"""
import json
from pathlib import Path

import streamlit as st

DATA_DIR = Path(__file__).parent.parent / "data"
LOCAL_FILE = DATA_DIR / "people.json"
BACKUP_DIR = DATA_DIR / "backup"
CONFIG_TAB = "config"

# 개인정보 없는 기본 골격 (실제 명단은 구글시트/로컬 people.json에 있음)
DEFAULT_CONFIG = {
    "people": [],
    "erp_switchers": [],
    "settings": {
        "default_sojung": 40,
        "sojung_basis": "주",
        "night_threshold_hour": 22,
        "limit_hours_per_week": 12,
    },
}


def _config_sheet_id():
    """설정 시트 ID 조회 (secrets 우선)"""
    try:
        if "config_sheet_id" in st.secrets:
            return str(st.secrets["config_sheet_id"])
    except Exception:
        pass
    cfg_path = DATA_DIR / "app_config.json"
    if cfg_path.exists():
        try:
            return json.loads(cfg_path.read_text(encoding="utf-8")).get("config_sheet_id")
        except Exception:
            return None
    return None


def _get_creds_info():
    """서비스계정 자격증명 (secrets dict 우선, 없으면 로컬 키파일 경로)"""
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"]), None
    except Exception:
        pass
    key_file = DATA_DIR / "gcp_service_account.json"
    if key_file.exists():
        return None, str(key_file)
    return None, None


def _use_gsheet():
    return bool(_config_sheet_id()) and any(_get_creds_info())


def _open_config_ws():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    info, path = _get_creds_info()
    if info:
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(path, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(_config_sheet_id())
    titles = [w.title for w in sh.worksheets()]
    if CONFIG_TAB in titles:
        return sh.worksheet(CONFIG_TAB)
    ws = sh.add_worksheet(title=CONFIG_TAB, rows=10, cols=2)
    try:
        ws.hide()  # 구성원 눈에 안 띄게 숨김 처리
    except Exception:
        pass
    return ws


def _seed_config():
    """초기 시드: 로컬 people.json이 있으면 그걸로, 없으면 빈 기본값."""
    if LOCAL_FILE.exists():
        try:
            return json.loads(LOCAL_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_CONFIG))


# ── 공개 API ──────────────────────────────────────────────────────────
def load_config():
    """현재 설정 로드. 구글시트 모드면 시트에서, 아니면 로컬에서. 비어 있으면 시드."""
    if _use_gsheet():
        try:
            ws = _open_config_ws()
            raw = ws.acell("A1").value
            if raw:
                return json.loads(raw)
            seed = _seed_config()       # 비어 있으면 시드
            save_config(seed)
            return seed
        except Exception:
            pass  # 실패 시 로컬 폴백
    return _load_local()


def save_config(cfg):
    """현재 설정 저장 (+로컬 백업). 구글시트 모드면 시트에 기록."""
    if _use_gsheet():
        try:
            ws = _open_config_ws()
            ws.update_acell("A1", json.dumps(cfg, ensure_ascii=False))
            return
        except Exception:
            pass  # 실패 시 로컬 폴백
    _save_local(cfg)


def storage_mode():
    """현재 저장 위치 안내용 문자열"""
    return "구글시트" if _use_gsheet() else "로컬파일"


def _load_local():
    if not LOCAL_FILE.exists():
        LOCAL_FILE.parent.mkdir(parents=True, exist_ok=True)
        LOCAL_FILE.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
    return json.loads(LOCAL_FILE.read_text(encoding="utf-8"))


def _save_local(cfg):
    from datetime import datetime
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    if LOCAL_FILE.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        (BACKUP_DIR / f"people_{ts}.json").write_text(
            LOCAL_FILE.read_text(encoding="utf-8"), encoding="utf-8")
    LOCAL_FILE.parent.mkdir(parents=True, exist_ok=True)
    LOCAL_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
