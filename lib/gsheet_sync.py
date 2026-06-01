"""
구글시트 연동 — 매월 결과를 구성원 공유용 구글시트에 새 탭으로 자동 기록.

인증: 서비스 계정(JSON 키). data/gcp_service_account.json 에 키를 두고,
대상 시트를 서비스계정 이메일에 '편집자'로 공유해야 동작.

구성원 공유용 포맷(기존 4월 탭과 동일): No·이름·소정근무시간(주)·4주차·초과·야간·출처
대상 인원: 연봉제 제외(비연봉제만).
"""
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
_LOCAL_KEY = Path(__file__).parent.parent / "data" / "gcp_service_account.json"


def _resolve_credentials(creds_path=None):
    """st.secrets['gcp_service_account'] 우선 → 로컬 키파일 폴백."""
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            return Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    except Exception:
        pass
    p = creds_path or _LOCAL_KEY
    if Path(p).exists():
        return Credentials.from_service_account_file(str(p), scopes=SCOPES)
    raise FileNotFoundError("서비스계정 키를 찾을 수 없습니다 (secrets 또는 data/gcp_service_account.json)")


def creds_available():
    try:
        _resolve_credentials()
        return True
    except Exception:
        return False

# 구성원 공유용 컬럼 순서 (계약종류 없음)
SHARE_COLUMNS = ['No', '이름', '소정근무시간(주)']  # + 주차들 + 뒤 3컬럼 (런타임 구성)
TAIL_COLUMNS = ['초과근무(시간)', '야간근무(시간)', '데이터 출처']


def get_service_account_email(creds_path=None):
    """서비스계정 이메일 추출 (secrets 우선, 공유 안내용)"""
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            return st.secrets["gcp_service_account"].get("client_email", "")
    except Exception:
        pass
    import json
    p = creds_path or _LOCAL_KEY
    if Path(p).exists():
        return json.loads(Path(p).read_text(encoding="utf-8")).get("client_email", "")
    return ""


def get_client(creds_path=None):
    return gspread.authorize(_resolve_credentials(creds_path))


def build_share_rows(summary, weeks, week_ranges, exclude_contracts=('연봉제',)):
    """summary(전체) → 구성원시트용 헤더+행렬. 연봉제 제외, No 재부여."""
    header1 = SHARE_COLUMNS + [w for w in weeks] + TAIL_COLUMNS
    header2 = ['', '', '(h)'] + [f'({r})' for r in week_ranges] + ['', '', '']
    rows = []
    no = 0
    for s in summary:
        if s.get('contract') in exclude_contracts:
            continue
        no += 1
        row = [no, s['이름'], s['소정근무시간(주)']]
        row += [s[w] for w in weeks]
        row += [s['초과근무(시간)'], s['야간근무(시간)'], s['데이터 출처']]
        rows.append(row)
    return header1, header2, rows


def push_month_tab(creds_path, spreadsheet_id, tab_name,
                   summary, weeks, week_ranges,
                   exclude_contracts=('연봉제',), overwrite=False):
    """구글시트에 새 탭(tab_name) 생성 후 데이터 기록.
    overwrite=False이고 동일 탭이 있으면 ValueError.
    반환: dict(tab_name, num_rows, url)
    """
    header1, header2, rows = build_share_rows(summary, weeks, week_ranges, exclude_contracts)
    values = [header1, header2] + rows
    ncols = len(header1)

    gc = get_client(creds_path)
    sh = gc.open_by_key(spreadsheet_id)
    existing = [ws.title for ws in sh.worksheets()]

    if tab_name in existing:
        if not overwrite:
            raise ValueError(f"'{tab_name}' 탭이 이미 존재합니다. 덮어쓰기를 선택하거나 다른 이름을 쓰세요.")
        ws = sh.worksheet(tab_name)
        ws.clear()
    else:
        ws = sh.add_worksheet(title=tab_name, rows=len(values) + 10, cols=ncols + 2)

    ws.update(values=values, range_name='A1')

    # 헤더 2줄 굵게 + 가운데 정렬
    last_col = gspread.utils.rowcol_to_a1(1, ncols).rstrip('1')
    ws.format(f'A1:{last_col}2', {
        'textFormat': {'bold': True},
        'horizontalAlignment': 'CENTER',
        'backgroundColor': {'red': 0.86, 'green': 0.92, 'blue': 0.97},
    })

    return {
        'tab_name': tab_name,
        'num_rows': len(rows),
        'url': f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit",
    }
