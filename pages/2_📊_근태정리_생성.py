"""근태현황 생성 — 파일 업로드 → 처리 → 다운로드"""
import streamlit as st
import pandas as pd
import json
import sys
import calendar
from pathlib import Path
from datetime import datetime, date, timedelta

# Add lib to path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
from lib.attendance_builder import build_excel

st.set_page_config(page_title="근태정리 생성", page_icon="📊", layout="wide")

from lib.auth import require_login
require_login()

DATA_DIR = ROOT / "data"
UPLOAD_DIR = ROOT / "uploads"
OUTPUT_DIR = ROOT / "outputs"
HISTORY_DIR = DATA_DIR / "history"
PEOPLE_FILE = DATA_DIR / "people.json"
GSHEET_CONFIG = DATA_DIR / "gsheet_config.json"
CREDS_FILE = DATA_DIR / "gcp_service_account.json"

for d in [UPLOAD_DIR, OUTPUT_DIR, HISTORY_DIR]:
    d.mkdir(exist_ok=True, parents=True)


def load_gsheet_config():
    if GSHEET_CONFIG.exists():
        with open(GSHEET_CONFIG, encoding='utf-8') as f:
            return json.load(f)
    return {"spreadsheet_id": "1OHDcoW9I9wTZg3Zkbp3b8IwI0CXlpwRFwiLu6H6iZQs"}


def save_gsheet_config(cfg):
    with open(GSHEET_CONFIG, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

st.title("📊 근태정리 생성")
st.caption("원본 파일 3종을 업로드하면 5시트 구조의 정리 엑셀이 자동 생성됩니다.")

# Load config (로컬 파일 또는 구글시트 — config_store가 자동 판단)
from lib.config_store import load_config
config = load_config()

# Build maps
active_people = [p for p in config['people'] if p.get('status', '활성') == '활성']
targets = [p['name'] for p in active_people]
sojung_map = {p['name']: p.get('sojung_h', 40) for p in active_people}
contract_map = {p['name']: p.get('contract_type', '시급제') for p in active_people}
dept_map = {p['name']: p.get('department', '') for p in active_people}
erp_switchers = {e['name']: e for e in config.get('erp_switchers', [])}

# ============================================================
# Step 1: 인원 명단 확인
# ============================================================
with st.expander(f"👥 Step 1. 대상 인원 확인 (현재 활성 {len(targets)}명)", expanded=False):
    st.dataframe(
        pd.DataFrame(active_people)[['name', 'sojung_h', 'contract_type', 'department']],
        hide_index=True,
        use_container_width=True
    )
    st.info("👈 좌측 메뉴 **인원 관리**에서 명단 수정 가능")

# ============================================================
# Step 2: 파일 업로드
# ============================================================
st.subheader("📁 Step 2. 원본 파일 업로드")

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**① 원티드 주차 집계**")
    st.caption("이삼오구_YYYY-MM-DD-...xlsx (16컬럼)")
    file1 = st.file_uploader("원티드 주차", type=['xlsx'], key='f1', label_visibility='collapsed')
with col2:
    st.markdown("**② 원티드 일자 상세**")
    st.caption("이삼오구_..._EI*.xlsx (87컬럼)")
    file2 = st.file_uploader("원티드 일자", type=['xlsx'], key='f2', label_visibility='collapsed')
with col3:
    st.markdown("**③ FLEX 일자 리포트** (선택)")
    st.caption("전체_일별_근태기록리포트_*.xlsx")
    file3 = st.file_uploader("FLEX (없으면 생략)", type=['xlsx'], key='f3', label_visibility='collapsed')

# ============================================================
# Step 3: 기간 선택
# ============================================================
st.subheader("📅 Step 3. 처리 월 (달력 한 달 전체)")

col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    sel_year = int(st.number_input("연도", min_value=2024, max_value=2035, value=2026, step=1))
with col2:
    sel_month = int(st.selectbox("월", list(range(1, 13)), index=4,
                                 format_func=lambda m: f"{m}월"))
with col3:
    _ndays = calendar.monthrange(sel_year, sel_month)[1]
    st.metric("처리 기간 (자동)",
              f"{sel_year}-{sel_month:02d}-01 ~ {sel_year}-{sel_month:02d}-{_ndays:02d}")

# ============================================================
# Step 4: 생성 버튼
# ============================================================
st.subheader("⚙️ Step 4. 생성")

if st.button("🚀 근태정리 생성", type="primary", use_container_width=True):
    if not file1 or not file2:
        st.error("❌ 원티드 주차 + 일자 상세 파일은 필수입니다.")
    else:
        with st.spinner("처리 중... (10~30초 소요)"):
            try:
                # 파일 임시 저장
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                f1_path = UPLOAD_DIR / f'wonted_weekly_{ts}.xlsx'
                f2_path = UPLOAD_DIR / f'wonted_daily_{ts}.xlsx'
                f3_path = UPLOAD_DIR / f'flex_{ts}.xlsx' if file3 else None

                with open(f1_path, 'wb') as f: f.write(file1.read())
                with open(f2_path, 'wb') as f: f.write(file2.read())
                if file3:
                    with open(f3_path, 'wb') as f: f.write(file3.read())

                # 데이터 로드
                df1 = pd.read_excel(f1_path, header=0)
                df2 = pd.read_excel(f2_path, header=0)
                df3 = None
                if f3_path:
                    df3 = pd.read_excel(f3_path, header=0)
                    df3['날짜'] = pd.to_datetime(df3['날짜'], errors='coerce')

                # 빌드
                output_name = f'근태현황_{sel_year}-{sel_month:02d}.xlsx'
                output_path = OUTPUT_DIR / output_name

                result = build_excel(
                    targets=targets,
                    sojung_map=sojung_map,
                    erp_switchers=erp_switchers,
                    contract_map=contract_map,
                    df1=df1, df2=df2, df3=df3,
                    year=sel_year, month=sel_month,
                    output_path=output_path,
                    dept_map=dept_map,
                    erp_source_file=file3.name if file3 else '',
                )

                # 이력 저장
                history_entry = {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'period': result.get('period_label', f'{sel_year}-{sel_month:02d}'),
                    'num_people': result['num_people'],
                    'output_file': output_name,
                    'status': '✅ 성공' if not result['issues'] else f'⚠ 경고 {len(result["warnings"])}건',
                }
                hist_file = HISTORY_DIR / f'run_{ts}.json'
                with open(hist_file, 'w', encoding='utf-8') as f:
                    json.dump(history_entry, f, ensure_ascii=False, indent=2)

                # 세션에 결과 저장 (다운로드/미리보기/구글시트 공유가 rerun 후에도 사용)
                st.session_state['last_result'] = {
                    'summary': result['summary'],
                    'weeks': result['weeks'],
                    'week_ranges': result['week_ranges'],
                    'period_month': result['period_month'],
                    'output_name': output_name,
                    'output_path': str(output_path),
                    'num_people': result['num_people'],
                }

                st.success(f"✅ 근태현황 생성 완료! ({result['num_people']}명 · {result.get('period_label','')})")

                # 검증 결과
                if result['issues']:
                    st.error("🚨 정합성 이슈:")
                    for issue in result['issues']:
                        st.write(f"- {issue}")

                if result['warnings']:
                    with st.expander(f"⚠️ 경고 {len(result['warnings'])}건 (한도 위반 등)", expanded=True):
                        for w in result['warnings']:
                            st.warning(w)
                else:
                    st.info("✅ 정합성 검증 통과 - 이상치 없음")

                # 다운로드
                with open(output_path, 'rb') as f:
                    st.download_button(
                        "📥 결과 엑셀 다운로드",
                        f.read(),
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )

                st.caption(f"📁 자동 저장: `outputs/{output_name}`")

            except Exception as e:
                st.error(f"❌ 처리 실패: {e}")
                import traceback
                with st.expander("상세 오류"):
                    st.code(traceback.format_exc())

# ============================================================
# Step 5: 결과 미리보기 + 구글시트 공유 (생성 후 표시)
# ============================================================
lr = st.session_state.get('last_result')
if lr:
    st.markdown("---")
    st.subheader("📄 결과 미리보기")
    prev = pd.DataFrame(lr['summary'])
    # 화면 표시용: contract 숨기고 컬럼 정리
    show_cols = ['이름', '소정근무시간(주)'] + lr['weeks'] + ['초과근무(시간)', '야간근무(시간)', '데이터 출처']
    prev_show = prev[show_cols].copy()
    prev_show.insert(0, 'No', range(1, len(prev_show) + 1))

    def _hl(row):
        sj = row['소정근무시간(주)'] or 0
        styles = []
        for c in prev_show.columns:
            s = ''
            if c in lr['weeks'] and sj:
                if row[c] > sj * 1.1: s = 'background-color:#FFC7CE'
                elif row[c] > sj: s = 'background-color:#FFEB9C'
            elif c == '초과근무(시간)' and sj and row[c] >= sj * 4 * 0.4:
                s = 'background-color:#FFC7CE'
            elif c == '야간근무(시간)' and sj and row[c] >= sj * 4 * 0.4:
                s = 'background-color:#FFC7CE'
            styles.append(s)
        return styles

    st.dataframe(prev_show.style.apply(_hl, axis=1), hide_index=True, use_container_width=True)
    st.caption("🔴 소정 110%↑ / 🟡 소정 초과 · 초과·야간은 (소정×4) 40%↑ 강조")

    st.markdown("---")
    st.subheader("📗 구글시트 공유 (구성원용)")
    st.caption("연봉제 제외 비연봉제 인원만, 구성원 공유용 포맷으로 새 탭을 자동 생성합니다.")

    gcfg = load_gsheet_config()
    from lib.gsheet_sync import creds_available, get_service_account_email, push_month_tab

    if not creds_available():
        st.warning("⚙️ 아직 구글 서비스계정 키가 등록되지 않았습니다.")
        with st.expander("🔑 서비스계정 키(JSON) 등록 방법", expanded=True):
            st.markdown("""
            1. Google Cloud Console에서 **서비스 계정** 생성 → **Google Sheets API** 사용 설정
            2. 서비스 계정 → 키 → **JSON 키 추가**로 다운로드
            3. 아래에 그 JSON 파일을 업로드 (한 번만 하면 됨)
            4. 업로드 후 표시되는 **서비스계정 이메일을 대상 구글시트에 '편집자'로 공유**
            """)
        key_up = st.file_uploader("서비스계정 JSON 키 업로드", type=['json'], key='gcp_key')
        if key_up:
            CREDS_FILE.write_bytes(key_up.read())
            st.success("✅ 키 저장 완료. 아래에서 이어서 진행하세요.")
            st.rerun()
    else:
        try:
            sa_email = get_service_account_email()
            st.info(f"🔑 서비스계정: `{sa_email}`\n\n→ 이 이메일이 대상 시트에 **편집자**로 공유되어 있어야 합니다.")

            col1, col2 = st.columns([2, 1])
            with col1:
                sid = st.text_input("스프레드시트 ID", value=gcfg.get('spreadsheet_id', ''))
            with col2:
                tab_name = st.text_input("새 탭 이름", value=f"{lr['period_month']}월")
            overwrite = st.checkbox("같은 이름 탭이 있으면 덮어쓰기", value=False)

            cbtn1, cbtn2 = st.columns([1, 3])
            with cbtn1:
                go = st.button("📤 구글시트에 탭 추가", type="primary", use_container_width=True)
            with cbtn2:
                if st.button("🔁 서비스계정 키 교체", use_container_width=True):
                    CREDS_FILE.unlink(missing_ok=True)
                    st.rerun()

            if go:
                with st.spinner("구글시트에 기록 중..."):
                    try:
                        if sid != gcfg.get('spreadsheet_id'):
                            gcfg['spreadsheet_id'] = sid
                            save_gsheet_config(gcfg)
                        out = push_month_tab(
                            None, sid, tab_name,
                            lr['summary'], lr['weeks'], lr['week_ranges'],
                            overwrite=overwrite,
                        )
                        st.success(f"✅ '{out['tab_name']}' 탭 생성 완료 ({out['num_rows']}명 기록)")
                        st.markdown(f"🔗 [구글시트 열기]({out['url']})")
                    except Exception as e:
                        st.error(f"❌ 구글시트 기록 실패: {e}")
                        with st.expander("상세 오류"):
                            import traceback
                            st.code(traceback.format_exc())
        except Exception as e:
            st.error(f"구글시트 모듈 로드 실패: {e}")
