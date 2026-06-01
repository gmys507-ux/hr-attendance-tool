"""인원 관리 페이지 — 추가/삭제/수정/일괄 import/export"""
import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="인원 관리", page_icon="👥", layout="wide")

import sys as _sys
_sys.path.insert(0, str(Path(__file__).parent.parent))
from lib.auth import require_login
require_login()

DATA_DIR = Path(__file__).parent.parent / "data"
BACKUP_DIR = DATA_DIR / "backup"
BACKUP_DIR.mkdir(exist_ok=True)

from lib.config_store import load_config, save_config, storage_mode

st.title("👥 인원 관리")
st.caption("매월 신규 입사/퇴사/소정시간 변경 등을 여기서 관리합니다.")
st.caption(f"🗄 저장 위치: **{storage_mode()}**")

config = load_config()

tab1, tab2, tab3 = st.tabs(["📋 활성 인원", "🔁 FLEX 전환자", "📥 일괄 가져오기/내보내기"])

# ============================================================
# Tab 1: 활성 인원 관리
# ============================================================
with tab1:
    df = pd.DataFrame(config['people'])
    # Ensure all columns exist
    for col in ['name', 'sojung_h', 'eng_name', 'contract_type', 'department', 'status', 'sabun', 'note']:
        if col not in df.columns:
            df[col] = None

    # Reorder columns for display
    display_cols = ['name', 'sojung_h', 'contract_type', 'department', 'eng_name', 'sabun', 'status', 'note']
    df = df[display_cols]

    # 필터
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        search = st.text_input("🔍 이름 검색", placeholder="이름 입력...")
    with col2:
        contract_filter = st.selectbox("계약종류 필터", ["전체"] + sorted(df['contract_type'].dropna().unique().tolist()))
    with col3:
        status_filter = st.selectbox("재직 상태", ["활성", "전체", "퇴사"])

    # 필터 적용
    filtered = df.copy()
    if search:
        filtered = filtered[filtered['name'].str.contains(search, na=False)]
    if contract_filter != "전체":
        filtered = filtered[filtered['contract_type'] == contract_filter]
    if status_filter == "활성":
        filtered = filtered[filtered['status'].fillna('활성') == '활성']
    elif status_filter == "퇴사":
        filtered = filtered[filtered['status'] == '퇴사']

    st.caption(f"📊 표시 {len(filtered)}명 / 전체 {len(df)}명")

    # Editable table
    edited_df = st.data_editor(
        filtered,
        column_config={
            "name": st.column_config.TextColumn("이름", width="medium", required=True),
            "sojung_h": st.column_config.NumberColumn("소정(주h)", min_value=0, max_value=60, step=1, width="small"),
            "contract_type": st.column_config.SelectboxColumn(
                "계약종류",
                options=["연봉제", "시급제", "사업소득제", "기타"],
                width="small",
                required=True
            ),
            "department": st.column_config.TextColumn("부서", width="medium"),
            "eng_name": st.column_config.TextColumn("영문명", width="medium", help="시스템 영문명이 있는 경우"),
            "sabun": st.column_config.TextColumn("사번", width="small"),
            "status": st.column_config.SelectboxColumn(
                "재직상태",
                options=["활성", "퇴사", "휴직"],
                width="small",
                default="활성"
            ),
            "note": st.column_config.TextColumn("비고", width="large"),
        },
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="people_editor"
    )

    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("💾 변경사항 저장", type="primary", use_container_width=True):
            # Merge edited rows back to full config (only those visible were edited)
            edited_records = edited_df.to_dict('records')
            # Replace by name
            existing_by_name = {p['name']: p for p in config['people']}
            for r in edited_records:
                # Clean empty strings to None
                for k, v in list(r.items()):
                    if isinstance(v, str) and v.strip() == '':
                        r[k] = None
                existing_by_name[r['name']] = r
            config['people'] = list(existing_by_name.values())
            save_config(config)
            st.success(f"✅ 저장 완료. 백업도 자동 생성됨.")
            st.rerun()

    with col2:
        if st.button("↻ 초기화", use_container_width=True):
            st.rerun()

# ============================================================
# Tab 2: FLEX 전환자
# ============================================================
with tab2:
    st.caption("원티드스페이스 → FLEX로 전환한 인원의 전환 날짜를 관리합니다.")

    erp_df = pd.DataFrame(config.get('erp_switchers', []))
    if not erp_df.empty:
        # DateColumn 편집을 위해 문자열 → datetime 변환 (저장 시 다시 문자열로 환원)
        if 'switch_date' in erp_df.columns:
            erp_df['switch_date'] = pd.to_datetime(erp_df['switch_date'], errors='coerce')
        edited_erp = st.data_editor(
            erp_df,
            column_config={
                "name": st.column_config.TextColumn("이름", required=True),
                "switch_date": st.column_config.DateColumn("전환일", format="YYYY-MM-DD", required=True),
                "erp_name": st.column_config.TextColumn("FLEX 등록명", help="시스템 등록된 이름 (한글 또는 영문)"),
            },
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            key="erp_editor"
        )

        if st.button("💾 FLEX 전환자 저장", type="primary"):
            # Convert dates to string
            records = []
            for r in edited_erp.to_dict('records'):
                if isinstance(r.get('switch_date'), (pd.Timestamp, datetime)):
                    r['switch_date'] = r['switch_date'].strftime('%Y-%m-%d')
                records.append(r)
            config['erp_switchers'] = records
            save_config(config)
            st.success("✅ FLEX 전환자 정보 저장 완료")
            st.rerun()
    else:
        st.info("FLEX 전환자가 없습니다. 행을 추가해주세요.")

# ============================================================
# Tab 3: Excel 일괄 가져오기/내보내기
# ============================================================
with tab3:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📤 내보내기 (Excel 백업)")
        st.caption("전체 인원 + FLEX 전환자 정보를 Excel 파일로 다운로드")
        if st.button("📥 Excel 다운로드 준비", use_container_width=True):
            import io
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pd.DataFrame(config['people']).to_excel(writer, sheet_name='활성_인원', index=False)
                pd.DataFrame(config.get('erp_switchers', [])).to_excel(writer, sheet_name='FLEX_전환자', index=False)
            st.download_button(
                "💾 Excel 파일 다운로드",
                buffer.getvalue(),
                file_name=f"people_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    with col2:
        st.subheader("📥 가져오기 (Excel → 시스템)")
        st.caption("⚠ 기존 데이터를 덮어씁니다. 백업 후 사용 권장.")
        upload = st.file_uploader("Excel 파일 선택", type=['xlsx'])
        if upload:
            try:
                people_df = pd.read_excel(upload, sheet_name='활성_인원')
                erp_df = pd.read_excel(upload, sheet_name='FLEX_전환자')
                st.write(f"활성 인원: **{len(people_df)}명**, FLEX 전환자: **{len(erp_df)}명**")
                if st.button("🔄 시스템에 반영", type="primary"):
                    config['people'] = people_df.to_dict('records')
                    config['erp_switchers'] = erp_df.to_dict('records')
                    save_config(config)
                    st.success("✅ 가져오기 완료")
                    st.rerun()
            except Exception as e:
                st.error(f"읽기 실패: {e}")

# ============================================================
# 사이드바: 백업 이력
# ============================================================
st.sidebar.title("🗄 백업 이력")
backups = sorted(BACKUP_DIR.glob("people_*.json"), reverse=True)[:10]
if backups:
    st.sidebar.caption(f"최근 백업 {len(backups)}건")
    for b in backups:
        ts = b.stem.replace('people_', '')
        st.sidebar.text(ts)
else:
    st.sidebar.info("백업 없음")
