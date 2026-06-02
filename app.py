"""
월간 근태 취합 자동화 도구 — 메인 페이지
이삼오구 HR팀 전용 / 매월 파견직·계약직 근태 정리 자동화
"""
import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime

# Page config
st.set_page_config(
    page_title="월간 근태 취합 자동화",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).parent))
from lib.auth import require_login
require_login()

# Custom CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #185FA5 0%, #0F6E56 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 24px; }
    .main-header p { color: rgba(255,255,255,0.9); margin: 4px 0 0; font-size: 14px; }
    .info-card {
        background: #f6f8fa;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        margin-bottom: 0.8rem;
        border-left: 4px solid #185FA5;
    }
    .info-card h3 { font-size: 16px; margin: 0 0 6px; color: #1f2328; }
    .info-card p { font-size: 13px; color: #6b7280; margin: 0; line-height: 1.5; }
    .step-num {
        display: inline-block;
        width: 22px; height: 22px;
        background: #185FA5;
        color: white;
        border-radius: 50%;
        text-align: center;
        line-height: 22px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 6px;
    }
</style>
""", unsafe_allow_html=True)

# Paths
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

from lib.config_store import load_config, storage_mode

# ====== Sidebar ======
st.sidebar.title("📊 월간 근태 취합")
st.sidebar.caption("이삼오구 HR팀 자동화")
st.sidebar.markdown("---")
st.sidebar.info("👈 좌측 메뉴에서 작업 선택")
st.sidebar.caption(f"🗄 명단 저장: {storage_mode()}")

# ====== Main Header ======
st.markdown("""
<div class="main-header">
    <h1>월간 근태 취합 자동화 도구</h1>
    <p>파견직·계약직 매월 4주 근태현황 분석 · 인원 관리 · 검증 리포트</p>
</div>
""", unsafe_allow_html=True)

# ====== Quick Stats ======
config = load_config()

active_people = [p for p in config['people'] if p.get('status', '활성') == '활성']
contract_dist = {}
for p in active_people:
    ct = p.get('contract_type', '미상')
    contract_dist[ct] = contract_dist.get(ct, 0) + 1

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("활성 인원", f"{len(active_people)} 명")
with col2:
    st.metric("연봉제", f"{contract_dist.get('연봉제', 0)} 명")
with col3:
    st.metric("시급제", f"{contract_dist.get('시급제', 0)} 명")
with col4:
    st.metric("FLEX 전환자", f"{len(config['erp_switchers'])} 명")

st.markdown("---")

# ====== 사용 안내 ======
st.subheader("📋 매월 작업 흐름")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="info-card">
        <h3><span class="step-num">1</span>인원 변경 확인</h3>
        <p>좌측 <b>인원 관리</b>에서 이번 달 신규 입사/퇴사/소정시간 변경 사항 반영</p>
    </div>

    <div class="info-card">
        <h3><span class="step-num">2</span>원본 파일 3종 업로드</h3>
        <p><b>근태정리 생성</b> 페이지에서 원티드 주차/일자 + FLEX 일자 파일 업로드</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="info-card">
        <h3><span class="step-num">3</span>기간 선택 + 생성</h3>
        <p>4주 기간 자동 인식 또는 직접 선택 → <b>"근태정리 생성"</b> 버튼 클릭</p>
    </div>

    <div class="info-card">
        <h3><span class="step-num">4</span>결과 다운로드</h3>
        <p>5개 시트 엑셀 + 검증 리포트 자동 생성 → 다운로드 또는 자동 저장</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ====== 최근 작업 이력 ======
st.subheader("📚 최근 작업 이력")
history_dir = DATA_DIR / "history"
history_dir.mkdir(exist_ok=True)
history_files = sorted(history_dir.glob("*.json"), reverse=True)[:5]

if history_files:
    rows = []
    for hf in history_files:
        try:
            with open(hf, 'r', encoding='utf-8') as f:
                d = json.load(f)
            rows.append({
                '실행 시각': d.get('timestamp', '-'),
                '대상 기간': d.get('period', '-'),
                '대상 인원': d.get('num_people', '-'),
                '결과 파일': d.get('output_file', '-'),
                '상태': d.get('status', '-'),
            })
        except: pass
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.info("아직 실행 이력이 없습니다.")
else:
    st.info("아직 실행 이력이 없습니다. 좌측 메뉴에서 첫 작업을 시작하세요.")

# ====== Footer ======
st.markdown("---")
st.caption(f"이삼오구 월간 근태 취합 자동화 도구 v1.0 · {datetime.now().strftime('%Y-%m-%d %H:%M')}")
