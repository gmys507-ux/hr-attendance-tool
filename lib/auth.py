"""
간단 비밀번호 잠금 (공유 비밀번호 1개).
- 우선순위: st.secrets['app_password'] (클라우드용) > data/app_config.json (로컬/사내망용)
- 둘 다 없으면 게이트 비활성 (로컬 개발 편의)
개인별 로그인 계정은 추후 확장.
"""
import json
from pathlib import Path
import streamlit as st

CONFIG = Path(__file__).parent.parent / "data" / "app_config.json"


def _expected_password():
    try:
        if "app_password" in st.secrets:
            return str(st.secrets["app_password"])
    except Exception:
        pass
    if CONFIG.exists():
        try:
            return json.loads(CONFIG.read_text(encoding="utf-8")).get("password")
        except Exception:
            return None
    return None


def require_login():
    """비밀번호가 설정돼 있으면 잠금 화면을 띄우고, 통과 전엔 페이지 실행 중단."""
    pw = _expected_password()
    if not pw:
        return  # 미설정 → 게이트 비활성
    if st.session_state.get("_authed"):
        return

    st.markdown("## 🔒 HR 근태정리 자동화 도구")
    st.caption("HR팀 전용 · 접속 비밀번호를 입력하세요.")
    with st.form("login_form"):
        entered = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("입장", type="primary")
    if submitted:
        if entered == pw:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("❌ 비밀번호가 올바르지 않습니다.")
    st.stop()
