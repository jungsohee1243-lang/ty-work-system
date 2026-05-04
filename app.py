import os
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from hdfc_processor import process_file

st.set_page_config(
    page_title="HDFC 메니보 자동처리",
    page_icon="📦",
    layout="wide",
)

st.markdown("""
<style>
.main {background-color: #f7f8fb;}
.block-container {padding-top: 2rem; padding-bottom: 2rem; max-width: 1100px;}
.hero {
    background: linear-gradient(135deg, #1f2937 0%, #334155 55%, #d8b76a 100%);
    color: white;
    padding: 28px 34px;
    border-radius: 22px;
    box-shadow: 0 12px 30px rgba(15, 23, 42, 0.14);
    margin-bottom: 22px;
}
.hero h1 {font-size: 34px; margin: 0 0 8px 0; font-weight: 800;}
.hero p {font-size: 15px; margin: 0; opacity: 0.92;}
.card {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 22px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}
.step-title {font-weight: 800; font-size: 18px; margin-bottom: 8px;}
.small-note {color: #64748b; font-size: 13px;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>HDFC 메니보 자동처리</h1>
  <p>엑셀 업로드 → 허용품목코드/ZIP/용도구분/중량/전화번호 중복/FTA 적용건 자동 처리 → 결과 다운로드</p>
</div>
""", unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="step-title">1. 엑셀 파일 업로드</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "메니보 엑셀 파일을 업로드하세요",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        label_visibility="collapsed",
    )

    st.markdown('<div class="step-title">2. 목표 총중량 입력</div>', unsafe_allow_html=True)
    target_text = st.text_input(
        "목표 총중량(kg) - 없으면 비워두기",
        placeholder="예: 6922.9 / 없으면 공란",
    )

    st.markdown('<div class="step-title">3. 변환 실행</div>', unsafe_allow_html=True)
    run_btn = st.button("변환 실행", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.divider()

if run_btn:
    if uploaded_file is None:
        st.error("먼저 엑셀 파일을 업로드해주세요.")
        st.stop()

    target_total = None
    if target_text.strip():
        try:
            target_total = float(target_text.strip().replace(",", ""))
        except ValueError:
            st.error("목표 총중량은 숫자로 입력해주세요. 예: 6922.9")
            st.stop()

    with st.spinner("파일 처리 중입니다. 잠시만 기다려주세요..."):
        try:
            suffix = Path(uploaded_file.name).suffix or ".xlsx"
            safe_name = Path(uploaded_file.name).stem.replace("/", "_").replace("\\", "_")
            with tempfile.TemporaryDirectory() as tmpdir:
                input_path = os.path.join(tmpdir, f"{safe_name}{suffix}")
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                final_path = process_file(input_path, target_total=target_total)

                if not final_path or not os.path.exists(final_path):
                    raise FileNotFoundError("결과 파일이 생성되지 않았습니다.")

                with open(final_path, "rb") as f:
                    result_bytes = f.read()

            output_name = f"{safe_name}_중량조정_최종_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.success("처리 완료되었습니다. 아래 버튼으로 결과 파일을 다운로드하세요.")
            st.download_button(
                "결과 엑셀 다운로드",
                data=result_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.info("컬럼명이 기존 양식과 다른 경우 오류가 날 수 있습니다. 업로드 파일의 컬럼명을 확인해주세요.")

with st.expander("처리 규칙 안내"):
    st.write("- 허용품목코드 앞자리가 1, 2, 30, 90이면 960719로 변환")
    st.write("- V=3이고 예외 HS 코드(900290, 900410, 902920)는 변환 제외")
    st.write("- ZIP CODE 4자리는 앞에 0을 붙여 5자리로 보정")
    st.write("- 지정 키워드 및 WIRELESS 문구에 따라 용도구분 V=1 → V=3 변경")
    st.write("- AF 2~5kg 대상은 1.5~1.9kg으로 재분배")
    st.write("- 목표 총중량 입력 시 0.1kg 단위로 추가 분배")
    st.write("- 전화번호 중복 + 총금액 합계 150 이상이면 V=3 변경")
    st.write("- FTA 적용건 HAWB 리스트를 메모 시트에 생성")
