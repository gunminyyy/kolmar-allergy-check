import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io

# 1. 화면 설정
st.set_page_config(page_title="콜마 83 알러지 통합 검토", layout="wide")

# 2. 공통 도구 함수
def get_cas_set(cas_val):
    if not cas_val: return frozenset()
    cas_list = re.findall(r'\d+-\d+-\d+', str(cas_val))
    return frozenset(cas.strip() for cas in cas_list)

# 3. 메인 UI 구성
st.title("🧪 콜마 83 ALLERGENS 검토 시스템 (다중 매칭)")
st.info("원본과 최종본 파일을 **동일한 순서**로 여러 개 업로드하세요. 순서대로 1:1 매칭되어 검증됩니다.")

# 양식 선택 라디오 버튼
mode = st.radio("📂 원본 파일 양식을 선택하세요", ["CFF 양식", "HP 양식"], horizontal=True)

st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    # accept_multiple_files=True 옵션 추가
    src_files = st.file_uploader(f"1. 원본({mode}) 파일들 업로드", type=["xlsx"], accept_multiple_files=True)
with col2:
    res_files = st.file_uploader("2. 최종본(Result) 파일들 업로드", type=["xlsx"], accept_multiple_files=True)

# 4. 검증 로직 실행
if src_files and res_files:
    # 두 리스트 중 개수가 적은 쪽에 맞춰서 반복 (짝이 안 맞으면 경고)
    if len(src_files) != len(res_files):
        st.warning(f"⚠️ 파일 개수가 일치하지 않습니다. (원본: {len(src_files)}개 / 최종본: {len(res_files)}개) 올린 순서대로 {min(len(src_files), len(res_files))}번까지만 비교합니다.")

    # zip을 사용하여 순서대로 짝꿍 만들기
    for idx, (src_f, res_f) in enumerate(zip(src_files, res_files), 1):
        with st.expander(f"📋 {idx}번 매칭 결과: {src_f.name} ↔ {res_f.name}", expanded=True):
            try:
                wb_src = load_workbook(src_f, data_only=True)
                wb_res = load_workbook(res_f, data_only=True)
                
                src_sheet_name = next((s for s in wb_src.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_src.sheetnames[0])
                res_sheet_name = next((s for s in wb_res.sheetnames if 'ALLERGY' in s.upper()), wb_res.sheetnames[0])
                
                ws_src = wb_src[src_sheet_name]
                ws_res = wb_res[res_sheet_name]

                src_data_map = {}
                res_data_map = {}

                # --- A. 원본 데이터 수집 ---
                if mode == "CFF 양식":
                    src_product = str(ws_src['D7'].value or "정보없음").strip()
                    src_date = str(ws_src['N9'].value or "날짜없음").split(' ')[0]
                    for r in range(13, 96):
                        c_set = get_cas_set(ws_src.cell(row=r, column=6).value)
                        val = ws_src.cell(row=r, column=12).value
                        if c_set and val is not None and val != 0:
                            src_data_map[c_set] = {"name": ws_src.cell(row=r, column=2).value, "val": float(val)}
                else:
                    src_product = str(ws_src['B10'].value or "정보없음").strip()
                    src_date = str(ws_src['E10'].value or "날짜없음").split(' ')[0]
                    for r in range(1, 400):
                        c_set = get_cas_set(ws_src.cell(row=r, column=2).value)
                        val = ws_src.cell(row=r, column=3).value
                        if c_set and val is not None and val != 0:
                            src_data_map[c_set] = {"name": ws_src.cell(row=r, column=1).value, "val": float(val)}

                # --- B. 최종본 데이터 수집 ---
                res_product = str(ws_res['B10'].value or "정보없음").strip()
                res_date = str(ws_res['E10'].value or "날짜없음").split(' ')[0]
                for r in range(1, 400):
                    c_set = get_cas_set(ws_res.cell(row=r, column=2).value)
                    val = ws_res.cell(row=r, column=3).value
                    if c_set and val is not None and val != 0:
                        res_data_map[c_set] = {"name": ws_res.cell(row=r, column=1).value, "val": float(val)}

                # --- C. 비교 결과 생성 ---
                all_cas_sets = set(src_data_map.keys()) | set(res_data_map.keys())
                table_data = []
                match_count = 0

                for i, c_set in enumerate(sorted(list(all_cas_sets), key=lambda x: list(x)[0] if x else ""), 1):
                    s_val = src_data_map.get(c_set, {}).get('val', "누락")
                    r_val = res_data_map.get(c_set, {}).get('val', "누락")
                    name = res_data_map.get(c_set, {}).get('name') or src_data_map.get(c_set, {}).get('name') or "Unknown"
                    is_match = (s_val != "누락" and r_val != "누락" and abs(s_val - r_val) < 0.0001)
                    if is_match: match_count += 1
                    
                    table_data.append({
                        "번호": i,
                        "CAS 번호": ", ".join(list(c_set)),
                        "물질명": name,
                        "원본 수치": s_val,
                        "최종 수치": r_val,
                        "상태": "✅ 일치" if is_match else "❌ 불일치"
                    })

                # --- D. 결과 출력 ---
                summ_col1, summ_col2 = st.columns(2)
                with summ_col1:
                    st.info(f"**원본 제품명:** {src_product}\n\n**원본 작성일:** {src_date}")
                with summ_col2:
                    st.info(f"**최종본 제품명:** {res_product}\n\n**최종본 작성일:** {res_date}")

                df = pd.DataFrame(table_data)
                st.dataframe(df, use_container_width=True, hide_index=True)
                
                mismatch_count = len(table_data) - match_count
                st.metric(f"매칭 {idx} 검증 요약", f"총 {len(table_data)}건", f"불일치 {mismatch_count}건", delta_color="inverse")

                wb_src.close(); wb_res.close()

            except Exception as e:
                st.error(f"{idx}번 파일 처리 중 오류 발생: {e}")
else:
    st.info("파일들을 업로드하면 순서대로 매칭하여 검토를 시작합니다.")
