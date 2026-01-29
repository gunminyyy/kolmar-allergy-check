import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
# 파일 순서 조정을 위한 라이브러리 추가
from streamlit_sortables import sort_items

# 1. 화면 설정
st.set_page_config(page_title="콜마 83 알러지 통합 검증", layout="wide")

# 2. 공통 도구 함수
def get_cas_set(cas_val):
    if not cas_val: return frozenset()
    cas_list = re.findall(r'\d+-\d+-\d+', str(cas_val))
    return frozenset(cas.strip() for cas in cas_list)

# 3. 메인 UI 구성
st.title("🧪 83 ALLERGENS 통합 검증 시스템")
st.info("양식을 선택하고 파일을 업로드하세요. 업로드 후 드래그하여 순서를 바꿀 수 있습니다.")

# 양식 선택
mode = st.radio("📂 원본 파일 양식을 선택하세요", ["CFF 양식", "HP 양식"], horizontal=True)

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader(f"1. 원본({mode}) 파일")
    uploaded_src_files = st.file_uploader("원본 파일들을 선택하세요 (다중 선택 가능)", type=["xlsx"], accept_multiple_files=True, key="src_upload")
    
    src_file_list = []
    if uploaded_src_files:
        # 파일명 앞에 순번(1, 2, 3...)과 드래그 표식 추가하여 리스트 생성
        file_display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src_files)]
        st.write("▼ 드래그하여 분석 순서를 변경하세요")
        sorted_display_names = sort_items(file_display_names)
        
        # 정렬된 순서에 맞게 실제 파일 객체 매핑
        for display_name in sorted_display_names:
            original_name = display_name.split(". ", 1)[1]
            actual_file = next(f for f in uploaded_src_files if f.name == original_name)
            src_file_list.append(actual_file)

with col2:
    st.subheader("2. 최종본(Result) 파일")
    uploaded_res_files = st.file_uploader("최종본 파일들을 선택하세요 (다중 선택 가능)", type=["xlsx"], accept_multiple_files=True, key="res_upload")
    
    res_file_list = []
    if uploaded_res_files:
        file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
        st.write("▼ 드래그하여 분석 순서를 변경하세요")
        sorted_display_names_res = sort_items(file_display_names_res)
        
        for display_name in sorted_display_names_res:
            original_name = display_name.split(". ", 1)[1]
            actual_file = next(f for f in uploaded_res_files if f.name == original_name)
            res_file_list.append(actual_file)

# 4. 검증 로직 실행 (첫 번째 쌍 위주로 예시 구현)
if src_file_list and res_file_list:
    # 예시로 정렬된 리스트의 첫 번째 파일들끼리 비교
    src_file = src_file_list[0]
    res_file = res_file_list[0]
    
    try:
        wb_src = load_workbook(src_file, data_only=True)
        wb_res = load_workbook(res_file, data_only=True)
        
        # (이하 기존 로직과 동일)
        src_sheet_name = next((s for s in wb_src.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_src.sheetnames[0])
        res_sheet_name = next((s for s in wb_res.sheetnames if 'ALLERGY' in s.upper()), wb_res.sheetnames[0])
        
        ws_src = wb_src[src_sheet_name]
        ws_res = wb_res[res_sheet_name]

        src_data_map = {}
        res_data_map = {}

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

        res_product = str(ws_res['B10'].value or "정보없음").strip()
        res_date = str(ws_res['E10'].value or "날짜없음").split(' ')[0]
        for r in range(1, 400):
            c_set = get_cas_set(ws_res.cell(row=r, column=2).value)
            val = ws_res.cell(row=r, column=3).value
            if c_set and val is not None and val != 0:
                res_data_map[c_set] = {"name": ws_res.cell(row=r, column=1).value, "val": float(val)}

        all_cas_sets = set(src_data_map.keys()) | set(res_data_map.keys())
        table_data = []
        match_count = 0

        for i, c_set in enumerate(sorted(list(all_cas_sets), key=lambda x: list(x)[0] if x else ""), 1):
            s_val = src_data_map.get(c_set, {}).get('val', "누락")
            r_val = res_data_map.get(c_set, {}).get('val', "누락")
            name = res_data_map.get(c_set, {}).get('name') or src_data_map.get(c_set, {}).get('name') or "Unknown"
            is_match = (s_val != "누락" and r_val != "누락" and abs(s_val - r_val) < 0.0001)
            if is_match: match_count += 1
            table_data.append({"번호": i, "CAS 번호": ", ".join(list(c_set)), "물질명": name, "원본 수치": s_val, "최종 수치": r_val, "상태": "✅ 일치" if is_match else "❌ 불일치"})

        st.success(f"현재 분석 대상: {src_file.name} vs {res_file.name}")
        summ_col1, summ_col2 = st.columns(2)
        with summ_col1: st.info(f"**원본 제품명:** {src_product}\n\n**원본 작성일:** {src_date}")
        with summ_col2: st.info(f"**최종본 제품명:** {res_product}\n\n**최종본 작성일:** {res_date}")

        st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)
        st.metric("검증 요약", f"총 {len(table_data)}건", f"불일치 {len(table_data) - match_count}건", delta_color="inverse")

    except Exception as e:
        st.error(f"에러 발생: {e}")
else:
    st.info("파일들을 업로드하면 순서대로 매칭하여 검토를 시작합니다.")

