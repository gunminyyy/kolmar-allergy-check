import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
from streamlit_sortables import sort_items

# 1. 화면 설정
st.set_page_config(page_title="알러지 자료 통합 검토", layout="wide")

# [수정] 파일 업로더 레이아웃 보정 및 목록 세로 정렬 CSS
st.markdown("""
    <style>
    /* 1. 파일 업로더 내부 요소가 배경 밖으로 나가지 않도록 조정 */
    [data-testid="stFileUploader"] {
        width: 100%;
    }
    [data-testid="stFileUploaderDropzone"] {
        padding: 1rem;  /* 내부 여유 공간 확보 */
        min-height: 150px;
    }
    /* 업로더 내부의 글씨 크기 및 간격 최적화 */
    [data-testid="stFileUploaderDropzone"] div div {
        gap: 0.5rem;
    }
    [data-testid="stFileUploaderDropzone"] small {
        display: block;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    /* 2. 정렬 목록(sort_items) 세로 꽉 차게 설정 */
    div[data-testid="stHorizontalBlock"] div div div div {
        display: block !important;
        width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 23 알러지 양식 검토 대상 CAS 리스트 (26종)
TARGET_23_CAS = {
    "127-51-5", "122-40-7", "101-85-9", "105-13-5", "100-51-6",
    "120-51-4", "103-41-3", "118-58-1", "104-55-2", "104-54-1",
    "5392-40-5", "106-22-9", "91-64-5", "5989-27-5", "97-53-0",
    "4602-84-0", "106-24-1", "101-86-0", "107-75-5", "97-54-1",
    "78-70-6", "31906-04-4", "80-54-6", "111-12-6", "90028-68-5", "90028-67-4"
}

def convert_xls_to_xlsx(uploaded_file):
    if uploaded_file.name.lower().endswith('.xls'):
        df_dict = pd.read_excel(uploaded_file, sheet_name=None, engine='xlrd')
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        return output
    return uploaded_file

def get_cas_set(cas_val):
    if not cas_val: return frozenset()
    cas_list = re.findall(r'\d+-\d+-\d+', str(cas_val))
    return frozenset(cas.strip() for cas in cas_list)

def check_name_match(file_name, product_name):
    clean_file_name = re.sub(r'\.(xlsx|xls)$', '', file_name, flags=re.IGNORECASE).strip()
    clean_product_name = str(product_name).strip()
    if clean_product_name in clean_file_name or clean_file_name in clean_product_name:
        return "✅ 일치"
    return "❌ 불일치"

# 3. 메인 UI 구성
st.title("ALLERGENS 자료 통합 검토 시스템(HP/CFF)")
st.info("검토할 원본과 양식 파일을 **동일한 순번**으로 배치하세요.")

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 목록")
    uploaded_src_files = st.file_uploader("원본 선택", type=["xlsx", "xls"], accept_multiple_files=True, key="src_upload")
    src_file_list = []
    if uploaded_src_files:
        file_display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src_files)]
        sorted_names = sort_items(file_display_names, direction="vertical", key="src_sort")
        for name in sorted_names:
            orig = name.split(". ", 1)[1]
            src_file_list.append(next(f for f in uploaded_src_files if f.name == orig))

with col2:
    st.subheader("2. 양식(Result) 파일 목록")
    uploaded_res_files = st.file_uploader("양식 선택", type=["xlsx", "xls"], accept_multiple_files=True, key="res_upload")
    res_file_list = []
    if uploaded_res_files:
        file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
        sorted_names_res = sort_items(file_display_names_res, direction="vertical", key="res_sort")
        for name in sorted_names_res:
            orig = name.split(". ", 1)[1]
            res_file_list.append(next(f for f in uploaded_res_files if f.name == orig))

st.markdown("---")

# 4. 검증 로직 및 결과 출력
if src_file_list and res_file_list:
    num_pairs = min(len(src_file_list), len(res_file_list))
    
    for idx in range(num_pairs):
        src_f_raw = src_file_list[idx]
        res_f_raw = res_file_list[idx]
        
        src_f = convert_xls_to_xlsx(src_f_raw)
        res_f = convert_xls_to_xlsx(res_f_raw)
        
        src_upper = src_f_raw.name.upper()
        res_upper = res_f_raw.name.upper()
        
        is_83_mode = "83" in res_upper
        mode_label = "83 알러지" if is_83_mode else "23 알러지"
        
        try:
            wb_s = load_workbook(src_f, data_only=True)
            wb_r = load_workbook(res_f, data_only=True)
            
            ws_s = wb_s.worksheets[0]
            ws_r = wb_r.worksheets[0]

            s_map, r_map = {}, {}
            
            # --- 1. 원본(Source) 데이터 추출 ---
            if "HPD" in src_upper:
                p_name, p_date = str(ws_s['C10'].value or "N/A"), str(ws_s['H10'].value or "N/A").split(' ')[0]
                for r in range(17, 99):
                    c = get_cas_set(ws_s.cell(row=r, column=3).value)
                    v = ws_s.cell(row=r, column=6).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=2).value, "v": float(v)}
            elif "HP" in src_upper:
                p_name, p_date = str(ws_s['B10'].value or "N/A"), str(ws_s['E10'].value or "N/A").split(' ')[0]
                for r in range(1, 401):
                    c = get_cas_set(ws_s.cell(row=r, column=2).value)
                    v = ws_s.cell(row=r, column=3).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=1).value, "v": float(v)}
            else: # CFF
                p_name, p_date = str(ws_s['D7'].value or "N/A"), str(ws_s['N9'].value or "N/A").split(' ')[0]
                for r in range(13, 96):
                    c = get_cas_set(ws_s.cell(row=r, column=6).value)
                    v = ws_s.cell(row=r, column=12).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=2).value, "v": float(v)}

            # --- 2. 양식(Result) 데이터 추출 ---
            if is_83_mode:
                rp_name, rp_date = str(ws_r['B10'].value or "N/A"), str(ws_r['E10'].value or "N/A").split(' ')[0]
                for r in range(1, 401):
                    c = get_cas_set(ws_r.cell(row=r, column=2).value)
                    v = ws_r.cell(row=r, column=3).value
                    if c and v is not None and v != 0: r_map[c] = {"n": ws_r.cell(row=r, column=1).value, "v": float(v)}
            else:
                rp_name, rp_date = str(ws_r['B12'].value or "N/A"), str(ws_r['E13'].value or "N/A").split(' ')[0]
                for r in range(18, 44):
                    c = get_cas_set(ws_r.cell(row=r, column=2).value)
                    v = ws_r.cell(row=r, column=3).value
                    if c and v is not None and v != 0: r_map[c] = {"n": ws_r.cell(row=r, column=1).value or "지정성분", "v": float(v)}

            # --- 3. 데이터 필터링 (23 알러지 모드 전용) ---
            if not is_83_mode:
                filtered_s_map = {}
                for cas_set, data in s_map.items():
                    if not cas_set.isdisjoint(TARGET_23_CAS):
                        filtered_s_map[cas_set] = data
                s_map = filtered_s_map

            # --- 4. 데이터 대조 ---
            rows, mismatch = [], 0
            all_s_cas = list(s_map.keys())
            all_r_cas = list(r_map.keys())
            matched_r_cas = set()
            
            for s_cas in all_s_cas:
                sv = s_map[s_cas]['v']
                found_r_cas = next((rc for rc in all_r_cas if not s_cas.isdisjoint(rc)), None)
                if found_r_cas:
                    rv = r_map[found_r_cas]['v']
                    matched_r_cas.add(found_r_cas)
                    match = abs(sv - rv) < 0.0001
                else:
                    rv = "누락"; match = False
                if not match: mismatch += 1
                rows.append({"번호": len(rows)+1, "CAS": ", ".join(list(s_cas)), "물질명": s_map[s_cas]['n'], "원본": sv, "양식": rv, "상태": "✅" if match else "❌"})

            for r_cas in all_r_cas:
                if r_cas not in matched_r_cas:
                    mismatch += 1
                    rows.append({"번호": len(rows)+1, "CAS": ", ".join(list(r_cas)), "물질명": r_map[r_cas]['n'], "원본": "누락", "양식": r_map[r_cas]['v'], "상태": "❌"})

            status_icon = "✅" if mismatch == 0 else "❌"
            expander_title = f"{status_icon} [{idx+1}번] {mode_label} | {res_f_raw.name} (불일치: {mismatch}건)"
            
            with st.expander(expander_title):
                m1, m2 = st.columns(2)
                with m1: 
                    st.success(f"**원본 제품명:** {p_name} ({check_name_match(src_f_raw.name, p_name)})\n\n**원본 작성일:** {p_date}")
                with m2: 
                    st.info(f"**양식 제품명:** {rp_name} ({check_name_match(res_f_raw.name, rp_name)})\n\n**양식 작성일:** {rp_date}")
                
                st.markdown("") 
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            
            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 파일 처리 중 오류: {e}")

    if len(src_file_list) != len(res_file_list):
        st.warning("⚠️ 파일 개수가 일치하지 않습니다.")
else:
    st.info("왼쪽과 오른쪽에 검토할 파일들을 업로드해 주세요.")
