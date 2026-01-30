import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
from streamlit_sortables import sort_items

# 1. 화면 설정
st.set_page_config(page_title="알러지 자료 통합 검토", layout="wide")

# [추가] xls 파일을 openpyxl이 읽을 수 있도록 메모리에서 변환하는 함수
def convert_xls_to_xlsx(uploaded_file):
    if uploaded_file.name.lower().endswith('.xls'):
        # xlrd 엔진을 사용하여 구형 엑셀 읽기
        df_dict = pd.read_excel(uploaded_file, sheet_name=None, engine='xlrd')
        output = io.BytesIO()
        # 메모리 내에서 최신 xlsx 형식으로 변환
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        return output
    return uploaded_file

# 2. 공통 도구 함수
def get_cas_set(cas_val):
    if not cas_val: return frozenset()
    cas_list = re.findall(r'\d+-\d+-\d+', str(cas_val))
    return frozenset(cas.strip() for cas in cas_list)

# 파일명과 제품명 비교 함수 (xls 확장자 대응 추가)
def check_name_match(file_name, product_name):
    clean_file_name = re.sub(r'\.(xlsx|xls)$', '', file_name, flags=re.IGNORECASE).strip()
    clean_product_name = str(product_name).strip()
    if clean_product_name in clean_file_name or clean_file_name in clean_product_name:
        return "✅ 일치"
    return "❌ 불일치"

# 3. 메인 UI 구성
st.title("ALLERGENS 자료 통합 검토 시스템(HP/CFF)")
st.info("원본과 양식 파일을 **동일한 순번**으로 배치하세요. .xls와 .xlsx 모두 지원합니다.")

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 목록")
    # [수정] type에 xls 추가
    uploaded_src_files = st.file_uploader("원본 선택 (다중 가능)", type=["xlsx", "xls"], accept_multiple_files=True, key="src_upload")
    src_file_list = []
    if uploaded_src_files:
        file_display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src_files)]
        st.caption("▼ 드래그하여 순서 조정")
        sorted_names = sort_items(file_display_names)
        for name in sorted_names:
            orig = name.split(". ", 1)[1]
            src_file_list.append(next(f for f in uploaded_src_files if f.name == orig))

with col2:
    st.subheader("2. 양식(Result) 파일 목록")
    # [수정] type에 xls 추가
    uploaded_res_files = st.file_uploader("양식 선택 (다중 가능)", type=["xlsx", "xls"], accept_multiple_files=True, key="res_upload")
    res_file_list = []
    if uploaded_res_files:
        file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
        st.caption("▼ 드래그하여 순서 조정")
        sorted_names_res = sort_items(file_display_names_res)
        for name in sorted_names_res:
            orig = name.split(". ", 1)[1]
            res_file_list.append(next(f for f in uploaded_res_files if f.name == orig))

st.markdown("---")

# 4. 검증 로직 및 결과 출력
if src_file_list and res_file_list:
    num_pairs = min(len(src_file_list), len(res_file_list))
    
    for idx in range(num_pairs):
        # [수정] 원본 파일 객체 보관 (이름 참조용)
        src_f_raw = src_file_list[idx]
        res_f_raw = res_file_list[idx]
        
        # [수정] xls인 경우 변환 로직 통과 후 처리
        src_f = convert_xls_to_xlsx(src_f_raw)
        res_f = convert_xls_to_xlsx(res_f_raw)
        
        target_name = src_f_raw.name.upper()
        # 모드 판별: '83'이 포함되면 기존 83 로직, 없으면 23 로직
        is_83_mode = "83" in target_name
        mode_label = "83 알러지" if is_83_mode else "23 알러지"
        
        try:
            wb_s = load_workbook(src_f, data_only=True)
            wb_r = load_workbook(res_f, data_only=True)
            
            # 시트 선택 로직
            ws_s = wb_s[next((s for s in wb_s.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_s.sheetnames[0])]
            ws_r = wb_r[next((s for s in wb_r.sheetnames if 'ALLERGY' in s.upper()), wb_r.sheetnames[0])]

            s_map, r_map = {}, {}
            
            # --- 원본(Source) 데이터 추출 ---
            if is_83_mode:
                sub_mode = "HP" if ("HP" in target_name or "HPD" in target_name) else "CFF"
                if sub_mode == "CFF":
                    p_name, p_date = str(ws_s['D7'].value or "N/A"), str(ws_s['N9'].value or "N/A").split(' ')[0]
                    for r in range(13, 96):
                        c = get_cas_set(ws_s.cell(row=r, column=6).value)
                        v = ws_s.cell(row=r, column=12).value
                        if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=2).value, "v": float(v)}
                else: # HP
                    p_name, p_date = str(ws_s['B10'].value or "N/A"), str(ws_s['E10'].value or "N/A").split(' ')[0]
                    for r in range(1, 401):
                        c = get_cas_set(ws_s.cell(row=r, column=2).value)
                        v = ws_s.cell(row=r, column=3).value
                        if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=1).value, "v": float(v)}
            else:
                p_name, p_date = str(ws_s['B12'].value or "N/A"), str(ws_s['E13'].value or "N/A").split(' ')[0]
                for r in range(18, 44):
                    c = get_cas_set(ws_s.cell(row=r, column=2).value)
                    v = ws_s.cell(row=r, column=3).value
                    if c and v is not None and v != 0: s_map[c] = {"n": "물질(23)", "v": float(v)}

            # --- 양식(Result) 데이터 추출 ---
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
                    if c and v is not None and v != 0: r_map[c] = {"n": "물질(23)", "v": float(v)}

            # 파일명-제품명 일치 여부 확인
            src_name_check = check_name_match(src_f_raw.name, p_name)
            res_name_check = check_name_match(res_f_raw.name, rp_name)

            # --- 데이터 대조 로직 ---
            rows = []
            mismatch = 0
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

            # --- 접이식 결과 섹션 ---
            status_icon = "✅" if mismatch == 0 else "❌"
            expander_title = f"{status_icon} [{idx+1}번] {mode_label} | {src_f_raw.name} (불일치: {mismatch}건)"
            
            with st.expander(expander_title):
                m1, m2 = st.columns(2)
                with m1:
                    st.success(f"**원본 제품명:** {p_name} ({src_name_check})  \n**원본 작성일:** {p_date}")
                with m2:
                    st.info(f"**양식 제품명:** {rp_name} ({res_name_check})  \n**양식 작성일:** {rp_date}")
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            
            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 파일 처리 중 오류: {e}")

    if len(src_file_list) != len(res_file_list):
        st.warning("⚠️ 원본과 양식의 파일 개수가 일치하지 않습니다.")
else:
    st.info("왼쪽과 오른쪽에 검토할 파일들을 업로드해 주세요.")
