import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
from streamlit_sortables import sort_items

# 1. 화면 설정
st.set_page_config(page_title="알러지 자료 통합 검토", layout="wide")

# [CSS] 파일 업로더 레이아웃 및 목록 세로 정렬
st.markdown("""
    <style>
    [data-testid="stFileUploader"] { width: 100%; }
    [data-testid="stFileUploaderDropzone"] { padding: 1rem; min-height: 150px; }
    [data-testid="stFileUploaderDropzone"] small { display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    div[data-testid="stHorizontalBlock"] div div div div { display: block !important; width: 100% !important; }
    </style>
    """, unsafe_allow_html=True)

# 23(26종) 알러지 양식 검토 대상 CAS 리스트
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

# 검토 모드 선택
mode = st.radio(
    "검토 방식 선택",
    ["원본 vs 83알러지", "원본 vs 26알러지", "83알러지 vs 26알러지", "원본 vs 83알러지 vs 26알러지"],
    horizontal=True
)

st.info("파일들을 **동일한 순번**으로 배치하세요.")
st.markdown("---")

# 업로드 박스 배치
files_A, files_B, files_C = [], [], []
if mode == "원본 vs 83알러지 vs 26알러지":
    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]
    labels = ["1. 원본 파일", "2. 83알러지 파일", "3. 26알러지 파일"]
else:
    col1, col2 = st.columns(2)
    cols = [col1, col2]
    labels = mode.split(" vs ")

def handle_upload(col, label, key):
    with col:
        st.subheader(label)
        uploaded = st.file_uploader(f"{label} 선택", type=["xlsx", "xls"], accept_multiple_files=True, key=key)
        sorted_list = []
        if uploaded:
            display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded)]
            sorted_names = sort_items(display_names, direction="vertical", key=f"sort_{key}")
            for name in sorted_names:
                orig = name.split(". ", 1)[1]
                sorted_list.append(next(f for f in uploaded if f.name == orig))
        return sorted_list

files_A = handle_upload(cols[0], labels[0], "upload_A")
files_B = handle_upload(cols[1], labels[1], "upload_B")
if mode == "원본 vs 83알러지 vs 26알러지":
    files_C = handle_upload(cols[2], labels[2], "upload_C")

st.markdown("---")

# 데이터 추출 함수 (중복 로직 제거용)
def extract_data(file_raw, is_23=False, is_83=False):
    f = convert_xls_to_xlsx(file_raw)
    wb = load_workbook(f, data_only=True)
    ws = wb.worksheets[0]
    name_upper = file_raw.name.upper()
    data_map = {}
    
    # 원본(HP/HPD/CFF) 및 83/26 양식 위치 로직
    if is_83: # 83알러지 양식
        p_name, p_date = str(ws['B10'].value or "N/A"), str(ws['E10'].value or "N/A").split(' ')[0]
        for r in range(1, 401):
            c = get_cas_set(ws.cell(row=r, column=2).value)
            v = ws.cell(row=r, column=3).value
            if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=1).value, "v": float(v)}
    elif is_23: # 26종 알러지 양식
        p_name, p_date = str(ws['B12'].value or "N/A"), str(ws['E13'].value or "N/A").split(' ')[0]
        for r in range(18, 44):
            c = get_cas_set(ws.cell(row=r, column=2).value)
            v = ws.cell(row=r, column=3).value
            if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=1).value or "지정성분", "v": float(v)}
    else: # 원본
        if "HPD" in name_upper:
            p_name, p_date = str(ws['C10'].value or "N/A"), str(ws['H10'].value or "N/A").split(' ')[0]
            for r in range(17, 99):
                c, v = get_cas_set(ws.cell(row=r, column=3).value), ws.cell(row=r, column=6).value
                if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=2).value, "v": float(v)}
        elif "HP" in name_upper:
            p_name, p_date = str(ws['B10'].value or "N/A"), str(ws['E10'].value or "N/A").split(' ')[0]
            for r in range(1, 401):
                c, v = get_cas_set(ws.cell(row=r, column=2).value), ws.cell(row=r, column=3).value
                if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=1).value, "v": float(v)}
        else: # CFF
            p_name, p_date = str(ws['D7'].value or "N/A"), str(ws['N9'].value or "N/A").split(' ')[0]
            for r in range(13, 96):
                c, v = get_cas_set(ws.cell(row=r, column=6).value), ws.cell(row=r, column=12).value
                if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=2).value, "v": float(v)}
    
    wb.close()
    return p_name, p_date, data_map

# 4. 검증 로직 및 결과 출력
ready = files_A and files_B
if mode == "원본 vs 83알러지 vs 26알러지": ready = ready and files_C

if ready:
    num_pairs = min(len(files_A), len(files_B), len(files_C)) if files_C else min(len(files_A), len(files_B))
    
    for idx in range(num_pairs):
        try:
            # 데이터 추출
            n1, d1, m1 = extract_data(files_A[idx], is_23=("26알러지" in labels[0]), is_83=("83알러지" in labels[0]))
            n2, d2, m2 = extract_data(files_B[idx], is_23=("26알러지" in labels[1]), is_83=("83알러지" in labels[1]))
            
            m3 = None
            if mode == "원본 vs 83알러지 vs 26알러지":
                n3, d3, m3 = extract_data(files_C[idx], is_23=True)
                # 3개 검토 시 원본(m1)에서 26종 CAS만 필터링
                m1 = {cas: data for cas, data in m1.items() if not cas.isdisjoint(TARGET_23_CAS)}
            elif "26알러지" in mode:
                # 2개 검토 중 26알러지가 포함된 경우 원본/83 필터링
                if "26알러지" in labels[1]: m1 = {cas: data for cas, data in m1.items() if not cas.isdisjoint(TARGET_23_CAS)}
                else: m2 = {cas: data for cas, data in m2.items() if not cas.isdisjoint(TARGET_23_CAS)}

            # 대조 로직
            rows, mismatch = [], 0
            base_map = m1
            compare_maps = [m2] if m3 is None else [m2, m3]
            
            for cas, data in base_map.items():
                v_base = data['v']
                res_vals = []
                row_match = True
                
                for target_map in compare_maps:
                    found_cas = next((tc for tc in target_map.keys() if not cas.isdisjoint(tc)), None)
                    if found_cas:
                        v_comp = target_map[found_cas]['v']
                        res_vals.append(v_comp)
                        if abs(v_base - v_comp) > 0.0001: row_match = False
                    else:
                        res_vals.append("누락")
                        row_match = False
                
                if not row_match: mismatch += 1
                
                row_data = {"번호": len(rows)+1, "CAS": ", ".join(list(cas)), "물질명": data['n'], labels[0]: v_base, labels[1]: res_vals[0]}
                if m3: row_data[labels[2]] = res_vals[1]
                row_data["상태"] = "✅" if row_match else "❌"
                rows.append(row_data)

            # 합계 행
            def get_sum(df_rows, key):
                return sum([r[key] for r in df_rows if isinstance(r[key], (int, float))])
            
            t_a, t_b = get_sum(rows, labels[0]), get_sum(rows, labels[1])
            total_match = abs(t_a - t_b) < 0.0001
            total_row = {"번호": "Total", "CAS": "-", "물질명": "합계", labels[0]: round(t_a, 6), labels[1]: round(t_b, 6)}
            if m3:
                t_c = get_sum(rows, labels[2])
                total_row[labels[2]] = round(t_c, 6)
                if abs(t_a - t_c) > 0.0001: total_match = False
            
            total_row["상태"] = "✅" if total_match else "❌"
            rows.append(total_row)

            # 출력
            st.expander(f"{'✅' if mismatch == 0 else '❌'} [{idx+1}번] {mode} | {files_A[idx].name}").dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            
        except Exception as e:
            st.error(f"{idx+1}번 처리 오류: {e}")
else:
    st.info("검토할 파일들을 모두 업로드해 주세요.")
