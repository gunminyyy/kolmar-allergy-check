import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
import os
import time
from streamlit_sortables import sort_items

# 1. 화면 설정
st.set_page_config(page_title="알러지 자료 통합 검토", layout="wide")

# [CSS] 파일 업로더 레이아웃 및 업로드 목록 숨김 처리 (중요!)
st.markdown("""
    <style>
    [data-testid="stFileUploader"] { width: 100%; }
    [data-testid="stFileUploaderDropzone"] { padding: 1rem; min-height: 150px; }
    /* 기본 업로더 아래에 생기는 지저분한 파일 목록을 숨깁니다 */
    [data-testid="stFileUploaderFileName"] { display: none; }
    [data-testid="stFileUploaderFileData"] { display: none; }
    div[data-testid="stHorizontalBlock"] div div div div { display: block !important; width: 100% !important; }
    </style>
    """, unsafe_allow_html=True)

# --- [종료 버튼 기능] ---
with st.sidebar:
    st.write("---")
    if st.button("❌ 프로그램 종료", type="primary"):
        st.warning("프로그램을 종료합니다. 창을 닫으셔도 됩니다.")
        time.sleep(1)
        os._exit(0) # 프로세스 강제 종료
# -----------------------

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

def handle_upload(col, label, key):
    with col:
        st.subheader(label)
        uploaded = st.file_uploader(f"{label} 선택", type=["xlsx", "xls"], accept_multiple_files=True, key=key)
        sorted_list = []
        if uploaded:
            # 1. 파일 이름 리스트 생성 (먼저 업로드한 파일이 1번으로 위로 오게 함)
            display_items = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded)]
            
            # 2. 정렬 컴포넌트 실행 (빨간 박스가 모든 파일에 대해 생기도록 함)
            # key값을 더 고유하게 만들어 렌더링 오류를 방지합니다.
            sorted_names = sort_items(display_items, direction="vertical", key=f"sort_v3_{key}_{len(uploaded)}")
            
            # 3. 정렬된 텍스트에서 원본 파일 객체 매칭
            for name in sorted_names:
                try:
                    orig_name = name.split(". ", 1)[1]
                    matched_file = next((f for f in uploaded if f.name == orig_name), None)
                    if matched_file:
                        sorted_list.append(matched_file)
                except (IndexError, StopIteration):
                    continue
        return sorted_list

def extract_data(file_raw, is_23=False, is_83=False):
    f = convert_xls_to_xlsx(file_raw)
    wb = load_workbook(f, data_only=True)
    ws = wb.worksheets[0]
    name_upper = file_raw.name.upper()
    data_map = {}
    product_name = "알 수 없음"
    
    val_a1 = str(ws.cell(row=1, column=1).value or "").strip()
    val_b1 = str(ws.cell(row=1, column=2).value or "").strip()
    
    if val_a1 == "성분코드" and val_b1 == "성분국문명":
        product_name = file_raw.name
        for r in range(2, 85):
            c, v = get_cas_set(ws.cell(row=r, column=6).value), ws.cell(row=r, column=8).value
            if c and v is not None and v != 0: 
                data_map[c] = {"n": ws.cell(row=r, column=2).value, "v": float(v)}
    elif is_83:
        product_name = ws.cell(row=10, column=2).value
        for r in range(1, 401):
            c, v = get_cas_set(ws.cell(row=r, column=2).value), ws.cell(row=r, column=3).value
            if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=1).value, "v": float(v)}
    elif is_23:
        product_name = ws.cell(row=12, column=2).value
        for r in range(18, 44):
            c, v = get_cas_set(ws.cell(row=r, column=2).value), ws.cell(row=r, column=3).value
            if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=1).value or "지정성분", "v": float(v)}
    else:
        if "HPD" in name_upper:
            product_name = ws.cell(row=10, column=3).value
            for r in range(17, 99):
                c, v = get_cas_set(ws.cell(row=r, column=3).value), ws.cell(row=r, column=6).value
                if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=2).value, "v": float(v)}
        elif "HP" in name_upper:
            product_name = ws.cell(row=10, column=2).value
            for r in range(1, 401):
                c, v = get_cas_set(ws.cell(row=r, column=2).value), ws.cell(row=r, column=3).value
                if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=1).value, "v": float(v)}
        else:
            product_name = ws.cell(row=7, column=4).value
            for r in range(13, 96):
                c, v = get_cas_set(ws.cell(row=r, column=6).value), ws.cell(row=r, column=12).value
                if c and v is not None and v != 0: data_map[c] = {"n": ws.cell(row=r, column=2).value, "v": float(v)}
    
    wb.close()
    return str(product_name).strip() if product_name else file_raw.name, data_map

# 3. 메인 UI 구성
st.title("ALLERGENS 자료 통합 검토 시스템(HP/CFF)")

mode = st.radio("검토 방식 선택", ["원본 vs 83알러지", "원본 vs 26알러지", "83알러지 vs 26알러지", "원본 vs 83알러지 vs 26알러지"], horizontal=True)
st.info("파일들을 **동일한 순번**으로 배치하세요. 동일 순번끼리 매칭되어 검토합니다.")
st.markdown("---")

files_A, files_B, files_C = [], [], []
if mode == "원본 vs 83알러지 vs 26알러지":
    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]
    labels = ["원본", "83알러지", "26알러지"]
else:
    col1, col2 = st.columns(2)
    cols = [col1, col2]
    labels = mode.split(" vs ")

files_A = handle_upload(cols[0], labels[0], "upload_A")
files_B = handle_upload(cols[1], labels[1], "upload_B")
if mode == "원본 vs 83알러지 vs 26알러지":
    files_C = handle_upload(cols[2], labels[2], "upload_C")

st.markdown("---")

# 4. 검증 로직 및 결과 출력
ready = files_A and files_B
if mode == "원본 vs 83알러지 vs 26알러지": ready = ready and files_C

if ready:
    num_pairs = min(len(files_A), len(files_B), len(files_C)) if (mode == "원본 vs 83알러지 vs 26알러지") else min(len(files_A), len(files_B))
    
    for idx in range(num_pairs):
        try:
            p_name_1, m1 = extract_data(files_A[idx], is_23=("26알러지" in labels[0]), is_83=("83알러지" in labels[0]))
            p_name_2, m2 = extract_data(files_B[idx], is_23=("26알러지" in labels[1]), is_83=("83알러지" in labels[1]))
            
            m3 = None
            p_name_3 = None
            if mode == "원본 vs 83알러지 vs 26알러지":
                p_name_3, m3 = extract_data(files_C[idx], is_23=True)

            display_p_name = p_name_2 if "알러지" in labels[1] else p_name_1
            if mode == "원본 vs 83알러지 vs 26알러지":
                display_p_name = p_name_2

            if "26알러지" in mode:
                m1 = {cas: d for cas, d in m1.items() if not cas.isdisjoint(TARGET_23_CAS)}
                m2 = {cas: d for cas, d in m2.items() if not cas.isdisjoint(TARGET_23_CAS)}
                if m3: m3 = {cas: d for cas, d in m3.items() if not cas.isdisjoint(TARGET_23_CAS)}

            all_cas_sets = set(m1.keys()) | set(m2.keys())
            if m3: all_cas_sets |= set(m3.keys())

            rows, mismatch = [], 0
            
            for cas in all_cas_sets:
                v1_data = next((m1[c] for c in m1 if not cas.isdisjoint(c)), None)
                v2_data = next((m2[c] for c in m2 if not cas.isdisjoint(c)), None)
                v3_data = next((m3[c] for c in m3 if not cas.isdisjoint(c)), None) if m3 is not None else None

                v1 = v1_data['v'] if v1_data else "누락"
                v2 = v2_data['v'] if v2_data else "누락"
                v3 = (v3_data['v'] if v3_data else "누락") if m3 is not None else None

                name = (v1_data or v2_data or v3_data)['n']

                match = True
                compare_vals = [v for v in [v1, v2, v3] if v is not None]
                
                if "누락" in compare_vals:
                    match = False
                else:
                    it = iter(compare_vals)
                    first = next(it)
                    if not all(abs(first - rest) < 0.0001 for rest in it):
                        match = False

                if not match: mismatch += 1
                
                row_data = {"번호": len(rows)+1, "CAS": ", ".join(list(cas)), "물질명": name, labels[0]: v1, labels[1]: v2}
                if m3 is not None: row_data[labels[2]] = v3
                row_data["상태"] = "✅" if match else "❌"
                rows.append(row_data)

            def get_sum(df_rows, key):
                return sum([r[key] for r in df_rows if isinstance(r[key], (int, float))])
            
            t_a, t_b = get_sum(rows, labels[0]), get_sum(rows, labels[1])
            total_match = abs(t_a - t_b) < 0.0001
            total_row = {"번호": "Total", "CAS": "-", "물질명": "합계", labels[0]: round(t_a, 6), labels[1]: round(t_b, 6)}
            if m3 is not None:
                t_c = get_sum(rows, labels[2])
                total_row[labels[2]] = round(t_c, 6)
                if abs(t_a - t_c) > 0.0001: total_match = False
            total_row["상태"] = "✅" if total_match else "❌"
            rows.append(total_row)

            # 결과 표 출력
            st.expander(f"{'✅' if mismatch == 0 else '❌'} [{idx+1}번] {display_p_name}").dataframe(
                pd.DataFrame(rows).astype(str),
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "CAS": st.column_config.TextColumn("CAS", width="medium", help="마우스를 올리면 전체 CAS 번호가 보입니다.")
                }
            )
            
        except Exception as e:
            st.error(f"{idx+1}번 처리 오류: {e}")
else:
    st.info("검토할 파일들을 모두 업로드해 주세요.")
