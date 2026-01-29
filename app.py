import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io
import zipfile
from streamlit_sortables import sort_items

# 1. 화면 설정
st.set_page_config(page_title="콜마 83 알러지 통합 검토", layout="wide")

# 2. 공통 도구 함수
def get_cas_set(cas_val):
    if not cas_val: return frozenset()
    cas_list = re.findall(r'\d+-\d+-\d+', str(cas_val))
    return frozenset(cas.strip() for cas in cas_list)

def check_name_match(file_name, product_name):
    clean_file_name = re.sub(r'\.xlsx$', '', file_name, flags=re.IGNORECASE).strip()
    clean_product_name = str(product_name).strip()
    if clean_product_name in clean_file_name or clean_file_name in clean_product_name:
        return "✅ 일치"
    return "❌ 불일치"

# 3. 메인 UI 구성 (멘트 유지)
st.title("콜마 83 ALLERGENS 통합 검토 시스템(HP,CFF)")
st.info("원본과 양식 파일을 **동일한 순번**으로 배치하세요. 순서대로 매칭되어 검토 및 수정본(엑셀) 저장이 가능합니다.")

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 목록")
    uploaded_src_files = st.file_uploader("원본 선택 (다중 가능)", type=["xlsx"], accept_multiple_files=True, key="src_upload")
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
    uploaded_res_files = st.file_uploader("양식 선택 (다중 가능)", type=["xlsx"], accept_multiple_files=True, key="res_upload")
    res_file_list = []
    if uploaded_res_files:
        file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
        st.caption("▼ 드래그하여 순서 조정")
        sorted_names_res = sort_items(file_display_names_res)
        for name in sorted_names_res:
            orig = name.split(". ", 1)[1]
            res_file_list.append(next(f for f in uploaded_res_files if f.name == orig))

st.markdown("---")

# 4. 검증 및 자동 수정 로직
if src_file_list and res_file_list:
    num_pairs = min(len(src_file_list), len(res_file_list))
    all_edited_files = [] 

    for idx in range(num_pairs):
        src_f = src_file_list[idx]
        res_f = res_file_list[idx]
        mode = "HP" if "HP" in src_f.name.upper() else "CFF"
        
        try:
            wb_s = load_workbook(src_f, data_only=True)
            wb_r = load_workbook(res_f)
            
            ws_s = wb_s[next((s for s in wb_s.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_s.sheetnames[0])]
            res_sheet_name = next((s for s in wb_r.sheetnames if 'ALLERGY' in s.upper()), wb_r.sheetnames[0])
            ws_r = wb_r[res_sheet_name]

            s_map = {}
            if mode == "CFF":
                p_name, p_date = str(ws_s['D7'].value or "N/A"), str(ws_s['N9'].value or "N/A").split(' ')[0]
                for r in range(13, 96):
                    c = get_cas_set(ws_s.cell(row=r, column=6).value)
                    v = ws_s.cell(row=r, column=12).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=2).value, "v": float(v)}
            else:
                p_name, p_date = str(ws_s['B10'].value or "N/A"), str(ws_s['E10'].value or "N/A").split(' ')[0]
                for r in range(1, 401):
                    c = get_cas_set(ws_s.cell(row=r, column=2).value)
                    v = ws_s.cell(row=r, column=3).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=1).value, "v": float(v)}

            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            rp_name, rp_date = str(ws_r['B10'].value or "N/A"), str(ws_r['E10'].value or "N/A").split(' ')[0]
            
            # 파일명-제품명 일치 확인
            src_name_check = check_name_match(src_f.name, p_name)
            res_name_check = check_name_match(res_f.name, rp_name)

            r_map = {}
            mismatch_count = 0
            
            for r in range(1, 401):
                cas_val = ws_r.cell(row=r, column=2).value
                c_set = get_cas_set(cas_val)
                if not c_set: continue
                
                curr_val = ws_r.cell(row=r, column=3).value
                if c_set in s_map:
                    src_val = s_map[c_set]['v']
                    try:
                        is_same = (curr_val is not None and abs(float(curr_val) - src_val) < 0.0001)
                    except:
                        is_same = False
                        
                    if not is_same:
                        ws_r.cell(row=r, column=3).value = src_val
                        ws_r.cell(row=r, column=3).fill = yellow_fill
                        mismatch_count += 1
                    r_map[c_set] = {"n": ws_r.cell(row=r, column=1).value, "v": src_val}
                else:
                    if curr_val is not None and curr_val != 0:
                        r_map[c_set] = {"n": ws_r.cell(row=r, column=1).value, "v": curr_val}
                        mismatch_count += 1

            all_cas = set(s_map.keys()) | set(r_map.keys())
            rows = []
            for i, c in enumerate(sorted(list(all_cas), key=lambda x: list(x)[0] if x else ""), 1):
                sv, rv = s_map.get(c, {}).get('v', "누락"), r_map.get(c, {}).get('v', "누락")
                match = (sv != "누락" and rv != "누락" and abs(float(sv if sv != "누락" else 0) - float(rv if rv != "누락" else 0)) < 0.0001)
                rows.append({"번호": i, "CAS": ", ".join(list(c)), "물질명": r_map.get(c,{}).get('n') or s_map.get(c,{}).get('n'), "원본": sv, "양식(수정후)": rv, "상태": "✅" if match else "⚠️ 수정됨"})

            out = io.BytesIO()
            wb_r.save(out)
            if mismatch_count > 0:
                all_edited_files.append({"name": f"수정본_{res_f.name}", "data": out.getvalue()})

            # --- 결과 섹션 ---
            status_icon = "✅" if mismatch_count == 0 else "❌"
            expander_title = f"{status_icon} [{idx+1}번] {src_f.name} (불일치: {mismatch_count}건)"
            
            with st.expander(expander_title):
                m1, m2 = st.columns(2)
                with m1: 
                    st.success(f"**원본 제품명:** \n{p_name} ({src_name_check})")
                    st.success(f"**원본 작성일:** \n{p_date}") # 디자인 통일 (검사 멘트 삭제)
                with m2: 
                    st.info(f"**양식 제품명:** \n{rp_name} ({res_name_check})")
                    st.info(f"**양식 작성일:** \n{rp_date}") # 디자인 통일 (검사 멘트 삭제)
                
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                
                if mismatch_count > 0:
                    st.download_button(f"💾 {idx+1}번 수정본 엑셀 다운로드", out.getvalue(), f"Edited_{res_f.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"btn_{idx}")
            
            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 파일 처리 중 오류: {e}")

    if all_edited_files:
        st.markdown("---")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for f in all_edited_files: zf.writestr(f["name"], f["data"])
        st.download_button("📥 모든 수정본 일괄 다운로드 (ZIP)", zip_buf.getvalue(), "Edited_All.zip", "application/zip")

    if len(src_file_list) != len(res_file_list):
        st.warning("⚠️ 원본과 양식의 파일 개수가 일치하지 않습니다.")
else:
    st.info("왼쪽과 오른쪽에 검토할 파일들을 업로드해 주세요.")
