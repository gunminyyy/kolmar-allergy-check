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
    return "✅ 일치" if clean_product_name in clean_file_name or clean_file_name in clean_product_name else "❌ 불일치"

# 3. 메인 UI 구성
st.title("🧪 콜마 83 ALLERGENS 검토 및 자동 수정 시스템")
st.info("불일치 항목이 있을 경우, 원본 수치를 양식 파일에 자동으로 기입한 '수정본 엑셀'을 생성합니다.")

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 목록")
    uploaded_src_files = st.file_uploader("원본 선택", type=["xlsx"], accept_multiple_files=True, key="src_upload")
    src_file_list = []
    if uploaded_src_files:
        file_display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src_files)]
        sorted_names = sort_items(file_display_names)
        for name in sorted_names:
            orig = name.split(". ", 1)[1]
            src_file_list.append(next(f for f in uploaded_src_files if f.name == orig))

with col2:
    st.subheader("2. 양식(Result) 파일 목록")
    uploaded_res_files = st.file_uploader("양식 선택", type=["xlsx"], accept_multiple_files=True, key="res_upload")
    res_file_list = []
    if uploaded_res_files:
        file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
        sorted_names_res = sort_items(file_display_names_res)
        for name in sorted_names_res:
            orig = name.split(". ", 1)[1]
            res_file_list.append(next(f for f in uploaded_res_files if f.name == orig))

st.markdown("---")

# 4. 검증 로직 및 결과 출력
if src_file_list and res_file_list:
    num_pairs = min(len(src_file_list), len(res_file_list))
    all_edited_files = [] # 수정된 파일 저장용 리스트

    for idx in range(num_pairs):
        src_f, res_f = src_file_list[idx], res_file_list[idx]
        mode = "HP" if "HP" in src_f.name.upper() else "CFF"
        
        try:
            # 수정을 위해 data_only=False로도 한 번 더 로드 (수식 유지 목적이나, 값 저장을 위해 일단 True 사용 후 처리)
            wb_s = load_workbook(src_f, data_only=True)
            wb_r = load_workbook(res_f) # 양식 파일은 수정을 위해 수식 유지 상태로 로드
            
            ws_s = wb_s[next((s for s in wb_s.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_s.sheetnames[0])]
            # 양식 파일 시트 찾기
            res_sheet_name = next((s for s in wb_r.sheetnames if 'ALLERGY' in s.upper()), wb_r.sheetnames[0])
            ws_r = wb_r[res_sheet_name]

            s_map = {}
            if mode == "CFF":
                p_name, p_date = str(ws_s['D7'].value or "N/A"), str(ws_s['N9'].value or "N/A").split(' ')[0]
                for r in range(13, 96):
                    c = get_cas_set(ws_s.cell(row=r, column=6).value)
                    v = ws_s.cell(row=r, column=12).value
                    if c and v is not None and v != 0: s_map[c] = {"v": float(v), "n": ws_s.cell(row=r, column=2).value}
            else:
                p_name, p_date = str(ws_s['B10'].value or "N/A"), str(ws_s['E10'].value or "N/A").split(' ')[0]
                for r in range(1, 401):
                    c = get_cas_set(ws_s.cell(row=r, column=2).value)
                    v = ws_s.cell(row=r, column=3).value
                    if c and v is not None and v != 0: s_map[c] = {"v": float(v), "n": ws_s.cell(row=r, column=1).value}

            # 양식 파일 데이터 읽기 및 수정 로직
            rp_name, rp_date = str(ws_r['B10'].value or "N/A"), str(ws_r['E10'].value or "N/A").split(' ')[0]
            r_map = {}
            rows = []
            mismatch_count = 0
            
            # 노란색 하이라이트 설정 (수정된 셀 표시용)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

            # 1단계: 양식 파일의 모든 행을 돌며 원본과 비교 및 수정
            for r in range(1, 401):
                cas_val = ws_r.cell(row=r, column=2).value
                c_set = get_cas_set(cas_val)
                if not c_set: continue
                
                curr_val = ws_r.cell(row=r, column=3).value
                # 원본에 해당 CAS가 있는지 확인
                if c_set in s_map:
                    src_val = s_map[c_set]['v']
                    # 수치가 다르거나 양식에 수치가 없는 경우 수정
                    if curr_val is None or abs(float(curr_val or 0) - src_val) > 0.0001:
                        ws_r.cell(row=r, column=3).value = src_val
                        ws_r.cell(row=r, column=3).fill = yellow_fill # 수정된 칸 표시
                        mismatch_count += 1
                    r_map[c_set] = {"v": src_val, "n": ws_r.cell(row=r, column=1).value, "status": "✅ 수정됨/일치"}
                else:
                    # 원본에 없는 물질이 양식에만 있는 경우
                    if curr_val is not None and curr_val != 0:
                        mismatch_count += 1
                        r_map[c_set] = {"v": curr_val, "n": ws_r.cell(row=r, column=1).value, "status": "❌ 원본누락"}
            
            # 2단계: 화면 출력을 위한 데이터 정리 (사용자님 기존 로직 유지)
            all_cas = set(s_map.keys()) | set(r_map.keys())
            for i, c in enumerate(sorted(list(all_cas), key=lambda x: list(x)[0] if x else ""), 1):
                sv = s_map.get(c, {}).get('v', "누락")
                rv = r_map.get(c, {}).get('v', "누락")
                match = (sv != "누락" and rv != "누락" and abs(float(sv) - float(rv)) < 0.0001)
                rows.append({"번호": i, "CAS": ", ".join(list(c)), "물질명": s_map.get(c,{}).get('n') or r_map.get(c,{}).get('n'), "원본": sv, "수정후": rv, "상태": "✅" if match else "⚠️ 수정됨"})

            # 수정된 파일 저장
            output = io.BytesIO()
            wb_r.save(output)
            all_edited_files.append({"name": f"Edited_{res_f.name}", "data": output.getvalue()})

            # --- 결과 섹션 ---
            status_icon = "✅" if mismatch_count == 0 else "🛠️"
            with st.expander(f"{status_icon} [{idx+1}번] {res_f.name} (수정됨: {mismatch_count}건)"):
                st.write(f"**제품명:** {rp_name} | **파일명 일치:** {check_name_match(res_f.name, rp_name)}")
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                st.download_button(f"💾 {idx+1}번 수정본 다운로드", output.getvalue(), f"Edited_{res_f.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{idx}")

            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 파일 처리 중 오류: {e}")

    # --- 일괄 다운로드 ---
    if all_edited_files:
        st.markdown("---")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for f in all_edited_files: zf.writestr(f["name"], f["data"])
        st.download_button("📥 모든 수정본 일괄 다운로드 (ZIP)", zip_buf.getvalue(), "Edited_All_Files.zip", "application/zip")
else:
    st.info("파일들을 업로드하면 검토 후 자동으로 수정한 파일을 생성합니다.")
