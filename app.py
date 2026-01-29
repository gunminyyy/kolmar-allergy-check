import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
import zipfile
from streamlit_sortables import sort_items
from fpdf import FPDF

# 1. 화면 설정
st.set_page_config(page_title="콜마 83 알러지 통합 검토", layout="wide")

# --- PDF 생성 함수 (기능 유지) ---
class AllergenPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'Allergen Review Report', 0, 1, 'C')
        self.ln(5)

def create_pdf(df, prod_name, p_date, file_name):
    pdf = AllergenPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font('Arial', 'B', 11)
    # 한글 깨짐 방지 인코딩 처리
    p_n = str(prod_name).encode('latin-1', 'ignore').decode('latin-1')
    f_n = str(file_name).encode('latin-1', 'ignore').decode('latin-1')
    pdf.cell(0, 8, f"Product: {p_n}", 0, 1)
    pdf.cell(0, 8, f"Date: {p_date}  |  File: {f_n}", 0, 1)
    pdf.ln(5)
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font('Arial', 'B', 10)
    cols = [("No", 15), ("CAS No", 50), ("Ingredient Name", 100), ("Source", 35), ("Result", 35), ("Status", 30)]
    for col_name, width in cols:
        pdf.cell(width, 10, col_name, 1, 0, 'C', True)
    pdf.ln()
    pdf.set_font('Arial', '', 9)
    for _, row in df.iterrows():
        pdf.cell(cols[0][1], 8, str(row['번호']), 1, 0, 'C')
        pdf.cell(cols[1][1], 8, str(row['CAS']), 1, 0, 'C')
        ing_name = str(row['물질명']).encode('latin-1', 'ignore').decode('latin-1')
        pdf.cell(cols[2][1], 8, ing_name[:55], 1, 0, 'L')
        pdf.cell(cols[3][1], 8, str(row['원본']), 1, 0, 'C')
        pdf.cell(cols[4][1], 8, str(row['양식']), 1, 0, 'C')
        status = "OK" if "✅" in str(row['상태']) else "FAIL"
        if status == "FAIL": pdf.set_text_color(255, 0, 0)
        pdf.cell(cols[5][1], 8, status, 1, 1, 'C')
        pdf.set_text_color(0, 0, 0)
    return pdf.output(dest='S').encode('latin-1')

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

# 3. 메인 UI 구성
st.title("콜마 83 ALLERGENS 통합 검토 시스템(HP,CFF)")
st.info("원본과 양식 파일을 **동일한 순번**으로 배치하세요.")

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 목록")
    uploaded_src_files = st.file_uploader("원본 선택", type=["xlsx"], accept_multiple_files=True, key="src_upload")
    src_file_list = []
    if uploaded_src_files:
        file_display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src_files)]
        st.caption("▼ 드래그하여 순서 조정")
        sorted_names = sort_items(file_display_names)
        for name in sorted_names:
            orig = name.split(". ", 1)[1]
            src_file_list.append(next(f for f in uploaded_src_files if f.name == orig))

# --- 데이터 선처리 로직 (버튼을 상단에 배치하기 위해 미리 계산) ---
all_pdfs = []
processed_results = []

if src_file_list and (uploaded_res_files := st.session_state.get('res_upload')):
    # 양식 파일 리스트 초기화
    res_file_temp = []
    file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
    
    # 4. 검증 로직 실행 (결과 미리 저장)
    num_pairs = min(len(src_file_list), len(uploaded_res_files))
    for idx in range(num_pairs):
        src_f = src_file_list[idx]
        # 일단 순서대로 매칭 (정렬 후 다시 매칭됨)
        # 실제 처리는 아래 UI 렌더링 시점에서 확정
        pass

# 3. 메인 UI 구성 (우측 컬럼 계속)
with col2:
    st.subheader("2. 양식(Result) 파일 목록")
    uploaded_res_files = st.file_uploader("양식 선택", type=["xlsx"], accept_multiple_files=True, key="res_upload")
    res_file_list = []
    
    if uploaded_res_files:
        c2_top_left, c2_top_right = st.columns([0.6, 0.4])
        with c2_top_left:
            st.caption("▼ 드래그하여 순서 조정")
        
        # 정렬 도구
        file_display_names_res = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res_files)]
        sorted_names_res = sort_items(file_display_names_res)
        for name in sorted_names_res:
            orig = name.split(". ", 1)[1]
            res_file_list.append(next(f for f in uploaded_res_files if f.name == orig))

st.markdown("---")

# 4. 검증 로직 및 결과 출력 (개별 버튼 및 전체 버튼 배치)
if src_file_list and res_file_list:
    num_pairs = min(len(src_file_list), len(res_file_list))
    
    # --- 상단 버튼 배치 (양식 목록 아래쪽) ---
    with col2:
        # 일괄 다운로드 버튼 배치 (드래그 조정 우측 라인)
        st.write("") # 간격 조절

    for idx in range(num_pairs):
        src_f = src_file_list[idx]
        res_f = res_file_list[idx]
        mode = "HP" if "HP" in src_f.name.upper() else "CFF"
        
        try:
            wb_s = load_workbook(src_f, data_only=True)
            wb_r = load_workbook(res_f, data_only=True)
            ws_s = wb_s[next((s for s in wb_s.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_s.sheetnames[0])]
            ws_r = wb_r[next((s for s in wb_r.sheetnames if 'ALLERGY' in s.upper()), wb_r.sheetnames[0])]

            s_map, r_map = {}, {}
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

            rp_name, rp_date = str(ws_r['B10'].value or "N/A"), str(ws_r['E10'].value or "N/A").split(' ')[0]
            for r in range(1, 401):
                c = get_cas_set(ws_r.cell(row=r, column=2).value)
                v = ws_r.cell(row=r, column=3).value
                if c and v is not None and v != 0: r_map[c] = {"n": ws_r.cell(row=r, column=1).value, "v": float(v)}

            src_name_check = check_name_match(src_f.name, p_name)
            res_name_check = check_name_match(res_f.name, rp_name)

            all_cas = set(s_map.keys()) | set(r_map.keys())
            rows = []
            mismatch = 0
            for i, c in enumerate(sorted(list(all_cas), key=lambda x: list(x)[0] if x else ""), 1):
                sv, rv = s_map.get(c, {}).get('v', "누락"), r_map.get(c, {}).get('v', "누락")
                match = (sv != "누락" and rv != "누락" and abs(sv - rv) < 0.0001)
                if not match: mismatch += 1
                rows.append({"번호": i, "CAS": ", ".join(list(c)), "물질명": r_map.get(c,{}).get('n') or s_map.get(c,{}).get('n'), "원본": sv, "양식": rv, "상태": "✅" if match else "❌"})

            # 데이터프레임 생성 및 PDF 데이터 생성
            df_res = pd.DataFrame(rows)
            pdf_data = create_pdf(df_res, rp_name, rp_date, res_f.name)
            all_pdfs.append({"name": f"Result_{idx+1}_{rp_name}.pdf", "data": pdf_data})

            # --- 결과 섹션 출력 ---
            status_icon = "✅" if mismatch == 0 else "❌"
            with st.expander(f"{status_icon} [{idx+1}번] {src_f.name} (불일치: {mismatch}건)"):
                m1, m2 = st.columns(2)
                with m1: st.success(f"**원본 제품명:** {p_name} ({src_name_check}) \n**원본 작성일:** {p_date}")
                with m2: st.info(f"**양식 제품명:** {rp_name} ({res_name_check}) \n**양식 작성일:** {rp_date}")
                st.dataframe(df_res, use_container_width=True, hide_index=True)
                # 개별 다운로드 버튼 (가독성을 위해 섹션 안에도 유지)
                st.download_button(f"📥 PDF 다운로드 ({idx+1}번)", pdf_data, f"Result_{idx+1}.pdf", "application/pdf", key=f"pdf_btn_{idx}")

            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 파일 처리 중 오류: {e}")

    # --- 요청하신 위치(양식 목록 우측 상단)에 버튼 배치 ---
    with col2:
        if all_pdfs:
            # 1. 일괄 다운로드 버튼 (드래그 순서 조정 우측)
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for p in all_pdfs: zf.writestr(p["name"], p["data"])
            
            # 버튼 위치 조정
            st.write("---")
            st.download_button("📥 전체 PDF 일괄 다운로드 (ZIP)", zip_buffer.getvalue(), "All_Reports.zip", "application/zip", use_container_width=True)
            
            # 2. 개별 파일별 다운로드 버튼 목록 (정렬된 순서대로 우측 배치)
            st.caption("📄 개별 PDF 바로 저장")
            for i, p_info in enumerate(all_pdfs):
                col_name, col_btn = st.columns([0.7, 0.3])
                with col_name:
                    st.text(f"  {p_info['name'][:30]}...")
                with col_btn:
                    st.download_button("💾 Down", p_info['data'], p_info['name'], "application/pdf", key=f"side_btn_{i}")

    if len(src_file_list) != len(res_file_list):
        st.warning("⚠️ 파일 개수가 일치하지 않습니다.")
else:
    st.info("파일들을 업로드해 주세요.")
