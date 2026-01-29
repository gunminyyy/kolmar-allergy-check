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

# --- PDF 생성 클래스 (fpdf 사용) ---
class AllergenPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'Allergen Review Report', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_pdf(df, prod_name, p_date, file_name):
    # L: Landscape(가로), mm: 밀리미터 단위, A4 용지
    pdf = AllergenPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font('Arial', '', 10)
    
    # 상단 요약 정보 (제품명 등)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 8, f"Product: {prod_name}", 0, 1)
    pdf.cell(0, 8, f"Date: {p_date}  |  File: {file_name}", 0, 1)
    pdf.ln(5)
    
    # 테이블 헤더 설정
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font('Arial', 'B', 10)
    # 컬럼 너비 설정 (합계 277mm 내외)
    cols = [("No", 15), ("CAS No", 50), ("Ingredient Name", 100), ("Source", 35), ("Result", 35), ("Status", 30)]
    
    for col_name, width in cols:
        pdf.cell(width, 10, col_name, 1, 0, 'C', True)
    pdf.ln()
    
    # 테이블 데이터 입력
    pdf.set_font('Arial', '', 9)
    for _, row in df.iterrows():
        pdf.cell(cols[0][1], 8, str(row['번호']), 1, 0, 'C')
        pdf.cell(cols[1][1], 8, str(row['CAS']), 1, 0, 'C')
        # 글자 너무 길면 잘림 방지 (간략화)
        ing_name = str(row['물질명']).encode('latin-1', 'ignore').decode('latin-1')
        pdf.cell(cols[2][1], 8, ing_name[:55], 1, 0, 'L')
        pdf.cell(cols[3][1], 8, str(row['원본']), 1, 0, 'C')
        pdf.cell(cols[4][1], 8, str(row['양식']), 1, 0, 'C')
        
        # 상태 표시 (OK/FAIL)
        status_text = "OK" if "✅" in str(row['상태']) else "FAIL"
        if status_text == "FAIL":
            pdf.set_text_color(255, 0, 0) # 불일치는 빨간색
        pdf.cell(cols[5][1], 8, status_text, 1, 1, 'C')
        pdf.set_text_color(0, 0, 0) # 다시 검정색으로

    return pdf.output(dest='S').encode('latin-1')

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
st.title("🧪 콜마 83 ALLERGENS 통합 검토 시스템")
st.info("파일 순서를 맞추면 동일 순번끼리 매칭됩니다. 검토 후 PDF로 저장하세요.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 원본 파일 목록")
    uploaded_src = st.file_uploader("원본 선택 (xlsx)", type=["xlsx"], accept_multiple_files=True, key="src")
    src_file_list = []
    if uploaded_src:
        sorted_src = sort_items([f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src)])
        for name in sorted_src:
            orig = name.split(". ", 1)[1]
            src_file_list.append(next(f for f in uploaded_src if f.name == orig))

with col2:
    st.subheader("2. 양식(Result) 파일 목록")
    uploaded_res = st.file_uploader("양식 선택 (xlsx)", type=["xlsx"], accept_multiple_files=True, key="res")
    res_file_list = []
    if uploaded_res:
        sorted_res = sort_items([f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_res)])
        for name in sorted_res:
            orig = name.split(". ", 1)[1]
            res_file_list.append(next(f for f in uploaded_res if f.name == orig))

st.markdown("---")

# 4. 검증 로직 및 결과 출력
if src_file_list and res_file_list:
    num_pairs = min(len(src_file_list), len(res_file_list))
    all_pdfs = [] # 일괄 다운로드용
    
    for idx in range(num_pairs):
        src_f, res_f = src_file_list[idx], res_file_list[idx]
        mode = "HP" if "HP" in src_f.name.upper() else "CFF"
        
        try:
            wb_s, wb_r = load_workbook(src_f, data_only=True), load_workbook(res_f, data_only=True)
            ws_s = wb_s[next((s for s in wb_s.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_s.sheetnames[0])]
            ws_r = wb_r[next((s for s in wb_r.sheetnames if 'ALLERGY' in s.upper()), wb_r.sheetnames[0])]

            # 데이터 맵 생성 (생략된 기존 로직과 동일)
            s_map, r_map = {}, {}
            if mode == "CFF":
                p_name, p_date = str(ws_s['D7'].value or "N/A"), str(ws_s['N9'].value or "N/A").split(' ')[0]
                for r in range(13, 96):
                    c, v = get_cas_set(ws_s.cell(row=r, column=6).value), ws_s.cell(row=r, column=12).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=2).value, "v": float(v)}
            else:
                p_name, p_date = str(ws_s['B10'].value or "N/A"), str(ws_s['E10'].value or "N/A").split(' ')[0]
                for r in range(1, 401):
                    c, v = get_cas_set(ws_s.cell(row=r, column=2).value), ws_s.cell(row=r, column=3).value
                    if c and v is not None and v != 0: s_map[c] = {"n": ws_s.cell(row=r, column=1).value, "v": float(v)}

            rp_name, rp_date = str(ws_r['B10'].value or "N/A"), str(ws_r['E10'].value or "N/A").split(' ')[0]
            for r in range(1, 401):
                c, v = get_cas_set(ws_r.cell(row=r, column=2).value), ws_r.cell(row=r, column=3).value
                if c and v is not None and v != 0: r_map[c] = {"n": ws_r.cell(row=r, column=1).value, "v": float(v)}

            all_cas = sorted(list(set(s_map.keys()) | set(r_map.keys())), key=lambda x: list(x)[0] if x else "")
            rows = []
            mismatch = 0
            for i, c in enumerate(all_cas, 1):
                sv, rv = s_map.get(c, {}).get('v', "누락"), r_map.get(c, {}).get('v', "누락")
                match = (sv != "누락" and rv != "누락" and abs(sv - rv) < 0.0001)
                if not match: mismatch += 1
                rows.append({"번호": i, "CAS": ", ".join(list(c)), "물질명": r_map.get(c,{}).get('n') or s_map.get(c,{}).get('n'), "원본": sv, "양식": rv, "상태": "✅" if match else "❌"})

            df_res = pd.DataFrame(rows)
            
            # --- 결과 화면 ---
            with st.expander(f"[{idx+1}번] {res_f.name} (불일치: {mismatch})"):
                st.dataframe(df_res, use_container_width=True, hide_index=True)
                
                # PDF 생성
                pdf_bytes = create_pdf(df_res, rp_name, rp_date, res_f.name)
                st.download_button(f"📄 {rp_name} PDF 저장", pdf_bytes, f"Result_{idx+1}.pdf", "application/pdf", key=f"btn_{idx}")
                all_pdfs.append({"name": f"Result_{idx+1}_{rp_name}.pdf", "data": pdf_bytes})

            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 파일 처리 중 오류: {e}")

    # --- 전체 다운로드 ---
    if all_pdfs:
        st.markdown("---")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for p in all_pdfs: zf.writestr(p["name"], p["data"])
        st.download_button("📥 모든 결과 PDF 일괄 다운로드 (ZIP)", zip_buf.getvalue(), "All_Allergy_Reports.zip", "application/zip")
