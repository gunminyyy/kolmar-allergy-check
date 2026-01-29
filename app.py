import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
import zipfile
from fpdf import FPDF

# 1. 화면 설정
st.set_page_config(page_title="콜마 83 알러지 통합 검토", layout="wide")

# --- PDF 생성 클래스 (fpdf 사용) ---
class AllergenPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'Allergen Review Report', 0, 1, 'C')
        self.ln(5)

def create_pdf(df, prod_name, p_date, file_name):
    # L: 가로방향 (열 맞춤을 위해 필수)
    pdf = AllergenPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 8, f"Product: {prod_name}", 0, 1)
    pdf.cell(0, 8, f"Date: {p_date}  |  File: {file_name}", 0, 1)
    pdf.ln(5)
    
    # 테이블 헤더
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font('Arial', 'B', 10)
    cols = [("No", 15), ("CAS No", 50), ("Ingredient Name", 100), ("Src Val", 35), ("Res Val", 35), ("Status", 30)]
    for col_name, width in cols:
        pdf.cell(width, 10, col_name, 1, 0, 'C', True)
    pdf.ln()
    
    # 테이블 데이터
    pdf.set_font('Arial', '', 9)
    for _, row in df.iterrows():
        pdf.cell(cols[0][1], 8, str(row['번호']), 1, 0, 'C')
        pdf.cell(cols[1][1], 8, str(row['CAS 번호']), 1, 0, 'C')
        # 한글 깨짐 방지를 위해 인코딩 처리 (데이터에 한글이 섞인 경우 공백 처리)
        ing_name = str(row['물질명']).encode('latin-1', 'ignore').decode('latin-1')
        pdf.cell(cols[2][1], 8, ing_name[:55], 1, 0, 'L')
        pdf.cell(cols[3][1], 8, str(row['원본 수치']), 1, 0, 'C')
        pdf.cell(cols[4][1], 8, str(row['최종 수치']), 1, 0, 'C')
        
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

# 3. 메인 UI 구성
st.title("🧪 콜마 83 ALLERGENS 검토 시스템(HP,CFF)")
st.info("원본과 최종본 파일을 **동일한 순서**로 업로드하세요. 순서대로 매칭되어 검증 및 PDF 저장이 가능합니다.")

mode = st.radio("📂 원본 파일 양식을 선택하세요", ["CFF 양식", "HP 양식"], horizontal=True)
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    src_files = st.file_uploader(f"1. 원본({mode}) 파일들 업로드", type=["xlsx"], accept_multiple_files=True)
with col2:
    res_files = st.file_uploader("2. 최종본(Result) 파일들 업로드", type=["xlsx"], accept_multiple_files=True)

# 4. 검증 로직 실행
if src_files and res_files:
    if len(src_files) != len(res_files):
        st.warning(f"⚠️ 파일 개수 불일치: {min(len(src_files), len(res_files))}번까지만 비교합니다.")

    all_pdf_data = [] # 일괄 다운로드용

    for idx, (src_f, res_f) in enumerate(zip(src_files, res_files), 1):
        with st.expander(f"📋 {idx}번 매칭 결과: {src_f.name} ↔ {res_f.name}", expanded=True):
            try:
                wb_src = load_workbook(src_f, data_only=True)
                wb_res = load_workbook(res_f, data_only=True)
                
                src_sheet = next((s for s in wb_src.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_src.sheetnames[0])
                res_sheet = next((s for s in wb_res.sheetnames if 'ALLERGY' in s.upper()), wb_res.sheetnames[0])
                
                ws_src, ws_res = wb_src[src_sheet], wb_res[res_sheet]
                src_map, res_map = {}, {}

                # 데이터 수집 (사용자님의 기존 로직 그대로)
                if mode == "CFF 양식":
                    src_p, src_d = str(ws_src['D7'].value or "N/A"), str(ws_src['N9'].value or "N/A").split(' ')[0]
                    for r in range(13, 96):
                        c = get_cas_set(ws_src.cell(row=r, column=6).value)
                        v = ws_src.cell(row=r, column=12).value
                        if c and v is not None and v != 0: src_map[c] = {"name": ws_src.cell(row=r, column=2).value, "val": float(v)}
                else:
                    src_p, src_d = str(ws_src['B10'].value or "N/A"), str(ws_src['E10'].value or "N/A").split(' ')[0]
                    for r in range(1, 400):
                        c = get_cas_set(ws_src.cell(row=r, column=2).value)
                        v = ws_src.cell(row=r, column=3).value
                        if c and v is not None and v != 0: src_map[c] = {"name": ws_src.cell(row=r, column=1).value, "val": float(v)}

                res_p, res_d = str(ws_res['B10'].value or "N/A"), str(ws_res['E10'].value or "N/A").split(' ')[0]
                for r in range(1, 400):
                    c = get_cas_set(ws_res.cell(row=r, column=2).value)
                    v = ws_res.cell(row=r, column=3).value
                    if c and v is not None and v != 0: res_map[c] = {"name": ws_res.cell(row=r, column=1).value, "val": float(v)}

                # 비교 결과 생성
                all_cas = sorted(list(set(src_map.keys()) | set(res_map.keys())), key=lambda x: list(x)[0] if x else "")
                table_data = []
                match_count = 0
                for i, c in enumerate(all_cas, 1):
                    s_v, r_v = src_map.get(c, {}).get('val', "누락"), res_map.get(c, {}).get('val', "누락")
                    is_match = (s_v != "누락" and r_v != "누락" and abs(s_v - r_v) < 0.0001)
                    if is_match: match_count += 1
                    table_data.append({
                        "번호": i, "CAS 번호": ", ".join(list(c)), 
                        "물질명": res_map.get(c,{}).get('name') or src_map.get(c,{}).get('name') or "Unknown",
                        "원본 수치": s_v, "최종 수치": r_v, "상태": "✅ 일치" if is_match else "❌ 불일치"
                    })

                # 화면 출력
                df = pd.DataFrame(table_data)
                st.info(f"**원본:** {src_p} ({src_d}) / **최종:** {res_p} ({res_d})")
                st.dataframe(df, use_container_width=True, hide_index=True)
                st.metric(f"매칭 {idx} 결과", f"총 {len(df)}건", f"불일치 {len(df)-match_count}건", delta_color="inverse")

                # 개별 PDF 다운로드 버튼
                pdf_bytes = create_pdf(df, res_p, res_d, res_f.name)
                st.download_button(f"📄 {idx}번 결과 PDF 저장", pdf_bytes, f"Result_{idx}.pdf", "application/pdf", key=f"dl_{idx}")
                all_pdf_data.append({"name": f"Result_{idx}_{res_p}.pdf", "data": pdf_bytes})

                wb_src.close(); wb_res.close()
            except Exception as e:
                st.error(f"{idx}번 처리 오류: {e}")

    # 일괄 다운로드 (ZIP)
    if all_pdf_data:
        st.markdown("---")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for p in all_pdf_data: zf.writestr(p["name"], p["data"])
        st.download_button("📥 모든 결과 PDF 일괄 다운로드 (ZIP)", zip_buf.getvalue(), "All_Reports.zip", "application/zip")
else:
    st.info("파일들을 업로드해 주세요.")

