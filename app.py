import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import io
import zipfile
from streamlit_sortables import sort_items
from xhtml2pdf import pisa  # HTML -> PDF 변환 라이브러리

# 1. 화면 설정
st.set_page_config(page_title="콜마 83 알러지 통합 검토", layout="wide")

# --- 엑셀 인쇄 양식을 재현하는 PDF 생성 함수 ---
def create_pdf_from_df(df, prod_name, p_date, file_name):
    # CSS를 사용하여 엑셀의 '인쇄 영역 설정'과 '열 맞춤'을 구현
    html_content = f"""
    <html>
    <head>
        <style>
            @page {{ 
                size: a4 landscape; /* 가로 출력으로 열 잘림 방지 */
                margin: 1cm; 
            }}
            body {{ font-family: helvetica; font-size: 9pt; color: #333; }}
            .header-title {{ text-align: center; font-size: 16pt; font-weight: bold; margin-bottom: 15px; }}
            .summary-table {{ width: 100%; margin-bottom: 15px; border-bottom: 1px solid #000; }}
            .summary-table td {{ padding: 3px; }}
            .main-table {{ width: 100%; border-collapse: collapse; }}
            .main-table th {{ background-color: #f2f2f2; border: 1px solid #444; padding: 5px; font-weight: bold; }}
            .main-table td {{ border: 1px solid #666; padding: 4px; text-align: center; }}
            .status-ok {{ color: green; }}
            .status-fail {{ color: red; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="header-title">Allergen Review Report</div>
        <table class="summary-table">
            <tr>
                <td><b>Product:</b> {prod_name}</td>
                <td><b>Date:</b> {p_date}</td>
            </tr>
            <tr>
                <td colspan="2"><b>Source File:</b> {file_name}</td>
            </tr>
        </table>
        <table class="main-table">
            <thead>
                <tr>
                    <th>No</th><th>CAS No</th><th>Ingredient Name</th><th>Source Val</th><th>Result Val</th><th>Status</th>
                </tr>
            </thead>
            <tbody>
                {"".join(f"<tr><td>{r['번호']}</td><td>{r['CAS']}</td><td>{r['물질명']}</td><td>{r['원본']}</td><td>{r['양식']}</td><td>{r['상태']}</td></tr>" for _, r in df.iterrows())}
            </tbody>
        </table>
    </body>
    </html>
    """
    pdf_buffer = io.BytesIO()
    pisa.CreatePDF(html_content, dest=pdf_buffer)
    return pdf_buffer.getvalue()

# 2. 공통 도구 함수 (기존과 동일)
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
st.info("파일 순서를 맞추면 동일 순번끼리 매칭됩니다. 검토 완료 후 PDF로 저장할 수 있습니다.")

st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 원본 파일 목록")
    uploaded_src_files = st.file_uploader("원본 선택 (다중)", type=["xlsx"], accept_multiple_files=True, key="src_upload")
    src_file_list = []
    if uploaded_src_files:
        file_display_names = [f"↕ {i+1}. {f.name}" for i, f in enumerate(uploaded_src_files)]
        sorted_names = sort_items(file_display_names)
        for name in sorted_names:
            orig = name.split(". ", 1)[1]
            src_file_list.append(next(f for f in uploaded_src_files if f.name == orig))

with col2:
    st.subheader("2. 양식(Result) 파일 목록")
    uploaded_res_files = st.file_uploader("양식 선택 (다중)", type=["xlsx"], accept_multiple_files=True, key="res_upload")
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
    all_pdfs = [] # 일괄 다운로드용 리스트
    
    for idx in range(num_pairs):
        src_f, res_f = src_file_list[idx], res_file_list[idx]
        mode = "HP" if "HP" in src_f.name.upper() else "CFF"
        
        try:
            wb_s, wb_r = load_workbook(src_f, data_only=True), load_workbook(res_f, data_only=True)
            ws_s = wb_s[next((s for s in wb_s.sheetnames if 'ALLERGEN' in s.upper() or 'Sheet' in s), wb_s.sheetnames[0])]
            ws_r = wb_r[next((s for s in wb_r.sheetnames if 'ALLERGY' in s.upper()), wb_r.sheetnames[0])]

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

            rows = []
            mismatch = 0
            for i, c in enumerate(sorted(list(set(s_map.keys())|set(r_map.keys())), key=lambda x: list(x)[0] if x else ""), 1):
                sv, rv = s_map.get(c,{}).get('v',"누락"), r_map.get(c,{}).get('v',"누락")
                match = (sv != "누락" and rv != "누락" and abs(sv - rv) < 0.0001)
                if not match: mismatch += 1
                rows.append({"번호": i, "CAS": ", ".join(list(c)), "물질명": r_map.get(c,{}).get('n') or s_map.get(c,{}).get('n'), "원본": sv, "양식": rv, "상태": "✅" if match else "❌"})

            df_res = pd.DataFrame(rows)
            
            # --- 결과 표시 ---
            status_icon = "✅" if mismatch == 0 else "❌"
            with st.expander(f"{status_icon} [{idx+1}번] {res_f.name} (불일치: {mismatch}건)"):
                m1, m2 = st.columns(2)
                with m1: st.success(f"**원본:** {p_name}\n\n**작성일:** {p_date}")
                with m2: st.info(f"**양식:** {rp_name}\n\n**작성일:** {rp_date}")
                
                st.dataframe(df_res, use_container_width=True, hide_index=True)
                
                # PDF 생성 및 다운로드 버튼
                pdf_data = create_pdf_from_df(df_res, rp_name, rp_date, res_f.name)
                st.download_button("📄 PDF로 저장 (양식 맞춤)", pdf_data, f"Result_{rp_name}.pdf", "application/pdf", key=f"dl_{idx}")
                all_pdfs.append({"name": f"Result_{rp_name}.pdf", "data": pdf_data})

            wb_s.close(); wb_r.close()
        except Exception as e:
            st.error(f"{idx+1}번 오류: {e}")

    # --- 일괄 다운로드 ---
    if all_pdfs:
        st.markdown("---")
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as zf:
            for p in all_pdfs: zf.writestr(p["name"], p["data"])
        st.download_button("📥 모든 결과 PDF 일괄 다운로드 (ZIP)", zip_io.getvalue(), "All_Allergy_Reports.zip", "application/zip")

else:
    st.info("파일들을 업로드해 주세요.")
