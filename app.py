import streamlit as st
import docx
import re
import html
import io
import zipfile

# ==========================================
# [기능 1] Word 파일을 읽어서 문항별로 쪼개는 기능 (문항 저장 로직 완벽 수정)
# ==========================================
def parse_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  # 문제 번호 매칭 (예: 1. 다음 중)
    opt_pattern = re.compile(r'^([①②③④⑤])\s*(.*)') # 선택지 매칭 (예: ① grows)
    meta_pattern = re.compile(r'^\[Chapter') # 단원 태그 매칭
    arrow_pattern = re.compile(r'^[→↳\s]+(.*)') # 17번 형태의 화살표 하위 문장 매칭
    
    all_elements = []
    
    # 1차 파싱: 일반 문단과 표 내부 요소를 순서대로 정밀 추출
    for element in doc.element.body:
        if element.tag.endswith('p'): # 일반 문단일 때
            p = docx.text.paragraph.Paragraph(element, doc)
            
            # 서식을 하나씩 체크하여 밑줄(underline)이 있다면 <u> 태그 삽입
            text_with_formatting = ""
            for run in p.runs:
                run_text = run.text
                if run.underline and run_text.strip():
                    text_with_formatting += f"<u>{run_text}</u>"
                else:
                    text_with_formatting += run_text
            
            txt = text_with_formatting.strip()
            if txt:
                all_elements.append({"type": "text", "text": txt})
                
        elif element.tag.endswith('tbl'): # 📦 표(네모 상자 지문)일 때
            t = docx.table.Table(element, doc)
            table_lines = []
            
            for row in t.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        cell_txt_formatted = ""
                        for run in paragraph.runs:
                            if run.underline and run.text.strip():
                                cell_txt_formatted += f"<u>{run.text}</u>"
                            else:
                                cell_txt_formatted += run.text
                                
                        ctxt = cell_txt_formatted.strip()
                        if ctxt and ctxt not in table_lines:
                            table_lines.append(ctxt)
            
            if table_lines:
                all_elements.append({"type": "table", "lines": table_lines})

    # 2차 파싱: 수집된 요소를 바탕으로 문항 객체 분할 정렬
    idx = 0
    while idx < len(all_elements):
        item = all_elements[idx]
        
        # [A] 표(네모 박스) 지문 처리
        if item["type"] == "table":
            if current_q is not None:
                for t_line in item["lines"]:
                    converted_t_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', t_line)
                    if "The following table:" in converted_t_line:
                        continue
                    current_q["sentence"].append(converted_t_line)
            idx += 1
            continue
            
        # [B] 일반 텍스트 라인 처리
        line = item["text"]
        
        # 단원 메타 정보 라인을 만났을 때
        if meta_pattern.match(line):
            if current_q and (current_q["num"] or current_q["title"]):
                questions.append(current_q)
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            idx += 1
            continue
            
        # ★ [핵심 수정] 신규 문제 번호를 만났을 때 처리 규칙
        q_match = q_pattern.match(line)
        if q_match:
            # 이미 수집 중이던 기존 문항이 있다면 안전하게 리스트에 마감 저장(append)
            if current_q and (current_q["num"] or current_q["title"]):
                questions.append(current_q)
            
            # 새 문항 방을 새로 개설
            current_q = {
                "num": q_match.group(1),
                "title": q_match.group(2),
                "sentence": [],
                "options": []
            }
            idx += 1
            continue
            
        # 정답 및 해설 라인은 지문이나 보기에 포함되지 않도록 패스
        if line.startswith("정답:"):
            idx += 1
            continue
            
        if current_q is None:
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            
        # 선택지 보기를 만났을 때 (①~⑤)
        opt_match = opt_pattern.match(line)
        if opt_match:
            label_char = opt_match.group(1)
            opt_text = opt_match.group(2)
            
            label_map = {"①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5"}
            label_num = label_map.get(label_char, "1")
            
            processed_opt = re.sub(r'\s{2,}', ' &nbsp;&nbsp;-&nbsp;&nbsp; ', opt_text)
            
            current_q["options"].append({
                "char": label_char,
                "num": label_num,
                "text": processed_opt
            })
            idx += 1
            continue
            
        # 화살표(→) 하위 문장 처리 규칙 (직전 보기에 합침)
        arrow_match = arrow_pattern.match(line)
        if arrow_match and current_q["options"]:
            current_q["options"][-1]["text"] += f" <br/> {line}"
            idx += 1
            continue
            
        if "The following table:" in line:
            idx += 1
            continue
        
        # 일반 지문 추가 및 공백 라인 제외
        if current_q is not None and line.strip():
            converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
            current_q["sentence"].append(converted_line)
            
        idx += 1
        
    # 마지막 문항 마감 보존 처리
    if current_q and (current_q["num"] or current_q["title"]):
        questions.append(current_q)
        
    return questions


# ==========================================
# [기능 2] '단 한 개의 문항'만 가지고 규격 HTML 서식을 만드는 기능
# ==========================================
def generate_single_html(q):
    html_content = """<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta charset="UTF-8" name="viewport" content="width=device-width, target-densitydpi=device-dpi" />
<meta HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no,minimum-scale=1.0,maximum-scale=1.0,target-densitydpi=device-dpi">
<title>Achievement Test</title>
<link rel="stylesheet" href="../../../common/css/common.css">
<link rel="stylesheet" href="../../../common/css/style.css">
<link rel="stylesheet" href="../../../common_ex/css/style.css">
<link rel="stylesheet" href="../../../common_ex/css/scroll.css">
<script type="text/javascript" charset="UTF-8" src="../../../common/js/jquery.js"></script>
<script type="text/javascript" charset="UTF-8" src="../../../common/js/common.js"></script>
<script type="text/javascript" charset="UTF-8" src="../../../common_ex/js/common.js"></script>
</head>
<body>
<div class="pageWrap">
\t<div class="listening_desc_box"><b>Question 1-20</b> 질문을 읽고 물음에 답하시오.</div>
\t<div id="L_question" class="STSection">
\t\t<div class="pageConts">
"""

    html_content += f"""\t\t\t<div class="q_box">
\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="2" gravity="top|left">
\t\t\t\t\t<tr>
\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t<td>{html.escape(q['title'])} <br></td>
\t\t\t\t\t</tr>\n"""
    
    if q["sentence"]:
        sentence_br = " <br/> \n".join(q["sentence"])
        html_content += f"""\t\t\t\t\t<tr>
\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t<td>
\t\t\t\t\t\t\t<div class="sentence">{sentence_br} \n\t\t\t\t\t\t\t</div>
\t\t\t\t\t\t</td>
\t\t\t\t\t</tr>\n"""
        
    for opt in q["options"]:
        html_content += f"""\t\t\t\t\t<tr>
\t\t\t\t\t\t<td></td>
\t\t\t\t\t\t<td><span class="STChoice" remarkable="true" label="{opt['num']}"><span class="label">{opt['char']}</span> {opt['text']} </span></td>
\t\t\t\t\t</tr>\n"""
        
    html_content += """\t\t\t\t</table>
\t\t\t</div>\n"""

    html_content += """\t\t</div>
\t</div>
</div>
</body>
</html>"""

    return html_content


# ==========================================
# [기능 3] 웹 화면 레이아웃 및 다운로드 제어
# ==========================================

st.set_page_config(page_title="주차 연동 및 색상 지정 시스템", layout="wide")

with st.sidebar:
    st.header("⚙️ 고정값 설정")
    st.text_input("service_code", value="SVC170")
    st.text_input("track_code", value="RSV_TRK01")
    st.text_input("top_cors_id", value="1879")
    st.text_input("level_code", value="TO_R_E_SP")
    st.text_input("component_code", value="COM170")
    st.text_input("book_code", value="SVC170")
    st.text_input("act_name", value="Vocabulary")

st.title("🗂️ 주차 연동 및 색상 지정 시스템")
st.caption("Word 정기평가 파일을 업로드하면 서식 교정 및 전 문항을 독립 폴더 압축파일(.zip)로 자동 분할 생성합니다.")

uploaded_file = st.file_uploader("워드 파일(.docx)을 업로드하세요", type=["docx"])
submit_button = st.button("🚀 번호별 폴더 구조로 분할 변환하기", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("서식 교정 및 전 문항 파싱 작업을 정밀 수행하고 있습니다..."):
            parsed_data = parse_docx(uploaded_file)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for q in parsed_data:
                    q_num = q["num"]
                    if not q_num:
                        continue
                    
                    single_html = generate_single_html(q)
                    folder_file_path = f"{q_num}/test.html"
                    zip_file.writestr(folder_file_path, single_html)
            
            zip_data = zip_buffer.getvalue()
            
        st.success(f"🎉 성공적으로 총 {len(parsed_data)}개의 전체 문항 분석을 완료하여 개별 독립 배치를 완수했습니다!")
        
        st.subheader("📂 패키지 파일 내보내기")
        st.download_button(
            label="📥 번호별 폴더 압축파일(.zip) 다운로드",
            data=zip_data,
            file_name="questions_folders.zip",
            mime="application/zip"
        )
        st.info(f"💡 다운로드 링크가 정상 표출되었습니다. 분석된 문항 수: 총 {len(parsed_data)}개")
            
    except Exception as e:
        st.error(f"⚠️ 시스템 오류가 발생했습니다: {e}")
