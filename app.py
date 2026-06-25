import streamlit as st
import docx
import re
import html
import io
import zipfile

# ==========================================
# [기능 1] Word 파일을 읽어서 문항별로 쪼개는 기능
# ==========================================
def parse_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  # 번호 매칭
    opt_pattern = re.compile(r'^([①②③④⑤])\s*(.*)') # 보기 매칭
    meta_pattern = re.compile(r'^\[Chapter')
    
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        if meta_pattern.match(line):
            if current_q:
                questions.append(current_q)
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            i += 1
            continue
            
        if current_q is None:
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            
        q_match = q_pattern.match(line)
        if q_match:
            current_q["num"] = q_match.group(1)
            current_q["title"] = q_match.group(2)
            i += 1
            continue
            
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
            i += 1
            continue
            
        if "The following table:" in line:
            i += 1
            if i < len(lines):
                sub_lines = [s.strip() for s in lines[i].split(',')]
                for sl in sub_lines:
                    if sl:
                        converted_sl = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', sl)
                        current_q["sentence"].append(converted_sl)
                i += 1
            continue
        
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
        current_q["sentence"].append(converted_line)
        i += 1
        
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

    # 해당 문제 1개만 바인딩
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

# [왼쪽 사이드바] 고정값 설정 영역
with st.sidebar:
    st.header("⚙️ 고정값 설정")
    st.text_input("service_code", value="SVC170")
    st.text_input("track_code", value="RSV_TRK01")
    st.text_input("top_cors_id", value="1879")
    st.text_input("level_code", value="TO_R_E_SP")
    st.text_input("component_code", value="COM170")
    st.text_input("book_code", value="SVC170")
    st.text_input("act_name", value="Vocabulary")

# [메인 화면]
st.title("🗂️ 주차 연동 및 색상 지정 시스템")
st.caption("Word 정기평가 파일을 업로드하면 문항별 폴더 구조를 가진 ZIP 압축 파일로 자동 분할 생성합니다.")

# 파일 업로드 상자
uploaded_file = st.file_uploader("워드 파일(.docx)을 업로드하세요", type=["docx"])

# 실행 버튼
submit_button = st.button("🚀 번호별 폴더 구조로 분할 변환하기", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("Word 파일을 쪼개어 번호별 독립 폴더 세트를 구축하는 중입니다..."):
            parsed_data = parse_docx(uploaded_file)
            
            # 📁 메모리 상에 ZIP 압축 파일을 만들기 위한 가상 상자 준비
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for q in parsed_data:
                    q_num = q["num"]
                    if not q_num:
                        continue
                    
                    # 1) 이 문항만의 개별 HTML 코드 생성
                    single_html = generate_single_html(q)
                    
                    # 2) 저장 경로를 '번호폴더/test.html' 형태로 동적 지정 (예: 1/test.html, 2/test.html ...)
                    folder_file_path = f"{q_num}/test.html"
                    
                    # 3) 압축 파일 내부에 폴더와 함께 저장
                    zip_file.writestr(folder_file_path, single_html)
