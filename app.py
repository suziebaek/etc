import streamlit as st
import docx
import re
import html

# ==========================================
# [기능 1] Word 파일을 읽어서 문항별로 쪼개는 기능
# ==========================================
def parse_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    # 글자들을 구별해내는 규칙 패턴들
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  # 번호 뒤에 점(.)이 없어도 인식하도록 보완
    opt_pattern = re.compile(r'^([①②③④⑤])\s*(.*)')
    meta_pattern = re.compile(r'^\[Chapter')
    
    # 문서에서 빈 줄을 제외하고 글자만 추출
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # [Chapter... 로 시작하면 새로운 문제 준비하기
        if meta_pattern.match(line):
            if current_q:
                questions.append(current_q)
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            i += 1
            continue
            
        if current_q is None:
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            
        # 1. 2. 같은 문제 번호와 발문 찾기
        q_match = q_pattern.match(line)
        if q_match:
            current_q["num"] = q_match.group(1)
            current_q["title"] = q_match.group(2)
            i += 1
            continue
            
        # ① ② 같은 보기 번호와 선택지 찾기
        opt_match = opt_pattern.match(line)
        if opt_match:
            label_char = opt_match.group(1)
            opt_text = opt_match.group(2)
            
            # 원문자를 숫자 시스템(1~5)으로 변경
            label_map = {"①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5"}
            label_num = label_map.get(label_char, "1")
            
            # 글자 사이의 넓은 공백을 시험지 스타일(  -  )로 예쁘게 정돈
            processed_opt = re.sub(r'\s{2,}', ' &nbsp;&nbsp;-&nbsp;&nbsp; ', opt_text)
            
            current_q["options"].append({
                "char": label_char,
                "num": label_num,
                "text": processed_opt
            })
            i += 1
            continue
            
        # 보기 지문(The following table:) 처리 및 밑줄(___)을 웹용 코드로 변경
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
        
        # 일반 지문 문장 처리
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
        current_q["sentence"].append(converted_line)
        i += 1
        
    if current_q and (current_q["num"] or current_q["title"]):
        questions.append(current_q)
        
    return questions


# ==========================================
# [기능 2] 쪼갠 문항들을 멋진 HTML 양식으로 조립하는 기능
# ==========================================
def generate_html(questions):
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

\t<div class="listening_desc_box"><b>Question 1-20</b> 질문을 읽고물음에 답하시오.</div>

\t\t<div id="L_question" class="STSection">
\t\t<div class="pageConts">
"""

    for q in questions:
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
\t\t\t</div>\n\n"""

    html_content += """\t\t</div>
\t</div>
\t</div>
</body>
</html>"""

    return html_content


# ==========================================
# [기능 3] 웹 화면 디자인 구역
# ==========================================

st.set_page_config(page_title="주차 연동 및 색상 지정 시스템", layout="wide")

# [왼쪽 사이드바] 기존 설정창 유지
with st.sidebar:
    st.header("⚙️ 고정값 설정")
    st.text_input("service_code", value="SVC170")
    st.text_input("track_code", value="RSV_TRK01")
    st.text_input("top_cors_id", value="1879")
    st.text_input("level_code", value="TO_R_E_SP")
    st.text_input("component_code", value="COM170")
    st.text_input("book_code", value="SVC170")
    st.text_input("act_name", value="Vocabulary")

# [메인 화면] 타이틀
st.title("🗂️ 주차 연동 및 색상 지정 시스템")
st.caption("Word 정기평가 파일을 업로드하고 버튼을 누르면 정해진 규격의 HTML 파일로 변환합니다.")

# 파일 업로드 상자
uploaded_file = st.file_uploader("워드 파일(.docx)을 업로드하세요", type=["docx"])

# ⭐ 초보자 전용 [HTML 변환 시작] 버튼 추가!
submit_button = st.button("🚀 HTML 코드로 변환하기", type="primary")

# 파일이 올라갔고 + '변환하기' 버튼을 눌렀을 때만 작동하도록 조건 변경
if uploaded_file is not None and submit_button:
    try:
        with st.spinner("Word 문서를 분석하여 웹 표준 코드로 바꾸는 중입니다..."):
            parsed_data = parse_docx(uploaded_file)
            final_html = generate_html(parsed_data)
            
        st.success(f"🎉 성공적으로 {len(parsed_data)}개의 문항을 변환했습니다! (1번 ~ {len(parsed_data)}번 완료)")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🖥️ 완성된 HTML 소스 코드")
            st.code(final_html, language="html", line_numbers=True)
            
        with col2:
            st.subheader("📂 시스템 파일 내보내기")
            st.download_button(
                label="📥 HTML 파일 다운로드 받기",
                data=final_html,
                file_name="test.html",
                mime="text/html"
            )
            st.info("💡 위 다운로드 버튼을 누르면 `test.html` 파일이 컴퓨터에 저장됩니다.")
            
    except Exception as e:
        st.error(f"⚠️ 파일 분석 중 오류가 발생했습니다: {e}")