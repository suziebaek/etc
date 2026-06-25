import streamlit as st
import docx
import re
import html
import io
import zipfile

# ==========================================
# [기능 1] Word 파일을 읽어서 문항별로 쪼개는 기능 (표 인식 오류 수정)
# ==========================================
def parse_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  # 문제 번호 매칭
    opt_pattern = re.compile(r'^([①②③④⑤])\s*(.*)') # 선택지 매칭
    meta_pattern = re.compile(r'^\[Chapter') # 단원 태그 매칭
    
    all_elements = []
    
    # 워드 문서 내부의 모든 요소(문단 또는 표)를 순서대로 순회합니다.
    for element in doc.element.body:
        if element.tag.endswith('p'): # 일반 문단일 때
            p = docx.text.paragraph.Paragraph(element, doc)
            txt = p.text.strip()
            if txt:
                all_elements.append({"type": "text", "text": txt})
        elif element.tag.endswith('tbl'): # 📦 표(네모 상자)일 때
            t = docx.table.Table(element, doc)
            table_lines = []
            
            # ★ [오류 수정 완료] 표의 행(row)을 돌며 각 칸(cell) 내부의 문단들을 안전하게 수집합니다.
            for row in t.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        ctxt = paragraph.text.strip()
                        if ctxt and ctxt not in table_lines:
                            table_lines.append(ctxt)
            
            # 표 내부 텍스트가 존재하면 하나의 '표 지문' 덩어리로 저장합니다.
            if table_lines:
                all_elements.append({"type": "table", "lines": table_lines})

    idx = 0
    while idx < len(all_elements):
        item = all_elements[idx]
        
        # 1. 표(네모 상자) 데이터 무리를 만났을 때 처리
        if item["type"] == "table":
            if current_q is not None:
                for t_line in item["lines"]:
                    # 밑줄(___)을 웹용 span 코드로 자동 치환
                    converted_t_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', t_line)
                    
                    # 쉼표(,)로 연결된 다중 지문이 있다면 개별 줄바꿈으로 쪼개어 담기
                    if "The following table:" in converted_t_line:
                        continue
                    if ',' in t_line and not opt_pattern.match(t_line):
                        for sub_t in t_line.split(','):
                            if sub_t.strip():
                                converted_sub = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', sub_t.strip())
                                current_q["sentence"].append(converted_sub)
                    else:
                        current_q["sentence"].append(converted_t_line)
            idx += 1
            continue
            
        # 2. 일반 텍스트 문장일 때 처리
        line = item["text"]
        
        if meta_pattern.match(line):
            if current_q:
                questions.append(current_q)
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            idx += 1
            continue
            
        if current_q is None:
            current_q = {"num": "", "title": "", "sentence": [], "options": []}
            
        q_match = q_pattern.match(line)
        if q_match:
            current_q["num"] = q_match.group(1)
            current_q["title"] = q_match.group(2)
            idx += 1
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
            idx += 1
            continue
            
        if "The following table:" in line:
            idx += 1
            if idx < len(all_elements) and all_elements[idx]["type"] == "text":
                sub_lines = [s.strip() for s in all_elements[idx]["text"].split(',')]
                for sl in sub_lines:
                    if sl:
                        converted_sl = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', sl)
                        current_q["sentence"].append(converted_sl)
                idx += 1
            continue
        
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
        current_q["sentence"].append(converted_line)
        idx += 1
        
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
st.caption("Word 정기평가 파일을 업로드하면 네모 상자(표) 지문까지 완벽히 파싱하여 ZIP 압축 파일로 자동 분할 생성합니다.")

uploaded_file = st.file_uploader("워드 파일(.docx)을 업로드하세요", type=["docx"])
submit_button = st.button("🚀 번호별 폴더 구조로 분할 변환하기", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("Word 파일을 쪼개어 번호별 독립 폴더 세트를 구축하는 중입니다..."):
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
            
        st.success(f"🎉 성공적으로 {len(parsed_data)}개의 문항을 분석하여 번호별 독립 폴더 구조 배치를 완료했습니다!")
        
        st.subheader("📂 패키지 파일 내보내기")
        st.download_button(
            label="📥 번호별 폴더 압축파일(.zip) 다운로드",
            data=zip_data,
            file_name="questions_folders.zip",
            mime="application/zip"
        )
        st.info("💡 압축을 풀면 표 내부 지문이 <div class='sentence'> 태그로 안전하게 감싸진 개별 test.html 파일들을 확인할 수 있습니다.")
            
    except Exception as e:
        st.error(f"⚠️ 시스템 오류가 발생했습니다: {e}")
