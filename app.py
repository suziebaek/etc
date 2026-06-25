import streamlit as st
import docx
import re
import html
import io
import zipfile

# ==========================================
# [기능 1] Word 파일을 읽어서 문항별로 쪼개는 기능 (강력한 필터링 보완)
# ==========================================
def parse_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    q_pattern = re.compile(r'^(\d+)\.?\s*(.*)')  # 문제 번호 매칭 (예: 1. 다음 중)
    opt_pattern = re.compile(r'^([①②③④⑤])\s*(.*)') # 선택지 매칭 (예: ① 달리다)
    arrow_pattern = re.compile(r'^[→↳\s]+(.*)') # 화살표 하위 문장 매칭
    
    # 💡 [보완] 문장 중간 어디든 출처 키워드가 포함되어 있으면 완전히 걸러내도록 무시 패턴 강화
    ignore_keywords = ["vocabulary", "reading", "chunking", "inside", "starter", "w1", "w2", "w3", "w4", "w5", "w6", "w7", "w8", "w9", "주차", "[part", "chapter"]
    
    all_elements = []
    for element in doc.element.body:
        # 최상위 body 바로 아래에 있는 요소만 순서대로 수집 (표 내부 문단 중복 처리 방지)
        if element.getparent() != doc.element.body:
            continue
            
        if element.tag.endswith('p'):  # 일반 문단일 때
            p = docx.text.paragraph.Paragraph(element, doc)
            text_with_formatting = ""
            for run in p.runs:
                if run.underline and run.text.strip():
                    text_with_formatting += f"<u>{run.text}</u>"
                else:
                    text_with_formatting += run.text
            
            txt = text_with_formatting.strip()
            if txt:
                all_elements.append({"type": "text", "text": txt})
                
        elif element.tag.endswith('tbl'):  # 📦 표(네모 상자 지문)일 때
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

    # 정렬된 요소를 바탕으로 문항 추출 수행
    for item in all_elements:
        if item["type"] == "table":
            if current_q is not None:
                for t_line in item["lines"]:
                    # 💡 표 내부 텍스트도 강력 필터링 검사 수행
                    t_line_lower = t_line.lower()
                    if any(k in t_line_lower for k in ignore_keywords):
                        continue
                        
                    converted_t_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', t_line)
                    if "the following table:" in converted_t_line.lower():
                        continue
                    current_q["sentence"].append(converted_t_line)
            continue

        line = item["text"].strip()
        line_lower = line.lower()

        # 1. 새로운 문제 번호를 만났을 때 (새로운 문제 방 개설)
        q_match = q_pattern.match(line)
        if q_match:
            if current_q is not None:
                questions.append(current_q)
            current_q = {
                "num": q_match.group(1),
                "title": q_match.group(2),
                "sentence": [],
                "options": []
            }
            continue

        # 💡 [핵심 보완] 방이 없거나, 줄바꿈 속에 출처 키워드가 단 하나라도 포함되어 있다면 즉시 통째로 패스!
        if current_q is None or any(k in line_lower for k in ignore_keywords):
            continue

        # 교사용 정답 라인 패스
        if line.startswith("정답:"):
            continue

        # 2. 선택지 보기(①~⑤)를 만났을 때
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
            continue

        # 3. 화살표(→) 하위 설명 문장을 만났을 때 (직전 보기에 병합)
        arrow_match = arrow_pattern.match(line)
        if arrow_match and current_q["options"]:
            current_q["options"][-1]["text"] += f" <br/> {line}"
            continue

        if "the following table:" in line_lower:
            continue

        # 4. 번호 안쪽에서 발견된 그 외의 모든 일반 문장은 지문으로 처리
        converted_line = re.sub(r'_{2,}', '<span class="underline" style="width:100px;"></span>', line)
        current_q["sentence"].append(converted_line)

    # 마지막 문항 추가 마감
    if current_q is not None:
        questions.append(current_q)

    return questions


# ==========================================
# [기능 2] '단 한 개의 문항'만 가지고 규격 HTML 서식을 만드는 기능 (태그 깨짐 수정)
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

    # 💡 [오류 수정] 타이틀 내부의 <u> 태그가 이스케이프되어 깨지지 않도록 안전 변환 로직을 우회/복원합니다.
    safe_title = html.escape(q['title'])
    safe_title = safe_title.replace("&lt;u&gt;", "<u>").replace("&lt;/u&gt;", "</u>")

    html_content += f"""\t\t\t<div class="q_box">
\t\t\t\t<table class="answer_txt STChooseAnAnswer L_tableQuestion" scale="190" answer="2" gravity="top|left">
\t\t\t\t\t<tr>
\t\t\t\t\t\t<td><span class="num STCorrectness">{q['num']}.</span></td>
\t\t\t\t\t\t<td>{safe_title} <br></td>
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
st.caption("Word 정기평가 파일을 업로드하면 불필요한 출처 기호를 완전히 청소하고 오직 문제 정보만 정밀 분할 생성합니다.")

uploaded_file = st.file_uploader("워드 파일(.docx)을 업로드하세요", type=["docx"])
submit_button = st.button("🚀 번호별 폴더 구조로 분할 변환하기", type="primary")

if uploaded_file is not None and submit_button:
    try:
        with st.spinner("출처 데이터 완전 필터링 및 지시문 서식 복원 중..."):
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
            
        st.success(f"🎉 출처 노이즈 제거 및 지시문 밑줄 서식 매칭 복원을 완료했습니다! (총 {len(parsed_data)}개 문항)")
        
        st.subheader("📂 패키지 파일 내보내기")
        st.download_button(
            label="📥 번호별 폴더 압축파일(.zip) 다운로드",
            data=zip_data,
            file_name="questions_folders.zip",
            mime="application/zip"
        )
        st.info(f"💡 이제 6번 이후에도 지문 상단에 출처명이 섞이지 않으며, 지시문의 '<u>않은</u>' 서식이 정상 표출됩니다.")
            
    except Exception as e:
        st.error(f"⚠️ 시스템 오류가 발생했습니다: {e}")
