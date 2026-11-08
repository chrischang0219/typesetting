import streamlit as st
import base64
import re
import time
import io
import os
import requests
import urllib3
import fitz  # PyMuPDF
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_TAB_ALIGNMENT

# 忽略 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# ⚙️ 全域字體與模型設定
# ==========================================
DOC_TITLE_FONT = "王漢宗顏楷體繁"
DOC_CHINESE_FONT = "標楷體"
DOC_ENGLISH_FONT = "Times New Roman"
VALID_MODELS = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]

st.set_page_config(page_title="考卷排版神器 Pro", page_icon="📄", layout="centered")

# ==========================================
# 🛠️ 核心排版與處理函式
# ==========================================
def parse_page_selection(s, max_p):
    if not s or not s.strip():
        return list(range(max_p))
    selected = set()
    try:
        for part in s.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                for p in range(max(1, start), min(max_p, end) + 1):
                    selected.add(p - 1)
            else:
                p = int(part)
                if 1 <= p <= max_p:
                    selected.add(p - 1)
        return sorted(list(selected)) if selected else list(range(max_p))
    except Exception:
        return list(range(max_p))

def custom_format_text(text):
    text = re.sub(r'</?[ub*]+>', '', text)
    
    # 統一轉換選項格式為 (A) (B) (C) (D)
    text = re.sub(r'^\s*([①1]|[\(\(]1[\)\)])\s*[.．、\s]*', '(A)', text)
    text = re.sub(r'^\s*([②2]|[\(\(]2[\)\)])\s*[.．、\s]*', '(B)', text)
    text = re.sub(r'^\s*([③3]|[\(\(]3[\)\)])\s*[.．、\s]*', '(C)', text)
    text = re.sub(r'^\s*([④4]|[\(\(]4[\)\)])\s*[.．、\s]*', '(D)', text)
    
    text = re.sub(r'^\s*[（\(]?A[）\)]?[.．、\s]*', '(A)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]?B[）\)]?[.．、\s]*', '(B)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]?C[）\)]?[.．、\s]*', '(C)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]?D[）\)]?[.．、\s]*', '(D)', text, flags=re.IGNORECASE)

    # 標點符號全形轉換
    fullwidth_map = {';': '；', '~': '～', ',': '，', '?': '？', '!': '！', ':': '：', '.': '．', '*': '×', '/': '÷'}
    for k, v in fullwidth_map.items():
        text = text.replace(k, v)
    text = text.replace('（', '(').replace('）', ')').replace(' ', ' ')
    return text

def set_run_font(run, size, is_english_mode, bold=False, use_title_font=False):
    run.font.size = Pt(size)
    run.bold = bold
    
    ch_font = DOC_TITLE_FONT if use_title_font else DOC_CHINESE_FONT
    ascii_font = DOC_ENGLISH_FONT if is_english_mode else ch_font
    
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    f = rPr.get_or_add_rFonts()
    f.set(qn('w:eastAsia'), ch_font)
    f.set(qn('w:ascii'), ascii_font)
    f.set(qn('w:hAnsi'), ascii_font)

def create_word_document(text, is_english_mode):
    doc = Document()
    settings = doc.settings.element
    compat = settings.find(qn('w:compat')) or OxmlElement('w:compat')
    settings.append(compat)
    ult = OxmlElement('w:ulTrailSpace')
    compat.append(ult)

    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1)

    header = "考卷排版結果"
    h_match = re.search(r"HEADER:\s*(.*)", text)
    if h_match:
        header = h_match.group(1).strip()
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if "|" in header:
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Cm(14), WD_TAB_ALIGNMENT.CENTER)
        tab_stops.add_tab_stop(Cm(19), WD_TAB_ALIGNMENT.RIGHT)
        parts = header.split("|")
        display_header = f"{parts[0].strip()}\t{parts[1].strip()}\t\xa0"
    else:
        display_header = header.replace(';', '；').replace('~', '～')

    run_h = p.add_run(display_header)
    run_h.underline = WD_UNDERLINE.SINGLE
    set_run_font(run_h, size=16, is_english_mode=is_english_mode, bold=True, use_title_font=True)

    content_match = re.search(r"CONTENT_START(.*)CONTENT_END", text, re.DOTALL)
    main_content = content_match.group(1).strip() if content_match else text

    for line in main_content.split('\n'):
        line = custom_format_text(line)
        if not line.strip():
            continue
        
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(2)

        match_q = re.match(r'^\s*(\d+)\s*[.、．]\s*(.*)', line)
        match_opt = re.match(r'^\s*([A-Da-d①②③④1-4]\s*[.．、\)]|[\(\(][A-Da-d1-4][\)\)])\s*(.*)', line)

        if match_q:
            q_num = int(match_q.group(1))
            q_content = match_q.group(2).strip()
            
            is_drawing = any(k in q_content for k in ["作圖", "畫出", "計算", "證明"])
            is_fill = "＿＿" in q_content or "{BLANK}" in q_content
            need_bracket = "(   )" in line or (not is_drawing and not is_fill)

            if need_bracket and not q_content.startswith("(   )"):
                q_content = f"(   ){q_content}"

            if q_num <= 9:
                line = f" {q_num}.{q_content}"
            else:
                line = f"{q_num}.{q_content}"
            
            indent_pt = 48 if need_bracket else 24
            p.paragraph_format.left_indent = Pt(indent_pt)
            p.paragraph_format.first_line_indent = -Pt(indent_pt)

        elif match_opt:
            line = " " * 8 + line.strip()

        run = p.add_run(line)
        set_run_font(run, size=12, is_english_mode=is_english_mode)

    # 將 Word 檔案存進記憶體中以便網頁下載
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def call_api_with_retry(parts, model_id, api_key, status_box):
    models_to_try = [model_id] + [m for m in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"] if m != model_id]
    
    for mid in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{mid}:generateContent?key={api_key}"
        status_box.write(f"📡 正在嘗試連線: `{mid}`...")
        
        for attempt in range(2):
            try:
                res = requests.post(
                    url, 
                    json={"contents": [{"parts": parts}], "generationConfig": {"temperature": 0.1}}, 
                    timeout=600, 
                    verify=False
                )
                if res.status_code == 200:
                    return res.json()['candidates'][0]['content']['parts'][0]['text']
                elif res.status_code == 429:
                    status_box.write(f"⚠️ `{mid}` 頻率限制 (429)，等待 20 秒...")
                    time.sleep(20)
                elif res.status_code == 503:
                    status_box.write(f"⚠️ `{mid}` 忙碌 (503)，等待 10 秒...")
                    time.sleep(10)
                else:
                    break
            except Exception:
                break
        status_box.write(f"🔄 `{mid}` 失敗，自動切換下一個可用模型...")
    return None

# ==========================================
# 🎨 網頁介面 (Streamlit)
# ==========================================
st.title("📄 考卷排版神器 Pro")
st.caption("格式還原加速版 — 上傳 PDF 自動轉換為標準 Word 考卷格式")

# 從後台 Secrets 讀取 API Key
api_key = st.secrets.get("GEMINI_API_KEY", "").strip()
if not api_key:
    st.error("⚠️ 系統尚未設定 API Key，請管理員至 Streamlit Cloud 後台 Secrets 新增 `GEMINI_API_KEY`。")
    st.stop()

# 側邊欄設定
with st.sidebar:
    st.header("⚙️ 排版設定")
    selected_model = st.selectbox("🧠 AI 模型選擇", VALID_MODELS, index=0)
    is_english_mode = st.checkbox("英文考卷模式 (啟用 Times New Roman)", value=False)

# 主畫面：上傳檔案
uploaded_file = st.file_uploader("📁 選擇 PDF 考卷檔案", type=["pdf"])

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    
    st.info(f"已載入：**{uploaded_file.name}**（共 {total_pages} 頁）")
    page_input = st.text_input("選擇頁數 (輸入頁碼或範圍如 `1-3`、`1,3`，留空則全部轉換)：", value="")

    if st.button("⚡ 開始 AI 解析並排版", type="primary", use_container_width=True):
        indices = parse_page_selection(page_input, total_pages)
        
        with st.status("正在處理考卷，請稍候...", expanded=True) as status:
            prompt = """
                    請分析圖片並遵循規則排版：
                    請分析考卷圖片，嚴格遵守以下排版規則：
                    1. HEADER: 強制重組標題為 `[年份]([學期])[學校]                  [考試名稱]                            `。
                       - [考試名稱]後面的空格一定要加上
                       - 絕對不要加上 <u> 或 ** 等任何 HTML 或 Markdown 標籤
                       - 學校只需要名稱前兩個字，考試只需要標示第一次、第二次、或第三次段考試題並且要在整行標題包含後面空格加上底線，然後要靠左對齊。
                       - 範例：`113(一)中和                  第一次段考試題                            `。
                    2. **選擇題 / 是非題**：
                       - 必須在題號後加上括號 (   )。
                       - 範例：`1.(   )這是題目`。
                    3. **填充題 / 作圖題 / 計算題**：
                       - 題號後 **不要** 加括號，直接空格接題目。
                       - 範例：`2. 請畫出三角形` 或 `3. ＿＿＿`。
                    4. **選項處理**：
                       - 若選項是 1,2,3,4 或 (1)(2)(3)(4)，請轉為 ①②③④。
                       - 若選項是 ABCD，保留原樣。
                       - **務必將每個選項 (A)(B)(C)(D) 或 ①②③④ 放在新的一行**。
                    5. 重要：忽略圖片、圖表，不要輸出圖片描述。
                    6. 如果找不到標題，請把第一行文字當作標題。
                    格式回傳：
                    HEADER: [標題]
                    CONTENT_START
                    ...
                    CONTENT_END
                    """
            
            parts = [{"text": prompt}]
            status.write(f"🖼️ 正在將第 {[i+1 for i in indices]} 頁轉換為圖片...")
            for idx in indices:
                pix = doc[idx].get_pixmap(dpi=140)
                b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64}})
            doc.close()

            txt = call_api_with_retry(parts, selected_model, api_key, status)
            
            if txt:
                status.write("✅ 解析成功！正在套用職人排版格式...")
                docx_bytes = create_word_document(txt, is_english_mode)
                
                # 將結果存入 session_state，避免下載按鈕消失
                base_name = os.path.splitext(uploaded_file.name)[0]
                st.session_state["docx_bytes"] = docx_bytes
                st.session_state["docx_name"] = f"{base_name}_排版完成.docx"
                
                status.update(label="🎉 全部排版完成！請點擊下方按鈕下載", state="complete", expanded=False)
            else:
                status.update(label="❌ 轉換失敗，請更換模型或稍後重試。", state="error")

# 顯示下載按鈕
if "docx_bytes" in st.session_state:
    st.download_button(
        label="💾 下載 Word 檔案 (.docx)",
        data=st.session_state["docx_bytes"],
        file_name=st.session_state["docx_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )