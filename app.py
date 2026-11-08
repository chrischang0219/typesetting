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
VALID_MODELS = [
    "gemini-3.8-flash",      # 推薦：最新主力，辨識精準度高 (5 RPM / 250K TPM)
    "gemini-3.6-flash",      # 備用主力：同等級高精準度 (5 RPM / 250K TPM)
    "gemini-3.5-flash-lite", # 省額度首選：速度極快、每分鐘上限高 (15 RPM / 250K TPM)
    "gemini-3.1-flash-lite", # 輕量備援：同樣享有高每分鐘次數 (15 RPM / 250K TPM)
    "gemini-2.5-flash-lite", # 穩定輕量版 (10 RPM / 250K TPM)
    "gemini-2.5-flash"       # 舊版備援 (5 RPM / 250K TPM)
]

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
    
    # 清除 LaTeX 數學符號 ($...$, \frac, 次方等)
    text = text.replace('$$', '').replace('$', '')
    text = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'\1__UNITSLASH__\2', text)
    sup_map = {'0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹'}
    text = re.sub(r'\^\{?([0-9])\}?', lambda m: sup_map.get(m.group(1), m.group(0)), text)
    text = text.replace('\\times', '×').replace('\\div', '÷').replace('\\pm', '±')
    text = text.replace('\\leq', '≤').replace('\\geq', '≥').replace('\\neq', '≠').replace('\\approx', '≈')
    
    # 統一轉換選項格式為 (A) (B) (C) (D)
    # 僅針對 ①②③④ 或 (1)(2)(3)(4)，避免誤殺題號 1. 2. 10.
    text = re.sub(r'^\s*(?:①|[（\(]1[）\)])\s*[.．、\s]*', '(A)', text)
    text = re.sub(r'^\s*(?:②|[（\(]2[）\)])\s*[.．、\s]*', '(B)', text)
    text = re.sub(r'^\s*(?:③|[（\(]3[）\)])\s*[.．、\s]*', '(C)', text)
    text = re.sub(r'^\s*(?:④|[（\(]4[）\)])\s*[.．、\s]*', '(D)', text)
    
    # 針對 A~D 選項，必須有括號或後方接標點/空白，避免誤殺英文單字開頭 (如 CONTENT_START)
    text = re.sub(r'^\s*[（\(]A[）\)]\s*[.．、\s]*|^\s*A[.．、]\s*', '(A)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]B[）\)]\s*[.．、\s]*|^\s*B[.．、]\s*', '(B)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]C[）\)]\s*[.．、\s]*|^\s*C[.．、]\s*', '(C)', text, flags=re.IGNORECASE)
    text = re.sub(r'^\s*[（\(]D[）\)]\s*[.．、\s]*|^\s*D[.．、]\s*', '(D)', text, flags=re.IGNORECASE)

    # 修正英文縮寫撇號後方多餘空白 (如 It’ s -> It's)
    text = re.sub(r"([a-zA-Z])['’]\s+([a-zA-Z])", r"\1'\2", text)

    # 標點符號全形轉換（保護小數點、題號點、單位斜線）
    # 1. 先將受保護的符號暫存
    text = re.sub(r'^(\s*\d+)\.', r'\1__QDOT__', text)          # 題號的點 (如 7.)
    text = re.sub(r'(?<=\d)\.(?=\d)', '__DECDOT__', text)        # 小數點 (如 100.0)
    text = re.sub(r'(?<=[a-zA-Z0-9³²])/(?=[a-zA-Z0-9³²])', '__UNITSLASH__', text) # 單位斜線 (如 g/cm³)

    fullwidth_map = {';': '；', '~': '～', ',': '，', '?': '？', '!': '！', ':': '：', '.': '．', '*': '×', '/': '÷'}
    for k, v in fullwidth_map.items():
        text = text.replace(k, v)

    # 2. 還原受保護的符號
    text = text.replace('__QDOT__', '.').replace('__DECDOT__', '.').replace('__UNITSLASH__', '/')
    text = text.replace('（', '(').replace('）', ')')
    return text

def set_run_font(run, size, is_english_mode, bold=False, use_title_font=False):
    run.font.size = Pt(size)
    run.bold = bold
    
    ch_font = DOC_TITLE_FONT if use_title_font else DOC_CHINESE_FONT
    if use_title_font:
        # 標題的英數字使用標楷體，避免王漢宗顏楷體繁缺少英數字元時退回 Times New Roman
        ascii_font = DOC_CHINESE_FONT
    elif is_english_mode:
        ascii_font = DOC_ENGLISH_FONT
    else:
        ascii_font = ch_font
    
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    f = rPr.get_or_add_rFonts()
    f.set(qn('w:eastAsia'), ch_font)
    f.set(qn('w:ascii'), ascii_font)
    f.set(qn('w:hAnsi'), ascii_font)
    f.set(qn('w:cs'), ch_font)

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
    # A4 寬度 21cm，左右邊界各 1cm，內文寬度為 19cm，正中央為 9.5cm
    tab_stops = p.paragraph_format.tab_stops
    if "|" in header:
        tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.CENTER)
        tab_stops.add_tab_stop(Cm(19), WD_TAB_ALIGNMENT.RIGHT)
        parts = header.split("|")
        display_header = f"{parts[0].strip()}\t{parts[1].strip()}\t\xa0"
    else:
        clean_header = header.replace(';', '；').replace('~', '～').rstrip()
        # 自動偵測「學校」與「第X次段考試題」，將考試名稱置中於 9.5cm
        m_hdr = re.match(r'^(.*?)\s+((?:第[一二三]次|期中|期末).*)$', clean_header)
        tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.CENTER)
        tab_stops.add_tab_stop(Cm(19), WD_TAB_ALIGNMENT.RIGHT)
        if m_hdr:
            display_header = f"{m_hdr.group(1).strip()}\t{m_hdr.group(2).strip()}\t\xa0"
        else:
            display_header = f"{clean_header}\t\xa0"

    run_h = p.add_run(display_header)
    run_h.underline = WD_UNDERLINE.SINGLE
    set_run_font(run_h, size=16, is_english_mode=is_english_mode, bold=False, use_title_font=True)

    # 萃取內文並防呆清除殘留標籤
    content_match = re.search(r"CONTENT_START(.*?)CONTENT_END", text, re.DOTALL)
    if content_match:
        main_content = content_match.group(1).strip()
    else:
        main_content = re.sub(r"HEADER:.*\n?", "", text)
        main_content = main_content.replace("CONTENT_START", "").replace("CONTENT_END", "").strip()

    current_q_indent = 0  # 追蹤目前題目的縮排寬度，讓延續行(如對話)能對齊

    for line in main_content.split('\n'):
        line = custom_format_text(line)
        if not line.strip():
            continue
        
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(2)

        match_sec = re.match(r'^\s*[一二三四五六七八九十]+、', line)
        match_q = re.match(r'^\s*(\d+)\s*[.、．]\s*(.*)', line)
        match_opt = re.match(r'^\s*\([A-Da-d]\)\s*(.*)', line)

        if match_sec:
            # 大題標題（如 五、文法及對話選擇 20%）靠最左對齊，並重置題目縮排
            current_q_indent = 0
            line = line.strip()

        elif match_q:
            q_num = int(match_q.group(1))
            q_content = match_q.group(2).strip()
            
            is_drawing = any(k in q_content for k in ["作圖", "畫出", "計算", "證明"])
            is_fill = "＿＿" in q_content or "{BLANK}" in q_content
            need_bracket = "(   )" in line or (not is_drawing and not is_fill)

            if need_bracket:
                if re.match(r'^\(\s*\)', q_content):
                    q_content = re.sub(r'^\(\s*\)\s*', '(   )', q_content)
                else:
                    q_content = f"(   ){q_content}"

            if q_num <= 9:
                line = f" {q_num}.{q_content}"
            else:
                line = f"{q_num}.{q_content}"
            
            current_q_indent = 48 if need_bracket else 24
            p.paragraph_format.left_indent = Pt(current_q_indent)
            p.paragraph_format.first_line_indent = -Pt(current_q_indent)

        elif match_opt:
            line = line.strip()
            p.paragraph_format.left_indent = Pt(current_q_indent if current_q_indent > 0 else 48)

        else:
            # 延續行 / 對話行（例如英文題目的 Ryan : ...），對齊上一題的題幹起點
            if current_q_indent > 0:
                p.paragraph_format.left_indent = Pt(current_q_indent)

        run = p.add_run(line)
        set_run_font(run, size=12, is_english_mode=is_english_mode)

    # 將 Word 檔案存進記憶體中以便網頁下載
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


def get_global_usage(webhook_url):
    if not webhook_url:
        return None
    try:
        res = requests.get(webhook_url, timeout=5)
        if res.status_code == 200:
            return res.json()
    except Exception:
        pass
    return None

def log_usage_to_gsheet(webhook_url, filename, pages_count, model_id, tokens):
    if not webhook_url:
        return
    try:
        payload = {
            "filename": filename,
            "pages": pages_count,
            "model": model_id,
            "tokens": tokens
        }
        requests.post(webhook_url, json=payload, timeout=5)
    except Exception:
        pass

def call_api_with_retry(parts, model_id, api_key, status_box):
    models_to_try = [model_id] + [m for m in ["gemini-3.8-flash", "gemini-3.6-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-2.5-flash-lite", "gemini-2.5-flash"] if m != model_id]
    
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
                    data = res.json()
                    text_out = data['candidates'][0]['content']['parts'][0]['text']
                    usage = data.get('usageMetadata', {})
                    return text_out, usage
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
    return None, {}

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

# 通行密碼驗證（若 Secrets 有設定 APP_PASSWORD 才啟用）
app_password = st.secrets.get("APP_PASSWORD", "").strip()
if app_password:
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        st.info("🔒 此工具已設定存取密碼，請輸入密碼後繼續使用。")
        pwd_input = st.text_input("請輸入通行密碼", type="password")
        if st.button("🔓 登入", type="primary"):
            if pwd_input == app_password:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ 密碼錯誤，請重新輸入。")
        st.stop()

# 初始化本次連線用量統計（備援用）
if "total_tokens" not in st.session_state:
    st.session_state["total_tokens"] = 0
if "request_count" not in st.session_state:
    st.session_state["request_count"] = 0

gsheet_url = st.secrets.get("GSHEET_WEBHOOK_URL", "").strip()

# 側邊欄設定
with st.sidebar:
    st.header("⚙️ 排版設定")
    selected_model = st.selectbox("🧠 AI 模型選擇", VALID_MODELS, index=0)
    is_english_mode = st.checkbox("英文考卷模式 (啟用 Times New Roman)", value=False)
    
    st.divider()
    st.header("📊 今日額度與用量監控")
    
    global_stats = get_global_usage(gsheet_url) if gsheet_url else None
    if global_stats is not None:
        disp_count = global_stats.get("today_requests", global_stats.get("total_requests", 0))
        disp_tokens = global_stats.get("today_tokens", global_stats.get("total_tokens", 0))
        stats_note = "資料來源：Google 試算表（每日凌晨 00:00 自動歸零重算）。"
    else:
        disp_count = st.session_state["request_count"]
        disp_tokens = st.session_state["total_tokens"]
        stats_note = "註：尚未設定試算表同步，目前僅顯示本次開啟網頁後的消耗。"

    col1, col2 = st.columns(2)
    col1.metric("今日轉檔次數", f"{disp_count} 次")
    col2.metric("今日 Tokens", f"{disp_tokens:,}")
    
    # 以每日免費/參考額度 1,000,000 Tokens 作為直觀進度條基準
    ref_limit = 1_000_000
    usage_ratio = min(disp_tokens / ref_limit, 1.0)
    st.progress(usage_ratio, text=f"今日額度佔比 ({usage_ratio*100:.2f}%)")
    st.caption(stats_note)
    st.link_button("🔗 查看 Google 官方即時額度", "https://aistudio.google.com/usage", use_container_width=True)

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
                    請分析考卷圖片並嚴格遵守以下規則進行逐字轉錄與排版：
                    0. **最高原則（過濾手寫與精確轉錄）**：
                       - **完全忽略所有手寫筆跡**（包含題號旁的紅筆/藍筆/鉛筆答案、選項上的圈選或刪除線、空白處的計算過程或塗鴉），只辨識原本的**印刷體文字**。
                       - 必須**逐字精確轉錄**印刷體題目與選項，絕對不可自行改寫、腦補或省略。
                       - 保留小數點（如 `100.0`）與單位斜線（如 `g/cm³`）的原貌。
                       - **絕對不要使用 LaTeX 語法或 `$` 符號**（例如不要寫 `$x^2$` 或 `$\frac{1}{2}$`），請直接輸出純文字與 Unicode 符號（如 `x²`、`1/2`、`×`、`÷`）。
                    1. HEADER: 強制重組標題為 `[年份]([學期])[學校]                  [考試名稱]                            `。
                       - [考試名稱]後面的空格一定要加上
                       - 絕對不要加上 <u> 或 ** 等任何 HTML 或 Markdown 標籤
                       - 學校只需要名稱前兩個字，考試只需要標示第一次、第二次、或第三次段考試題並且要在整行標題包含後面空格加上底線，然後要靠左對齊。
                       - 範例：`113(一)中和                  第一次段考試題                            `。
                    2. **選擇題 / 是非題**：
                       - 必須在題號後加上括號 `(   )`，且括號後面**不要加空格**，直接接題目文字。
                       - 範例：`1.(   )這是題目`。
                    3. **填充題 / 作圖題 / 計算題**：
                       - 題號後 **不要** 加括號，直接空格接題目。
                       - 範例：`2. 請畫出三角形` 或 `3. ＿＿＿`。
                    4. **選項處理**：
                       - 若選項是 1,2,3,4 或 (1)(2)(3)(4)，請轉為 ①②③④。
                       - 若選項是 (A)(B)(C)(D)，保留原樣。
                       - **務必將每個選項 (A)(B)(C)(D) 或 ①②③④ 放在新的一行**。
                    5. 重要：忽略圖片、圖表，不要輸出圖片描述。
                    6. 如果找不到標題，請把第一行文字當作標題。
                    格式回傳（務必完整包含頭尾標記）：
                    HEADER: [標題]
                    CONTENT_START
                    ...
                    CONTENT_END
                    """
            
            parts = [{"text": prompt}]
            status.write(f"🖼️ 正在將第 {[i+1 for i in indices]} 頁轉換為圖片...")
            for idx in indices:
                pix = doc[idx].get_pixmap(dpi=250)
                b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64}})
            doc.close()

            txt, usage = call_api_with_retry(parts, selected_model, api_key, status)
            
            if txt:
                used_tokens = usage.get("totalTokenCount", 0)
                st.session_state["request_count"] += 1
                st.session_state["total_tokens"] += used_tokens
                if gsheet_url:
                    log_usage_to_gsheet(gsheet_url, uploaded_file.name, len(indices), selected_model, used_tokens)
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
