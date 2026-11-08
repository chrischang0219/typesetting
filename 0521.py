import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext, simpledialog
import base64
import json
import re
import time
import threading
import io
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
# ⚙️ 全域路徑與字體設定
# ==========================================
DOC_TITLE_FONT = "王漢宗顏楷體繁"
DOC_CHINESE_FONT = "標楷體"       
DOC_ENGLISH_FONT = "Times New Roman"
UI_FONT_FAMILY = "Microsoft JhengHei UI"
UI_FONT_LARGE = (UI_FONT_FAMILY, 16, "bold")

app_data_dir = os.getenv('APPDATA') or os.path.expanduser("~")
CONFIG_DIR = os.path.join(app_data_dir, "ExamFormatterPro")
if not os.path.exists(CONFIG_DIR): os.makedirs(CONFIG_DIR)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

# 模型清單
VALID_MODELS = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]

class ExamFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("考卷排版神器 Pro - 格式還原加速版")
        self.root.geometry("950x900")
        self.root.configure(bg="#f5f5f5")
        
        self.api_key_var = tk.StringVar()
        self.model_var = tk.StringVar(value=VALID_MODELS[0]) 
        self.is_english_mode = tk.BooleanVar(value=False) 
        self.current_doc = None 
        self.current_filename = "output"

        self.setup_ui()
        self.load_api_key()

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"> {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def setup_ui(self):
        # 頂部選單
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        model_menu = tk.Menu(self.menu_bar, tearoff=0)
        for model in VALID_MODELS:
            model_menu.add_radiobutton(label=model, variable=self.model_var, value=model)
        self.menu_bar.add_cascade(label="🧠 模型選擇", menu=model_menu)
        self.menu_bar.add_command(label="🔑 設定 API Key", command=self.ask_api_key)

        frame = tk.Frame(self.root, bg="#ffffff", padx=30, pady=20)
        frame.pack(fill="both", expand=True)

        self.btn_upload = tk.Button(frame, text="📁 選擇 PDF 並開始排版 (格式還原加速模式)", 
                                   command=self.start_process, font=UI_FONT_LARGE,
                                   bg="#1a73e8", fg="white", pady=20, relief="flat")
        self.btn_upload.pack(fill="x", pady=(0, 10))

        chk_english = tk.Checkbutton(frame, text="英文考卷模式 (啟用 Times New Roman)", variable=self.is_english_mode, bg="#ffffff")
        chk_english.pack(pady=5)

        self.log_text = scrolledtext.ScrolledText(frame, height=20, font=("Consolas", 10), state='disabled', bg="#f8f9fa")
        self.log_text.pack(fill="both", expand=True)

        self.btn_save = tk.Button(frame, text="💾 儲存 Word 檔案", command=self.manual_save, 
                                 state="disabled", font=(UI_FONT_FAMILY, 12, "bold"), pady=10)
        self.btn_save.pack(fill="x", pady=(15, 0))

    def start_process(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not file_path: return
        
        doc = fitz.open(file_path)
        total = len(doc)
        doc.close()
        
        s = simpledialog.askstring("選擇頁數", f"PDF 共 {total} 頁，輸入頁碼或範圍 (如 1-3) 或留空全轉:", parent=self.root)
        indices = self.parse_page_selection(s, total)
        if indices is None: return

        self.btn_upload.config(state="disabled", text="⚡ AI 解析並排版中...")
        threading.Thread(target=self.process_pdf_logic, args=(file_path, indices)).start()

    def process_pdf_logic(self, file_path, selected_indices):
        current_model = self.model_var.get().split(' ')[0]
        try:
            doc = fitz.open(file_path)
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
            for idx in selected_indices:
                pix = doc[idx].get_pixmap(dpi=140)
                b64 = base64.b64encode(pix.tobytes("jpeg")).decode('utf-8')
                parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64}})
            doc.close()

            # 呼叫 API
            txt = self.call_api_with_retry(parts, current_model)
            
            if txt:
                self.log("✅ 解析成功！正在套用職人排版格式...")
                self.create_word_document(txt, os.path.splitext(os.path.basename(file_path))[0])
            else:
                self.log("❌ 轉換失敗，請更換模型或稍後重試。")
        except Exception as e:
            self.log(f"❌ 嚴重錯誤: {e}")
        finally:
            self.root.after(0, lambda: self.btn_upload.config(state="normal", text="📁 選擇 PDF 並開始排版 (格式還原加速模式)"))

    def call_api_with_retry(self, parts, model_id):
        api_key = self.api_key_var.get().strip()
        models_to_try = [model_id, "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
        
        for mid in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{mid}:generateContent?key={api_key}"
            self.log(f"📡 正在嘗試連線: {mid}...")
            
            for attempt in range(2):
                try:
                    res = requests.post(url, json={"contents": [{"parts": parts}], "generationConfig": {"temperature": 0.1}}, timeout=600, verify=False)
                    if res.status_code == 200:
                        return res.json()['candidates'][0]['content']['parts'][0]['text']
                    elif res.status_code == 429:
                        self.log(f"⚠️ {mid} 頻率限制 (429)，等待 20 秒...")
                        time.sleep(20)
                    elif res.status_code == 503:
                        self.log(f"⚠️ {mid} 忙碌 (503)，等待 10 秒...")
                        time.sleep(10)
                    else:
                        break
                except:
                    break
            self.log(f"🔄 {mid} 失敗，自動切換下一個可用模型...")
        return None

    def create_word_document(self, text, filename):
        doc = Document()
        # 強制底線空格畫出 (ulTrailSpace)
        settings = doc.settings.element
        compat = settings.find(qn('w:compat')) or OxmlElement('w:compat')
        settings.append(compat)
        ult = OxmlElement('w:ulTrailSpace'); compat.append(ult)

        sec = doc.sections[0]
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1)

        # 標題與定位點處理
        header = "考卷排版結果"
        h_match = re.search(r"HEADER:\s*(.*)", text)
        if h_match: header = h_match.group(1).strip()
        
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
        self.set_run_font(run_h, size=16, bold=True, use_title_font=True)

        # 正文與格式處理
        content_match = re.search(r"CONTENT_START(.*)CONTENT_END", text, re.DOTALL)
        main_content = content_match.group(1).strip() if content_match else text

        for line in main_content.split('\n'):
            line = self.custom_format_text(line)
            if not line.strip(): continue
            
            # 建立新的段落
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)

            # -----------------------------------------------------------
            # 🔍 核心修改：判斷是否為題號開頭 (例如: "1.(   )題目" 或 "12. 題目")
            # -----------------------------------------------------------
            match_q = re.match(r'^\s*(\d+)\s*[.、．]\s*(.*)', line)
            
            # 🔍 核心修改：判斷是否為選項開頭 (例如: "(A)"、"A."、"①"、"1."開頭但已經是選項)
            # 這裡匹配 常見的 A., (A), ① ② ③ ④ 或是由 AI 整理出來的獨立行選項
            match_opt = re.match(r'^\s*([A-Da-d①②③④1-4]\s*[.．、\)]|[\(\(][A-Da-d1-4][\)\)])\s*(.*)', line)

            if match_q:
                q_num = int(match_q.group(1)) # 抓出題號數字
                q_content = match_q.group(2).strip()
                
                # 判斷這題是不是選擇題（內容包含括號，或上一行/本行有選擇題特徵）
                is_drawing = any(k in q_content for k in ["作圖", "畫出", "計算", "證明"])
                is_fill = "＿＿" in q_content or "{BLANK}" in q_content
                need_bracket = "(   )" in line or (not is_drawing and not is_fill)

                # 依據你的需求自動補上括號與空格
                if need_bracket and not q_content.startswith("(   )"):
                    q_content = f"(   ){q_content}"

                # ✨ 條件 1：1-9題題號後空一格，10題之後不空
                # 注意：這裡的空一格是指「題號.」後面與「括號/題目」之間的距離
                if q_num <= 9:
                    line = f" {q_num}.{q_content}"  # 有留一個半形空格
                else:
                    line = f"{q_num}.{q_content}"   # 10題之後緊接，不留空格
                
                # 處理凸排，讓第二行自動對齊
                indent_pt = 48 if need_bracket else 24
                p.paragraph_format.left_indent = Pt(indent_pt)
                p.paragraph_format.first_line_indent = -Pt(indent_pt)

            elif match_opt:
                # ✨ 條件 2：如果是獨立一行的選項，前面強制空 8 格
                # 用 8 個半形空格（" " * 8）來確保對齊
                line = " " * 8 + line.strip()

            # -----------------------------------------------------------
            
            run = p.add_run(line)
            self.set_run_font(run, size=12)
            
        self.current_doc = doc
        self.current_filename = filename
        self.btn_save.config(state="normal", bg="#34a853", fg="white")
        self.log("🎉 全部排版完成！")

    def custom_format_text(self, text):
        # 移除可能干擾的 HTML 標籤
        text = re.sub(r'</?[ub*]+>', '', text)
        
        # 1️⃣ 核心修改：將各種選項格式統一轉換為 (A) (B) (C) (D)
        # 匹配 ①②③④ 或 1. 2. 3. 4. 或 (1)(2)(3)(4)
        text = re.sub(r'^\s*([①1]|[\(\(]1[\)\)])\s*[.．、\s]*', '(A)', text)
        text = re.sub(r'^\s*([②2]|[\(\(]2[\)\)])\s*[.．、\s]*', '(B)', text)
        text = re.sub(r'^\s*([③3]|[\(\(]3[\)\)])\s*[.．、\s]*', '(C)', text)
        text = re.sub(r'^\s*([④4]|[\(\(]4[\)\)])\s*[.．、\s]*', '(D)', text)
        
        # 匹配原本就是 A. B. C. D. 或 A) B) C) D) 沒加左括號的，統一補齊
        text = re.sub(r'^\s*[（\(]?A[）\)]?[.．、\s]*', '(A)', text, flags=re.IGNORECASE)
        text = re.sub(r'^\s*[（\(]?B[）\)]?[.．、\s]*', '(B)', text, flags=re.IGNORECASE)
        text = re.sub(r'^\s*[（\(]?C[）\)]?[.．、\s]*', '(C)', text, flags=re.IGNORECASE)
        text = re.sub(r'^\s*[（\(]?D[）\)]?[.．、\s]*', '(D)', text, flags=re.IGNORECASE)

        # 2️⃣ 其他標點符號全形轉換
        fullwidth_map = {';': '；', '~': '～', ',': '，', '?': '？', '!': '！', ':': '：', '.': '．', '*': '×', '/': '÷'}
        for k, v in fullwidth_map.items(): 
            text = text.replace(k, v)
        text = text.replace('（', '(').replace('）', ')').replace('　', ' ')
        
        return text

    def set_run_font(self, run, size, bold=False, use_title_font=False):
        run.font.size = Pt(size)
        run.bold = bold
        
        # 判斷中文字體該用「標楷體」還是「顏楷體」
        ch_font = DOC_TITLE_FONT if use_title_font else DOC_CHINESE_FONT
        ascii_font = DOC_ENGLISH_FONT if self.is_english_mode.get() else ch_font
        
        run.font.name = ascii_font
        rPr = run._element.get_or_add_rPr()
        f = rPr.get_or_add_rFonts()
        f.set(qn('w:eastAsia'), ch_font)  # 確保中文字體正確綁定
        f.set(qn('w:ascii'), ascii_font)
        f.set(qn('w:hAnsi'), ascii_font)

    def parse_page_selection(self, s, max_p):
        if s is None: return None
        if not s.strip(): return list(range(max_p))
        selected = set()
        try:
            for part in s.split(','):
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    for p in range(max(1, start), min(max_p, end) + 1): selected.add(p - 1)
                else:
                    p = int(part)
                    if 1 <= p <= max_p: selected.add(p - 1)
            return sorted(list(selected)) if selected else list(range(max_p))
        except: return list(range(max_p))

    def manual_save(self):
        p = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=self.current_filename)
        if p: self.current_doc.save(p); messagebox.showinfo("成功", "儲存成功")

    def ask_api_key(self):
        k = simpledialog.askstring("API Key", "輸入 Key:", initialvalue=self.api_key_var.get())
        if k:
            self.api_key_var.set(k.strip())
            with open(CONFIG_FILE, "w") as f: json.dump({"api_key": k.strip()}, f)

    def load_api_key(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f: self.api_key_var.set(json.load(f).get("api_key", ""))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExamFormatterApp(root)
    root.mainloop()