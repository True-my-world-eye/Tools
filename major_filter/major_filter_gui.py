import os
import re
import json
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def to_halfwidth(s: str) -> str:
    r = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            code = 32
        elif 0xFF01 <= code <= 0xFF5E:
            code -= 0xFEE0
        r.append(chr(code))
    return "".join(r)

def normalize_text(s: str) -> str:
    s = to_halfwidth(s)
    s = s.strip().lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", s)
    return s

def extract_code(s: str):
    m = re.search(r"(\d{4,6}[A-Z]{0,3})", s)
    if not m:
        return None
    return re.sub(r"[^0-9]", "", m.group(1))

def read_text_file(path: str):
    for enc in ("utf-8", "gbk", "utf-8-sig"):
        try:
            with open(path, "r", encoding=enc) as f:
                return [line.rstrip("\n") for line in f]
        except Exception:
            continue
    with open(path, "r", errors="ignore") as f:
        return [line.rstrip("\n") for line in f]

def parse_requirements(lines):
    items = []
    for raw in lines:
        t = raw.strip()
        if not t:
            continue
        name = None
        code = None
        m = re.search(r"^(.*?)（\s*([0-9]{4,6}[A-Z]{0,3})\s*）", t)
        if m:
            name = m.group(1).strip()
            code = re.sub(r"[^0-9]", "", m.group(2))
        else:
            parts = re.split(r"\s+", t)
            if parts:
                name = parts[-1]
                c = extract_code(t)
                if c:
                    code = c
        if not name:
            continue
        items.append({
            "raw": raw,
            "name": name,
            "norm": normalize_text(name),
            "code": code or ""
        })
    seen = set()
    result = []
    for it in items:
        key = (it["norm"], it["code"])
        if key in seen:
            continue
        seen.add(key)
        result.append(it)
    return result

def similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()

def best_match(major: str, reqs):
    if not major:
        return None, 0.0
    major_norm = normalize_text(major)
    major_code = extract_code(major) or ""
    best = None
    best_score = 0.0
    for r in reqs:
        score = 0.0
        if major_code and r["code"] and major_code == r["code"]:
            score = 1.0
        elif major_norm == r["norm"]:
            score = 0.95
        elif major_norm and r["norm"] and (major_norm in r["norm"] or r["norm"] in major_norm):
            score = 0.9
        else:
            score = similarity(major_norm, r["norm"])
        if score > best_score:
            best_score = score
            best = r
    return best, best_score

def dedup_dataframe(df, col_major: str, key: str | None):
    if df is None or len(df) == 0:
        return df
    if key and key in df.columns:
        return df.drop_duplicates(subset=[key]).copy()
    keys = []
    for i in range(len(df)):
        v = str(df.iloc[i][col_major]) if col_major in df.columns else ""
        code = str(df.iloc[i]["_matched_code"]) if "_matched_code" in df.columns else ""
        keys.append((normalize_text(v), code))
    df["_dedup_key"] = keys
    res = df.drop_duplicates(subset=["_dedup_key"]).copy()
    res.drop(columns=["_dedup_key"], inplace=True)
    return res

def write_output(df, out_path: str, append: bool = False, dedup: bool = False, dedup_key: str | None = None, col_major: str = "Major") -> str:
    ext = os.path.splitext(out_path)[1].lower()
    if append and os.path.exists(out_path):
        try:
            import pandas as pd
            if ext in (".xlsx", ".xls"):
                old = pd.read_excel(out_path)
            else:
                old = pd.read_csv(out_path)
            merged = pd.concat([old, df], ignore_index=True)
            if dedup:
                merged = dedup_dataframe(merged, col_major, dedup_key)
            df = merged
        except Exception:
            pass
    if ext in (".xlsx", ".xls"):
        try:
            df.to_excel(out_path, index=False)
            return out_path
        except Exception:
            csv_path = os.path.splitext(out_path)[0] + ".csv"
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            return csv_path
    else:
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
        return out_path

def process_single(excel_path: str, require_path: str, col_major: str, threshold: float, out_path: str | None, sheet: str | None, progress_step: int, limit: int | None, append: bool, dedup: bool, dedup_key: str | None, progress_cb=None, log_cb=None):
    try:
        import pandas as pd
    except Exception:
        raise RuntimeError("需要安装pandas，请运行: pip install pandas openpyxl")
    lines = read_text_file(require_path)
    reqs = parse_requirements(lines)
    if not reqs:
        raise RuntimeError("专业要求解析为空")
    if sheet:
        if limit and limit > 0:
            df = pd.read_excel(excel_path, sheet_name=sheet, nrows=limit)
        else:
            df = pd.read_excel(excel_path, sheet_name=sheet)
    else:
        if limit and limit > 0:
            df = pd.read_excel(excel_path, nrows=limit)
        else:
            df = pd.read_excel(excel_path)
    if col_major not in df.columns:
        raise RuntimeError(f"未找到列: {col_major}")
    matches = []
    majors = df[col_major].astype(str).fillna("").tolist()
    total = len(majors)
    for i, v in enumerate(majors, start=1):
        m, score = best_match(v, reqs)
        if m and score >= threshold:
            matches.append({"match": True, "matched_name": m["name"], "matched_code": m["code"], "score": round(score, 4)})
        else:
            matches.append({"match": False, "matched_name": "", "matched_code": "", "score": round(score, 4)})
        if progress_cb and progress_step and progress_step > 0 and i % progress_step == 0:
            progress_cb(i, total)
    df["_match"] = [x["match"] for x in matches]
    df["_matched_name"] = [x["matched_name"] for x in matches]
    df["_matched_code"] = [x["matched_code"] for x in matches]
    df["_score"] = [x["score"] for x in matches]
    out_df = df[df["_match"] == True].copy()
    if out_path is None or out_path == "":
        base = os.path.splitext(os.path.basename(excel_path))[0]
        out_path = os.path.join(os.path.dirname(excel_path), f"{base}_filtered.xlsx")
    count = len(out_df)
    saved = write_output(out_df, out_path, append=append, dedup=dedup, dedup_key=dedup_key, col_major=col_major)
    if log_cb:
        log_cb(f"筛选完成：{os.path.basename(excel_path)} 命中 {count} 条 → {saved}")
    return saved, count

class MajorFilterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("专业筛选数据分析")
        self.files = []
        self.require_path = tk.StringVar()
        self.major_col = tk.StringVar(value="Major")
        self.threshold = tk.DoubleVar(value=0.8)
        self.threshold_input = tk.StringVar(value="0.80")
        self._syncing = False
        self._modal_open = False
        self.progress_step = tk.IntVar(value=500)
        self.limit = tk.StringVar(value="")
        self.sheet = tk.StringVar(value="")
        self.out_dir = tk.StringVar(value="")
        self.merge_out = tk.StringVar(value="")
        self.append_mode = tk.BooleanVar(value=False)
        self.dedup = tk.BooleanVar(value=False)
        self.dedup_key = tk.StringVar(value="")
        self.log_queue = queue.Queue()
        self.total_count = 0
        self.setup_style()
        self.create_widgets()
        self.load_config()
        self.root.after(100, self.consume_logs)

    def setup_style(self):
        try:
            style = ttk.Style()
            themes = style.theme_names()
            if "vista" in themes and os.name == "nt":
                style.theme_use("vista")
            elif "aqua" in themes and os.name != "nt":
                style.theme_use("aqua")
            else:
                style.theme_use("clam")
            style.configure("TLabel", padding=4)
            style.configure("TEntry", padding=4)
            style.configure("TButton", padding=4)
            style.configure("TLabelframe", padding=8)
            style.configure("TLabelframe.Label", padding=4)
        except Exception:
            pass
        try:
            self.root.minsize(900, 650)
        except Exception:
            pass

    def create_widgets(self):
        outer = ttk.Frame(self.root, padding=0)
        outer.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        canvas = tk.Canvas(outer, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        container = ttk.Frame(canvas, padding=10)
        self._canvas_window = canvas.create_window((0, 0), window=container, anchor="nw")
        self._canvas = canvas
        def _on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            try:
                canvas.itemconfigure(self._canvas_window, width=canvas.winfo_width())
            except Exception:
                pass
        container.bind("<Configure>", _on_configure)
        self.root.bind_all("<MouseWheel>", self._on_main_mousewheel)
        self.root.bind_all("<Button-4>", self._on_main_mousewheel_up)
        self.root.bind_all("<Button-5>", self._on_main_mousewheel_down)
        container.columnconfigure(0, weight=1)
        # 文件分区
        files_frame = ttk.Labelframe(container, text="文件")
        files_frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        files_frame.columnconfigure(1, weight=1)
        ttk.Label(files_frame, text="Excel文件").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        ttk.Button(files_frame, text="添加文件", command=self.add_files).grid(row=0, column=2, sticky="e", padx=4, pady=2)
        lst_frame = ttk.Frame(files_frame)
        lst_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=4, pady=2)
        files_frame.rowconfigure(1, weight=1)
        lst_frame.columnconfigure(0, weight=1)
        self.lst = tk.Listbox(lst_frame, height=6)
        self.lst.grid(row=0, column=0, sticky="nsew")
        sb_list = ttk.Scrollbar(lst_frame, orient="vertical", command=self.lst.yview)
        sb_list.grid(row=0, column=1, sticky="ns")
        self.lst.configure(yscrollcommand=sb_list.set)
        ttk.Button(files_frame, text="移除选中", command=self.remove_selected).grid(row=2, column=2, sticky="e", padx=4, pady=2)
        ttk.Label(files_frame, text="require.txt").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        ttk.Entry(files_frame, textvariable=self.require_path).grid(row=3, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(files_frame, text="选择", command=self.pick_require).grid(row=3, column=2, sticky="e", padx=4, pady=2)
        # 参数分区
        params = ttk.Labelframe(container, text="参数")
        params.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        for c in range(3):
            params.columnconfigure(c, weight=1 if c == 1 else 0)
        r = 0
        ttk.Label(params, text="专业列名").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.major_col).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(params, text="Sheet").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.sheet).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(params, text="阈值").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        scl = ttk.Scale(params, from_=0.0, to=1.0, variable=self.threshold, command=self.on_threshold_slider)
        scl.grid(row=r, column=1, sticky="ew", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.threshold_input).grid(row=r, column=2, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(params, text="参考：≤0.70宽松，0.75~0.85均衡，≥0.90严格；可填0~1或百分比").grid(row=r, column=0, columnspan=3, sticky="w", padx=4, pady=2); r+=1
        ttk.Label(params, text="进度步长").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.progress_step).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(params, text="limit").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.limit).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(params, text="输出目录").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.out_dir).grid(row=r, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(params, text="选择", command=self.pick_out_dir).grid(row=r, column=2, sticky="e", padx=4, pady=2); r+=1
        ttk.Label(params, text="合并输出文件").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.merge_out).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(params, text="去重键").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(params, textvariable=self.dedup_key).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Checkbutton(params, text="追加模式", variable=self.append_mode).grid(row=r, column=0, sticky="w", padx=4, pady=2)
        ttk.Checkbutton(params, text="开启去重", variable=self.dedup).grid(row=r, column=1, sticky="w", padx=4, pady=2); r+=1
        # 控制分区
        actions = ttk.Frame(container)
        actions.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
        actions.columnconfigure(0, weight=1)
        left = ttk.Frame(actions)
        left.grid(row=0, column=0, sticky="w")
        right = ttk.Frame(actions)
        right.grid(row=0, column=1, sticky="e")
        ttk.Button(left, text="保存配置", command=self.save_config).grid(row=0, column=0, padx=4)
        ttk.Button(left, text="清除本地缓存", command=self.clear_cache).grid(row=0, column=1, padx=4)
        ttk.Button(left, text="软件使用须知", command=self.show_usage_notice).grid(row=0, column=2, padx=4)
        ttk.Button(right, text="开始处理", command=self.start_processing).grid(row=0, column=0, padx=4)
        # 反馈分区
        feedback = ttk.Labelframe(container, text="进度与日志")
        feedback.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)
        feedback.columnconfigure(0, weight=1)
        feedback.rowconfigure(1, weight=1)
        self.progress = ttk.Progressbar(feedback, orient="horizontal", mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=4, pady=2)
        log_frame = ttk.Frame(feedback)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=2)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=12)
        self.log.grid(row=0, column=0, sticky="nsew")
        sb_log = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        sb_log.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=sb_log.set)
        try:
            self.threshold_input.trace_add("write", self.on_threshold_input_change)
        except Exception:
            pass

    def on_threshold_slider(self, val):
        if self._syncing:
            return
        try:
            v = float(val)
            v = max(0.0, min(1.0, v))
            self._syncing = True
            self.threshold.set(v)
            self.threshold_input.set(f"{v:.2f}")
        except Exception:
            pass
        finally:
            self._syncing = False

    def parse_threshold_value(self, s: str):
        try:
            if not s:
                return None
            t = s.strip().replace("%", "")
            v = float(t)
            if v > 1.0:
                v = v / 100.0
            v = max(0.0, min(1.0, v))
            return v
        except Exception:
            return None

    def on_threshold_input_change(self, *args):
        if self._syncing:
            return
        v = self.parse_threshold_value(self.threshold_input.get())
        if v is None:
            return
        try:
            self._syncing = True
            self.threshold.set(v)
            self.threshold_input.set(f"{v:.2f}")
        finally:
            self._syncing = False

    def add_files(self):
        paths = filedialog.askopenfilenames(title="选择Excel文件", filetypes=[("Excel", ".xlsx .xls .csv"), ("All files", "*.*")])
        for p in paths:
            if p and p not in self.files:
                self.files.append(p)
                self.lst.insert(tk.END, p)

    def remove_selected(self):
        sel = list(self.lst.curselection())
        sel.reverse()
        for i in sel:
            p = self.lst.get(i)
            self.files.remove(p)
            self.lst.delete(i)

    def pick_require(self):
        p = filedialog.askopenfilename(title="选择require.txt", filetypes=[("Text", ".txt"), ("All files", "*.*")])
        if p:
            self.require_path.set(p)

    def pick_out_dir(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.out_dir.set(p)

    def log_cb(self, msg: str):
        self.log_queue.put(msg)

    def progress_cb(self, done: int, total: int):
        if total <= 0:
            self.progress["value"] = 0
            self.progress["maximum"] = 1
        else:
            self.progress["maximum"] = total
            self.progress["value"] = done

    def consume_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log.insert(tk.END, msg + "\n")
                self.log.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self.consume_logs)

    def start_processing(self):
        if not self.files:
            messagebox.showwarning("提示", "请添加至少一个Excel文件")
            return
        req = self.require_path.get().strip()
        if not req:
            messagebox.showwarning("提示", "请指定require.txt文件")
            return
        try:
            import pandas as pd  # noqa
        except Exception:
            messagebox.showerror("错误", "需要安装pandas与openpyxl:\n pip install pandas openpyxl")
            return
        self.total_count = 0
        self.progress["value"] = 0
        self.log_cb("开始处理")
        t = threading.Thread(target=self.run_processing, daemon=True)
        t.start()

    def show_usage_notice(self):
        top = tk.Toplevel(self.root)
        top.title("软件使用须知")
        top.geometry("600x500")
        top.transient(self.root)
        top.grab_set()
        self._modal_open = True
        top.bind("<Destroy>", lambda e: self._on_modal_closed())
        frm = ttk.Frame(top, padding=10)
        frm.pack(fill="both", expand=True)
        txt = tk.Text(frm, wrap="word")
        txt.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        sb.pack(fill="y", side="right")
        txt.configure(yscrollcommand=sb.set)
        content = (
            "软件主要目的：\n"
            "用于根据require.txt中的专业要求，对所选Excel文件的专业列进行模糊匹配与筛选，"
            "生成筛选结果并支持批量合并导出、追加与去重。\n\n"
            "使用方法：\n"
            "1. 添加文件：选择一个或多个Excel文件。\n"
            "2. 移除选中：从列表中移除不需要处理的文件。\n"
            "3. 选择require.txt：指定专业要求文本文件。\n"
            "4. 参数设置：\n"
            "   - 专业列名：Excel中专业所在列（默认Major）。\n"
            "   - Sheet：指定要读取的工作表（可留空读取默认）。\n"
            "   - 阈值：匹配相似度阈值，数值越高越严格；可输入0~1或百分比。\n"
            "           参考：≤0.70宽松，0.75~0.85均衡，≥0.90严格。\n"
            "   - 进度步长：每处理N行输出一次进度。\n"
            "   - limit：仅读取前N行用于调试或大文件处理（留空为不限制）。\n"
            "   - 输出目录：逐文件筛选结果的保存目录。\n"
            "   - 合并输出文件：批量处理后合并导出文件名（留空使用默认）。\n"
            "   - 去重键：指定唯一键进行去重（不填则按规范化Major+编码）。\n"
            "   - 追加模式：勾选后在已存在的输出文件上追加写入，可结合去重使用。\n"
            "   - 开启去重：对结果进行去重处理。\n"
            "5. 保存配置：将当前参数保存，下次启动自动加载。\n"
            "6. 开始处理：后台执行筛选，查看进度与日志；完成后弹窗提示总计筛选条数。\n\n"
            "配置与缓存：\n"
            "配置文件保存位置：程序同目录的“major_filter_gui.json”。\n"
            "可点击主界面“清除本地缓存”按钮删除该配置并恢复默认设置。\n\n"
            "\nrequire.txt格式要求：\n"
            "1. 每行一个专业条目，可包含空行与前后空白（将自动忽略）。\n"
            "2. 推荐格式：专业名称（编码），使用中文全角括号“（ ）”。\n"
            "   例如：信息资源管理（120503）。\n"
            "3. 也支持列式条目：如“397 计算机类 080901 计算机科学与技术”。\n"
            "4. 编码形态：4~6位数字 + 可选字母后缀（如T/K/TK）。编码匹配优先，准确度更高。\n"
            "5. 若不使用中文括号，程序将尝试以最后一个词作为名称并抽取编码，可能影响名称识别效果，建议使用中文括号或确保编码存在。\n"
            "6. 名称建议使用正式中文，不混写多个专业于同一行；同义写法脚本会做规范化，但规范格式更有利于匹配。\n\n"
            "作者信息：\n"
            "作者：True my world eye\n"
            "Wechat：Truemwe\n"
            "E-mail：hwzhang0722@163.com\n"
        )
        txt.insert("1.0", content)
        txt.config(state="disabled")
        btn = ttk.Button(top, text="关闭", command=top.destroy)
        btn.pack(anchor="e", padx=10, pady=8)

    def clear_cache(self):
        try:
            p = "major_filter_gui.json"
            if os.path.exists(p):
                os.remove(p)
            self.files = []
            self.lst.delete(0, tk.END)
            self.require_path.set("")
            self.major_col.set("Major")
            self.threshold.set(0.8)
            self.threshold_input.set("0.80")
            self.progress_step.set(500)
            self.limit.set("")
            self.sheet.set("")
            self.out_dir.set("")
            self.merge_out.set("")
            self.append_mode.set(False)
            self.dedup.set(False)
            self.dedup_key.set("")
            messagebox.showinfo("提示", "本地缓存已清除，设置已恢复默认")
        except Exception as e:
            messagebox.showerror("错误", str(e))
    def _on_modal_closed(self):
        self._modal_open = False

    def _on_main_mousewheel(self, event):
        if self._modal_open:
            return
        try:
            delta = -1 * int(event.delta / 120)
            self._canvas.yview_scroll(delta, "units")
        except Exception:
            pass

    def _on_main_mousewheel_up(self, event):
        if self._modal_open:
            return
        try:
            self._canvas.yview_scroll(-1, "units")
        except Exception:
            pass

    def _on_main_mousewheel_down(self, event):
        if self._modal_open:
            return
        try:
            self._canvas.yview_scroll(1, "units")
        except Exception:
            pass
    def run_processing(self):
        req = self.require_path.get().strip()
        col_major = self.major_col.get().strip() or "Major"
        threshold = float(self.threshold.get())
        progress_step = int(self.progress_step.get())
        limit_str = self.limit.get().strip()
        limit = int(limit_str) if limit_str.isdigit() else None
        sheet = self.sheet.get().strip() or None
        out_dir = self.out_dir.get().strip() or None
        merge_out = self.merge_out.get().strip() or None
        append = bool(self.append_mode.get())
        dedup = bool(self.dedup.get())
        dedup_key = self.dedup_key.get().strip() or None
        outputs = []
        merged_parts = []
        total_count = 0
        try:
            import pandas as pd
            for pth in self.files:
                base = os.path.splitext(os.path.basename(pth))[0]
                out_path = None
                if out_dir:
                    out_path = os.path.join(out_dir, f"{base}_filtered.xlsx")
                saved, count = process_single(pth, req, col_major, threshold, out_path, sheet, progress_step, limit, append, dedup, dedup_key, progress_cb=self.progress_cb, log_cb=self.log_cb)
                outputs.append(saved)
                total_count += count
                try:
                    if saved.lower().endswith(".xlsx"):
                        part = pd.read_excel(saved)
                    else:
                        part = pd.read_csv(saved)
                    if dedup:
                        part = dedup_dataframe(part, col_major, dedup_key)
                    merged_parts.append(part)
                except Exception:
                    pass
            if merged_parts:
                all_df = pd.concat(merged_parts, ignore_index=True)
                if dedup:
                    all_df = dedup_dataframe(all_df, col_major, dedup_key)
                if merge_out is None:
                    first_dir = os.path.dirname(self.files[0]) if self.files else os.getcwd()
                    merge_out = os.path.join(first_dir, "merged_filtered.xlsx")
                saved = write_output(all_df, merge_out, append=append, dedup=dedup, dedup_key=dedup_key, col_major=col_major)
                self.log_cb(f"总计筛选 {total_count} 条；合并后共 {len(all_df)} 条 → {saved}")
            else:
                self.log_cb(f"总计筛选 {total_count} 条")
            messagebox.showinfo("完成", f"处理完成，共筛选 {total_count} 条")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.log_cb(f"错误：{str(e)}")

    def save_config(self):
        cfg = {
            "files": self.files,
            "require_path": self.require_path.get(),
            "major_col": self.major_col.get(),
            "threshold": float(self.threshold.get()),
            "progress_step": int(self.progress_step.get()),
            "limit": self.limit.get(),
            "sheet": self.sheet.get(),
            "out_dir": self.out_dir.get(),
            "merge_out": self.merge_out.get(),
            "append_mode": bool(self.append_mode.get()),
            "dedup": bool(self.dedup.get()),
            "dedup_key": self.dedup_key.get()
        }
        try:
            with open("major_filter_gui.json", "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("提示", "配置已保存")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def load_config(self):
        try:
            with open("major_filter_gui.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.files = cfg.get("files", [])
            self.lst.delete(0, tk.END)
            for p in self.files:
                self.lst.insert(tk.END, p)
            self.require_path.set(cfg.get("require_path", ""))
            self.major_col.set(cfg.get("major_col", "Major"))
            self.threshold.set(cfg.get("threshold", 0.8))
            try:
                self.threshold_input.set(f"{float(self.threshold.get()):.2f}")
            except Exception:
                pass
            self.progress_step.set(cfg.get("progress_step", 500))
            self.limit.set(cfg.get("limit", ""))
            self.sheet.set(cfg.get("sheet", ""))
            self.out_dir.set(cfg.get("out_dir", ""))
            self.merge_out.set(cfg.get("merge_out", ""))
            self.append_mode.set(cfg.get("append_mode", False))
            self.dedup.set(cfg.get("dedup", False))
            self.dedup_key.set(cfg.get("dedup_key", ""))
        except Exception:
            pass

def main():
    root = tk.Tk()
    app = MajorFilterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
