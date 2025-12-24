import os
import re
import json
import threading
import queue
import time
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

def parse_options_local(s: str):
    opts = {}
    if not s:
        return opts
    for part in s.split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            k, v = part.split("=", 1)
            opts[k.strip()] = v.strip()
        else:
            opts[part] = "true"
    return opts

def apply_condition_local(val: str, cond: dict):
    t = cond.get("type", "")
    op = cond.get("operator", "")
    value = cond.get("value", "")
    th_raw = cond.get("threshold", "")
    opts = parse_options_local(cond.get("options", ""))
    norm = str(opts.get("normalize","false")).lower() == "true"
    ignore_case = str(opts.get("ignore_case","false")).lower() == "true"
    code_prefer = str(opts.get("code_prefer","false")).lower() == "true"
    if ignore_case:
        val_cmp = val.lower()
        value_cmp = value.lower()
    else:
        val_cmp = val
        value_cmp = value
    if norm:
        val_cmp = normalize_text(val_cmp)
        value_cmp = normalize_text(value_cmp)
    if t == "text":
        if op == "equals":
            return (val_cmp == value_cmp, 1.0 if val_cmp == value_cmp else 0.0)
        if op == "contains":
            return (value_cmp in val_cmp, 1.0 if value_cmp in val_cmp else 0.0)
        if op == "startswith":
            return (val_cmp.startswith(value_cmp), 1.0 if val_cmp.startswith(value_cmp) else 0.0)
        if op == "endswith":
            return (val_cmp.endswith(value_cmp), 1.0 if val_cmp.endswith(value_cmp) else 0.0)
        return (False, 0.0)
    if t == "enum":
        values = [x.strip() for x in value.split(";") if x.strip()]
        hit = val_cmp in values
        return (hit, 1.0 if hit else 0.0)
    if t == "regex":
        try:
            hit = re.search(value, val) is not None
            return (hit, 1.0 if hit else 0.0)
        except Exception:
            return (False, 0.0)
    if t == "boolean":
        truth = value_cmp in ("true", "1", "yes", "y", "t")
        val_truth = val_cmp in ("true", "1", "yes", "y", "t")
        hit = (val_truth == truth)
        return (hit, 1.0 if hit else 0.0)
    if t == "number":
        try:
            v = float(val.strip())
        except Exception:
            return (False, 0.0)
        if op == "between":
            try:
                parts = value.replace(" ", "").split("-")
                lo = float(parts[0]); hi = float(parts[1])
                hit = (lo <= v <= hi)
                return (hit, 1.0 if hit else 0.0)
            except Exception:
                return (False, 0.0)
        if op == "min":
            try:
                lo = float(value)
                hit = v >= lo
                return (hit, 1.0 if hit else 0.0)
            except Exception:
                return (False, 0.0)
        if op == "max":
            try:
                hi = float(value)
                hit = v <= hi
                return (hit, 1.0 if hit else 0.0)
            except Exception:
                return (False, 0.0)
        if op == "equals":
            try:
                eq = float(value)
                hit = v == eq
                return (hit, 1.0 if hit else 0.0)
            except Exception:
                return (False, 0.0)
        return (False, 0.0)
    if t == "code":
        val_code = extract_code(val) or ""
        value_code = re.sub(r"[^0-9]", "", value)
        hit = (val_code and value_code and val_code == value_code)
        return (hit, 1.0 if hit else 0.0)
    if t == "fuzzy":
        th = 0.0
        try:
            if th_raw:
                if th_raw.strip().endswith("%"):
                    th = float(th_raw.strip()[:-1]) / 100.0
                else:
                    th = float(th_raw.strip())
        except Exception:
            th = 0.0
        if code_prefer:
            val_code = extract_code(val) or ""
            tgt_code = extract_code(value) or ""
            if val_code and tgt_code and val_code == tgt_code:
                return (True, 1.0)
        a = normalize_text(val)
        b = normalize_text(value)
        from difflib import SequenceMatcher
        s = SequenceMatcher(None, a, b).ratio()
        return (s >= th, s)
    return (False, 0.0)

def evaluate_conditions_row_local(row: dict, conditions: list, combine_mode: str, combine_threshold: float):
    # 性能优化：将同列且相同选项的text/contains合并为“任意命中”组
    groups = []
    merged_keys = {}
    for cond in conditions:
        t = cond.get("type", "")
        op = cond.get("operator", "")
        col = cond.get("column", "")
        opts = cond.get("options", "")
        if t == "text" and op == "contains":
            key = (col, t, op, opts)
            g = merged_keys.get(key)
            if not g:
                g = {"kind": "contains_any", "column": col, "options": opts, "tokens": set(), "weight": 1.0}
                try:
                    g["weight"] = float(cond.get("weight", "") or "1")
                except Exception:
                    g["weight"] = 1.0
                merged_keys[key] = g
                groups.append(g)
            g["tokens"].add(cond.get("value", ""))
        else:
            groups.append({"kind": "single", "cond": cond})
    # Aho-Corasick（可选）构建
    automata = {}
    try:
        import ahocorasick  # type: ignore
        for g in groups:
            if g.get("kind") == "contains_any":
                A = ahocorasick.Automaton()
                for tok in g["tokens"]:
                    if tok:
                        A.add_word(tok, tok)
                A.make_automaton()
                automata[id(g)] = A
    except Exception:
        automata = {}
    # 评估
    details = []
    total = 0.0
    any_hit = False
    all_hit = True
    for idx, g in enumerate(groups, start=1):
        if g.get("kind") == "contains_any":
            col = g["column"]
            val = str(row.get(col, ""))
            opts = parse_options_local(g.get("options", ""))
            ignore_case = str(opts.get("ignore_case", "false")).lower() == "true"
            target = val.lower() if ignore_case else val
            hit = False
            score = 0.0
            A = automata.get(id(g))
            if A:
                try:
                    for _ in A.iter(target):
                        hit = True
                        score = 1.0
                        break
                except Exception:
                    hit = False
            else:
                # 回退：分批正则（避免巨型pattern）
                toks = [tok for tok in g["tokens"] if tok]
                if ignore_case:
                    toks = [tok.lower() for tok in toks]
                batch = 500
                import re as _re
                for i in range(0, len(toks), batch):
                    part = toks[i:i+batch]
                    try:
                        pat = "|".join([_re.escape(x) for x in part])
                        if _re.search(pat, target) is not None:
                            hit = True
                            score = 1.0
                            break
                    except Exception:
                        continue
            details.append((hit, score, f"{col}:contains_any({len(g['tokens'])})"))
            w = g.get("weight", 1.0)
            total += score * w
            any_hit = any_hit or hit
            all_hit = all_hit and hit
        else:
            cond = g["cond"]
            col = cond.get("column", "")
            val = str(row.get(col, ""))
            hit, score = apply_condition_local(val, cond)
            details.append((hit, score, f"{col}:{cond.get('type','')}/{cond.get('operator','')}={cond.get('value','')}"))
            w = 1.0
            try:
                w = float(cond.get("weight", "") or "1")
            except Exception:
                w = 1.0
            total += score * w
            any_hit = any_hit or hit
            all_hit = all_hit and hit
    if combine_mode == "AND":
        return (all_hit, total, details)
    if combine_mode == "OR":
        return (any_hit, total, details)
    return (total >= combine_threshold, total, details)
def read_excel_merged_local(pd, excel_path: str, sheet: str | None, limit: int | None):
    if not sheet or sheet.strip() == "":
        return pd.read_excel(excel_path, nrows=limit if limit else None)
    s = sheet.strip()
    if s == "*":
        dfs = pd.read_excel(excel_path, sheet_name=None, nrows=limit if limit else None)
        parts = list(dfs.values())
        if not parts:
            raise RuntimeError("未找到任何工作表")
        return pd.concat(parts, ignore_index=True)
    names = [x.strip() for x in s.split(",") if x.strip()]
    if len(names) == 1:
        return pd.read_excel(excel_path, sheet_name=names[0], nrows=limit if limit else None)
    parts = []
    for nm in names:
        try:
            dfp = pd.read_excel(excel_path, sheet_name=nm, nrows=limit if limit else None)
            parts.append(dfp)
        except Exception:
            pass
    if not parts:
        raise RuntimeError("指定的工作表均不存在")
    return pd.concat(parts, ignore_index=True)

def process_single(excel_path: str, require_path: str, col_major: str, threshold: float, out_path: str | None, sheet: str | None, progress_step: int, limit: int | None, append: bool, dedup: bool, dedup_key: str | None, progress_cb=None, log_cb=None, progress_text_cb=None):
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
    matched_so_far = 0
    t0 = time.time()
    for i, v in enumerate(majors, start=1):
        m, score = best_match(v, reqs)
        if m and score >= threshold:
            matches.append({"match": True, "matched_name": m["name"], "matched_code": m["code"], "score": round(score, 4)})
            matched_so_far += 1
        else:
            matches.append({"match": False, "matched_name": "", "matched_code": "", "score": round(score, 4)})
        if progress_cb and progress_step and progress_step > 0 and i % progress_step == 0:
            progress_cb(i, total)
            if progress_text_cb:
                bar = progress_text_cb("render", i, total)
                progress_text_cb("log", f"{bar} 已处理 {i}/{total} 行 | 已运行 {time.strftime('%H:%M:%S', time.gmtime(time.time()-t0))} | 命中 {matched_so_far} 行")
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
        self.only_merge = tk.BooleanVar(value=False)
        self.append_mode = tk.BooleanVar(value=False)
        self.dedup = tk.BooleanVar(value=False)
        self.dedup_key = tk.StringVar(value="")
        self.active_mode = tk.StringVar(value="multi")
        self.combine_mode = tk.StringVar(value="AND")
        self.combine_threshold = tk.StringVar(value="0.80")
        self.write_audit = tk.BooleanVar(value=False)
        self.conditions = []
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.running = False
        self.total_count = 0
        self.setup_style()
        self.create_widgets()
        self.load_config()
        self.root.after(100, self.consume_logs)
        self.root.after(100, self.consume_progress)

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
        # require.txt 仅旧版使用，移动到“专业列筛选（旧版）”分区
        # 模式Tabs
        tabs = ttk.Notebook(container)
        tabs.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        container.rowconfigure(1, weight=1)
        # 多条件筛选 Tab
        cond_tab = ttk.Frame(tabs)
        tabs.add(cond_tab, text="多条件筛选")
        cond_tab.columnconfigure(0, weight=1)
        cond_frame = ttk.Labelframe(cond_tab, text="筛选条件（CSV导入/编辑/导出）")
        cond_frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        cond_frame.columnconfigure(0, weight=1)
        btns = ttk.Frame(cond_frame)
        btns.grid(row=0, column=0, sticky="ew")
        ttk.Button(btns, text="导入条件CSV", command=self.import_conditions).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="新增条件", command=self.add_condition_dialog).grid(row=0, column=1, padx=4)
        ttk.Button(btns, text="删除选中", command=self.delete_selected_condition).grid(row=0, column=2, padx=4)
        ttk.Button(btns, text="导出条件CSV", command=self.export_conditions).grid(row=0, column=3, padx=4)
        mode_bar = ttk.Frame(cond_frame)
        mode_bar.grid(row=1, column=0, sticky="ew", pady=4)
        ttk.Label(mode_bar, text="组合模式").grid(row=0, column=0, sticky="w")
        mode_cb = ttk.Combobox(mode_bar, values=["AND","OR","WEIGHTED"], textvariable=self.combine_mode, state="readonly", width=12)
        mode_cb.grid(row=0, column=1, padx=6)
        ttk.Label(mode_bar, text="总阈值(加权)").grid(row=0, column=2, sticky="w")
        ttk.Entry(mode_bar, textvariable=self.combine_threshold, width=10).grid(row=0, column=3, padx=6)
        self.cond_view = ttk.Treeview(cond_frame, columns=("column","type","operator","value","threshold","priority","weight","options"), show="headings", height=8)
        for c in ("column","type","operator","value","threshold","priority","weight","options"):
            self.cond_view.heading(c, text=c)
            self.cond_view.column(c, width=120, anchor="w")
        self.cond_view.grid(row=2, column=0, sticky="nsew")
        cond_frame.rowconfigure(2, weight=1)
        sb_cond = ttk.Scrollbar(cond_frame, orient="vertical", command=self.cond_view.yview)
        sb_cond.grid(row=2, column=1, sticky="ns")
        self.cond_view.configure(yscrollcommand=sb_cond.set)
        # 专业列筛选 Tab
        major_tab = ttk.Frame(tabs)
        tabs.add(major_tab, text="专业列筛选（旧版）")
        for c in range(3):
            major_tab.columnconfigure(c, weight=1 if c == 1 else 0)
        r2 = 0
        ttk.Label(major_tab, text="专业列名").grid(row=r2, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(major_tab, textvariable=self.major_col).grid(row=r2, column=1, sticky="ew", padx=4, pady=2); r2+=1
        ttk.Label(major_tab, text="阈值").grid(row=r2, column=0, sticky="e", padx=4, pady=2)
        scl = ttk.Scale(major_tab, from_=0.0, to=1.0, variable=self.threshold, command=self.on_threshold_slider)
        scl.grid(row=r2, column=1, sticky="ew", padx=4, pady=2)
        ttk.Entry(major_tab, textvariable=self.threshold_input).grid(row=r2, column=2, sticky="ew", padx=4, pady=2); r2+=1
        ttk.Label(major_tab, text="参考：≤0.70宽松，0.75~0.85均衡，≥0.90严格；可填0~1或百分比").grid(row=r2, column=0, columnspan=3, sticky="w", padx=4, pady=2); r2+=1
        ttk.Label(major_tab, text="require.txt（旧版）").grid(row=r2, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(major_tab, textvariable=self.require_path).grid(row=r2, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(major_tab, text="选择", command=self.pick_require).grid(row=r2, column=2, sticky="e", padx=4, pady=2); r2+=1
        # 通用参数分区
        general = ttk.Labelframe(container, text="通用参数")
        general.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        for c in range(3):
            general.columnconfigure(c, weight=1 if c == 1 else 0)
        r = 0
        ttk.Label(general, text="Sheet").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(general, textvariable=self.sheet).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(general, text="进度步长").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(general, textvariable=self.progress_step).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        ttk.Label(general, text="limit").grid(row=r, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(general, textvariable=self.limit).grid(row=r, column=1, sticky="ew", padx=4, pady=2); r+=1
        # 输出设置分区
        output = ttk.Labelframe(container, text="输出设置")
        output.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)
        output.columnconfigure(1, weight=1)
        ttk.Label(output, text="输出目录").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(output, textvariable=self.out_dir).grid(row=0, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(output, text="选择", command=self.pick_out_dir).grid(row=0, column=2, sticky="e", padx=4, pady=2)
        ttk.Label(output, text="合并输出文件").grid(row=1, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(output, textvariable=self.merge_out).grid(row=1, column=1, sticky="ew", padx=4, pady=2)
        ttk.Checkbutton(output, text="仅合并输出（不写逐文件）", variable=self.only_merge).grid(row=2, column=0, sticky="w", padx=4, pady=2)
        # 处理选项分区
        options = ttk.Labelframe(container, text="处理选项")
        options.grid(row=4, column=0, sticky="nsew", padx=4, pady=4)
        options.columnconfigure(1, weight=1)
        ttk.Label(options, text="去重键").grid(row=0, column=0, sticky="e", padx=4, pady=2)
        ttk.Entry(options, textvariable=self.dedup_key).grid(row=0, column=1, sticky="ew", padx=4, pady=2)
        ttk.Checkbutton(options, text="追加模式", variable=self.append_mode).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Checkbutton(options, text="开启去重", variable=self.dedup).grid(row=1, column=1, sticky="w", padx=4, pady=2)
        ttk.Checkbutton(options, text="写出审计列", variable=self.write_audit).grid(row=2, column=0, sticky="w", padx=4, pady=2)
        # 监听Tab变化
        def on_tab_changed(event):
            idx = tabs.index(tabs.select())
            self.active_mode.set("multi" if idx == 0 else "major")
        tabs.bind("<<NotebookTabChanged>>", on_tab_changed)
        # 控制分区
        actions = ttk.Frame(container)
        actions.grid(row=5, column=0, sticky="ew", padx=4, pady=4)
        actions.columnconfigure(0, weight=1)
        left = ttk.Frame(actions)
        left.grid(row=0, column=0, sticky="w")
        right = ttk.Frame(actions)
        right.grid(row=0, column=1, sticky="e")
        ttk.Button(left, text="保存配置", command=self.save_config).grid(row=0, column=0, padx=4)
        ttk.Button(left, text="清除本地缓存", command=self.clear_cache).grid(row=0, column=1, padx=4)
        ttk.Button(left, text="软件使用须知", command=self.show_usage_notice).grid(row=0, column=2, padx=4)
        self.btn_start = ttk.Button(right, text="开始处理", command=self.start_processing)
        self.btn_start.grid(row=0, column=0, padx=4)
        self.btn_stop = ttk.Button(right, text="取消运行", command=self.stop_processing)
        self.btn_stop.grid(row=0, column=1, padx=4)
        self.btn_stop.configure(state="disabled")
        # 反馈分区
        feedback = ttk.Labelframe(container, text="进度与日志")
        feedback.grid(row=6, column=0, sticky="nsew", padx=4, pady=4)
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
    def import_conditions(self):
        p = filedialog.askopenfilename(title="导入条件CSV", filetypes=[("CSV/Excel", ".csv .xlsx .xls"), ("All files", "*.*")])
        if not p:
            return
        try:
            import pandas as pd
            ext = os.path.splitext(p)[1].lower()
            if ext in (".xlsx", ".xls"):
                df = pd.read_excel(p)
            else:
                df = pd.read_csv(p)
            required = ["column","type","operator","value","threshold","priority","weight","options"]
            for c in required:
                if c not in df.columns:
                    messagebox.showerror("错误", f"条件文件缺少列: {c}")
                    return
            self.conditions = []
            self.cond_view.delete(*self.cond_view.get_children())
            for _, r in df.fillna("").iterrows():
                item = {k: str(r[k]).strip() for k in required}
                self.conditions.append(item)
                self.cond_view.insert("", "end", values=tuple(item[k] for k in required))
            messagebox.showinfo("提示", "条件CSV已导入")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def add_condition_dialog(self):
        top = tk.Toplevel(self.root)
        top.title("新增条件")
        frm = ttk.Frame(top, padding=10)
        frm.pack(fill="both", expand=True)
        fields = ["column","type","operator","value","threshold","priority","weight","options"]
        vars = {f: tk.StringVar() for f in fields}
        for i, f in enumerate(fields):
            ttk.Label(frm, text=f).grid(row=i, column=0, sticky="e", padx=4, pady=2)
            ttk.Entry(frm, textvariable=vars[f]).grid(row=i, column=1, sticky="ew", padx=4, pady=2)
        frm.columnconfigure(1, weight=1)
        def ok():
            item = {f: vars[f].get().strip() for f in fields}
            self.conditions.append(item)
            self.cond_view.insert("", "end", values=tuple(item[f] for f in fields))
            top.destroy()
        ttk.Button(frm, text="确定", command=ok).grid(row=len(fields), column=1, sticky="e", padx=4, pady=8)

    def delete_selected_condition(self):
        sel = self.cond_view.selection()
        if not sel:
            return
        idxs = []
        for s in sel:
            vals = self.cond_view.item(s, "values")
            # 简单删除：按第一个匹配项删除
            for i, it in enumerate(self.conditions):
                if tuple(it[k] for k in ("column","type","operator","value","threshold","priority","weight","options")) == vals:
                    idxs.append(i); break
            self.cond_view.delete(s)
        for i in sorted(set(idxs), reverse=True):
            del self.conditions[i]

    def export_conditions(self):
        p = filedialog.asksaveasfilename(title="导出条件CSV", defaultextension=".csv", filetypes=[("CSV", ".csv")])
        if not p:
            return
        try:
            import pandas as pd
            df = pd.DataFrame(self.conditions)
            df.to_csv(p, index=False, encoding="utf-8-sig")
            messagebox.showinfo("提示", "条件CSV已导出")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def pick_out_dir(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.out_dir.set(p)

    def log_cb(self, msg: str):
        self.log_queue.put(msg)

    def progress_cb(self, done: int, total: int):
        try:
            self.progress_queue.put((done, total))
        except Exception:
            pass

    def _format_time(self, secs: float) -> str:
        secs = int(secs)
        h = secs // 3600
        m = (secs % 3600) // 60
        s = secs % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    def _render_progress(self, done: int, total: int, width: int = 30) -> str:
        if total <= 0 or done < 0:
            bar = "-" * width
            return f"[{bar}] ??%"
        pct = max(0, min(100, int(done * 100 / total)))
        filled = int(width * pct / 100)
        bar = "#" * filled + "-" * (width - filled)
        return f"[{bar}] {pct:02d}%"

    def consume_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log.insert(tk.END, msg + "\n")
                self.log.see(tk.END)
                # 控制日志大小，保留最后50,000字符
                try:
                    content_len = len(self.log.get("1.0", "end-1c"))
                except Exception:
                    content_len = 0
                if content_len > 50000:
                    self.log.delete("1.0", "end-40000c")
        except queue.Empty:
            pass
        self.root.after(100, self.consume_logs)

    def consume_progress(self):
        try:
            while True:
                done, total = self.progress_queue.get_nowait()
                if total <= 0:
                    self.progress["value"] = 0
                    self.progress["maximum"] = 1
                else:
                    self.progress["maximum"] = total
                    self.progress["value"] = done
        except queue.Empty:
            pass
        self.root.after(100, self.consume_progress)

    def start_processing(self):
        if not self.files:
            messagebox.showwarning("提示", "请添加至少一个Excel文件")
            return
        req = self.require_path.get().strip()
        mode = self.active_mode.get().strip()
        if mode == "major" and not req:
            messagebox.showwarning("提示", "请在旧版模式下指定require.txt文件")
            return
        if mode == "multi" and not self.conditions and not req:
            messagebox.showwarning("提示", "请导入条件CSV；如需回退到旧版，请同时指定require.txt")
            return
        try:
            import pandas as pd  # noqa
        except Exception:
            messagebox.showerror("错误", "需要安装pandas与openpyxl:\n pip install pandas openpyxl")
            return
        try:
            self.running = True
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(state="normal")
        except Exception:
            pass
        self.total_count = 0
        self.progress["value"] = 0
        self.log_cb("开始处理")
        t = threading.Thread(target=self.run_processing, daemon=True)
        t.start()

    def stop_processing(self):
        try:
            self.running = False
            self.btn_stop.configure(state="disabled")
        except Exception:
            pass
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
            "用于根据CSV条件文件，对所选Excel文件的多列进行多条件筛选（支持AND/OR/加权），\n"
            "保留原专业模糊匹配能力，并支持批量合并导出、追加与去重。\n\n"
            "使用方法：\n"
            "1. 添加文件：选择一个或多个Excel文件。\n"
            "2. 移除选中：从列表中移除不需要处理的文件。\n"
            "3. 选择require.txt（仅旧版）：在“专业列筛选（旧版）”标签页中指定专业要求文本文件。\n"
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
            "条件文件（CSV）规范：\n"
            "第一行必须是表头：column,type,operator,value,threshold,priority,weight,options\n"
            "每行一条条件，例如：\n"
            "Major,fuzzy,similar,信息资源管理,0.85,3,1.0,normalize=true;code_prefer=true\n"
            "GPA,number,between,3.5-4.0,,1,1.0,\n"
            "Email,regex,match,^\\S+@\\S+\\.\\S+$,,1,1.0,\n"
            "组合模式：AND（全部命中）、OR（任意命中）、WEIGHTED（加权总分超过阈值即命中）。\n\n"
            "Sheet支持多表：\n"
            "- 留空：读每个文件的首个工作表。\n"
            "- 逗号分隔多个工作表名：如“Sheet1,Sheet2”，按顺序读取并合并。\n"
            "- 星号“*”：合并该文件的所有工作表。\n\n"
            "字段说明与可填项：\n"
            "- column（必填）：Excel列名；缺列时该条件视为不命中。\n"
            "- type（必填）：text|number|regex|fuzzy|enum|boolean|code。\n"
            "- operator（必填）：随type变化：\n"
            "  · text：equals（完全相等）|contains（包含）|startswith（以…开头）|endswith（以…结尾）\n"
            "  · number：between（在“最小-最大”范围内）|min（≥最小值）|max（≤最大值）|equals（等于指定值）\n"
            "  · regex：match（满足正则表达式）\n"
            "  · fuzzy：similar（按相似度匹配，配合threshold判定）\n"
            "  · enum：in（属于给定集合）\n"
            "  · boolean：is（与true/false匹配）\n"
            "  · code：equals（编码数字部分完全一致）\n"
            "- value（必填）：按类型填写：\n"
            "  · text：任意字符串（可配合 ignore_case/normalize）\n"
            "  · number-between：min-max；min/max/equals 填单值\n"
            "  · regex-match：正则表达式（如邮箱的 ^\\S+@\\S+\\.\\S+$）\n"
            "  · fuzzy-similar：目标字符串（如“信息资源管理”“Python”）\n"
            "  · enum-in：分号分隔（如“本科;硕士;博士”）\n"
            "  · code-equals：建议填纯数字编码（4~6位），系统抽取数字部分比对\n"
            "- threshold（选填，仅fuzzy）：0~1或百分比（如0.80/80%）；参考：≤0.70宽松、0.75~0.85均衡、≥0.90严格\n"
            "- priority（选填）：用于展示排序，当前不参与命中计算\n"
            "- weight（选填，用于WEIGHTED）：默认1.0，可设0（不计分）\n"
            "- options（选填，键值对；以“;”分隔）：\n"
            "  · ignore_case=true|false：文本是否忽略大小写（默认false）\n"
            "  · normalize=true|false：文本是否规范化（默认false；半角化、去空白与标点、统一小写）\n"
            "  · code_prefer=true|false：模糊匹配时是否优先以编码全等判定命中（默认false）\n"
            "  · 书写示例：ignore_case=true;normalize=true;code_prefer=true\n\n"
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
        combine_mode = self.combine_mode.get().strip() or "AND"
        try:
            ct = self.combine_threshold.get().strip()
            if ct.endswith("%"):
                combine_threshold = float(ct[:-1]) / 100.0
            else:
                combine_threshold = float(ct)
        except Exception:
            combine_threshold = 0.8
        outputs = []
        merged_parts = []
        total_count = 0
        try:
            import pandas as pd
            for pth in self.files:
                base = os.path.splitext(os.path.basename(pth))[0]
                out_path = None if bool(self.only_merge.get()) else (os.path.join(out_dir, f"{base}_filtered.xlsx") if out_dir else None)
                mode = self.active_mode.get()
                if mode == "multi":
                    df = read_excel_merged_local(pd, pth, sheet, limit)
                    if not self.conditions:
                        self.log_cb("未配置条件，已回退到专业列筛选")
                        saved, count = process_single(pth, req, col_major, threshold, out_path, sheet, progress_step, limit, append, dedup, dedup_key, progress_cb=self.progress_cb, log_cb=self.log_cb, progress_text_cb=lambda kind, *args: (self._render_progress(args[0], args[1]) if kind=="render" else self.log_cb(args[0])))
                    else:
                        total = len(df)
                        hits = []
                        scores = []
                        file_start = time.time()
                        file_matched = 0
                        for i in range(len(df)):
                            if not self.running:
                                break
                            row = {c: str(df.iloc[i][c]) if c in df.columns else "" for c in df.columns}
                            hit, score_all, ds = evaluate_conditions_row_local(row, self.conditions, combine_mode, combine_threshold)
                            hits.append(hit)
                            scores.append(round(score_all, 4))
                            if hit:
                                file_matched += 1
                            if bool(self.write_audit.get()):
                                for j, (h, s, desc) in enumerate(ds, start=1):
                                    col_hit = f"_cond_{j}_match"; col_score = f"_cond_{j}_score"; col_desc = f"_cond_{j}_desc"
                                    if col_hit not in df.columns:
                                        df[col_hit] = ""
                                    if col_score not in df.columns:
                                        df[col_score] = 0.0
                                    if col_desc not in df.columns:
                                        df[col_desc] = ""
                                    df.at[i, col_hit] = h
                                    df.at[i, col_score] = round(s, 4)
                                    df.at[i, col_desc] = desc
                            if progress_step and progress_step > 0 and (i + 1) % progress_step == 0:
                                self.progress_cb(i + 1, total)
                                bar = self._render_progress(i + 1, total)
                                self.log_cb(f"{bar} 已处理 {i+1}/{total} 行 | 已运行 {self._format_time(time.time()-file_start)} | 命中 {file_matched} 行")
                        df["_match_all"] = hits
                        df["_score_all"] = scores
                        out_df = df[df["_match_all"] == True].copy()
                        if bool(self.only_merge.get()):
                            # 仅合并输出：不写逐文件，直接入合并池（可先局部去重以降低内存）
                            part = out_df
                            if dedup:
                                part = dedup_dataframe(part, col_major, dedup_key)
                            merged_parts.append(part)
                            count = len(part)
                            saved = "(仅合并)"
                            self.log_cb(f"筛选完成：{os.path.basename(pth)} 命中 {count} 条（已加入合并）")
                        else:
                            if out_path is None or out_path == "":
                                out_path = os.path.join(os.path.dirname(pth), f"{base}_filtered.xlsx")
                            saved = write_output(out_df, out_path, append=append, dedup=dedup, dedup_key=dedup_key, col_major=col_major)
                            count = len(out_df)
                            self.log_cb(f"筛选完成：{os.path.basename(pth)} 命中 {count} 条 → {saved}")
                else:
                    saved, count = process_single(pth, req, col_major, threshold, out_path, sheet, progress_step, limit, append, dedup, dedup_key, progress_cb=self.progress_cb, log_cb=self.log_cb)
                outputs.append(saved)
                total_count += count
                try:
                    if not bool(self.only_merge.get()):
                        if isinstance(saved, str) and (saved.lower().endswith(".xlsx") or saved.lower().endswith(".csv")):
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
            self.root.after(0, lambda: messagebox.showinfo("完成", f"处理完成，共筛选 {total_count} 条"))
        except Exception as e:
            import traceback
            err = f"{e}\n{traceback.format_exc()}"
            self.log_cb(f"错误：{err}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            try:
                self.running = False
                self.root.after(0, lambda: self.btn_start.configure(state="normal"))
                self.root.after(0, lambda: self.btn_stop.configure(state="disabled"))
            except Exception:
                pass

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
            "only_merge": bool(self.only_merge.get()),
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
            try:
                self.only_merge.set(cfg.get("only_merge", False))
            except Exception:
                pass
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
