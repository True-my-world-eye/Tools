import tkinter as tk
import random
import os
import json
import sys
from tkinter import font, ttk, filedialog, messagebox
from dataset_loader import (
    load_md, load_json, export_json, export_md, load_builtin,
    export_bootstrap_jsons, get_datasets_dir
)

class EnglishPracticeApp:
    def __init__(self, root):
        # 初始化主窗口与基础样式配置
        self.root = root
        self.root.title("英语口语练习")
        self.root.geometry("900x760")
        self.root.resizable(False, False)
        self.theme_color = "#2196F3"
        self.secondary_color = "#64B5F6"
        self.bg_color = "#F5F5F5"
        self.text_color = "#212121"
        self.title_font = font.Font(family="Microsoft YaHei", size=24, weight="bold")
        self.font_config = font.Font(family="Microsoft YaHei", size=14)
        self.button_font = font.Font(family="Microsoft YaHei", size=12, weight="bold")
        self.root.configure(bg=self.bg_color)
        try:
            def resource_path(rel):
                base = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
                return os.path.join(base, rel)
            ico = resource_path('assets/app.ico')
            png = resource_path('assets/软件图标.png')
            if os.path.exists(ico):
                self.root.iconbitmap(ico)
            if os.path.exists(png):
                try:
                    img = tk.PhotoImage(file=png)
                    self.root.iconphoto(False, img)
                except Exception:
                    pass
        except Exception:
            pass
        # 打包优化：所有持久化文件放入用户数据目录
        self.base_dir = os.path.dirname(__file__)
        export_bootstrap_jsons()
        self.datasets_dir = get_datasets_dir()
        self.registry_path = os.path.join(self.datasets_dir, "registry.json")
        self.display_mode = tk.StringVar(value="zh")
        self.practice_count_var = tk.IntVar(value=3)
        # 默认题库选择（支持内置与JSON）
        self.selected_dataset_name = tk.StringVar(value="大英3")
        self.dataset_sources = {
            "大英2": ("builtin", "d2"),
            "大英3": ("builtin", "d3"),
            "大英4": ("builtin", "d4")
        }
        self.load_registry_into_sources()
        self.dataset = self.load_selected_dataset(initial=True)
        self.create_widgets()

    def create_widgets(self):
        # 创建界面组件：工具栏、Notebook页签、文本框与状态信息
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            style.theme_use("clam")
        style.configure('Toolbar.TButton', padding=(10, 6))
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=24, pady=20)
        title_label = tk.Label(main_frame, text="英语口语练习", font=self.title_font, bg=self.bg_color, fg=self.theme_color)
        title_label.pack(pady=(0, 10))
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill="x", pady=6)
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入题库", command=self.import_dataset)
        file_menu.add_command(label="导出当前题库", command=self.export_dataset)
        file_menu.add_command(label="打开存储目录", command=self.open_storage_dir)
        menubar.add_cascade(label="文件", menu=file_menu)
        self.root.config(menu=menubar)
        ttk.Label(toolbar, text="题库").pack(side=tk.LEFT, padx=(0, 6))
        self.dataset_box = ttk.Combobox(toolbar, state="readonly", textvariable=self.selected_dataset_name,
                                        values=list(self.dataset_sources.keys()), width=16)
        self.dataset_box.pack(side=tk.LEFT)
        self.dataset_box.bind("<<ComboboxSelected>>", lambda e: self.load_selected_dataset())
        ttk.Label(toolbar, text="显示模式").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Radiobutton(toolbar, text="中文", variable=self.display_mode, value="zh").pack(side=tk.LEFT)
        ttk.Radiobutton(toolbar, text="英文", variable=self.display_mode, value="en").pack(side=tk.LEFT)
        ttk.Radiobutton(toolbar, text="中英", variable=self.display_mode, value="both").pack(side=tk.LEFT)
        ttk.Label(toolbar, text="计时(秒)").pack(side=tk.LEFT, padx=(12, 6))
        self.seconds_var = tk.IntVar(value=30)
        self.seconds_spin = ttk.Spinbox(toolbar, from_=5, to=180, textvariable=self.seconds_var, width=5)
        self.seconds_spin.pack(side=tk.LEFT)
        ttk.Button(toolbar, text="查看全部", command=self.view_all, style='Toolbar.TButton').pack(side=tk.RIGHT, padx=6)
        self.end_button = ttk.Button(toolbar, text="结束练习", command=self.show_english, state=tk.DISABLED, style='Toolbar.TButton')
        self.end_button.pack(side=tk.RIGHT, padx=6)
        ttk.Button(toolbar, text="开始练习", command=self.start_practice, style='Toolbar.TButton').pack(side=tk.RIGHT, padx=6)
        instruction = ttk.Label(main_frame, text="选择题库与显示模式后，点击开始练习。30秒后自动显示中英对照。",
                                font=self.font_config)
        instruction.pack(fill="x", pady=(6, 12))
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True, pady=(0, 10))
        practice_frame = ttk.Frame(notebook)
        browse_frame = ttk.Frame(notebook)
        notebook.add(practice_frame, text="练习")
        notebook.add(browse_frame, text="题库浏览")
        text_frame = tk.Frame(practice_frame, bg=self.theme_color, padx=2, pady=2)
        text_frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sentence_text = tk.Text(text_frame, width=70, height=10, font=self.font_config, wrap=tk.WORD,
                                     bg="white", fg=self.text_color, relief="flat", padx=15, pady=15,
                                     yscrollcommand=scrollbar.set)
        self.sentence_text.pack(fill="both", expand=True, side=tk.LEFT)
        scrollbar.config(command=self.sentence_text.yview)
        self.sentence_text.config(state=tk.DISABLED)
        browse_text_frame = tk.Frame(browse_frame, bg=self.theme_color, padx=2, pady=2)
        browse_text_frame.pack(fill="both", expand=True)
        bscroll = tk.Scrollbar(browse_text_frame)
        bscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.browse_text = tk.Text(browse_text_frame, width=80, height=20, font=self.font_config, wrap=tk.WORD,
                                   bg="white", fg=self.text_color, relief="flat", padx=15, pady=15,
                                   yscrollcommand=bscroll.set)
        self.browse_text.pack(fill="both", expand=True, side=tk.LEFT)
        bscroll.config(command=self.browse_text.yview)
        self.browse_text.config(state=tk.DISABLED)
        self.timer_label = tk.Label(practice_frame, text="剩余时间: 30秒", font=self.font_config, bg=self.bg_color, fg=self.text_color)
        self.timer_label.pack(pady=(8, 4))
        self.status_label = tk.Label(practice_frame, text="准备就绪", font=self.font_config, bg=self.bg_color, fg=self.text_color)
        self.status_label.pack(pady=(0, 8))
        # 右下角说明图标入口，始终可见
        self.info_btn = ttk.Button(self.root, text="ℹ️", command=self.show_info, width=2)
        self.info_btn.place(relx=1.0, rely=1.0, x=-12, y=-12, anchor='se')
        self.info_label = ttk.Label(self.root, text="软件说明", foreground=self.text_color)
        self.info_label.place(relx=1.0, rely=1.0, x=-60, y=-12, anchor='se')

    def start_practice(self):
        # 开始练习：按显示模式展示随机抽样句子并启动倒计时
        mode = self.display_mode.get()
        count = max(1, int(self.practice_count_var.get()))
        if not self.dataset:
            self.status_label.config(text="题库为空")
            return
        sample_count = min(len(self.dataset), count)
        samples = random.sample(self.dataset, sample_count)
        self.selected_pairs = samples
        secs = int(self.seconds_var.get())
        self.remaining_time = secs
        self.timer_label.config(text=f"剩余时间: {self.remaining_time}秒")
        self.sentence_text.config(state=tk.NORMAL)
        self.sentence_text.delete(1.0, tk.END)
        if mode == "zh":
            for item in samples:
                self.sentence_text.insert(tk.END, f"{item['zh']}\n\n")
            self.status_label.config(text="正在显示中文句子...")
            self.end_button.config(state=tk.NORMAL)
            self.timer_id = self.root.after(1000, self.update_timer)
        elif mode == "en":
            for item in samples:
                self.sentence_text.insert(tk.END, f"{item['en']}\n\n")
            self.status_label.config(text="正在显示英文句子...")
            self.end_button.config(state=tk.NORMAL)
            self.timer_id = self.root.after(1000, self.update_timer)
        else:
            for item in samples:
                self.sentence_text.insert(tk.END, f"{item['zh']}\n{item['en']}\n\n")
            self.status_label.config(text="已显示中英")
            self.end_button.config(state=tk.DISABLED)
        self.sentence_text.config(state=tk.DISABLED)

    def update_timer(self):
        # 每秒更新倒计时，归零后显示中英对照
        if self.remaining_time > 0:
            self.remaining_time -= 1
            self.timer_label.config(text=f"剩余时间: {self.remaining_time}秒")
            self.timer_id = self.root.after(1000, self.update_timer)
        else:
            self.show_english()

    def show_english(self):
        # 展示中英对照，并关闭倒计时
        self.remaining_time = 0
        self.timer_label.config(text="时间到！")
        if hasattr(self, 'timer_id'):
            try:
                self.root.after_cancel(self.timer_id)
            except Exception:
                pass
        self.status_label.config(text="正在显示中英对照...")
        self.sentence_text.config(state=tk.NORMAL)
        self.sentence_text.delete(1.0, tk.END)
        pairs = getattr(self, 'selected_pairs', [])
        for item in pairs:
            self.sentence_text.insert(tk.END, f"{item['zh']}\n{item['en']}\n\n")
        self.sentence_text.config(state=tk.DISABLED)
        self.end_button.config(state=tk.DISABLED)

    def load_selected_dataset(self, initial=False):
        # 根据下拉选择加载题库（内置/builtin 或 JSON/md）
        name = self.selected_dataset_name.get()
        if not name and initial:
            name = "大英3"
            self.selected_dataset_name.set(name)
        kind, path = self.dataset_sources.get(name, (None, None))
        data = []
        if kind == "json":
            data = load_json(path)
        elif kind == "md":
            data = load_md(path)
        elif kind == "builtin":
            data = load_builtin(path)
        self.dataset = data
        return data

    def view_all(self):
        # 查看全部：按当前显示模式渲染整库内容
        mode = self.display_mode.get()
        self.browse_text.config(state=tk.NORMAL)
        self.browse_text.delete(1.0, tk.END)
        for item in self.dataset:
            if mode == "zh":
                self.browse_text.insert(tk.END, f"{item['zh']}\n\n")
            elif mode == "en":
                self.browse_text.insert(tk.END, f"{item['en']}\n\n")
            else:
                self.browse_text.insert(tk.END, f"{item['zh']}\n{item['en']}\n\n")
        self.browse_text.config(state=tk.DISABLED)

    def import_dataset(self):
        # 导入题库：解析md/json，保存为user_*.json，并写入registry持久化
        file_path = filedialog.askopenfilename(filetypes=[("题库文件", "*.json;*.md"), ("JSON", "*.json"), ("Markdown", "*.md")])
        if not file_path:
            return
        name = os.path.splitext(os.path.basename(file_path))[0]
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".json":
            data = load_json(file_path)
            kind = "json"
        else:
            data = load_md(file_path)
            kind = "md"
        if not data:
            messagebox.showwarning("导入失败", "未解析到有效数据")
            return
        safe_name = name
        base_dir = self.datasets_dir
        os.makedirs(base_dir, exist_ok=True)
        save_name = f"user_{safe_name}.json"
        target_path = os.path.join(base_dir, save_name)
        idx = 1
        while os.path.exists(target_path):
            save_name = f"user_{safe_name}_{idx}.json"
            target_path = os.path.join(base_dir, save_name)
            idx += 1
        export_json(target_path, data)
        self.append_registry({"name": safe_name, "path": target_path, "kind": "json"})
        self.dataset_sources[safe_name] = ("json", target_path)
        self.dataset_box["values"] = list(self.dataset_sources.keys())
        self.selected_dataset_name.set(name)
        self.dataset = data
        self.status_label.config(text=f"已导入题库：{name}")

    def export_dataset(self):
        # 导出当前题库为JSON或Markdown
        if not self.dataset:
            messagebox.showinfo("导出", "当前题库为空")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".json",
                                                filetypes=[("JSON", "*.json"), ("Markdown", "*.md")])
        if not save_path:
            return
        ext = os.path.splitext(save_path)[1].lower()
        if ext == ".md":
            export_md(save_path, self.dataset)
        else:
            export_json(save_path, self.dataset)
        messagebox.showinfo("导出", "导出完成")

    def copy_current(self):
        # 复制当前练习区文本到剪贴板
        try:
            text = self.sentence_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_label.config(text="已复制到剪贴板")
        except Exception:
            pass

    def load_registry_into_sources(self):
        # 启动时读取registry.json，将用户导入题库加载到下拉选项
        try:
            if os.path.exists(self.registry_path):
                with open(self.registry_path, 'r', encoding='utf-8') as f:
                    items = json.load(f)
                for it in items:
                    nm = it.get('name')
                    p = it.get('path')
                    if nm and p and os.path.exists(p):
                        self.dataset_sources[nm] = ("json", p)
        except Exception:
            pass

    def append_registry(self, item):
        # 将新导入题库写入registry.json，实现持久化
        items = []
        try:
            if os.path.exists(self.registry_path):
                with open(self.registry_path, 'r', encoding='utf-8') as f:
                    items = json.load(f)
        except Exception:
            items = []
        items.append(item)
        os.makedirs(os.path.dirname(self.registry_path), exist_ok=True)
        with open(self.registry_path, 'w', encoding='utf-8') as f:
            json.dump(items, f, ensure_ascii=False, indent=2)

    def show_info(self):
        # 弹出软件说明与作者信息（导入规则与使用方式详解）
        info = (
            "【软件简介】\n"
            "多题库英语口语练习工具，支持随机练习、倒计时、中文/英文/中英三种显示模式，含题库浏览、导入/导出与复制。\n\n"
            "【使用方式】\n"
            "1. 选择题库：《大英2》《大英3》《大英4》为内置题库，直接可用。\n"
            "2. 选择显示模式：中文/英文/中英。中文/英文模式下倒计时结束会自动显示中英对照；中英模式直接显示对照。\n"
            "3. 设置计时(秒)：默认30秒，可在5~180秒间调整。\n"
            "4. 开始练习：随机抽取若干句（当前为3句），展示并倒计时。\n"
            "5. 查看全部：按当前显示模式一次性浏览整个题库。\n"
            "6. 复制当前显示：将练习区文本复制到剪贴板，便于备份或分享。\n\n"
            "【菜单与按钮布局】\n"
            "- 顶部菜单栏：文件 → 导入题库、导出当前题库、打开存储目录。\n"
            "- 工具栏：左侧显示‘复制当前显示’与‘说明’；右侧显示‘查看全部’、‘结束/显示英文’、‘开始练习’。\n\n"
            "【导入题库规则】\n"
            "1. 支持导入格式：Markdown(.md) 与 JSON(.json)。\n"
            "2. Markdown规则：\n"
            "   - 连续中文行会被合并为一个中文段，紧随其后的一行英文作为对应翻译；空行与编号可忽略。\n"
            "   - 示例：\n"
            "     中文第1行\n"
            "     中文第2行（与上一行合并）\n"
            "     English line for the above Chinese\n"
            "3. JSON结构：数组形式，每项含 zh/en 字段，如：\n"
            "   [ {\"zh\": \"中文\", \"en\": \"English\" }, ... ]\n"
            "4. 导入后的保存：自动解析并保存为 user_名称.json 到用户数据目录（%APPDATA%/EnglishPracticeApp/datasets），同时登记到 registry.json，重启后仍可在下拉中选择。\n"
            "5. 重名处理：若存在同名文件，会自动追加序号（如 user_名称_1.json）。\n"
            "6. 移除导入题库：当前版本暂未提供界面移除；如需删除，可手动删除 datasets 目录下对应 user_*.json，并编辑或删除 registry.json。\n\n"
            "【内置题库说明】\n"
            "- 内置《大英2/3/4》随软件打包，不依赖外部文件；首次运行会将其写出为 d2/d3/d4.json 以便分享或导出，但默认读取走内置数据，不重复从JSON读取。\n\n"
            "【导出说明】\n"
            "- 支持导出为 JSON(.json) 或 Markdown(.md)。\n"
            "- Markdown导出格式为：每条中文一行、对应英文一行，之间空一行。\n\n"
            "【存储目录】\n"
            "- 用户数据路径：%APPDATA%/EnglishPracticeApp/datasets（若无APPDATA则为用户主目录/.EnglishPracticeApp/datasets）。\n"
            "- 可在‘文件’菜单选择‘打开存储目录’直接查看。\n\n"
            "【版权与作者】\n"
            "本软件免费开放传播。\n"
            "作者：True my world eye\n"
            "微信号：Truemwe\n"
            "邮箱：hwzhang0722@163.com"
        )
        messagebox.showinfo("软件说明", info)

    def open_storage_dir(self):
        # 打开用户数据存储目录
        try:
            os.startfile(self.datasets_dir)
        except Exception:
            messagebox.showinfo("打开目录", self.datasets_dir)

if __name__ == "__main__":
    root = tk.Tk()
    app = EnglishPracticeApp(root)
    root.mainloop()
