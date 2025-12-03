import json
import os
from builtin_datasets import get_builtin

# 说明：
# 本模块负责题库的读取（Markdown/JSON/内置）与写出（JSON），
# 同时提供将内置题库首次写出到用户数据目录（适合打包为exe后运行）。
# 为了避免打包后无法写入程序所在目录，所有持久化文件均写入到用户数据目录：
# Windows：%APPDATA%/EnglishPracticeApp/datasets
# 其他系统：~/.EnglishPracticeApp/datasets

def _has_chinese(s: str) -> bool:
    return any('\u4e00' <= ch <= '\u9fff' for ch in s)

def load_md(path: str):
    # 解析Markdown文件，将连续中文行合并，并在遇到英文行时与前中文配对
    data = []
    zh_buf = []
    if not os.path.exists(path):
        return data
    with open(path, 'r', encoding='utf-8') as f:
        for raw in f:
            line = raw.strip()
            if not line:
                continue
            if _has_chinese(line):
                zh_buf.append(line)
            else:
                en = line
                if zh_buf:
                    zh = ''.join(zh_buf)
                    data.append({'zh': zh, 'en': en})
                    zh_buf = []
    return data

def load_json(path: str):
    # 读取统一结构的JSON题库
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        obj = json.load(f)
    out = []
    for item in obj:
        zh = item.get('zh', '')
        en = item.get('en', '')
        if isinstance(zh, str) and isinstance(en, str) and zh and en:
            out.append({'zh': zh, 'en': en})
    return out

def export_json(path: str, data):
    # 将题库写出为JSON文件（UTF-8，缩进2）
    base = os.path.dirname(path)
    if base and not os.path.exists(base):
        os.makedirs(base, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def export_md(path: str, data):
    # 将题库写出为Markdown文件：每条中文+英文，之间空行
    base = os.path.dirname(path)
    if base and not os.path.exists(base):
        os.makedirs(base, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        for item in data:
            zh = item.get('zh', '').strip()
            en = item.get('en', '').strip()
            if not zh or not en:
                continue
            f.write(f"{zh}\n{en}\n\n")

def load_builtin(name: str):
    # 加载代码内置题库
    return get_builtin(name)

def get_app_data_dir():
    # 返回用户数据目录（用于保存datasets与registry等持久化文件）
    appdata = os.getenv('APPDATA')
    if appdata:
        base = os.path.join(appdata, 'EnglishPracticeApp')
    else:
        base = os.path.join(os.path.expanduser('~'), '.EnglishPracticeApp')
    os.makedirs(base, exist_ok=True)
    return base

def get_datasets_dir():
    # 返回datasets目录路径
    base = get_app_data_dir()
    datasets_dir = os.path.join(base, 'datasets')
    os.makedirs(datasets_dir, exist_ok=True)
    return datasets_dir

def export_bootstrap_jsons():
    # 首次运行将内置题库写出到用户数据目录的datasets中
    datasets_dir = get_datasets_dir()
    d2 = get_builtin('d2')
    d3 = get_builtin('d3')
    d4 = get_builtin('d4')
    if d2:
        export_json(os.path.join(datasets_dir, 'd2.json'), d2)
    if d3:
        export_json(os.path.join(datasets_dir, 'd3.json'), d3)
    if d4:
        export_json(os.path.join(datasets_dir, 'd4.json'), d4)
