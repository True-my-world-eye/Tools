import argparse
import os
import re
import sys
from difflib import SequenceMatcher
from typing import List, Tuple, Optional, Dict

"""
本脚本用于从Excel中根据require.txt给出的专业要求进行模糊匹配筛选：
1. 解析require.txt中“专业名称（编码）”或带编号行的条目，生成要求列表
2. 对Excel指定列（默认Major）逐行匹配：优先按编码匹配，其次规范化文本匹配，再相似度匹配
3. 支持多文件批量处理、进度可视化、合并导出、以及可选去重
4. 输出包含命中标记/命中名称/编码/分数，便于回溯与调参
参数默认值可在下方常量处直接修改，便于统一配置。
"""

# 默认输入Excel文件列表（支持多文件）
DEFAULT_EXCELS = ["TMT_FIGUREINFO1.xlsx"]
# 专业要求文本路径
DEFAULT_REQUIRE = "require.txt"
# Excel中专业列名
DEFAULT_MAJOR_COL = "Major"
# 匹配阈值（0~1，越大越严格）
DEFAULT_THRESHOLD = 0.8
# 进度输出的步长（每处理N行打印一次）
DEFAULT_PROGRESS_STEP = 100
# 限制读取的最大行数（None表示不限制；可设为1000用于调试）
DEFAULT_LIMIT = None

def to_halfwidth(s: str) -> str:
    """
    将全角字符转换为半角，统一中英文符号形态，减少格式差异。
    """
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
    """
    将字符串标准化：
    - 半角化
    - 去首尾空白、转小写
    - 去所有空白
    - 仅保留中文、字母、数字
    用于降低匹配过程中的格式噪声。
    """
    s = to_halfwidth(s)
    s = s.strip().lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", s)
    return s

def extract_code(s: str) -> Optional[str]:
    """
    从字符串中提取专业编码（如080914TK/120503等），
    并返回其数字部分（去字母后缀），用于编码优先匹配。
    """
    m = re.search(r"(\d{4,6}[A-Z]{0,3})", s)
    if not m:
        return None
    return re.sub(r"[^0-9]", "", m.group(1))

def read_text_file(path: str) -> List[str]:
    """
    尝试多种编码读取文本文件（utf-8、gbk、utf-8-sig），
    读不到则忽略错误尽力读取，返回行列表。
    """
    for enc in ("utf-8", "gbk", "utf-8-sig"):
        try:
            with open(path, "r", encoding=enc) as f:
                return [line.rstrip("\n") for line in f]
        except Exception:
            continue
    with open(path, "r", errors="ignore") as f:
        return [line.rstrip("\n") for line in f]

def resolve_file_path(p: Optional[str]) -> str:
    """
    解析文件路径：支持相对路径、绝对路径。
    优先按当前工作目录，其次按脚本所在目录进行查找。
    """
    if not p:
        raise FileNotFoundError("未提供文件路径")
    if os.path.isabs(p) and os.path.exists(p):
        return p
    if os.path.exists(p):
        return p
    script_dir = os.path.dirname(os.path.abspath(__file__))
    alt = os.path.join(script_dir, p)
    if os.path.exists(alt):
        return alt
    raise FileNotFoundError(f"未找到文件: {p}（尝试于工作目录与脚本目录）")

def parse_requirements(lines: List[str]) -> List[Dict[str, str]]:
    """
    解析require.txt：
    - 支持“专业名称（编码）”格式，例如：信息资源管理（120503）
    - 也支持带编号列的行，例如：397 计算机类 080901 计算机科学与技术
    产出去重后的要求项：raw原文、name专业名称、norm规范化名称、code编码数字。
    """
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
    """
    基于SequenceMatcher计算相似度，范围0~1。
    """
    return SequenceMatcher(None, a, b).ratio()

def best_match(major: str, reqs: List[Dict[str, str]]) -> Tuple[Optional[Dict[str, str]], float]:
    """
    为给定Excel行的专业字段major寻找最佳匹配：
    - 若两边编码一致，分数=1.0
    - 若规范化文本完全相等，分数=0.95
    - 若互为子串，分数=0.9
    - 否则使用相似度分数
    返回命中项及分数。
    """
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

def ensure_pandas() -> Optional[object]:
    """
    安全导入pandas，失败返回None以提示安装。
    """
    try:
        import pandas as pd  # type: ignore
        return pd
    except Exception:
        return None

def write_output(df, out_path: str, append: bool = False, dedup: bool = False, dedup_key: Optional[str] = None, col_major: str = DEFAULT_MAJOR_COL) -> str:
    """
    将结果DataFrame写出：
    - 优先写Excel（xlsx/xls）
    - 写失败自动降级为CSV（utf-8-sig）
    - 若append为True且目标文件已存在，则读取旧文件与新结果合并后写出
      （如启用dedup则按指定规则进行去重）
    返回最终写出路径。
    """
    ext = os.path.splitext(out_path)[1].lower()
    if append and os.path.exists(out_path):
        try:
            import pandas as pd
            if ext in (".xlsx", ".xls"):
                old = pd.read_excel(out_path)
            else:
                old = pd.read_csv(out_path)
            new_df = df
            merged = pd.concat([old, new_df], ignore_index=True)
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

def print_progress(done: int, total: int):
    """
    文本进度条展示：每处理progress-step行时打印一次。
    """
    pct = 0 if total == 0 else int(done * 100 / total)
    bar_len = 20
    filled = int(bar_len * pct / 100)
    bar = "#" * filled + "-" * (bar_len - filled)
    print(f"[{bar}] {done}/{total} {pct}%")

def process(excel_path: str, require_path: str, col_major: str, threshold: float, out_path: Optional[str], sheet: Optional[str], progress_step: int, limit: Optional[int], append: bool = False, dedup: bool = False, dedup_key: Optional[str] = None) -> str:
    """
    处理单个Excel：
    - 读取require.txt并解析要求项
    - 读取Excel（可选指定sheet与limit）
    - 对major列逐行匹配，生成命中信息与分数
    - 根据阈值筛选命中行并写出结果
    返回输出文件路径。
    """
    pd = ensure_pandas()
    if pd is None:
        raise RuntimeError("需要安装pandas，请运行: pip install pandas openpyxl")
    require_path = resolve_file_path(require_path)
    lines = read_text_file(require_path)
    reqs = parse_requirements(lines)
    if not reqs:
        raise RuntimeError("专业要求解析为空")
    excel_path = resolve_file_path(excel_path)
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
        if progress_step and progress_step > 0 and i % progress_step == 0:
            print_progress(i, total)
    df["_match"] = [x["match"] for x in matches]
    df["_matched_name"] = [x["matched_name"] for x in matches]
    df["_matched_code"] = [x["matched_code"] for x in matches]
    df["_score"] = [x["score"] for x in matches]
    out_df = df[df["_match"] == True].copy()
    if out_path is None:
        base = os.path.splitext(os.path.basename(excel_path))[0]
        out_path = os.path.join(os.path.dirname(excel_path), f"{base}_filtered.xlsx")
    count = len(out_df)
    saved = write_output(out_df, out_path, append=append, dedup=dedup, dedup_key=dedup_key, col_major=col_major)
    print(f"筛选完成：{os.path.basename(excel_path)} 命中 {count} 条 → {saved}")
    return saved

def dedup_dataframe(df, col_major: str, key: Optional[str]):
    """
    对结果DataFrame去重：
    - 若指定去重列key且存在，则按该列去重
    - 否则按规范化Major+匹配编码生成唯一键去重
    """
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

def process_many(excel_paths: List[str], require_path: str, col_major: str, threshold: float, out_dir: Optional[str], sheet: Optional[str], progress_step: int, limit: Optional[int], dedup: bool, dedup_key: Optional[str], merge_out: Optional[str], append: bool = False) -> List[str]:
    """
    批量处理多个Excel：
    - 逐文件调用process生成各自的筛选结果（可按out_dir落盘）
    - 读取各结果并（可选）去重后合并
    - 将合并结果写出到merge_out（默认目录下merged_filtered.xlsx）
    返回所有输出路径列表（包含合并结果）。
    """
    outputs = []
    merged = []
    total_count = 0
    require_path = resolve_file_path(require_path)
    for pth in excel_paths:
        base = os.path.splitext(os.path.basename(pth))[0]
        if out_dir:
            out_path = os.path.join(out_dir, f"{base}_filtered.xlsx")
        else:
            out_path = None
        saved = process(pth, require_path, col_major, threshold, out_path, sheet, progress_step, limit, append=append, dedup=dedup, dedup_key=dedup_key)
        outputs.append(saved)
        try:
            import pandas as pd
            if saved.lower().endswith(".xlsx"):
                part = pd.read_excel(saved)
            else:
                part = pd.read_csv(saved)
            if dedup:
                part = dedup_dataframe(part, col_major, dedup_key)
            merged.append(part)
            total_count += len(part)
        except Exception:
            continue
    if merged:
        import pandas as pd
        all_df = pd.concat(merged, ignore_index=True)
        if dedup:
            all_df = dedup_dataframe(all_df, col_major, dedup_key)
        if merge_out is None:
            first_dir = os.path.dirname(excel_paths[0]) if excel_paths else os.getcwd()
            merge_out = os.path.join(first_dir, "merged_filtered.xlsx")
        saved = write_output(all_df, merge_out, append=append, dedup=dedup, dedup_key=dedup_key, col_major=col_major)
        outputs.append(saved)
        print(f"总计筛选 {total_count} 条；合并后共 {len(all_df)} 条 → {saved}")
    else:
        print(f"总计筛选 {total_count} 条")
    return outputs

def build_arg_parser():
    """
    构建命令行参数：
    - 支持多文件输入、合并导出、去重、进度步长、limit等
    - 也包含默认值，便于“开箱即用”
    """
    p = argparse.ArgumentParser(
        description="根据require.txt对Excel的专业列进行模糊匹配筛选，支持多文件、合并导出、进度显示与去重。"
    )
    p.add_argument("--excel", nargs="+", default=DEFAULT_EXCELS, help="输入Excel文件路径，支持多个文件")
    p.add_argument("--require", default=DEFAULT_REQUIRE, help="专业要求文本文件路径（默认require.txt）")
    p.add_argument("--major-col", default=DEFAULT_MAJOR_COL, help="Excel中的专业列列名（默认Major）")
    p.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD, help="匹配阈值(0~1)，越大越严格（默认0.8）")
    p.add_argument("--out", default=None, help="单文件模式下的输出文件路径（不指定则按源文件生成*_filtered.xlsx）")
    p.add_argument("--out-dir", default=None, help="批量模式逐文件输出目录（未指定则各自默认路径）")
    p.add_argument("--merge-out", default=None, help="批量模式合并后的输出文件路径（默认merged_filtered.xlsx）")
    p.add_argument("--sheet", default=None, help="读取的工作表名称（不指定则读取默认/首个工作表）")
    p.add_argument("--progress-step", type=int, default=DEFAULT_PROGRESS_STEP, help="每处理N行打印一次进度（默认500）")
    p.add_argument("--limit", type=int, default=DEFAULT_LIMIT, help="仅读取前N行，用于调试或大文件处理（默认不限制）")
    p.add_argument("--dedup", action="store_true", help="开启去重功能（按指定列或规范化Major+编码）")
    p.add_argument("--dedup-key", default=None, help="指定去重字段名（如StudentID；不指定则用规范化Major+编码）")
    p.add_argument("--append", action="store_true", help="若输出文件已存在，选择追加模式（默认覆盖重写）")
    return p

def main(argv: List[str]) -> int:
    """
    命令行入口：根据输入文件数量选择单文件或批量处理流程。
    """
    args = build_arg_parser().parse_args(argv)
    try:
        if len(args.excel) == 1 and args.out_dir is None:
            process(args.excel[0], args.require, args.major_col, args.threshold, args.out, args.sheet, args.progress_step, args.limit, append=args.append, dedup=args.dedup, dedup_key=args.dedup_key)
        else:
            process_many(args.excel, args.require, args.major_col, args.threshold, args.out_dir, args.sheet, args.progress_step, args.limit, args.dedup, args.dedup_key, args.merge_out, append=args.append)
        print("处理完成")
        return 0
    except Exception as e:
        print(str(e))
        return 1

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
