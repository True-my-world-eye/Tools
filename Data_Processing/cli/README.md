**概述**
- 这是一个跨平台的命令行批量筛选工具，替代原GUI，面向百万级数据行与≤500条条件
- 支持Excel/CSV输入、多文件处理、Sheet多表合并、AND/OR/WEIGHTED组合、去重与追加
- 终端输出进度与日志，无图形界面依赖，Windows/Mac均可直接运行

**环境与依赖**
- Python 版本：建议 3.9+
- 必需：`pandas`、`openpyxl`
- 可选：`rapidfuzz`（提升模糊匹配性能）
- 安装（Windows）
  - `python -m venv .venv`
  - `.venv\Scripts\activate`
  - `pip install -r ../requirements.txt`
- 安装（macOS）
  - `python3 -m venv .venv`
  - `source .venv/bin/activate`
  - `pip install -r ../requirements.txt`

**使用**
- 编辑 `cli/filter_cli.py` 顶部“配置区域”，按需设置：
  - `EXCEL_FILES`：输入文件列表（支持多文件）
  - `SHEET`：`""`（首个工作表）/`"Sheet1,Sheet2"`（多个）/`"*"`（全部工作表）
  - `CONDITIONS_CSV`：条件文件路径（为空时回退旧版Major逻辑）
  - 组合与阈值：`COMBINE_MODE`（AND/OR/WEIGHTED）、`COMBINE_THRESHOLD`（0~1）
  - 输出与去重：`OUT_DIR`、`MERGE_OUT`、`APPEND`、`DEDUP`、`DEDUP_KEY`
  - 性能与日志：`CHUNK_SIZE`（建议5万~10万）、`PROGRESS_STEP`、`WRITE_AUDIT_COLUMNS`
- 运行：
  - `python cli/filter_cli.py`

**配置项详解（filter_cli.py 顶部）**
- `EXCEL_FILES`：输入文件列表
  - 填写多个完整路径，支持 `.xlsx/.xls/.csv`
  - 文件不存在会被跳过并提示
- `SHEET`：工作表选择与合并
  - `""`：每个 Excel 文件读取首个工作表
  - `"Sheet1,Sheet2"`：指定多个工作表并纵向合并
  - `"*"`：合并该文件所有工作表
  - 缺失工作表仅提示并跳过，不影响其他表处理
- `CONDITIONS_CSV`：条件文件路径
  - 填写 CSV/Excel 条件文件路径；为空将回退到“仅Major列”的旧逻辑（不建议）
  - 条件文件规范见下文“条件文件规范（CSV）”
- `MAJOR_COL`：仅旧逻辑使用的专业列名
  - 当 `CONDITIONS_CSV` 为空时作为筛选列
- `MAJOR_THRESHOLD`：旧逻辑的模糊阈值
  - 0~1 或百分比；旧逻辑不建议在CLI使用，推荐提供条件CSV
- `COMBINE_MODE`：条件组合模式
  - `AND`：所有条件命中才判为命中
  - `OR`：任意条件命中即判为命中
  - `WEIGHTED`：按 `score*weight` 求和，与 `COMBINE_THRESHOLD` 比较判定命中
- `COMBINE_THRESHOLD`：加权模式总阈值
  - 0~1；例如 0.8 表示加权总分达 0.8 即命中
- `OUT_DIR`：逐文件输出目录
  - `None`：按源文件所在目录写出
  - 非空：统一写到该目录
- `MERGE_OUT`：合并输出文件路径
  - `None`：在首个输入文件所在目录写出 `merged_filtered.xlsx`
  - 非空：按指定路径写出
- `APPEND`：追加模式
  - `True`：若目标文件存在，先读旧结果，与新结果合并后写出
  - `False`：覆盖写出，不合并旧结果
- `DEDUP`：是否去重
  - `True`：启用去重；按 `DEDUP_KEY` 或回退键进行唯一化
  - `False`：不做去重
- `DEDUP_KEY`：去重键列名
  - 填写且该列存在时，按该列值去重（推荐使用唯一ID，如学号/员工号）
  - 未填写或列不存在时，回退为“规范化Major+编码”组合键（旧逻辑兼容）
- `CHUNK_SIZE`：分块行数
  - 推荐 50,000~100,000；越大内存占用越高，但IO次数更少
  - 对CSV使用 `read_csv(chunksize)`；对Excel使用 `openpyxl` 流式读取
- `PROGRESS_STEP`：进度输出步长
  - 每处理该行数输出一次当前文件进度、总计行数、处理速率
  - 设置为与 `CHUNK_SIZE` 相近或其整数倍能获得较稳定的进度输出
- `WRITE_AUDIT_COLUMNS`：是否写出审计列
  - `True`：在输出中包含每条条件的命中与分数以及条件描述
  - `False`：仅写出总命中与总分，输出更轻量

**Sheet合并**
- `""`：每个文件读取首个工作表
- `"Sheet1,Sheet2"`：指定多个工作表，纵向合并后处理
- `"*"`：合并该文件所有工作表
- 缺失工作表只提示并跳过，不会中断处理

**条件文件规范（CSV）**
- 表头固定：`column,type,operator,value,threshold,priority,weight,options`
- 字段含义与取值：
  - `column`（必填）：Excel列名；缺列时该条件不命中但不中断
  - `type`（必填）：`text|number|regex|fuzzy|enum|boolean|code`
  - `operator`（必填，随type变化）：
    - `text`：`equals|contains|startswith|endswith`
    - `number`：`between|min|max|equals`
    - `regex`：`match`
    - `fuzzy`：`similar`
    - `enum`：`in`
    - `boolean`：`is`
    - `code`：`equals`
  - `value`（必填）：按类型填写（例如`min-max`、正则表达式、枚举值分号分隔等）
  - `threshold`（选填，仅`fuzzy`）：`0~1`或百分比；参考≤0.70宽松、0.75~0.85均衡、≥0.90严格
  - `priority`（选填）：展示排序，不参与命中计算
  - `weight`（选填，用于WEIGHTED）：默认`1.0`
  - `options`（选填；以`;`分隔）：`ignore_case`、`normalize`、`code_prefer`等

**组合模式**
- AND：所有条件命中 → `_match_all=true`
- OR：任意条件命中 → `_match_all=true`
- WEIGHTED：`sum(score_i * weight_i) ≥ COMBINE_THRESHOLD` → `_match_all=true`，总分写入 `_score_all`

**输出与去重**
- 逐文件输出：默认写`<源文件名>_filtered.xlsx`到`OUT_DIR`；写失败自动降级CSV
- 合并输出：按`MERGE_OUT`写出全量合并结果（在终端显示绝对路径）
- 去重：
  - 指定 `DEDUP_KEY` 且列存在，则按该列去重
  - 未指定时，回退“规范化Major+编码”组合键（旧版兼容）
- 追加：若文件存在且 `APPEND=true`，将新结果与旧结果合并后写出

**性能建议（百万行）**
- 使用 `CHUNK_SIZE` 分块处理，避免一次性读入整个文件
- 条件≤500条时，文本包含类已做合并与向量化；合理设置 `ignore_case/normalize`
- 模糊匹配昂贵：建议优先使用 `code_prefer=true` 精确编码命中；安装 `rapidfuzz` 可显著提速
- 合并写出优先CSV（Excel在大数据量下较慢）

**运行日志**
- 每处理`PROGRESS_STEP`行输出一次进度（去除速度、显示已运行时间、占比与已命中数量）
  - 样式：`[##########--------------------] 33% 已处理 165000/500000 行 | 已运行 00:07:12 | 命中 7213 行`
- 对异常条件与缺失列进行警告提示，不中断整体处理

**常见问题**
- 依赖未安装：脚本会提示 `pip install pandas openpyxl`
- Excel格式异常：工作表为空或标题行缺失将被跳过
- 条件错误：不合法正则等条件会被忽略并提示

**示例**
- 将条件写入 `conditions.csv` 后：
  - 编辑 `EXCEL_FILES` 与其他配置
  - 运行：`python new/filter_cli.py`


**作者信息**
- 作者：True my world eye
- Wechat：Truemwe
- E-mail：hwzhang0722@163.com