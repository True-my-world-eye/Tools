**项目概述**
- 目标：从 Excel 的专业列中，依据 `require.txt` 的专业要求做模糊匹配与筛选，输出结果并支持多文件合并、追加与去重。
- 组成：
  - `major_filter.py`：命令行脚本，提供数据清洗与筛选的核心逻辑与批处理。
  - `major_filter_gui.py`：桌面 GUI 应用，提供可视化配置、进度与日志、合并输出，跨 Windows/Mac。
- 适用场景：原始 Excel 数据未清洗、专业名称存在格式差异或含编码后缀时的快速筛选与导出。

**主要特性**
- 规范化匹配：半角化、去空格与标点、统一小写，仅保留中文、字母、数字。
- 编码优先：提取 4~6 位数字编码（允许 T/K/TK 后缀），编码一致即判定命中。
- 相似度与子串：互为子串容错，剩余情况使用 `SequenceMatcher` 相似度评分。
- 批量处理：支持多 Excel 输入；逐文件导出外，提供合并导出并可选去重。
- 进度与统计：按步长输出进度条；处理完成输出命中条数与汇总。
- 输出模式：覆盖或追加；Excel 写出失败自动降级为 UTF-8-SIG CSV。
- GUI 操作：多文件选择、参数控件、进度条与日志、配置保存/清除、本地化提示。

**核心架构**
- 数据层（复用于 CLI 与 GUI）
  - 文本读取：多编码尝试（`utf-8`、`gbk`、`utf-8-sig`），忽略错误兜底。
  - 规范化：`to_halfwidth`、`normalize_text`。
  - 编码抽取：`extract_code`，将 `080914TK` 标准化为 `080914`。
  - 要求解析：`parse_requirements`，支持“专业名称（编码）”与“列式条目”两种格式；去重。
  - 匹配与评分：`best_match`，优先编码→全等→互为子串→相似度。
  - 去重：指定列或以“规范化专业+编码”生成唯一键。
  - 写出：优先 Excel，异常降级 CSV；支持追加并结合去重。
- CLI（`major_filter.py`）
  - 参数：`--excel`（多文件）、`--require`、`--major-col`、`--threshold`、`--progress-step`、`--limit`、`--sheet`、`--out`、`--out-dir`、`--merge-out`、`--append`、`--dedup`、`--dedup-key`
  - 路径解析：优先工作目录，找不到时回退脚本目录。
  - 输出规范：单文件打印“筛选完成：源文件 命中X条→路径”；批量打印总计与合并后条数。
- GUI（`major_filter_gui.py`）
  - 框架：`tkinter` + `ttk`，跨平台，无第三方 UI 依赖。
  - 线程：后台处理线程；`queue` + `after()` 更新进度与日志。
  - 布局：`LabelFrame` 分区（文件/参数/控制/进度与日志）；整页滚动容器，宽度自适应；弹窗模态。
  - 阈值控件：滑块 + 手动输入（支持 0~1 或百分比），双向同步。
  - 配置持久化：保存为 `major_filter_gui.json`，启动自动加载；支持一键清除并恢复默认。

**require.txt 格式要求**
- 1. 每行一个专业条目，允许空行与前后空白（自动忽略）。
- 2. 推荐写法：专业名称（编码），使用中文全角括号，如 `信息资源管理（120503）`。
- 3. 支持列式条目：如 `397 计算机类 080901 计算机科学与技术`。
- 4. 编码形态：4~6 位数字 + 可选字母后缀（如 T/K/TK），编码匹配优先、准确度更高。
- 5. 若不使用中文括号，将以最后一个词作为名称并抽取编码；建议使用中文括号或确保编码存在。
- 6. 名称建议使用正式中文，不混写多个专业于同一行；同义写法会做规范化，但规范格式更利于匹配。

**使用指南（CLI）**
- 基础运行（默认参数在脚本内已设置）：
  - `python major_filter.py`
- 指定参数示例：
  - 单文件处理并显示进度：`python major_filter.py --excel "TMT_FIGUREINFO.xlsx" --progress-step 500`
  - 只处理前 1000 行：`python major_filter.py --limit 1000`
  - 多文件并合并导出、去重：`python major_filter.py --excel "A.xlsx" "B.xlsx" --out-dir "." --merge-out "merged.xlsx" --dedup`
  - 追加写出（存在则合并后写出）：`python major_filter.py --out "result.xlsx" --append`
- 输出列说明：`_match`、`_matched_name`、`_matched_code`、`_score`

**使用指南（GUI）**
- 启动：`python major_filter_gui.py`
- 操作：
  - 文件区：添加/移除 Excel；选择 `require.txt`
  - 参数区：专业列、Sheet、阈值（滑块与输入框）、进度步长、`limit`、输出目录、合并输出文件、去重键、追加/去重开关
  - 控制区：开始处理、保存配置、清除本地缓存、软件使用须知
  - 反馈区：进度条、日志滚动窗口
- 配置文件：`major_filter_gui.json`（与程序同目录）

**实现要点**
- 规范化策略：
  - 半角化（全角标点转半角）、去空格、统一小写、剔除非中英数字符
  - 兼容“信息资源管理（120503）”“信息资源管理 120503”等变体
- 编码优先与评分规则：
  - 编码一致：`score = 1.0`
  - 规范化文本全等：`score = 0.95`
  - 互为子串：`score = 0.9`
  - 其他情况：`SequenceMatcher` 相似度
- 追加与合并：
  - 追加写出时先读旧文件，与新结果拼接；去重后再写出
  - 多文件合并时统一去重并写出到 `merge_out`
- 大文件优化：
  - `--sheet` 指定工作表、`--limit` 逐步验证、`--progress-step` 控制输出频率

**跨平台打包**
- 依赖：`pyinstaller`、`pandas`、`openpyxl`
- Windows 打包：
  - 环境：`python -m venv .venv && .venv\Scripts\activate && pip install pyinstaller pandas openpyxl`
  - 指令（隐藏控制台、单文件）：`pyinstaller -F -w major_filter_gui.py`
  - 打包资源：`pyinstaller -F -w --add-data "require.txt;." major_filter_gui.py`
  - 产物：`dist/major_filter_gui.exe`
- macOS 打包：
  - 环境：`python3 -m venv .venv && source .venv/bin/activate && pip install pyinstaller pandas openpyxl`
  - 指令（图形应用、单文件）：`pyinstaller --windowed --onefile major_filter_gui.py`
  - 打包资源：`pyinstaller --windowed --onefile --add-data "require.txt:." major_filter_gui.py`
  - 产物：`dist/MajorFilterGUI.app`
- 注意事项：
  - 必须在目标平台上打包（Windows 生成 exe；Mac 生成 app）
  - 资源路径分隔符：Windows 用 `;`，Mac/Linux 用 `:`
  - GUI隐藏控制台：Windows 用 `-w`，Mac 用 `--windowed`
  - macOS 签名与公证（推荐）：
    - 签名：`codesign --deep --force --verify --verbose --sign "Developer ID Application: 名称 (TEAMID)" dist/MajorFilterGUI.app`
    - 公证：`xcrun notarytool submit dist/MajorFilterGUI.app --keychain-profile "ProfileName" --wait`

**目录结构**

- `require.txt`：专业要求列表
- `major_filter.py`：CLI 核心
- `major_filter_gui.py`：GUI 核心
- `major_filter_gui.json`：GUI 配置文件（运行后生成）
- `*.xlsx/*.csv`：筛选输出文件

**开发与扩展**
- 阈值说明：
  - ≤0.70：宽松；0.75~0.85：均衡；≥0.90：严格
  - GUI 支持输入 0~1 或百分比（如 `80%`）
- 配置位置优化：
  - 当前保存于程序同目录；打包后可改为用户目录（Windows `%APPDATA%`，Mac `~/Library/Application Support`）以避免不可写路径
- 资源路径：
  - 如需更稳健，可在打包版本中检测 `sys._MEIPASS` 获取临时解包目录
- 性能：
  - 大文件场景建议使用 `--limit` 验证、再批量运行；合并与去重使用 `pandas`，注意内存占用

**常见问题**
- 未安装依赖：GUI 开始处理会提示安装命令 `pip install pandas openpyxl`
- Excel 写出失败：自动降级为 CSV；编码为 `utf-8-sig`
- 滚动与弹窗：
  - 主页面为整页滚动容器，鼠标滚轮在任意区域都可滚动
  - 使用须知弹窗为模态，滚轮事件不影响主页面

**作者信息**
- 作者：True my world eye
- Wechat：Truemwe
- E-mail：hwzhang0722@163.com
