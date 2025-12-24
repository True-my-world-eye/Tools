**项目概述**
- 目标：在 Excel 的专业列及其他属性列上进行筛选，输出结果并支持多文件合并、追加与去重。
- 组成：
  - `major_filter_gui.py`：桌面 GUI 应用，提供可视化配置、进度与日志、合并输出，跨 Windows/Mac。
- 适用场景：原始 Excel 数据未清洗、专业名称存在格式差异或含编码后缀时的筛选与导出。

**主要特性**
- 规范化匹配：半角化、去空格与标点、统一小写，仅保留中文、字母、数字。
- 编码优先：提取 4~6 位数字编码（允许 T/K/TK 后缀），编码一致即判定命中。
- 相似度与子串：互为子串容错，剩余情况使用 `SequenceMatcher` 相似度评分。
- 批量处理：支持多 Excel 输入；逐文件导出外，提供合并导出并可选去重。
- 进度与统计：按步长输出进度条；处理完成输出命中条数与汇总。
- 输出模式：覆盖或追加；Excel 写出失败自动降级为 UTF-8-SIG CSV。
- GUI 操作：多文件选择、参数控件、进度条与日志、配置保存/清除、本地化提示。

**核心架构（GUI）**
- 框架：`tkinter` + `ttk`，跨平台，无第三方 UI 依赖。
- 线程：后台处理线程；`queue` + `after()` 更新进度与日志。
- 布局：`LabelFrame` 分区（文件/参数/控制/进度与日志）；整页滚动容器，宽度自适应；弹窗模态。
- 阈值控件：滑块 + 手动输入（支持 0~1 或百分比），双向同步。
- 配置持久化：保存为 `major_filter_gui.json`，启动自动加载；支持一键清除并恢复默认。
 - 模式切换：顶部标签页在“多条件筛选”与“专业列筛选（旧版）”之间切换，互斥显示。
 - 稳定性优化：主线程消费进度与日志；运行时禁用开始按钮、支持取消运行；日志长度控制；多表合并读取。

**环境与依赖**
- Python 版本：建议 3.9+（含 `tkinter` 标准库）
- 依赖文件：`requirements.txt`
- 依赖列表：
  - `pandas`：数据读取与处理
  - `openpyxl`：Excel 读写支持
  - `pyinstaller`：可选，用于打包为可执行文件
  - `rapidfuzz`：可选，用于加速模糊匹配
  - `pyahocorasick`：可选，用于多关键词高效匹配
- 安装方式：
  - Windows：
    - `python -m venv .venv`
    - `.venv\Scripts\activate`
    - `pip install -r requirements.txt`
  - macOS：
    - `python3 -m venv .venv`
    - `source .venv/bin/activate`
    - `pip install -r requirements.txt`

**多条件筛选与条件文件（CSV）规范**
- 统一用一个可读的CSV文件描述筛选条件；表头固定为：`column,type,operator,value,threshold,priority,weight,options`
- 字段解释（简明）：  
  - `column`：列名；`type`：数据类别（text/number/regex/fuzzy/enum/boolean/code）；`operator`：匹配方式；`value`：匹配值；`threshold`：模糊阈值；`priority/weight`：加权模式使用；`options`：`ignore_case/normalize/code_prefer`等
- 示例如下（同样适用于 `require.txt` 与 `resume_require.txt`）：  
```
column,type,operator,value,threshold,priority,weight,options
Major,fuzzy,similar,信息资源管理,0.85,3,1.0,normalize=true;code_prefer=true
Major,code,equals,120503,,2,1.0,
School,text,contains,大学,,1,1.0,ignore_case=true
GPA,number,between,3.5-4.0,,2,1.0,
Skills,fuzzy,similar,Python,0.8,2,1.0,normalize=true
Email,regex,match,^\S+@\S+\.\S+$,,1,1.0,
Degree,enum,in,本科;硕士;博士,,1,1.0,
```
- 组合模式（运行时选择）：`AND`（都命中）、`OR`（任意命中）、`WEIGHTED`（加权总分达阈值命中）
- 审计输出：包含每条条件的命中与分数，以及整体命中 `_match_all` 与总分 `_score_all`（加权）

**字段说明与取值范围**
- `column`（必填）
  - 含义：Excel中的列名；多文件时建议使用列并集中的通用列
  - 缺列行为：该条件记为不命中（`hit=0`，`score=0`），不会中断处理
- `type`（必填）
  - 取值：`text|number|regex|fuzzy|enum|boolean|code`
- `operator`（必填，随`type`变化）
  - `text`：`equals`（完全相等）、`contains`（包含子串）、`startswith`（以指定内容开头）、`endswith`（以指定内容结尾）
  - `number`：`between`（数值在“最小-最大”范围内）、`min`（大于等于最小值）、`max`（小于等于最大值）、`equals`（数值等于指定值）
  - `regex`：`match`（满足给定正则表达式）
  - `fuzzy`：`similar`（按相似度匹配，配合`threshold`判定）
  - `enum`：`in`（值属于给定集合）
  - `boolean`：`is`（布尔值匹配 true/false）
  - `code`：`equals`（编码数字部分完全一致）
- `value`（必填，按类型填写）
  - `text`：任意字符串（可配合 `options.ignore_case=true` 忽略大小写；`options.normalize=true` 进行规范化）
  - `number-between`：用“`最小-最大`”表示范围（如`3.5-4.0`）；`min/max/equals` 填单个数字（如`3.5`）
  - `regex-match`：正则表达式（如`^\\S+@\\S+\\.\\S+$`，匹配邮箱）
  - `fuzzy-similar`：目标字符串（如`信息资源管理`、`Python`），系统按相似度计算
  - `enum-in`：枚举值集合，使用分号分隔（如`本科;硕士;博士`）
  - `code-equals`：建议填写纯数字编码（4~6位，如`120503`）；系统会自动抽取数字部分比对
- `threshold`（选填，仅`fuzzy`使用）
  - 取值：`0~1` 或百分比（如`0.80`或`80%`），留空视为`0`
  - 参考：`≤0.70`宽松，`0.75~0.85`均衡，`≥0.90`严格
- `priority`（选填）
  - 效果：仅用于人为排序或展示，不参与命中计算（当前逻辑）
- `weight`（选填，用于`WEIGHTED`）
  - 取值：浮点数，默认`1.0`，可设`0`（相当于不计分）
- `options`（选填，键值对；多个以`;`分隔）
  - `ignore_case=true|false`：文本匹配是否忽略大小写（默认 false）
  - `normalize=true|false`：文本是否规范化（默认 false）。规范化包括：半角化、去空白及标点、统一小写
  - `code_prefer=true|false`：在模糊匹配时是否优先使用“编码完全一致”判定为命中（默认 false）
  - 书写示例：`ignore_case=true;normalize=true;code_prefer=true`

**组合模式与阈值**
- `AND`：所有条件命中 → `_match_all=true`
- `OR`：任意条件命中 → `_match_all=true`
- `WEIGHTED`：`sum(score_i * weight_i) ≥ combine_threshold` → `_match_all=true`
  - `combine_threshold` 在GUI/CLI设置，支持`0~1`或百分比；默认`0.8`
  - 输出总分 `_score_all` 便于回溯与调参

**额外说明**
- 留空规则：除 `column/type/operator/value` 外的字段均可留空；`threshold`留空视为`0`（模糊最宽松）；`weight`留空视为`1.0`
- 缺列容错：若条件所指列在当前Excel不存在，该条件记为不命中，但不会中断整体筛选
- 正则容错：正则表达式无效时，该条件记为不命中；建议先在小样本上验证

**使用指南（GUI）**
- 启动：`python major_filter_gui.py`
- 操作：
  - 文件区：添加/移除 Excel；选择 `require.txt`（仅旧版标签页）
  - 条件区：导入/新增/删除/导出条件CSV；组合模式（AND/OR/WEIGHTED）与总阈值（加权）
  - Sheet 多表支持：留空读首个；填写`Sheet1,Sheet2`合并指定多个；填写`*`合并所有工作表
  - 参数区：专业列、Sheet、阈值（滑块与输入框）、进度步长、`limit`、输出目录、合并输出文件
  - 处理选项：去重键、追加模式、开启去重、写出审计列（可选，减少内存占用）
  - 输出设置：勾选“仅合并输出（不写逐文件）”时，单文件结果不会写出，仅生成合并文件
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
- 未安装依赖：GUI 开始处理会提示安装命令 `pip install -r requirements.txt`
- Excel 写出失败：自动降级为 CSV；编码为 `utf-8-sig`
- 滚动与弹窗：
  - 主页面为整页滚动容器，鼠标滚轮在任意区域都可滚动
  - 使用须知弹窗为模态，滚轮事件不影响主页面
- 运行卡顿：开启线程运行、禁用开始按钮、支持取消运行；适当调小进度步长以减少刷新；关闭“写出审计列”降低内存
- 运行日志：每处理“进度步长”行，在底部日志区域输出一次进度（与CLI一致的样式），例如：`[##########--------------------] 33% 已处理 165000/500000 行 | 已运行 00:07:12 | 命中 7213 行`

**作者信息**
- 作者：True my world eye
- Wechat：Truemwe
- E-mail：hwzhang0722@163.com
