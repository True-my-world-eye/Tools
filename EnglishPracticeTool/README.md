# 英语口语练习（增强版）

## 概览
本项目是一个基于 `tkinter/ttk` 构建的桌面应用，支持多题库英语口语练习、倒计时显示、中英模式切换、题库浏览、题库导入/导出与文本复制。内置《大英2/3/4》题库，同时支持从 Markdown/JSON 文件导入并持久化保存。

## 技术栈
- UI框架：`tkinter` 标准库 + `ttk` 主题组件
- 数据解析：自定义 Markdown 合并中文行解析、标准 JSON 结构读取
- 持久化：用户数据目录 `%APPDATA%/EnglishPracticeApp/datasets` 下写入 `d2/d3/d4.json` 与用户导入的 `user_*.json`，并使用 `registry.json` 记录清单以便重启后保留
- 打包：`PyInstaller -F -w` 生成单文件 Windows 可执行程序

## 核心模块
- `翻译练习_增强版.py`：应用主界面与交互逻辑
  - 题库选择、显示模式、计时设置、开始练习、查看全部、复制、说明入口
  - 关键位置：
    - 加载题库来源与默认选择：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:31-41`
    - 工具栏与按钮：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:53-78`
    - 练习流程（倒计时/展示）：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:110-168`
    - 题库浏览：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:186-197`
    - 导入/导出与持久化：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:199-275`
    - 软件说明：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:305-327`

- `dataset_loader.py`：数据加载与写出
  - `load_md(path)`：解析 Markdown，将连续中文行合并，并与紧随其后的英文行配对
  - `load_json(path)`：读取标准结构 `[{"zh": "...", "en": "..."}]`
  - `export_json(path, data)`：UTF-8/缩进2写出
  - `export_md(path, data)`：导出Markdown（中文行+英文行，条目间空行）
  - `load_builtin(name)`：从内置模块读取 `d2/d3/d4`
  - `get_app_data_dir()/get_datasets_dir()`：用户数据与 `datasets` 目录定位
  - `export_bootstrap_jsons()`：首次运行将内置题库写出为 `d2/d3/d4.json`
  - 关键位置：`f:\AAAAclass\python\翻译练习\dataset_loader.py:4-66`

- `builtin_datasets.py`：内置题库
  - 常量：`D2_DATA/D3_DATA/D4_DATA`，代码内置，无需依赖外部文件
  - `get_builtin(name)`：按名称返回对应题库
  - 关键位置：`f:\AAAAclass\python\翻译练习\builtin_datasets.py:1-157`

## 显示模式与练习逻辑
- 中文：先显示中文，倒计时结束后显示中英对照
- 英文：先显示英文，倒计时结束后显示中英对照
- 中英：直接显示中英对照，不启用倒计时
- 抽样数量：当前为3句（可扩展为界面配置）

## 题库导入与持久化
- 支持格式：Markdown(.md)、JSON(.json)
- Markdown 解析规则：
  - 连续中文行自动合并为一个中文段，随后第一行英文作为对应翻译
  - 空行与编号可忽略，保留原有标点与空格
- JSON 结构：数组形式，每项含 `zh`/`en` 字段
- 导入保存：解析后自动保存为 `user_名称.json` 到 `%APPDATA%/EnglishPracticeApp/datasets`，并登记至 `registry.json`
- 重名处理：若存在同名，自动追加序号（如 `user_名称_1.json`）
- 移除导入题库：当前版本不提供界面移除；可手动删除 `user_*.json` 并更新或删除 `registry.json`

## 内置题库策略
- 内置《大英2/3/4》在代码层面始终可用；首次运行会写出 `d2/d3/d4.json` 供分享与导出
- 默认加载走内置数据，避免重复读取 JSON 版本

## 菜单与按钮布局
- 顶部菜单栏“文件”：导入题库、导出当前题库（支持JSON/Markdown）、打开存储目录。
- 工具栏仅保留三项：查看全部、结束练习、开始练习（由右向左排列）。

## 打包与运行
- 打包要求：`Python 3.12+`、`PyInstaller 6+`
- 安装打包工具：
  - `python -m pip install pyinstaller`
- 生成可执行文件：
  - 在项目根目录执行：
    - `pyinstaller -F -w -n EnglishPracticeApp "f:\AAAAclass\python\翻译练习\翻译练习_增强版.py"`
  - 生成位置：`dist/EnglishPracticeApp.exe`
- 首次运行：自动在 `%APPDATA%/EnglishPracticeApp/datasets` 写出 `d2/d3/d4.json`、创建 `registry.json`

## 自定义程序图标
- 准备图标文件：将 `assets/app.ico` 放置到 `翻译练习` 目录下（建议包含 16/32/48/64/128/256 尺寸）。
- 打包时指定图标：
  - `pyinstaller -F -w -n EnglishPracticeApp --icon "f:\AAAAclass\python\翻译练习\assets\app.ico" "f:\AAAAclass\python\翻译练习\翻译练习_增强版.py"`
- 窗口图标：程序已在启动时尝试加载 `assets/app.ico` 作为窗口图标（支持打包后的 `_MEIPASS` 路径）。
- 若使用 PNG：可替换为 `iconphoto`，并在打包时添加资源：
  - `pyinstaller -F -w --add-data "f:\AAAAclass\python\翻译练习\assets\app.png;assets" ...`

## 新增与变更（重要）
- 菜单与按钮布局：
  - 顶部“文件”菜单包含：导入题库、导出当前题库（支持JSON/Markdown）、打开存储目录。
  - 工具栏仅保留三项：查看全部、结束练习、开始练习（由右向左排列）。
- 导入/导出：
  - 导入支持 `.md/.json`；导出支持 `.md/.json`，Markdown为“中文一行 + 英文一行 + 空行”。
  - 存储目录：`%APPDATA%/EnglishPracticeApp/datasets`（无APPDATA时为用户主目录下 `.EnglishPracticeApp/datasets`）。
  - 可通过“文件→打开存储目录”直接进入。
- 图标支持：
  - 运行时窗口图标加载顺序：`assets/app.ico` → `assets/软件图标.png`（在打包后同样生效）。
  - 新增脚本：`f:\AAAAclass\python\翻译练习\make_ico.py`，用于将PNG转换为多尺寸ICO。

## 快速打包（使用PNG图标）
1. 将你的PNG图标放在：`f:\AAAAclass\python\翻译练习\软件图标.png`
2. 生成ICO（多尺寸）：
   - 运行：`python f:\AAAAclass\python\翻译练习\make_ico.py`
   - 输出：`f:\AAAAclass\python\翻译练习\assets\app.ico`
3. 打包命令（嵌入EXE图标并打包PNG资源）：
   - `pyinstaller -F -w -n EnglishPracticeApp --icon "f:\AAAAclass\python\翻译练习\assets\app.ico" --add-data "f:\AAAAclass\python\翻译练习\软件图标.png;assets" "f:\AAAAclass\python\翻译练习\翻译练习_增强版.py"`
4. 若需要清理后重打包：
   - `pyinstaller --clean -F -w -n EnglishPracticeApp --icon "f:\AAAAclass\python\翻译练习\assets\app.ico" --add-data "f:\AAAAclass\python\翻译练习\软件图标.png;assets" "f:\AAAAclass\python\翻译练习\翻译练习_增强版.py"`
5. 注意：Windows资源管理器可能缓存旧图标。若EXE图标未更新，可尝试重命名EXE、移动目录或刷新图标缓存（重启资源管理器或系统）。

## 关键代码位置（便于查阅）
- 窗口图标加载：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:24`（`iconbitmap`/`iconphoto`，兼容打包后的 `_MEIPASS`）
- 导出Markdown实现：`f:\AAAAclass\python\翻译练习\dataset_loader.py:47`
- 打开存储目录：`f:\AAAAclass\python\翻译练习\翻译练习_增强版.py:327`

## 重要注意事项
- 写入路径：所有持久化文件写入用户数据目录，兼容打包后的写权限限制
- 字体与样式：使用 `ttk` 主题并对工具栏按钮增加 padding，避免文字遮挡（导入按钮已加宽）
- 说明入口：工具栏“说明”按钮与右下角“ℹ️ 软件说明”均可打开说明对话框

## 导出说明
- 导出支持 JSON(.json) 与 Markdown(.md) 两种格式。
- Markdown导出格式：每条中文一行、对应英文一行，条目之间空一行。

## 常见问题
- “导入后丢失”：导入后的题库保存到 `%APPDATA%/EnglishPracticeApp/datasets` 并登记到 `registry.json`，重启仍在下拉可选
- “md格式不识别”：请确保英文行紧随中文段之后且为英文字符为主；空行与编号会被忽略
- “写入权限”问题：若无法写入，请检查用户目录权限或以管理员身份运行

## 版权与作者
本软件免费开放传播。
作者：True my world eye
微信号：Truemwe
邮箱：hwzhang0722@163.com
