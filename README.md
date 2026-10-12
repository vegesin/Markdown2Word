# Markdown2Word

**将包含公式、图片、表格和代码的 Markdown 转换为可编辑的 Word 文档。**

Markdown2Word 是一款面向 Windows 的桌面转换工具，适用于技术报告、研究文档及其他需要以 DOCX 交付的 Markdown 内容。选择源文件、可选 Word 模板和保存位置，即可完成转换，得到最终 Word 文档。

支持 Word 原生公式和 MathType 公式两种模式。原生模式无需启动 Office；MathType 模式可生成可编辑的 MathType 公式，并为行间公式添加右编号。

[主要功能](#主要功能) · [环境要求](#环境要求) · [快速开始](#快速开始) · [使用说明](#使用说明) · [Markdown 编写示例](#markdown-编写示例) · [问题排查](#问题排查) · [命令行使用](#命令行使用)

## 主要功能

| 功能 | 转换效果 |
| --- | --- |
| Markdown 转 DOCX | 转换标题、正文、列表、表格等内容，支持选择 Word 模板 |
| 从要求创建模板 | 导入 TXT、Markdown、DOCX 格式要求，生成 JSON 规范、干净的 Word 模板和样张 |
| AI 辅助规范解析 | GUI 内配置外部 API 并对话，内置提示词约束标准输出；支持人工编辑及离线规则解析 |
| 可编辑公式 | 支持行内、行间公式，可选择 Word 原生公式或 MathType 公式 |
| 公式编号 | MathType 模式支持行间公式右编号及多种编号格式 |
| 图片自动插入 | 根据 Markdown 图片路径嵌入图片，过宽图片按页面可用宽度缩小 |
| 引用与超链接 | 可选数字引用上标；超链接显示为蓝色并保留跳转 |
| 代码块 | 保留为可编辑文本，使用等宽字体、底纹和适合代码的行距 |
| 表格与题注 | 支持表格排版及自动图表题注，可选择保留模板表格格式 |
| 日志与文件保护 | 显示警告和错误，支持复制完整日志，覆盖前检查新结果 |
| Mermaid 转 Visio | 实验性支持将部分流程图转换为可编辑 Visio 对象 |

界面支持深浅主题和窗口大小自适应。转换主要在本机执行，界面资源随工具安装；Markdown 引用网络图片等资源时，转换过程可能访问相应网址。只有主动使用 AI 对话时，格式要求和会话内容会发送到用户配置的 API。

## 环境要求

当前使用和验证以 Windows 桌面环境为主。

| 组件 | 何时需要 | 说明 |
| --- | --- | --- |
| Python 3.10+ | 所有运行方式 | 建议使用独立虚拟环境 |
| Pandoc | 所有转换 | 推荐通过下方的 `[pandoc]` 安装方式一并安装 |
| Microsoft Edge WebView2 Runtime | 桌面界面 | 如未安装，使用 [Microsoft 官方安装入口](https://developer.microsoft.com/microsoft-edge/webview2/) |
| Microsoft Word + MathType 6 | MathType 模式 | 需要已激活的 MathType 和可用的 Word 加载项 |
| Microsoft Visio | Mermaid 转 Visio | 可选，功能仍为实验性 |

Word 原生公式模式无需启动 Word 或 MathType。运行桌面界面不需要 Node.js，也不需要启动前端开发服务器。

## 快速开始

### 安装与启动

下载或克隆仓库后，在**仓库根目录**打开 PowerShell。以下命令创建独立环境、安装依赖并启动工具，无需激活脚本：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[pandoc]"
.\.venv\Scripts\python.exe run_gui.py
```

`[pandoc]` 会安装包含 Pandoc 可执行文件的 `pypandoc-binary`。如果已有可用的系统 Pandoc，可以将安装命令改为：

```powershell
.\.venv\Scripts\python.exe -m pip install -e .
```

单独安装 `pypandoc` 不保证包含 Pandoc 程序。工具会自动查找 Pandoc，也可以在“高级设置”中指定其路径。

如果使用已有虚拟环境或 Conda 环境，先激活自己的环境，再运行：

```powershell
python -m pip install -e ".[pandoc]"
python run_gui.py
```

后续启动时使用同一环境。安装后也可在已激活的环境中执行 `markdown2word`。

### 完成第一次转换

- 新建 `report.md`，写入一段标题和正文，或使用下方的最小示例。
- 启动工具，在“选择文档”中选择这个 Markdown 文件。
- “Word 模板”暂时留空，使用内置模板。
- 检查自动填写的保存位置，首次验证可选择“Word 原生公式”。
- 点击“开始转换”，等待结果卡片显示完成。
- 点击“打开文档”检查成品；如果显示警告，先阅读警告和运行日志。

正常完成后，保存位置会出现一个 `.docx` 文件。若目标文件被占用而触发另存，请以结果卡片中的实际路径为准。

## 使用说明

### 源文件、模板与保存位置

选择 `.md` 或 `.markdown` 文件后，工具会自动填写同名 DOCX 路径，也可以手工指定其他位置。

“Word 模板”接受 `.docx/.dotx/.docm` 文件，留空且未选择格式协议时使用内置模板。模板作为文档样式参考，用于控制字体、标题等排版；模板中的封面或正文内容不会作为整份文档直接复制到输出中。复杂版式仍建议在转换后检查。

输入 Markdown 和模板不会被修改。“允许覆盖已有文档”默认关闭，需要替换已有输出时再开启。

### 根据格式要求创建模板

在 GUI 的“创建模板”中：

1. 导入格式要求文件或 JSON 协议。可以先用 `tests/fixtures/requirements.md` 或 `requirements.docx` 体验。
2. 使用本地解析，或配置 API 基础地址、模型和密钥后点击“发送给 AI”。内置提示词要求返回统一 JSON 草稿；可继续对话修正规则。
3. 查看和编辑草稿，点击“校验并查看完整规范”，解决 `issues` 中的冲突和不支持要求。可以先保存草稿，稍后再导入继续编辑。
4. 核对默认值，填写一个尚不存在的模板目录，生成模板与 `preview.docx` 样张。
5. 生成成功后，GUI 自动选中该项目的 `spec.json`，后续 Markdown 转换无需再调用 AI。

格式协议与外部 Word 模板二选一。使用协议时，表格、题注和公式编号由协议控制；目录域需在 Word 中更新实际页码。PDF/OCR、复杂分节和真正按章公式编号暂未实现。完整字段、接口配置、安全边界与样例见 [格式规范与模板生成](docs/format-spec.md)。

### 选择公式模式

| 对比项 | Word 原生公式 | MathType 公式 |
| --- | --- | --- |
| 输出形式 | Word 可编辑公式（OMML） | 可编辑 MathType 对象 |
| 转换时需要 Office | 无需启动 Office | 需要 Microsoft Word 和 MathType |
| 行间公式右编号 | 当前不提供 | 支持 MathType 原生右编号 |
| 适用场景 | 日常文档、快速转换 | 明确要求 MathType 公式及编号的文档 |

MathType 模式支持 `(1)`、`(1-1)`、`(1.1)` 等编号格式；后两种当前按整篇单章处理，不会根据标题自动划分章节。后续编辑公式后，可使用 MathType 的“更新公式编号”。

MathType 转换前会清理公式中的 `\tag{...}` 和 `\tag*{...}`，原 tag 内容不作为编号保留，编号由工具选项统一生成。此清理仅适用于 MathType 模式。

使用 MathType 时，请确保 Word 与工具以相同权限运行，并在转换期间暂停操作 Word 和 MathType。

### 引用、图片、链接与代码

- **引用上标**：开启“参考文献引用上标”后，处理正文中的 `[1]`、`[1, 2]`、`[1-3]` 等数字引用。参考文献列表、代码和公式不参与处理。这是文本规则，不会自动核对文献条目，其他数字方括号也可能被匹配。
- **图片**：支持相对路径、绝对路径、URL 编码路径和 `file://` 路径，包括中文及空格文件名。图片文件须实际存在；缺失图片会产生警告，应补齐后重新转换。
- **超链接**：转换后统一显示为蓝色，保留点击跳转和已有上标等格式。
- **代码块**：保留缩进、换行及 Pandoc 支持的语法高亮，允许跨页，不执行代码。

### 高级设置

“高级设置”提供自动图表题注、保留模板表格格式、Word 自动启动、等待超时及 Python/Pandoc 路径等选项。默认等待超时为每个后台操作 1800 秒。

“检测运行环境”会检查 Python 依赖、Pandoc 和 Office 注册信息，并将结果写入日志。环境检测不启动 Office；MathType 加载项能否实际执行，需要在转换时或手动测试中确认。

Mermaid 转 Visio 需要桌面版 Visio，当前支持部分常见 `flowchart` / `graph` 语法，复杂图形可能失败。不启用此选项时，Mermaid 内容保留为代码块。

详细交互说明见 [GUI 操作说明](docs/gui-usage.md)。

## Markdown 编写示例

### 最小文档

将下面内容保存为 UTF-8 编码的 `report.md`。这个示例无需额外图片文件即可转换：

````markdown
# 示例技术报告

这是正文，参考文献引用示例为[1]。

## 公式

行内公式：$E=mc^2$。

行间公式：

$$
y = ax + b
$$

## 表格

| 参数 | 数值 |
| --- | --- |
| a | 2 |
| b | 1 |

## 代码与链接

```python
def predict(x):
    return 2 * x + 1
```

相关文档：[Pandoc](https://pandoc.org/)。

## 参考文献

[1] 示例文献条目。
````

转换后，公式和代码均可编辑，链接可点击；开启引用上标选项后，正文中的 `[1]` 会显示为上标。

### 插入图片

例如将图片和 Markdown 按以下结构存放：

```text
report/
├── report.md
└── figures/
    └── architecture.png
```

在 `report.md` 中写入：

```markdown
![系统架构](figures/architecture.png)
```

相对路径从 **Markdown 所在目录** 解析。移动文档时，应同时保留对应的图片文件和目录结构。

### 公式书写规则

行内公式使用 `$...$`，行间公式推荐将首尾 `$$` 独占一行，公式内部不留空白行。

对于明确的独立 `$$` 公式块，工具会兼容内部空白行，并在日志中提示原始行号，不修改源 Markdown。未配对的分隔符、空公式或未被识别的公式会报错，应根据行号修正后再转换。

代码块、行内代码和转义美元符号保持原样。需要展示公式源码或美元符号示例时，可以使用代码块或行内代码包裹。

## 输出与文件保护

- 正常转换只向保存位置输出最终 DOCX，内部临时文件会在任务结束后清理。
- 开启覆盖后，也会等待新文档通过转换和检查，再替换旧文件。转换失败或取消时保留原有成品。
- 保存时遇到文件占用或拒绝访问，会短暂重试；仍无法覆盖时，在同目录另存带“转换结果”和唯一后缀的 DOCX。
- 输出目录无法写入时，会尝试保存到 `%LOCALAPPDATA%\Markdown2Word\logs\recovered`。
- 另存成功时显示警告，结果卡片、“打开文档”和“打开文件夹”均指向实际成品位置。

点击取消后，工具先请求后台清理；后台长时间未响应时，会终止本次转换的后台进程。若 Office 存在模态对话框，可能仍需手动处理。

## 问题排查

### 先查看哪里

- **结果卡片**：查看失败原因、警告和实际成品位置。
- **运行日志**：展开日志，点击“复制”复制完整日志；文件不可用时复制当前缓存。也可选择文字后按 `Ctrl+C`。
- **完整日志**：点击“完整日志”，打开本次任务的详细记录。
- **环境检测**：保存当前错误日志后，再到“高级设置”点击“检测运行环境”。

| 文件或目录 | 用途 |
| --- | --- |
| `%LOCALAPPDATA%\Markdown2Word\logs\latest.log` | 最新转换或环境检测日志 |
| `%LOCALAPPDATA%\Markdown2Word\logs\startup.log` | 桌面窗口启动失败时的日志 |
| `%LOCALAPPDATA%\Markdown2Word\logs\recovered` | 原输出目录无法写入时的成品恢复位置 |

`latest.log` 会被后续转换或环境检测覆盖。排查和反馈问题前，请先复制保存当前日志。

### 常见问题

| 现象 | 优先检查 | 处理方式 |
| --- | --- | --- |
| 窗口无法启动，提示 WebView2 | WebView2 是否安装，当前 Python 是否正确 | 安装或修复运行时；查看启动弹窗和 `startup.log` 中的具体原因 |
| 找不到 Pandoc | 是否存在 Pandoc 可执行程序 | 使用 `python -m pip install -e ".[pandoc]"` 安装，或在高级设置中指定路径 |
| 提示缺少 `docx`、`lxml` 等模块 | 安装依赖与启动工具是否使用同一环境 | 在启动工具所用的 Python 环境中重新安装项目依赖 |
| 图片缺失或只显示描述 | 图片文件是否存在，相对路径是否正确 | 补齐图片或修正 Markdown 中的路径，再重新转换 |
| 公式显示为 `$$` 或原始文本 | 分隔符和日志提示的行号 | 修正公式格式；确认内容没有被代码块包裹 |
| MathType 转换失败或卡住 | MathType 激活状态、Word 加载项和未关闭的弹窗 | 在 Word 中选中 `$x^2$`，手动测试 Toggle TeX；关闭阻塞弹窗，检查完整日志 |
| 转换等待超时 | Office 是否有弹窗，文档是否包含大量公式 | 先处理阻塞；确实需要更长时间时，在高级设置中增加超时 |
| 保存时出现“拒绝访问” | Word 是否打开目标文件，文件和目录是否允许写入 | 关闭目标文档后重试；若已另存，直接打开结果卡片所示的新成品 |
| 转换完成但找不到文件 | 是否触发自动另存或恢复保存 | 使用结果卡片中的“打开文档”或“打开文件夹” |
| 正文完整性或结构检查失败 | 是否在转换期间操作 Word，日志中的差异上下文 | 暂停 Office 操作后从原 Markdown 重试；仍失败时提供最小复现文档和日志 |

### 反馈问题时提供什么

- 完整错误信息或保存后的日志。
- 使用的公式模式，以及是否启用编号、引用上标或 Visio。
- Python、Pandoc 及相关 Office / MathType 版本。
- 能触发问题的最小 Markdown 片段；图片问题需说明对应文件路径。
- 实际结果和期望结果，必要时附截图。

分享前请移除文档内容、日志和路径中的敏感信息。可先用最小示例或只保留出错段落，判断问题是否与特定内容有关。

## 命令行使用

在安装项目的环境中执行。命令行与桌面界面共用转换逻辑，不启动 GUI。

```powershell
# Word 原生公式，无需启动 Office
python -m core.cli report.md report.docx --no-formulas

# MathType 公式及右编号，需要 Word 和 MathType
python -m core.cli report.md report.docx

# 使用模板，并处理正文数字引用上标
python -m core.cli report.md report.docx --template template.docx --reference-superscript

# 查看全部参数
python -m core.cli --help
```

如果使用前面的未激活虚拟环境，将 `python` 替换为 `.\.venv\Scripts\python.exe`。

| 参数 | 作用 |
| --- | --- |
| `--no-formulas` | 使用 Word 原生公式，不进行 MathType 转换；公式仍保留 |
| `--template PATH` | 指定 DOCX 样式模板 |
| `--format-spec PATH` | 使用标准 JSON 格式规范，与 `--template` 互斥 |
| `--pandoc PATH` | 指定 Pandoc 可执行程序 |
| `--no-numbers` | 关闭 MathType 行间公式右编号 |
| `--number-format Simple/ChapterDash/ChapterDot` | 选择编号格式，三个值任选其一 |
| `--reference-superscript` | 开启正文数字引用上标 |
| `--preserve-table-style` | 保留模板表格格式 |
| `--no-auto-captions` | 关闭自动图表题注 |
| `--visio` | 启用实验性的 Mermaid 转 Visio |
| `--no-auto-word` | 关闭 Word 自动启动，需要事先手动打开 Word |
| `--timeout SECONDS` | 设置每个后台操作的等待超时 |
| `--force` | 允许覆盖已有成品 |

## 开发与相关文档

桌面界面使用 pywebview 和本地 HTML/CSS/JavaScript，Python 后端负责文档转换与 Office 自动化。前端资源位于 `src/gui/assets`，无需构建即可随包加载。

源码直接放在 `src` 下，包含 `core`、`gui`、`backends`、`formatting`、`requirements`，不再嵌套 `markdown2word` 包。安装命令和 `python run_gui.py` 启动方式不变；旧的 `python -m markdown2word.core.cli` 改为 `python -m core.cli`。在未安装的源码环境运行模块时，先设置 `$env:PYTHONPATH = "$PWD/src"`。

```powershell
# 根据要求生成模板，输出文件/目录需要尚不存在
python -m core.template_cli parse tests/fixtures/requirements.md draft.json
python -m core.template_cli build draft.json report-template
python -m core.cli tests/fixtures/report.md report.docx --format-spec report-template/spec.json --no-formulas
```

基础检查可在开发环境中执行：

```powershell
python -m compileall -q src
python -m core.cli --help
python -m core.template_cli --help
python -m unittest discover -s tests

# 以下仅前端开发需要 Node.js
npm ci
npm run check
npm test
```

`tests/` 和 `tests/fixtures/` 随仓库提供，包括 Word 格式要求文件、冲突要求、中间协议及转换输入。安装 `[pandoc]` 后可运行真实原生公式转换回归；AI 测试使用模拟 HTTP 响应。真实 API、MathType / Visio 自动化需另行配置相应服务或 Windows 软件环境。

- [GUI 操作说明](docs/gui-usage.md)：界面设置、日志和任务操作。
- [架构说明](docs/architecture.md)：前后端职责及后台任务组织。
- [内部转换契约](docs/pipeline-contract.md)：公式、图形和文档检查的实现约定。
- [格式规范与模板生成](docs/format-spec.md)：协议、AI 对话、模板项目及版本边界。
