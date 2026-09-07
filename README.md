# Markdown2Word GUI

Markdown2Word GUI 是一个面向 Windows 的独立桌面工具，用于把 Markdown 技术报告转换为套用指定 Word 模板的 DOCX，并继续完成 MathType 公式、公式编号和可选 Visio 流程图处理。

本仓库不再按 Codex skill 组织。普通用户只通过图形界面选择文件和转换选项，不需要安装 skill，也不需要手工执行任何转换脚本。

## 主要功能

- 选择 Markdown、Word 模板和输出文件，无需输入命令行路径。
- 保留 Markdown 一至四级标题并映射 Word 模板样式。
- 生成三线表、图注和表注，并使用 Word Caption/SEQ 域编号。
- 识别 Markdown 中常见的 LaTeX 定界符；进入 Word/MathType 阶段时统一为行内 `$...$` 和行间 `\[...\]`，再转换为可编辑的 MathType `Equation.DSMT4` 对象。
- 分别处理行内公式和行间公式，并自动应对表格公式触发的 MathType 行内/行间选择框。
- 可为行间公式添加 MathType 原生右编号：`(1)`、`(1-1)` 或 `(1.1)`。
- 提取 Mermaid 源码，并可选生成可编辑的 Visio 图形后嵌入 Word。
- 显示实时日志，保留阶段文件，支持取消和失败排查。

## 运行环境

必需组件：

- Windows 10/11；
- Python 3.10 或更高版本，并包含 Tkinter/Tcl-Tk；
- Pandoc；
- 桌面版 Microsoft Word；
- MathType 6，已激活且 Word 加载项可用；
- Python 包 `python-docx` 和 `lxml`。

仅启用 Visio 转换时需要桌面版 Microsoft Visio。WPS 不能替代 Microsoft Word，因为公式转换依赖 Word COM 和 MathType Word 加载项。

## 快速开始

安装 Python 依赖：

```powershell
python -m pip install -e .
```

启动 GUI：

```powershell
python run_gui.py
```

启动后：

1. 选择 Markdown 文件。
2. 选择 Word 模板；留空时使用程序自带模板。
3. 选择输出 DOCX。
4. 根据需要启用 MathType 转换、公式编号和 Visio 转换。
5. 点击“开始转换”，完成前不要操作正在自动化的 Word 或 MathType 窗口。

## 转换阶段

完整任务的主文档转换分为两个阶段：

1. Markdown 经 Pandoc 转换为 DOCX，并套用模板和文档格式。
2. 在 Word 中从文档末尾向前逐个转换 LaTeX；启用编号时，每个行间公式转换完成后立即插入 MathType 原生编号，最后统一更新编号。

Visio 是主文档转换后的可选处理：把支持的 Mermaid 流程图转换为可编辑 VSDX，并替换 Word 中对应的占位符。

程序不会修改 Markdown 原文或 Word 模板。默认保留下列中间结果：

- `*_stage1_latex.docx`：已套模板、公式仍为 LaTeX 的文档；
- 行间公式在转换为 MathType 对象后立即插入 MathType 原生编号，不再生成独立的 `*_mathtype_unnumbered.docx` 中间文件；
- `*_visio_sources/`：Mermaid 源文件、清单及生成的 VSDX。

这些文件用于排查问题和从失败阶段恢复。

## 公式转换与编号逻辑

### 公式转换

- Stage 1 输出的 DOCX 暂时保留 LaTeX 文本：行内公式为 `$...$`，独立行间公式为 `\[...\]`。
- Stage 2 使用 Word 中 MathType 6 加载项的 Toggle TeX，把每个精确文本范围转换成 `Equation.DSMT4` OLE 对象，不会调用 Word 原生 OMML 公式。
- 公式按文档坐标从后向前处理，避免前面的替换使尚未处理的范围失效。
- 表格中的行内公式如果触发 “Convert To Display Style?”，GUI 自动选择 `Create Inline Style Equation`。
- 只有独占一个段落的行间公式参与编号；正文和表格中的行内公式不会编号。

### MathType 原生编号

启用“添加公式编号”后，不再先生成一个完整的未编号 DOCX 再遍历编号。每个 `\[...\]` 转换成 MathType 对象后，程序立即执行 MathType 的直接调用宏 `MTCommand_InsertEqnNum`，在同一 `MTDisplayEquation` 段落右侧插入编号。全部公式完成后执行 `MTCommand_UpdateEqns` 更新整个文档。

编号不是普通文本，也不是旧脚本使用的 `SEQ Equation`。它由 MathType 管理的 Word 域组成：外层是 `MACROBUTTON MTPlaceRef`，内部使用 `SEQ MTEqn`，因此可以继续使用 MathType 选项卡的格式化、更新、浏览和公式引用功能。

GUI 编号格式与 MathType 格式串的对应关系如下：

| GUI 选项 | 显示示例 | MathType 格式 | 初始化分隔符 |
| --- | --- | --- | --- |
| 连续编号 `Simple` | `(1)` | `(#E1)` | 文档开头插入隐藏的“下一节”分隔符 |
| 章-公式 `ChapterDash` | `(1-1)` | `(#C1-#E1)` | 文档开头插入隐藏的“下一章”分隔符 |
| 章.公式 `ChapterDot` | `(1.1)` | `(#C1.#E1)` | 文档开头插入隐藏的“下一章”分隔符 |

这些分隔符通过 `MTCommand_InsertNextSection` 或 `MTCommand_InsertNextChapter` 创建，使用 MathType 的 `MTEquationSection`/`MTEditEquationSection` 隐藏域保存编号状态。`Simple` 虽然不显示节号，仍使用隐藏节分隔符初始化和重置公式序号。不要把这些隐藏域当作乱码或多余正文删除。

### 后续手工新增或修改公式

- 修改已有公式：双击 MathType 公式对象编辑即可；公式编号域与公式对象相互独立，通常不需要重新插入编号。
- 在现有编号序列中新增公式：优先使用 Word 的 MathType 选项卡插入“右编号行间公式”；也可以先插入行间公式，再把光标放在同一公式段落右侧并使用“插入编号”。
- 新增、删除或移动编号公式后：使用 MathType 选项卡的“更新公式编号”，不要只按 Word 的普通“更新域”。
- 修改 `(1)`、`(1-1)`、`(1.1)` 等格式：打开 MathType 的“格式化公式编号”，选择应用到整个文档，然后再执行“更新公式编号”。
- 开始新章或新节：在第一个新公式之前使用 MathType 的“插入下一章分隔符”或“插入下一节分隔符”。如果要指定非连续的章/节值，使用 MathType 的“插入分隔符…”或“修改分隔符…”。
- 查看分隔符：可在 Word 中开启“显示/隐藏编辑标记”。分隔符属于 MathType 编号基础设施，即使不打印也应保留。
- 不要手工创建 `SEQ Equation`、普通括号数字或删除 `MTPlaceRef` 域，否则 MathType 无法统一格式化和更新这些编号。

## 仓库结构

```text
Markdown2WordByTemplate/
├─ run_gui.py                       # 源码环境启动入口
├─ pyproject.toml                   # Python 项目和依赖配置
├─ src/markdown2word/
│  ├─ __main__.py                   # python -m markdown2word
│  ├─ gui/                          # Tkinter 图形界面
│  ├─ core/                         # 参数校验、任务编排、日志和取消
│  ├─ win32/                        # MathType 弹窗等 Windows 辅助逻辑
│  └─ backends/
│     ├─ stage1/                    # Markdown、Pandoc、DOCX 格式化
│     ├─ mathtype/                  # Word/MathType 转换与编号
│     ├─ pipeline/                  # 当前兼容后端编排
│     └─ visio/                     # Mermaid/Visio 后处理
├─ tests/                           # 自动化测试
└─ docs/                            # 架构、操作和阶段契约文档
```

`src/markdown2word/backends` 中暂时保留了 PowerShell 兼容实现，因为现有 Word/MathType 自动化逻辑已经过实际文档验证。它们属于应用内部实现，用户不需要调用。后续会逐阶段迁移为 Python COM 模块，而 GUI 和配置接口保持不变。

## 开发与测试

运行测试：

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m unittest discover -s tests -v
```

主要测试范围包括参数校验、命令编排、错误分类和 MathType 弹窗匹配。涉及 Word、MathType 和 Visio 的端到端测试必须在安装对应桌面软件的 Windows 交互式会话中执行。

## 常见问题

### Word 或 MathType 无法连接

- 确认使用桌面版 Microsoft Word，而不是 WPS。
- 确认 Word 与 GUI 以相同权限级别运行，通常都应使用普通权限。
- 在 Word 中手动执行一次 MathType Toggle TeX，确认加载项与许可证正常。
- 关闭正在占用输入或输出 DOCX 的 Word/WPS 窗口。

### 表格公式弹出 “Convert To Display Style?”

GUI 会识别该窗口并选择 `Create Inline Style Equation`。日志应显示“已检测到 MathType 行内/行间选择框”和“已自动选择”。如果只能检测而不能点击，重点检查 Word 与 GUI 是否处于不同权限级别。

### 转换失败但只有笼统提示

查看日志中最后一个明确的阶段和错误。阶段文件不会自动删除：Stage 1 文件用于检查 Markdown/Pandoc；最终 DOCX 在核心公式与编号审计通过后保存检查点。如果日志明确显示检查点已经保存，文档可用于检查，但重新自动转换时仍建议从 Markdown 或 Stage 1 DOCX 开始，不要把部分转换、部分编号的 DOCX 当作新的输入。

### Mermaid 无法转换

当前优先支持 `flowchart`/`graph` 的常用节点和 `-->` 连接。复杂子图、HTML 标签和高级样式应先简化，原始 `.mmd` 文件会保留供人工编辑。

## 后续重构方向

- 将流水线编排迁移为独立 Python Worker，并使用结构化事件向 GUI 返回进度。
- 将 Word/MathType 转换和编号逐步迁移到 `pywin32`。
- 将 DOCX 自定义属性与编号格式处理迁移为纯 Python XML 模块。
- 将 Visio COM 操作迁移为 Python。
- 增加 PyInstaller 构建，使最终用户可以直接运行 `Markdown2Word.exe`。

详细边界见 [架构说明](docs/architecture.md) 和 [阶段契约](docs/pipeline-contract.md)。
