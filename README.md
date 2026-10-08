# Markdown2Word

面向 Windows 的端到端 Markdown 技术报告转换工具。选择 Markdown、可选 Word 模板和输出位置，即可生成最终 DOCX，支持可编辑公式、表格格式和图表题注。

## 安装与启动

需要 Python 3.10+ 和 Microsoft Edge WebView2 Runtime。新版桌面界面使用 pywebview + 本地 HTML/CSS/JavaScript，不再依赖 Tkinter。应用运行不需要 Node.js，也不需要启动前端开发服务器。在实际使用的环境中安装：

```powershell
conda activate pack_cpu
python -m pip install -e ".[pandoc]"
python run_gui.py
```

`[pandoc]` 安装包含 Pandoc 可执行文件的 `pypandoc-binary`。已有系统 Pandoc 时可以执行 `python -m pip install -e .`。单独安装 `pypandoc` 不保证包含 Pandoc 程序。

工具使用启动程序的 Python，自动从 PATH、常见安装位置或 pypandoc 查找 Pandoc。高级设置允许指定路径并检测环境。

Windows 使用 Edge Chromium 渲染器，不降级到旧版 IE。缺少 WebView2 时需要安装 [Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)。界面资源随 Python 包一起安装，无 CDN 和远程界面请求。

## 使用

1. 在“选择文档”卡片选择 Markdown、可选模板和输出 DOCX。
2. 选择公式模式及是否覆盖已有输出。
3. 点击“开始转换”，在结果卡片查看状态。完成后直接打开文档或所在文件夹；错误和警告在结果区显示。

界面使用浅色圆角卡片和蓝色主操作，支持深色主题。宽窗口为设置 / 结果双栏，小窗口自动切换单栏，底部操作栏固定可见。运行日志可展开并拖动调整高度，查看旧日志时不会被新输出强行滚回底部。工具路径、超时和图表设置收在高级设置中。旧版偏好设置自动兼容读取。

转换前会检查 `$$` 行间公式。对于首尾分隔符独占一行的明确公式块，自动忽略内部空白行，日志会提示原始行号；此兼容同时适用于 MathType 和 Word 原生公式，不修改源 Markdown。代码块、行内代码及转义美元符号保持原样。分隔符未配对、空公式或仍未被 Pandoc 识别的公式会报告错误，避免静默输出残缺文本。

公式有两种模式：

- **使用 MathType 公式**：需要桌面版 Microsoft Word、已激活的 MathType 6 和可用的 Word 加载项。支持给行间公式添加 MathType 原生右编号。
- **关闭 MathType**：生成可编辑的 Word 原生公式（OMML），无需启动 Office。当前不提供此模式的右编号。

MathType 编号格式支持 `(1)`、`(1-1)`、`(1.1)`；后两种为整篇单章格式，不会按标题自动划章。后续编辑公式后可使用 MathType“更新公式编号”。转换时不要操作 Word 或 MathType，Office 和工具应使用相同权限。

高级设置还包括保留模板表格格式、自动图表题注、等待超时（每个后台操作默认 1800 秒）及实验性的 Mermaid → Visio 对象转换。Visio 模式需要桌面版 Visio，目前支持常见 flowchart/graph 节点和连接，复杂语法可能失败。不启用 Visio 时保留 Mermaid 源代码，不生成无内容的图形占位符。

图片自动嵌入 DOCX，支持相对 Markdown 的路径、绝对路径、URL 编码路径和 `file://` 文件路径，包括中文及空格文件名；过宽图片保持比例缩小到模板可打印页宽。图片文件须实际存在。缺失图片或 Pandoc 降级转换会明确报告警告；请补齐资源后重新转换。

MathType 转换前自动删除数学表达式中的 `\tag{...}` 和 `\tag*{...}`（支持嵌套花括号），右编号由 MathType 编号选项统一生成，原来的 tag 内容不作为编号保留。正文和代码中的同名文字不改动。此清理只应用于 MathType 模式。

“正文参考文献引用上标”默认关闭，开启后将正文中的 `[1]`、`[1, 2]`、`[1-3]` 等数字引用转换为上标，保留引用链接。代码和公式不参与处理；参考文献章节及以 `[数字]` 开头的文献条目保持原样。这是文本规则，不校验引用是否对应真实文献，其他正文数字方括号也可能匹配。

三反引号或波浪线围栏代码块保留为可编辑文本，保留缩进、换行和 Pandoc 支持的语法高亮，使用等宽字体、浅色底纹和单倍行距，长代码块可跨页。不会把代码中的美元符号当公式或把 `[1]` 当引用，也不自动执行代码、转换为截图或表格。Mermaid 仅在启用 Visio 时走图形转换。

## 输出与日志

输出位置只生成最终 DOCX。临时文档、图形源文件和内部清单存放在系统临时目录，成功、失败或取消后自动清理。界面和命令行不再提供中间阶段选择或续跑入口。

完整日志保存在 `%LOCALAPPDATA%\Markdown2Word\logs\latest.log`，每次转换覆盖上一份，可在日志面板打开。“复制”通过系统剪贴板复制完整日志；文件不可用时复制当前缓存。日志及错误详情可直接选中文字后按 Ctrl+C，追加日志不会重建已有文字节点。页面显示内容和桥接层日志缓存均有大小限制，完整内容以日志文件为准。窗口通过原生标题栏拖动。历史版本已经生成的任务目录不会自动删除。

输入和模板保持只读。默认禁止覆盖成品；允许覆盖时也会等待新结果通过公式、正文和结构检查后再替换。失败或取消不替换已有输出。取消会先请求后台清理，10 秒未响应再终止后台进程树。外部 Office 若有模态窗口，需要关闭后再重试。

## 无 GUI 运行

```powershell
# 原生 Word 公式，无需 Office
python -m markdown2word.core.cli report.md report.docx --no-formulas

# MathType 和右编号
python -m markdown2word.core.cli report.md report.docx --force
```

其他选项包括 `--template`、`--pandoc`、`--no-numbers`、`--number-format`、`--reference-superscript`、`--preserve-table-style`、`--no-auto-captions`、`--visio`、`--no-auto-word` 和 `--timeout`。`--no-formulas` 为兼容命名，含义是不用 MathType，公式仍会保留为可编辑的原生公式。

## 开发与验证

```powershell
python -m unittest discover -s tests -v
python -m compileall -q src tests

# 仅开发测试需要 Node.js；不打开 GUI 或浏览器
npm ci
npm run check
npm test
```

回归测试使用真实 Pandoc，覆盖数学 AST、代码/金额识别、原生公式、表格、题注、缺图警告、文件保护、临时文件清理、取消和超时。Windows 上还解析 PowerShell 脚本并测试纯 Mermaid 解析及正文指纹逻辑。测试不启动 GUI 或 Office，真实 MathType/Visio 自动化需要在相应桌面环境验证。

桌面接口测试覆盖设置迁移、并发任务拒绝、退出清理和通过桥接层的真实 Pandoc 转换。前端使用 jsdom 验证路径联动、选项提交、运行中禁用控件及错误/日志的纯文本显示。HTML/CSS/JS 位于 `src/markdown2word/gui/assets`，无需编译即可随包加载。

详见 [架构说明](docs/architecture.md)、[内部转换契约](docs/pipeline-contract.md)、[GUI 操作](docs/gui-usage.md)。
