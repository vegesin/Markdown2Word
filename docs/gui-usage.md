# GUI 操作说明

## 启动

开发环境可以运行根目录的 `run_gui.py`，安装项目后也可以执行 `markdown2word`。

GUI 使用 Tkinter。若启动时报 Tcl/Tk 错误，请改用带 Tkinter 的标准 Python 或 Anaconda Python。

## 文件区域

- **Markdown**：待转换的 `.md` 或 `.markdown` 文件。
- **Word 模板**：可选的 `.docx`、`.dotx` 或 `.docm`；留空时使用内置模板。
- **输出 DOCX**：最终文档路径。

程序不会修改 Markdown 和模板。启用 MathType 时会额外保留 Stage 1 和未编号文档。

## 转换选项

- **转换为可编辑 MathType 公式**：逐个调用 MathType Toggle TeX。
- **给行间公式添加右编号**：只编号独立行间 MathType 对象。
- **编号格式**：支持 `(1)`、`(1-1)` 和 `(1.1)`。
- **自动启动/连接 Microsoft Word**：找不到 Word 会话时尝试启动 Word。
- **允许覆盖已有输出**：仅覆盖明确选择的输出和阶段文件。
- **转换 Mermaid 为 Visio**：需要桌面版 Visio。
- **跳过正文完整性检查**：只用于恢复或诊断，不建议日常启用。

## 日志和取消

日志窗口显示当前阶段、公式数量、转换序号、检查结果和建议。点击取消会终止当前后端进程，但已经生成的阶段文件会保留。

如果 Word 正显示模态对话框，取消后还需要关闭该对话框，之后再开始新任务。

## MathType 注意事项

- Word 和 GUI 必须处于同一 Windows 桌面会话和相同权限级别。
- 转换期间不要点击 Word、MathType 或公式对象。
- MathType 许可证必须允许 Word 加载项执行 Toggle TeX。
- 表格内行内公式的选择弹窗由 GUI 守护器自动处理。
