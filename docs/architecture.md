# 架构说明

## 产品边界

Markdown2Word 是独立 Windows GUI 应用，不依赖 Codex skill 安装目录。界面、配置、转换后端和资源全部从 `src/markdown2word` 加载。

## 模块职责

- `gui`：收集参数、展示日志、处理用户取消，不包含文档转换规则。
- `core`：校验配置、生成阶段任务、启动后端和归类错误。
- `win32`：处理只与 Windows 桌面交互有关的功能。
- `backends/stage1`：Markdown、Pandoc、模板、标题、表格和题注。
- `backends/mathtype`：Word COM、MathType OLE、公式编号和审计。
- `backends/visio`：创建 VSDX 并嵌入 Word。

## 当前兼容层

Word/MathType 与 Visio 自动化目前仍由内部 PowerShell 后端执行。保留它们是为了避免在仓库改造时同时重写已经验证过的 COM 逻辑。GUI 只引用应用包内部路径，不再引用 skill 名称或用户的 `.codex/skills` 目录。

## 迁移约束

后续 Python 化必须保持以下行为：

- 原始 Markdown 和模板只读；
- 按阶段保存独立输出；
- 公式从文档末尾向前转换；
- 每次转换后检查 MathType 对象增量；
- 表格、段落、正文和公式总数接受审计；
- 只有行间 MathType 对象可以编号；
- GUI 进程不能因 Word 模态框永久失去响应。

推荐把 Office 自动化放入独立 Python Worker，而不是 Tkinter 主线程。Worker 通过结构化事件报告 `stage`、`progress`、`warning`、`error` 和 `artifact`。
