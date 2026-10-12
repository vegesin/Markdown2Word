# 架构说明

源码平铺在 `src` 下，`core`、`gui`、`backends`、`formatting`、`requirements` 是独立 Python 包。桌面启动仍由 `run_gui.py` 或安装后的 `markdown2word` 命令完成。

`requirements` 负责只读提取格式要求、本地规则解析、内置提示词及 Chat Completions 兼容 API；AI 只产生带来源和问题清单的规范草稿，没有文件操作或工具执行能力。`formatting` 负责 JSON Schema 校验、继承与默认值展开、Word 模板生成和格式校验。`core.template_cli` 提供独立模板命令。

`gui/template_api.py` 与转换任务共用线程互斥和退出等待机制，模板任务不阻塞前端。AI 对话历史仅保存在当前窗口，Key 不进入持久配置；取消等待当前 HTTP 请求返回/超时并丢弃结果。模板项目保存规范、样张、来源报告和哈希，不是转换流水线的恢复接口。

使用 `format_spec` 时，核心保存完整规范快照，Stage1 由受控基础文档构建 Pandoc 参考模板，并在转换后应用动态规则。MathType 接收公式行距和段间距参数。最终发布前执行协议格式检查和输入哈希检查；不指定协议则使用兼容的旧转换行为。

`gui` 收集参数、显示结果和日志；`core.pipeline` 统一校验、编排后台进程、处理取消和超时以及发布成品；`core.cli` 提供无 GUI 入口。`backends/stage1` 是内部 AST/DOCX 构建实现，`backends/mathtype` 和 `backends/visio` 分别负责 Office 自动化。`office_session.ps1` 共享 Word 发现和连接逻辑。

`gui/app.py` 负责 pywebview 桌面窗口和前端资源加载；`gui/bridge.py` 提供文件对话框、配置迁移、任务管理、增量日志与退出清理接口；`gui/assets` 为原生 HTML/CSS/JavaScript 界面。转换核心不依赖 GUI，实现单独维护。

界面资源加载到本地窗口，不使用 CDN 或前端 HTTP API。JavaScript 通过 pywebview Promise 通信桥调用 Python，周期性拉取任务快照和增量日志。后台使用独立工作线程，桥接层以锁保护共享状态、有界缓存存储日志，不阻塞页面线程。关闭运行中的窗口时，先询问是否清理退出，再异步等待任务结束。运行中禁用转换设置并防止重复提交。

Windows 明确使用 Edge Chromium，不允许降级到 IE。窗口按照逻辑像素和屏幕可用区域确定大小，CSS 在窄窗口切换单栏，固定底部操作栏；日志可展开并调整高度。主题和偏好保存到应用设置目录，兼容旧 `auto_word` 和编号格式标签。页面文本、错误和日志使用 textContent，不将文档内容作为 HTML 执行。CSP 禁止外部连接；pywebview 动态创建通信包装函数，因此脚本策略保留其所需的 unsafe-eval。

每次转换创建系统临时工作区，处理 Markdown AST、模板、公式和图形。全部成功后将最终文档复制到输出目录内的临时文件，再原子发布；无覆盖模式不替换已有文件。输入/模板的哈希在发布前重新核对。工作区由上下文管理器清理，不维护持久任务清单或阶段复用接口。仅保留应用目录下的一份最新日志。

子进程 stdout 由独立线程读取，监控线程负责取消、超时和进程清理。`M2W_CANCEL_FILE` 位于临时目录，后端在公式/图形边界检查。宽限 10 秒后强制终止后端进程树。外部 Office 进程不属于普通子进程，强制中止后可能需要用户关闭模态窗口。

Word 发现支持 32/64 位注册表 App Paths、Click-to-Run、PATH 和常见 Office 安装目录。环境检查只检查程序路径和 MathType 注册信息，不实例化 Office。实际任务才连接 Word，必要时启动。

`M2W_EVENT <JSON>` 传递准备结果、公式进度、警告和错误。原始消息完整记录到日志，GUI 展示中文操作名称、失败原因和警告。删除了重复的 PowerShell 总入口、独立编号脚本、Lua 替换逻辑、独立审计命令和旧桌面弹窗守卫。必要审计集中在构建器和对应后端，MathType 弹窗只在具体公式转换期间按进程及窗口归属处理。
