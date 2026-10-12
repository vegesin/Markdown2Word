# 格式规范与模板生成

从 0.5.0 开始，可以由自然语言格式要求生成规范，再生成干净的 Word 模板。转换时同时读取规范，避免图表、代码和公式的格式被后处理覆盖。

## 快速体验

安装项目后，在仓库根目录执行：

```powershell
python -m core.template_cli parse tests/fixtures/requirements.md draft.json
python -m core.template_cli validate draft.json
python -m core.template_cli build draft.json report-template
python -m core.cli tests/fixtures/report.md report.docx --format-spec report-template/spec.json --no-formulas
```

输出草稿和模板目录不能已存在。第二次实验请改名。也可以跳过自然语言解析，直接使用 `tests/fixtures/report-spec.json` 创建模板。GUI 中使用“创建模板”，可导入要求、与 AI 对话、编辑 JSON、保存草稿、校验和生成样张。

## 三类数据

1. **格式要求**：TXT、Markdown、DOCX。DOCX 只读取 XML 中的段落及表格，不启动 Word，不执行宏，不跟随外部资源。单文件最多 10 MB、提取文字最多 80000 字符；超限必须拆分，不能静默截断。
2. **草稿**：`{"spec": {...}, "evidence": [...], "issues": [...]}`。含来源和未解决问题，用于 AI 往返及人工审核。
3. **执行规范**：`spec.json`，只包含格式规则。`resolved-spec.json` 展开默认值及继承。转换命令的 `--format-spec` 接收执行规范，不接收草稿外壳。

Schema 位于 `src/formatting/schema.json`，使用 JSON Schema 2020-12。未知字段、错误类型、未知版本、非法数值及循环继承均拒绝执行。未指定字段继承 `m2w-default@1`；零和 false 是明确值，null 不代表默认。

## 最小协议

```json
{
  "schema_version": "1.0",
  "name": "技术报告",
  "styles": {
    "body": {
      "font": {"east_asia": "宋体", "latin": "Times New Roman", "size_pt": 12},
      "paragraph": {
        "first_line_indent": {"value": 2, "unit": "char"},
        "line_spacing": {"mode": "multiple", "value": 1.5}
      }
    }
  }
}
```

## 本版支持范围

| 字段 | 支持内容 |
| --- | --- |
| `page` | A4/Letter、横向/纵向、四边页边距 mm |
| `styles` | 正文、首段、1～9 级标题、图题、表题、表头、表正文、代码、引用、公式段落 |
| `font` | 中西文字体、字号 pt、粗体、斜体、六位 RGB 颜色 |
| `paragraph` | 对齐、字符/pt 首行缩进、单倍/多倍/固定/最小行距、段前后 pt、分页及段落保持 |
| `headings` | 无编号或十进制多级编号 1、1.1、1.1.1，可配置深度 |
| `table` | 三线表、网格、无边框、重复表头、单元格内边距 |
| `captions` | 独立图表题注样式、标签、启用开关；图下表上、全文连续编号 |
| `equations` | 含公式段落的最小行距；无编号或现有 MathType 编号格式 |
| `toc` | 插入指定层级的 Word 目录域；实际页码须在 Word 中更新 |
| `header_footer` | 简单页眉文字、页脚页码 |
| `hyperlink_color` | 链接颜色 |

角色名称与 Word 样式映射见 `src/formatting/spec.py` 中的 `STYLE_NAMES`。样式的 `based_on` 引用语义角色，例如 `body`，不是 Word 界面的本地化名称。`First Paragraph`、`Body Text`、`Compact` 的正文映射由编译器处理。字体字号按 Word 的半磅精度表达（0.5pt 的倍数）；公式样式控制所在段落，不设置 OMML/MathType 对象内部的数学字体。

公式段落的 `equations.inline_min_pt/display_min_pt` 优先于普通段落行距；这是明确的公式例外规则。MathType 后端要求公式居中且无首行缩进，编号使用制表位布局。`ChapterDash/ChapterDot` 仍是整篇单章格式，不能表示根据标题重启编号。原生公式暂不支持右编号；选择不兼容后端时会报错。

PDF/OCR、复杂分节、奇偶页眉、中文章编号、按章题注、原生公式右编号、独立脚注/参考文献/列表格式、交叉引用和复杂封面尚未实现。AI 应把这些要求写入 `issues`，不能转换为近似支持的字段。

## 解析与审核

本地规则支持示例文件中的明确表达，如“正文：宋体，小四，1.5倍行距”。每个未完整识别的段落都会列出，标题和说明文字也可能需要人工确认。不承诺理解任意自然语言。

`evidence` 每项包含 JSON Pointer `path`、原文位置 `source` 和原文片段 `quote`。AI 回答的引用必须能在输入块找到。`issues` 的类型为 `conflict`、`unresolved`、`unsupported`。存在问题时可以保存草稿，不能生成模板；编辑实际规范并解决问题后，再移除对应条目。缺省字段的实际值通过 GUI“校验并查看完整规范”展示。

自然语言理解仍需人工核对。通过 Schema 校验只能证明结构可执行，不能证明 AI 已完整理解原文。

## AI 对话接口

内置提示词：`src/requirements/prompt.md`。发送时注入同一份 Schema，要求 AI 仅返回完整草稿。适配器位于 `src/requirements/ai.py`，目前使用 Chat Completions 兼容的 `POST {base_url}/chat/completions`，不含工具调用或 Agent 执行能力。

GUI 填写 API 基础地址（通常以 `/v1` 结尾）、模型名，配置 Key 环境变量或本次会话密码框。地址需 HTTPS，本机 localhost 服务可 HTTP。服务不支持 JSON 模式时不要勾选。只有点击“发送给 AI”才发送要求文件的提取文字、当前草稿和会话历史；普通转换不会调用 AI。

Key 不写入 settings.json、不进入提示词和日志；本次会话输入留在密码框内，窗口关闭后清除。环境变量配置的值也不会返回前端。地址、模型、环境变量名称等非敏感配置可持久化。

CLI 使用配置文件：

```powershell
# 编辑 tests/fixtures/ai-config.example.json 的地址和模型，另存为 ai-config.json
$env:M2W_AI_API_KEY = "你的服务密钥"
python -m core.template_cli parse tests/fixtures/requirements.docx ai-draft.json --ai-config ai-config.json
```

网络错误、输出截断、非法 JSON、来源伪造均报错，不替换当前草稿。可修正配置后重新发送；本版没有自动网络重试或自动修复调用，避免不透明的额外费用。取消不强制中断正在进行的 HTTP 请求，会等待请求返回或超时后丢弃结果；GUI 请求超时为 60 秒。

## 模板项目和转换优先级

生成目录包含 `spec.json`、`resolved-spec.json`、`reference.docx`、`preview.docx`、`report.json`、`manifest.json`。只从受控基础文档建立样式，不复制用户参考材料中的宏、对象或样式。发布前检查模板包，整个目录生成成功才发布，不覆盖已有目录。

推荐转换时选择 `spec.json`：它同时控制模板与转换规则。单独选择 `reference.docx` 只提供样式参考，不能包含目录插入等行为。外部 Word 模板与格式规范互斥；未选择规范时维持旧流程。

选择规范后，规范控制题注、表格和公式编号，GUI 中对应旧开关不参与覆盖。公式后端、Visio、引用上标、输出位置等仍由转换选项控制。转换使用规范快照，结束前检查原文件哈希，Office 后处理后再检查最终文档的样式和布局。

规范和字体能被写入 DOCX，不保证目标电脑安装了对应字体；请在交付环境打开 `preview.docx` 检查。目录通过 Word“更新域/更新整个目录”计算实际页码。

## 测试资料

- `tests/fixtures/requirements.md`：本地解析可完整识别的要求。
- `tests/fixtures/requirements.docx`：包含说明段落和表格的真实 Word 要求文件。
- `tests/fixtures/conflicting-requirements.md`：相互矛盾及尚不支持的要求。
- `tests/fixtures/report-spec.json`：可直接执行的中间层协议。
- `tests/fixtures/report.md`：包含标题、表格、代码、引用和公式的转换输入。
- `tests/fixtures/ai-config.example.json`：无密钥的配置样例。

测试命令见 README。AI 测试使用模拟 HTTP 响应，不访问付费服务；真实服务联调需要自行填写地址、模型和凭据。
