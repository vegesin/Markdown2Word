from __future__ import annotations

import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable


PACKAGE_ROOT = Path(__file__).resolve().parents[1]
BACKEND_ROOT = PACKAGE_ROOT / "backends"
PIPELINE_SCRIPT = BACKEND_ROOT / "pipeline" / "convert_md_to_mathtype_docx.ps1"
STAGE1_BACKEND = BACKEND_ROOT / "stage1"
MATHTYPE_BACKEND = BACKEND_ROOT / "mathtype"
STAGE1_SCRIPT = STAGE1_BACKEND / "run_md_to_template_docx.ps1"
VISIO_SCRIPT = BACKEND_ROOT / "visio" / "convert_mermaid_to_visio.ps1"

NUMBER_FORMATS = {
    "(1), (2), ...": "Simple",
    "(1-1), (1-2), ...": "ChapterDash",
    "(1.1), (1.2), ...": "ChapterDot",
}


class ConfigurationError(ValueError):
    pass


@dataclass(frozen=True)
class ConversionConfig:
    markdown: Path
    output: Path
    template: Path | None = None
    python: Path | None = None
    pandoc: Path | None = None
    convert_formulas: bool = True
    add_numbers: bool = True
    number_format: str = "Simple"
    create_visio: bool = False
    auto_start_word: bool = True
    overwrite: bool = True
    skip_text_integrity_check: bool = False

    @property
    def stage1_path(self) -> Path:
        if not self.convert_formulas:
            return self.output
        return self.output.with_name(f"{self.output.stem}_stage1_latex.docx")

    @property
    def visio_directory(self) -> Path:
        return self.output.with_name(f"{self.output.stem}_visio_sources")


def find_powershell() -> str:
    return shutil.which("powershell") or "powershell.exe"


def find_python() -> Path:
    candidates = [
        Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies" / "python" / "python.exe",
        Path(sys.executable),
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    found = shutil.which("python") or shutil.which("py")
    return Path(found) if found else Path(sys.executable)


def find_pandoc() -> Path:
    found = shutil.which("pandoc")
    candidates = [
        Path(found) if found else None,
        Path(os.environ.get("LOCALAPPDATA", "")) / "Pandoc" / "pandoc.exe",
        Path(os.environ.get("ProgramFiles", "")) / "Pandoc" / "pandoc.exe",
    ]
    for candidate in candidates:
        if candidate and candidate.is_file():
            return candidate.resolve()
    return Path(candidates[1])


def normalize_output(path: Path) -> Path:
    return path if path.suffix.lower() == ".docx" else path.with_suffix(".docx")


def validate(config: ConversionConfig) -> ConversionConfig:
    if str(config.markdown).strip() in {"", "."}:
        raise ConfigurationError("请选择 Markdown 文件。")
    if str(config.output).strip() in {"", "."}:
        raise ConfigurationError("请选择输出 DOCX。")
    markdown = config.markdown.expanduser().resolve()
    output = normalize_output(config.output.expanduser()).resolve()
    template = config.template.expanduser().resolve() if config.template else None
    python = (config.python or find_python()).expanduser().resolve()
    pandoc = (config.pandoc or find_pandoc()).expanduser().resolve()

    if not markdown.is_file() or markdown.suffix.lower() not in {".md", ".markdown"}:
        raise ConfigurationError(f"Markdown 文件不存在或扩展名不正确：{markdown}")
    if template and (not template.is_file() or template.suffix.lower() not in {".docx", ".dotx", ".docm"}):
        raise ConfigurationError(f"Word 模板不存在或格式不支持：{template}")
    if not python.is_file():
        raise ConfigurationError(f"未找到 Python：{python}")
    if not pandoc.is_file():
        raise ConfigurationError(f"未找到 Pandoc：{pandoc}")
    if config.number_format not in NUMBER_FORMATS.values():
        raise ConfigurationError(f"不支持的公式编号格式：{config.number_format}")
    if output == markdown or (template and output == template):
        raise ConfigurationError("输出文件不能覆盖 Markdown 原文或 Word 模板。")
    if output.exists() and not config.overwrite:
        raise ConfigurationError(f"输出文件已存在，请勾选“允许覆盖”：{output}")
    if config.add_numbers and not config.convert_formulas:
        raise ConfigurationError("添加公式编号需要先启用 MathType 公式转换。")
    if config.create_visio and not VISIO_SCRIPT.is_file():
        raise ConfigurationError(f"缺少 Visio 后处理脚本：{VISIO_SCRIPT}")

    return ConversionConfig(
        markdown=markdown,
        output=output,
        template=template,
        python=python,
        pandoc=pandoc,
        convert_formulas=config.convert_formulas,
        add_numbers=config.add_numbers,
        number_format=config.number_format,
        create_visio=config.create_visio,
        auto_start_word=config.auto_start_word,
        overwrite=config.overwrite,
        skip_text_integrity_check=config.skip_text_integrity_check,
    )


def build_commands(config: ConversionConfig) -> list[list[str]]:
    config = validate(config)
    powershell = find_powershell()
    common = [powershell, "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File"]
    if config.convert_formulas:
        command = common + [
            str(PIPELINE_SCRIPT),
            "-MarkdownPath", str(config.markdown),
            "-OutputPath", str(config.output),
            "-Stage1Path", str(config.stage1_path),
            "-VisioDirectory", str(config.visio_directory),
            "-Stage1BackendDirectory", str(STAGE1_BACKEND),
            "-MathTypeBackendDirectory", str(MATHTYPE_BACKEND),
            "-PythonPath", str(config.python),
            "-PandocPath", str(config.pandoc),
        ]
        if config.template:
            command += ["-TemplatePath", str(config.template)]
        if config.add_numbers:
            command += ["-AddEquationNumbers", "-EquationNumberFormat", config.number_format]
        if config.auto_start_word:
            command.append("-StartWordIfNeeded")
        if config.skip_text_integrity_check:
            command.append("-SkipTextIntegrityCheck")
        if config.overwrite:
            command.append("-Force")
    else:
        command = common + [
            str(STAGE1_SCRIPT),
            "-MarkdownPath", str(config.markdown),
            "-OutputPath", str(config.output),
            "-BackendDirectory", str(STAGE1_BACKEND),
            "-PythonPath", str(config.python),
            "-PandocPath", str(config.pandoc),
            "-VisioDirectory", str(config.visio_directory),
        ]
        if config.template:
            command += ["-TemplatePath", str(config.template)]

    commands = [command]
    if config.create_visio:
        commands.append(common + [
            str(VISIO_SCRIPT),
            "-DocxPath", str(config.output),
            "-VisioDirectory", str(config.visio_directory),
            *( ["-StartWordIfNeeded"] if config.auto_start_word else [] ),
        ])
    return commands


def diagnostic_hints(output: str) -> list[str]:
    text = output.lower()
    if "disp_e_paramnotoptional" in text or "0x8002000f" in text:
        return [
            "调用到了需要 Ribbon 控件参数的 MathType OnAction 回调，或 Word.Run 参数没有正确封送；这不是公式语法或 MathType 激活问题。",
            "请确认正在运行使用 MTCommand_InsertEqnNum/MTCommand_InsertNextSection 等直接调用入口的新版本后端。",
            "保留 Stage-1 DOCX，并从 GUI 重新执行转换。",
        ]
    if "formula conversion failed" in text or "validate mathtype object count" in text:
        return [
            "MathType 已返回，但当前公式没有生成新的 Equation.DSMT4 对象；工具会等待对象创建并对未变化的公式安全重试一次。",
            "如果安全重试仍失败，请检查日志中的 Formula、Start 和 InTable，并从保留的未编号 DOCX 继续排查。",
            "确认转换期间没有操作 Word；Word 与 GUI 应使用相同权限级别。",
        ]

    hints: list[str] = []
    if "pandoc" in text and ("not found" in text or "cannot find" in text or "failed" in text):
        hints.append("检查 Pandoc 路径；也可能是 Markdown 语法、图片路径或编码导致 Pandoc 失败。")
    if "80070520" in text or "logon session" in text:
        hints.append("当前进程没有可用的交互式 Office 登录会话；请在正常桌面登录会话中启动 GUI。")
    if "word.application" in text or "microsoft word" in text:
        hints.append("确认桌面版 Microsoft Word 已安装激活，输出未被 Word/WPS 占用，且权限级别与 GUI 相同。")
    if "mathtype" in text or "mtcommand" in text or "equation.dsmt4" in text:
        hints.append("确认 MathType 6 已激活、Word 加载项已启用，并可手动执行一次 Toggle TeX。")
    if "template" in text or "style" in text:
        hints.append("检查模板是否可正常打开、目标样式是否存在；受保护视图或 VBA/宏安全策略也可能阻止自动化。")
    if "visio" in text:
        hints.append("确认桌面版 Visio 已安装；复杂 Mermaid 语法请先简化为 flowchart/graph 常用节点和 --> 连接。")
    if "access" in text or "permission" in text or "占用" in output:
        hints.append("关闭输入/输出 DOCX 和 WPS，并确认输出目录具有写权限。")
    return list(dict.fromkeys(hints))


class PipelineRunner:
    def __init__(self) -> None:
        self.process: subprocess.Popen[str] | None = None
        self.cancelled = False

    def run(self, config: ConversionConfig, log: Callable[[str], None]) -> int:
        self.cancelled = False
        commands = build_commands(config)
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        for index, command in enumerate(commands, 1):
            if self.cancelled:
                return 130
            log(f"\n--- 执行阶段 {index}/{len(commands)} ---\n")
            self.process = subprocess.Popen(
                command,
                cwd=str(config.markdown.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
                creationflags=creationflags,
            )
            assert self.process.stdout is not None
            captured: list[str] = []
            for line in self.process.stdout:
                captured.append(line)
                log(line)
            code = self.process.wait()
            self.process = None
            if code != 0:
                log(f"\n阶段失败，退出码：{code}\n")
                hints = diagnostic_hints("".join(captured))
                if hints:
                    log("\n可能原因与处理建议：\n")
                    for hint in hints:
                        log(f"  - {hint}\n")
                log("  - 已完成的阶段文件不会删除，可用于检查或从断点恢复。\n")
                return code
        return 0

    def cancel(self) -> None:
        self.cancelled = True
        process = self.process
        if process and process.poll() is None:
            if os.name == "nt":
                subprocess.run(
                    ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            else:
                process.terminate()
