from __future__ import annotations

import os
import shutil
import subprocess
import sys
import json
import hashlib
import queue
import re
import threading
import time
import tempfile
from collections import deque
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable


PACKAGE_ROOT = Path(__file__).resolve().parents[1]
BACKEND_ROOT = PACKAGE_ROOT / "backends"
STAGE1_BACKEND = BACKEND_ROOT / "stage1"
MATHTYPE_BACKEND = BACKEND_ROOT / "mathtype"
VISIO_SCRIPT = BACKEND_ROOT / "visio" / "convert_mermaid_to_visio.ps1"

NUMBER_FORMATS = {
    "(1), (2), ...": "Simple",
    "(1-1), (1-2), ...（单章）": "ChapterDash",
    "(1.1), (1.2), ...（单章）": "ChapterDot",
}


class ConfigurationError(ValueError):
    pass


@dataclass(frozen=True)
class ConversionConfig:
    markdown: Path
    output: Path
    template: Path | None = None
    format_spec: Path | None = None
    python: Path | None = None
    pandoc: Path | None = None
    convert_formulas: bool = True
    add_numbers: bool = True
    number_format: str = "Simple"
    create_visio: bool = False
    auto_start_word: bool = True
    overwrite: bool = False
    preserve_table_style: bool = False
    auto_captions: bool = True
    reference_superscript: bool = False
    stage_timeout: int = 1800

    @property
    def intermediate_path(self) -> Path:
        if not self.convert_formulas:
            return self.output
        return self.output.with_name("prepared.docx")

    @property
    def visio_directory(self) -> Path:
        return self.output.with_name("diagrams")


def find_powershell() -> str:
    return shutil.which("powershell") or "powershell.exe"


def find_python() -> Path:
    return Path(sys.executable).resolve()


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
    try:
        import pypandoc
        candidate = Path(pypandoc.get_pandoc_path())
        if os.name == "nt" and not candidate.is_file():
            candidate = candidate.with_suffix(".exe")
        if candidate.is_file():
            return candidate.resolve()
    except (ImportError, OSError, RuntimeError):
        pass
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
    format_spec = config.format_spec.expanduser().resolve() if config.format_spec else None
    if format_spec:
        from formatting.spec import load_spec
        if template:
            raise ConfigurationError('格式协议与外部 Word 模板不能同时选择。')
        if not format_spec.is_file():
            raise ConfigurationError('格式协议不存在。')
        spec = load_spec(format_spec)
        if config.convert_formulas:
            equation = spec['styles']['equation']['paragraph']
            if equation['alignment'] != 'center' or equation['first_line_indent']['value'] != 0:
                raise ConfigurationError('MathType 后端目前要求公式居中且首行不缩进。')
        numbering = spec['equations']['numbering']
        if numbering != 'none' and not config.convert_formulas:
            raise ConfigurationError('此协议要求 MathType 编号，请选择 MathType 模式。')
        config = replace(config, add_numbers=numbering != 'none', number_format=numbering if numbering != 'none' else 'Simple')
    effective_template = template or (STAGE1_BACKEND / "assets" / "default-template.docx").resolve()
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
    if output in {markdown, effective_template, format_spec}:
        raise ConfigurationError("输出文件不能覆盖 Markdown 原文或 Word 模板。")
    if output.exists() and not output.is_file():
        raise ConfigurationError(f"输出路径不是文件：{output}")
    if not 30 <= config.stage_timeout <= 86400:
        raise ConfigurationError("等待超时应为 30～86400 秒。")
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
        format_spec=format_spec,
        python=python,
        pandoc=pandoc,
        convert_formulas=config.convert_formulas,
        add_numbers=config.add_numbers,
        number_format=config.number_format,
        create_visio=config.create_visio,
        auto_start_word=config.auto_start_word,
        overwrite=config.overwrite,
        preserve_table_style=config.preserve_table_style,
        auto_captions=config.auto_captions,
        reference_superscript=config.reference_superscript,
        stage_timeout=config.stage_timeout,
    )


def build_commands(config: ConversionConfig) -> list[list[str]]:
    powershell = find_powershell()
    common = [powershell, "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File"]
    command = [str(config.python), str(STAGE1_BACKEND / "build_stage1_docx.py"), str(config.markdown)]
    if config.template:
        command.append(str(config.template))
    command += [str(config.intermediate_path), "--pandoc", str(config.pandoc)]
    if config.format_spec:
        command += ['--format-spec', str(config.format_spec)]
    if not config.convert_formulas:
        command.append('--native-math')
    if config.preserve_table_style:
        command.append("--preserve-table-style")
    if not config.auto_captions:
        command.append("--no-auto-captions")
    if config.reference_superscript:
        command.append('--reference-superscript')
    if config.create_visio:
        command += ["--write-visio-sources", "--visio-dir", str(config.visio_directory)]
    commands = [command]
    if config.convert_formulas:
        command = common + [str(MATHTYPE_BACKEND / "convert_word_latex_to_mathtype.ps1"),
                            "-InputPath", str(config.intermediate_path), "-OutputPath", str(config.output)]
        if config.add_numbers:
            command += ["-AddEquationNumbers", "-EquationNumberFormat", config.number_format]
        if config.auto_start_word:
            command.append("-StartWordIfNeeded")
        if config.format_spec:
            from formatting.spec import load_spec
            spec = load_spec(config.format_spec)
            paragraph = spec['styles']['equation']['paragraph']
            command += ['-DisplayMinPt', str(spec['equations']['display_min_pt']),
                        '-DisplaySpaceBeforePt', str(paragraph['space_before_pt']),
                        '-DisplaySpaceAfterPt', str(paragraph['space_after_pt'])]
        commands.append(command)
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
            "请查看错误详情，修复 Word 加载项后重新转换。",
        ]
    if "formula conversion failed" in text or "validate mathtype object count" in text:
        return [
            "MathType 已返回，但当前公式没有生成新的 Equation.DSMT4 对象；工具会等待对象创建并对未变化的公式安全重试一次。",
            "如果安全重试仍失败，请查看日志中的具体公式和位置。",
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


def describe_failure(output: str) -> str:
    text = output.lower()
    if 'word desktop executable was not found' in text or 'not found in app paths' in text:
        return '未找到桌面版 Microsoft Word。请检查 Office 安装；或关闭“使用 MathType 公式”，生成不需要 Office 的 Word 原生公式。'
    if 'no running word instance' in text or 'desktop word is not available' in text:
        return '无法连接 Microsoft Word。请关闭启动提示或模态对话框，并确保 Word 与本工具使用相同权限。'
    reasons = re.findall(r'(?:Reason\s*:\s*|(?:RuntimeError|ValueError|FileNotFoundError|PermissionError):\s*)([^\r\n]+)', output)
    detail = reasons[-1].strip() if reasons else next((line.strip() for line in reversed(output.splitlines()) if line.strip() and not line.strip().startswith(('+', 'at ', '所在位置', 'CategoryInfo', 'FullyQualifiedErrorId'))), '后端返回失败，请查看完整日志。')
    if 'formula conversion failed' in text:
        return 'MathType 公式转换失败。\n' + detail[:900]
    return detail[:1000]


class PipelineRunner:
    def __init__(self, *, log_directory: Path | None = None) -> None:
        self.process: subprocess.Popen[str] | None = None
        self._cancel = threading.Event()
        self.log_path = (log_directory or Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'Markdown2Word' / 'logs') / 'latest.log'
        self.last_error = ''
        self._cancel_path: Path | None = None
        self._publish_lock = threading.Lock()
        self.output_path: Path | None = None
        self.warnings: list[str] = []

    @property
    def cancelled(self) -> bool:
        return self._cancel.is_set()

    def cancel(self) -> None:
        # GUI callbacks never wait for taskkill or a COM call.
        with self._publish_lock:
            self._cancel.set()
            cancel_path = self._cancel_path
            if cancel_path:
                try:
                    cancel_path.touch(exist_ok=True)
                except OSError:
                    pass  # The monitor still enforces cancellation after grace.

    def _terminate(self, process) -> None:
        if process.poll() is not None:
            return
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           check=False, timeout=10,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        else:
            process.terminate()

    def _execute(self, command, config, log, event, logfile) -> int:
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        if self._cancel_path:
            env["M2W_CANCEL_FILE"] = str(self._cancel_path)
        process = subprocess.Popen(command, cwd=str(config.markdown.parent),
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
            encoding="utf-8", errors="replace", bufsize=1, env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        self.process = process
        lines: queue.Queue = queue.Queue(maxsize=1000)
        stop_reading = threading.Event()
        def read():
            try:
                for line in process.stdout:
                    if stop_reading.is_set():
                        break
                    lines.put(line)
            finally:
                process.stdout.close()
                if not stop_reading.is_set():
                    lines.put(None)
        reader = threading.Thread(target=read, daemon=True)
        reader.start()
        tail = deque(maxlen=200)
        started = time.monotonic()
        stopping = None
        timed_out = False
        ended = False
        exited_at = None
        try:
            while not ended:
                now = time.monotonic()
                if process.poll() is not None:
                    if exited_at is None:
                        exited_at = now
                    elif now - exited_at > 2 and reader.is_alive():
                        log("后端已退出，停止等待继承日志管道的子进程。\n")
                        break
                if not timed_out and now - started > config.stage_timeout:
                    timed_out = True
                    if self._cancel_path:
                        self._cancel_path.touch(exist_ok=True)
                    log("转换等待超时，正在请求后端退出。\n")
                if self.cancelled or timed_out:
                    if stopping is None:
                        stopping = now
                    if now - stopping > 10:
                        self._terminate(process)
                        if process.poll() is None:
                            raise RuntimeError("后端未能退出；请检查任务日志和 Office 会话。")
                try:
                    line = lines.get(timeout=0.1)
                except queue.Empty:
                    if process.poll() is not None and not reader.is_alive():
                        break
                    continue
                if line is None:
                    ended = True
                    continue
                logfile.write(line)
                logfile.flush()
                tail.append(line)
                if line.startswith("M2W_EVENT "):
                    try:
                        event(json.loads(line[len("M2W_EVENT "):]))
                    except (ValueError, TypeError):
                        log(line)
                else:
                    log(line)
            # stdout can close before the process exits; keep the watchdog active.
            while process.poll() is None:
                now = time.monotonic()
                if not timed_out and now - started > config.stage_timeout:
                    timed_out = True
                    if self._cancel_path:
                        self._cancel_path.touch(exist_ok=True)
                    log("转换等待超时，正在请求后端退出。\n")
                if self.cancelled or timed_out:
                    if stopping is None:
                        stopping = now
                    if now - stopping > 10:
                        self._terminate(process)
                time.sleep(0.1)
            code = process.wait()
            if self.cancelled:
                return 130
            if timed_out:
                self.last_error = '转换等待超时。请检查 Word / MathType 是否有未处理的对话框，或在高级设置中增加超时时间。'
                return 124
            if code:
                self.last_error = describe_failure(''.join(tail))
                for hint in diagnostic_hints("".join(tail)):
                    log(f"  - {hint}\n")
            return code
        finally:
            if process.poll() is None:
                self._terminate(process)
            stop_reading.set()
            if not reader.is_alive():
                process.stdout.close()
            self.process = None

    def _publish_result(self, source: Path, output: Path, overwrite: bool, record, event) -> Path | None:
        """Publish atomically; preserve a verified result if the destination is locked."""
        denied = None
        folders = [output.parent, self.log_path.parent / 'recovered']
        for index, folder in enumerate(folders):
            publication = None
            try:
                if self.cancelled:
                    return None
                folder.mkdir(parents=True, exist_ok=True)
                descriptor, filename = tempfile.mkstemp(prefix='.markdown2word-', suffix='.tmp', dir=folder)
                os.close(descriptor)
                publication = Path(filename)
                shutil.copyfile(source, publication)
                if index == 0:
                    for attempt in range(3):
                        try:
                            with self._publish_lock:
                                if self.cancelled:
                                    return None
                                if overwrite:
                                    publication.replace(output)
                                else:
                                    os.link(publication, output)
                                return output
                        except PermissionError as exc:
                            denied = exc
                            if attempt < 2:
                                record('输出文件暂时无法写入，正在重试保存…\n')
                                if self._cancel.wait(.35):
                                    return None
                # A unique final DOCX never replaces an existing user document.
                suffix = publication.stem.removeprefix('.markdown2word-')
                alternate = folder / f'{output.stem}_转换结果_{time.strftime("%Y%m%d_%H%M%S")}_{suffix}.docx'
                with self._publish_lock:
                    if self.cancelled:
                        return None
                    os.link(publication, alternate)
                message = (f'原输出文件被占用或没有写入权限，未能保存到：{output}。'
                           f'已将转换成品另存为：{alternate}。原文件未修改。'
                           '如需覆盖原文件，请先关闭 Word 中的该文档，并检查文件及目录权限。')
                record(f'原保存错误：{denied}\n')
                event({'type': 'warning', 'message': message})
                return alternate
            except PermissionError as exc:
                denied = exc
                if index == len(folders) - 1:
                    raise PermissionError(f'转换已完成，但输出目录及恢复目录均无法写入。请检查权限：{output.parent}；{folder}。原错误：{exc}') from exc
            finally:
                if publication is not None:
                    try:
                        publication.unlink(missing_ok=True)
                    except OSError as exc:
                        event({'type': 'warning', 'message': f'保存临时文件未能清理：{publication}；{exc}'})
        return None

    def run(self, config: ConversionConfig, log: Callable[[str], None],
            event: Callable[[dict], None] | None = None) -> int:
        callback = event or (lambda value: None)
        self.last_error = ''
        self.output_path = None
        self.warnings.clear()
        prepared = {}
        def event(value):
            if value.get("type") == "warning":
                self.warnings.append(value.get("message", "后端警告"))
                log(f"警告：{self.warnings[-1]}\n")
            elif value.get('type') == 'prepared':
                prepared.update(value)
            callback(value)
        config = validate(config)
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        def digest(path):
            return hashlib.sha256(path.read_bytes()).hexdigest()
        template = config.template or STAGE1_BACKEND / "assets" / "default-template.docx"
        source_hashes = (digest(config.markdown), digest(template), digest(config.format_spec) if config.format_spec else None)
        with self.log_path.open('w', encoding='utf-8') as logfile:
            def record(message):
                logfile.write(message)
                logfile.flush()
                log(message)
            record(f"Markdown：{config.markdown}\n输出：{config.output}\n日志：{self.log_path}\n")
            try:
                with tempfile.TemporaryDirectory(prefix='markdown2word-') as directory:
                    workspace = Path(directory)
                    self._cancel_path = workspace / 'cancel.request'
                    work = replace(config, output=workspace / 'result.docx', overwrite=False)
                    if config.format_spec:
                        from formatting.spec import load_spec, dump_json
                        snapshot = workspace / 'format-spec.json'
                        dump_json(snapshot, load_spec(config.format_spec))
                        work = replace(work, format_spec=snapshot)
                    if self.cancelled:
                        return 130
                    code = self._execute([str(config.python), '-c', 'import docx, lxml'], config, log, event, logfile)
                    if code:
                        if code != 130:
                            self.last_error = 'Python 环境检查失败：所选 Python 必须安装 python-docx 和 lxml。\n' + self.last_error
                        return code
                    commands = build_commands(work)
                    for index, command in enumerate(commands):
                        if self.cancelled:
                            return 130
                        stage = 'prepare' if index == 0 else ('mathtype' if config.convert_formulas and index == 1 else 'visio')
                        label = {'prepare': '处理正文、模板和图表', 'mathtype': '转换 MathType 公式', 'visio': '嵌入可编辑图形'}[stage]
                        event({'type': 'stage', 'stage': stage, 'label': label, 'current': index, 'total': len(commands)})
                        record(f"\n{label}…\n")
                        if stage == 'mathtype':
                            if not prepared:
                                raise RuntimeError('文档处理没有返回公式检查结果。')
                            inline, display = prepared['inline'], prepared['display']
                            if inline + display == 0:
                                shutil.copy2(work.intermediate_path, work.output)
                                command = []
                                record('文档中没有公式，跳过 MathType。\n')
                            else:
                                command += ['-ExpectedInline', str(inline), '-ExpectedDisplay', str(display)]
                        code = self._execute(command, config, log, event, logfile) if command else 0
                        if code:
                            return code
                    if not work.output.is_file():
                        raise RuntimeError('转换后没有生成有效的 Word 文档。')
                    if config.format_spec:
                        from docx import Document
                        from formatting.audit import audit_format
                        audit_format(Document(work.output), load_spec(snapshot))
                    if (digest(config.markdown), digest(template), digest(config.format_spec) if config.format_spec else None) != source_hashes:
                        raise ConfigurationError('转换期间 Markdown 或模板发生变化，请重新开始。')
                    self.output_path = self._publish_result(work.output, config.output, config.overwrite, record, event)
                    if self.output_path is None:
                        return 130
                    event({'type': 'progress', 'stage': 'complete', 'current': 1, 'total': 1})
                    record(f'转换完成：{self.output_path}\n')
                    return 0
            except Exception as exc:
                if self.cancelled:
                    return 130
                self.last_error = str(exc)
                return 1
            finally:
                self._cancel_path = None
                if self.last_error and not self.cancelled:
                    record(f'\n转换失败：{self.last_error}\n')
                    event({'type': 'error', 'message': self.last_error})
                elif self.cancelled:
                    record('\n转换已取消，临时文件已清理。\n')
