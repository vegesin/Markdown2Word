"""Desktop API for the bundled local frontend; no HTTP API."""
from __future__ import annotations

import json
import os
import subprocess
import threading
from collections import deque
from dataclasses import fields
from pathlib import Path
from core.version import __version__
from .template_api import TemplateApi
from core.pipeline import (
    BACKEND_ROOT, NUMBER_FORMATS, ConversionConfig, PipelineRunner,
    find_pandoc, find_powershell, find_python, validate,
)


def default_settings() -> dict:
    config = ConversionConfig(markdown=Path('.'), output=Path('.'))
    result = {field.name: getattr(config, field.name) for field in fields(config)
              if field.name not in {'markdown', 'output', 'template', 'format_spec', 'python', 'pandoc'}}
    return {**result, 'template': '', 'format_spec': '', 'python': str(find_python()),
            'pandoc': str(find_pandoc()), 'dark_mode': False,
            'ai_base_url':'', 'ai_model':'', 'ai_key_env':'M2W_AI_API_KEY', 'ai_json_mode':False}


def normalize_settings(data: object, defaults: dict) -> dict:
    result = defaults.copy()
    if not isinstance(data, dict):
        return result
    data = data.copy()
    if 'auto_start_word' not in data and 'auto_word' in data:
        data['auto_start_word'] = data['auto_word']
    for key, default in defaults.items():
        value = data.get(key)
        if isinstance(default, bool):
            if isinstance(value, bool):
                result[key] = value
        elif key == 'stage_timeout':
            try:
                if not isinstance(value, bool) and 30 <= int(value) <= 86400:
                    result[key] = int(value)
            except (ValueError, TypeError, OverflowError):
                pass
        elif isinstance(value, str):
            result[key] = value.strip()
    result['number_format'] = NUMBER_FORMATS.get(result['number_format'], result['number_format'])
    if result['number_format'] not in NUMBER_FORMATS.values():
        result['number_format'] = 'Simple'
    return result


def make_config(data: dict, defaults: dict) -> ConversionConfig:
    if not isinstance(data, dict):
        raise ValueError('转换参数格式不正确。')
    options = normalize_settings(data, defaults)
    try:
        timeout = int(data.get('stage_timeout', options['stage_timeout']))
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError('等待超时必须是 30～86400 之间的整数。') from exc
    values = {field.name: options[field.name] for field in fields(ConversionConfig)
              if field.name in options and field.name not in {'template', 'format_spec', 'python', 'pandoc'}}
    values['stage_timeout'] = timeout
    values['add_numbers'] = options['convert_formulas'] and options['add_numbers']
    for key in ('markdown', 'output', 'template', 'format_spec', 'python', 'pandoc'):
        raw = data.get(key, options.get(key, ''))
        if not isinstance(raw, str):
            raise ValueError(f'{key} 路径必须是文本。')
        values[key] = Path(raw.strip()) if raw.strip() else (Path('.') if key in {'markdown', 'output'} else None)
    return ConversionConfig(**values)


class DesktopApi(TemplateApi):
    def __init__(self, *, settings_path: Path | None = None, runner_factory=PipelineRunner):
        self._lock = threading.RLock()
        self._window = None
        self._runner_factory = runner_factory
        self._runner = runner_factory()
        self._worker = None
        self._closing = False
        self._defaults = default_settings()
        self._settings_path = settings_path or Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'Markdown2Word' / 'settings.json'
        try:
            saved = json.loads(self._settings_path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            saved = {}
        self._settings = normalize_settings(saved, self._defaults)
        for name in ('python', 'pandoc'):
            if not Path(self._settings[name]).is_file() or 'codex-runtimes' in self._settings[name]:
                self._settings[name] = self._defaults[name]
        self._logs = deque(maxlen=2000)
        self._seq = 0
        self._generation = 0
        self._progress_base = 0
        self._progress_span = 100
        self._state = {'state': 'idle', 'status': '准备好开始转换', 'progress': 0,
                       'error': '', 'warnings': [], 'output': '', 'close_requested': False}
        self._init_templates()

    def bootstrap(self) -> dict:
        with self._lock:
            return {'settings': self._settings.copy(), 'version': __version__}

    def save_settings(self, data: dict) -> dict:
        with self._lock:
            try:
                settings = normalize_settings(data, self._defaults)
                self._settings_path.parent.mkdir(parents=True, exist_ok=True)
                temporary = self._settings_path.with_suffix('.tmp')
                temporary.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding='utf-8')
                temporary.replace(self._settings_path)
                self._settings = settings
                return {'ok': True}
            except OSError as exc:
                return {'ok': False, 'error': f'偏好设置保存失败：{exc}'}

    def pick_file(self, kind: str, current: str = '') -> dict:
        import webview
        types = {'markdown': ('Markdown (*.md;*.markdown)',),
                 'requirements': ('格式要求 (*.txt;*.md;*.docx;*.json)',),
                 'format_spec': ('格式协议 (*.json)',),
                 'draft_output': ('规范草稿 (*.json)',),
                 'template': ('Word 模板 (*.docx;*.dotx;*.docm)',),
                 'output': ('Word 文档 (*.docx)',),
                 'python': ('Python (python.exe)',), 'pandoc': ('Pandoc (pandoc.exe)',)}
        if kind not in types or self._window is None:
            return {'ok': False, 'error': '无法打开文件选择窗口。'}
        try:
            path = Path(current) if current else Path.home()
            selected = self._window.create_file_dialog(
                webview.FileDialog.SAVE if kind in {'output','draft_output'} else webview.FileDialog.OPEN,
                directory=str(path if path.is_dir() else path.parent),
                save_filename=path.name if kind in {'output','draft_output'} and current else ('draft.json' if kind=='draft_output' else 'output.docx'),
                file_types=types[kind] + ('所有文件 (*.*)',))
            if not selected:
                return {'ok': True, 'path': None}
            path = Path(selected[0])
            if kind == 'output' and path.suffix.lower() != '.docx':
                path = path.with_suffix('.docx')
            if kind == 'draft_output' and path.suffix.lower() != '.json':
                path = path.with_suffix('.json')
            return {'ok': True, 'path': str(path)}
        except Exception as exc:
            return {'ok': False, 'error': str(exc)}

    def _busy(self) -> bool:
        return self._worker is not None and self._worker.is_alive()

    def _log(self, text: str) -> None:
        with self._lock:
            self._seq += 1
            self._logs.append({'seq': self._seq, 'text': str(text)[:8000]})

    def _event(self, event: dict) -> None:
        with self._lock:
            if event.get('type') in {'stage', 'progress'} and not self._runner.cancelled:
                total = max(1, event.get('total', 1))
                fraction = min(1, max(0, event.get('current', 0) / total))
                if event['type'] == 'stage':
                    self._progress_base = fraction * 100
                    self._progress_span = 100 / total
                    progress = self._progress_base
                else:
                    progress = 100 if event.get('stage') == 'complete' else self._progress_base + fraction * self._progress_span
                status = event.get('label') or f"正在转换公式与图形… {event.get('current', 0)}/{total}"
                self._state.update(progress=min(100, progress), status=status)

    def poll(self, after: int = 0) -> dict:
        try:
            after = int(after)
        except (ValueError, TypeError):
            after = 0
        with self._lock:
            return {**self._state, 'warnings': list(self._state['warnings']),
                    'busy': self._busy(), 'generation': self._generation, 'last_seq': self._seq,
                    'logs': [line.copy() for line in self._logs if line['seq'] > after],
                    'truncated': bool(self._logs and after and after < self._logs[0]['seq'] - 1)}

    def start_conversion(self, data: dict) -> dict:
        with self._lock:
            if self._busy() or self._closing:
                return {'ok': False, 'error': '当前任务尚未结束。'}
            try:
                config = validate(make_config(data, self._defaults))
            except (ValueError, OSError) as exc:
                return {'ok': False, 'error': str(exc)}
            saved = self.save_settings(data)
            self._runner = self._runner_factory()
            self._generation += 1
            self._progress_base, self._progress_span = 0, 100
            self._logs.clear()
            self._state.update(state='running', status='正在准备文档…', progress=0, error='', warnings=[], output='')
            if not saved['ok']:
                self._log(saved['error'] + '\n')
            self._worker = threading.Thread(target=self._convert, args=(config,), name='document-conversion', daemon=False)
            self._worker.start()
            return {'ok': True, 'output': str(config.output)}

    def _convert(self, config: ConversionConfig) -> None:
        try:
            code = self._runner.run(config, self._log, self._event)
            with self._lock:
                if code == 0:
                    warnings = list(self._runner.warnings)
                    self._state.update(state='warning' if warnings else 'success', status='转换完成，含警告' if warnings else '文档已准备就绪',
                                       progress=100, output=str(getattr(self._runner, 'output_path', None) or config.output), warnings=warnings)
                elif code == 130:
                    self._state.update(state='cancelled', status='已取消，临时文件已清理')
                else:
                    self._state.update(state='error', status='转换未能完成', error=self._runner.last_error or '请查看完整日志。')
        except Exception as exc:
            self._log(f'转换错误：{exc}\n')
            with self._lock:
                self._state.update(state='error', status='转换未能完成', error=str(exc))

    def cancel(self) -> dict:
        with self._lock:
            if self._state['state'] not in {'running', 'cancelling'}:
                return {'ok': False, 'error': '当前没有可取消的转换任务。'}
            self._runner.cancel()
            self._state.update(state='cancelling', status='正在取消，等待后台清理…')
            return {'ok': True}

    def diagnose(self, data: dict) -> dict:
        with self._lock:
            if self._busy() or self._closing:
                return {'ok': False, 'error': '请等待当前任务完成。'}
            settings = normalize_settings(data, self._defaults)
            self._generation += 1
            self._logs.clear()
            self._state.update(state='diagnosing', status='正在检查运行环境…', error='', warnings=[], output='', progress=0)
            self._worker = threading.Thread(target=self._diagnose, args=(settings,), name='environment-check', daemon=False)
            self._worker.start()
            return {'ok': True}

    def _diagnose(self, settings: dict) -> None:
        commands = [
            ('Python 及依赖', [settings['python'], '-c', 'import sys,docx,lxml; print(sys.executable); print("python-docx",docx.__version__); print("lxml",lxml.__version__)']),
            ('Pandoc', [settings['pandoc'], '--version']),
            ('Word / MathType', [find_powershell(), '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(BACKEND_ROOT / 'office_session.ps1'), '-CheckOnly']),
        ]
        failed = []
        try:
            self._runner.log_path.parent.mkdir(parents=True, exist_ok=True)
            self._runner.log_path.write_text('环境检测进行中…\n', encoding='utf-8')
        except OSError:
            pass  # Report any persistence failure with the final diagnostic result.
        try:
            for name, command in commands:
                try:
                    result = subprocess.run(command, capture_output=True, encoding='utf-8', errors='replace', timeout=15,
                                            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                    self._log(f'{name}：\n{result.stdout[:3000]}{result.stderr[:3000]}\n')
                    if result.returncode:
                        failed.append(name)
                    elif name == 'Word / MathType':
                        info = json.loads(result.stdout)
                        if not info.get('word_path') or not info.get('mathtype_registered'):
                            self._log('未发现完整 Word / MathType 环境；仍可使用 Word 原生公式模式。\n')
                except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
                    failed.append(name)
                    self._log(f'{name}：{exc}\n')
            self._log('环境检测不启动 Office；实际加载项连接在转换时检查。\n')
        finally:
            with self._lock:
                diagnostic_log = ''.join(line['text'] for line in self._logs)
            try:
                self._runner.log_path.write_text(diagnostic_log, encoding='utf-8')
            except OSError as exc:
                failed.append('日志保存')
                self._log(f'诊断日志保存失败：{exc}\n')
            with self._lock:
                self._state.update(state='error' if failed else 'idle', status='环境检测发现问题' if failed else '环境检测完成，请查看日志',
                                   error='以下检测未通过：' + '、'.join(failed) if failed else '')

    def open_output(self, folder: bool = False) -> dict:
        with self._lock:
            output = self._state['output']
        if not output:
            return {'ok': False, 'error': '尚未生成输出文档。'}
        return self._open(Path(output).parent if folder else Path(output))

    def open_log(self) -> dict:
        return self._open(self._runner.log_path)

    def copy_log(self) -> dict:
        """Copy the persisted full log, falling back to the current UI buffer."""
        with self._lock:
            text = ''.join(line['text'] for line in self._logs)
            path = self._runner.log_path
            waiting = self._busy() and not text
        if waiting:
            return {'ok': False, 'error': '日志尚未生成，请稍后重试。'}
        source = 'buffer'
        try:
            persisted = path.read_text(encoding='utf-8', errors='replace')
            if persisted:
                text, source = persisted, 'file'
        except OSError:
            pass
        if not text:
            return {'ok': False, 'error': '暂无可复制的日志。'}
        try:
            import pyperclip
            pyperclip.copy(text)
        except Exception as exc:
            return {'ok': False, 'error': f'复制日志失败：{exc}。可打开完整日志后复制。'}
        return {'ok': True, 'source': source}

    def _open(self, path: Path) -> dict:
        try:
            if not path.exists():
                raise FileNotFoundError(f'文件不存在：{path}')
            os.startfile(path)
            return {'ok': True}
        except OSError as exc:
            return {'ok': False, 'error': str(exc)}

    def _on_closing(self):
        with self._lock:
            if self._busy():
                if not self._closing:
                    self._state['close_requested'] = True
                return False
        return True

    def request_close(self, confirmed: bool = False) -> dict:
        with self._lock:
            self._state['close_requested'] = False
            if not confirmed or self._closing:
                return {'ok': True}
            self._closing = True
            if self._busy():
                if self._state['state'] == 'templating':
                    self._template_cancel.set()
                if self._state['state'] in {'running', 'cancelling'}:
                    self._runner.cancel()
                self._state.update(state='cancelling', status='正在清理，完成后退出…')
            worker = self._worker
        def finish():
            if worker:
                worker.join()
            if self._window:
                self._window.destroy()
        threading.Thread(target=finish, name='desktop-shutdown', daemon=False).start()
        return {'ok': True}
