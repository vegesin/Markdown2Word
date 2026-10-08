"""Desktop shell; conversion runs outside the webview UI thread."""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
from .bridge import DesktopApi


class WebViewRuntimeError(RuntimeError):
    """The initialized backend cannot provide the required Chromium renderer."""


def load_frontend() -> str:
    assets = Path(__file__).with_name('assets')
    return (assets / 'index.html').read_text(encoding='utf-8').replace(
        '<!-- APP_STYLE -->', '<style>' + (assets / 'app.css').read_text(encoding='utf-8') + '</style>'
    ).replace('<!-- APP_SCRIPT -->', '<script>' + (assets / 'app.js').read_text(encoding='utf-8') + '</script>')


def window_geometry(width: int, height: int) -> tuple[int, int]:
    """pywebview sizes are logical pixels, not physical screen pixels."""
    return min(1120, int(width * .9)), min(820, int(height * .9))


def main() -> None:
    try:
        import webview
    except ImportError as exc:
        raise SystemExit('请在当前 Python 环境执行：python -m pip install -e ".[pandoc]"') from exc
    try:
        if os.name == 'nt':
            os.environ['PYWEBVIEW_GUI'] = 'edgechromium'
        # screens is a lazy proxy: assigning it does not initialize the renderer.
        # Materialize it before inspecting renderer or native screen properties.
        screens = list(webview.screens)
        if os.name == 'nt' and getattr(webview, 'renderer', None) != 'edgechromium':
            raise WebViewRuntimeError('需要 Microsoft Edge WebView2 Runtime，无法使用旧版 IE 渲染器。')
        api = DesktopApi()
        screen = screens[0] if screens else None
        area = getattr(screen, 'frame', None)
        width, height = window_geometry(getattr(area, 'Width', screen.width), getattr(area, 'Height', screen.height)) if screen else (1100, 780)
        window = webview.create_window('Markdown2Word', html=load_frontend(), js_api=api,
            width=width, height=height, min_size=(min(540, width), min(420, height)),
            background_color='#F5F5F7', text_select=True, frameless=False, easy_drag=False)
        api._window = window
        window.events.closing += api._on_closing
        webview.start(gui='edgechromium' if os.name == 'nt' else None, debug=False)
    except Exception as exc:
        message = ('未能加载 Edge WebView2 运行时。请检查或修复 Microsoft Edge WebView2 Runtime。'
                   if isinstance(exc, WebViewRuntimeError) else '桌面界面启动失败。')
        message += f'\n\n{exc}\n\n当前 Python：{sys.executable}'
        log_path = Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'Markdown2Word' / 'logs' / 'startup.log'
        try:
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_path.write_text(message + '\n\n' + traceback.format_exc(), encoding='utf-8')
            message += f'\n启动日志：{log_path}'
        except OSError:
            pass
        if os.name == 'nt':
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, message, 'Markdown2Word', 0x10)
        raise SystemExit(message) from exc


if __name__ == '__main__':
    main()
