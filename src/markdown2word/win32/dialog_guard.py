from __future__ import annotations

import ctypes
import os
import re
import threading
import time
from ctypes import wintypes
from typing import Callable


_TITLE_RE = re.compile(r"^Convert\s+To\s+Display\s+Style\??$", re.IGNORECASE)
_INLINE_BUTTON_RE = re.compile(r"^Create\s+Inline\s+Style\s+Equation$", re.IGNORECASE)


def _normalize_control_text(text: str) -> str:
    # Native dialog captions may contain an invisible accelerator marker, for
    # example ``Create &Inline Style Equation``.  UI Automation usually strips
    # it while GetWindowText preserves it.
    return re.sub(r"\s+", " ", text.replace("&", "")).strip()


def is_inline_dialog_title(text: str) -> bool:
    return bool(_TITLE_RE.fullmatch(_normalize_control_text(text)))


def is_inline_button(text: str) -> bool:
    return bool(_INLINE_BUTTON_RE.fullmatch(_normalize_control_text(text)))


class MathTypeInlineDialogGuard:
    """Click only MathType's inline-style choice while a conversion is active."""

    BM_CLICK = 0x00F5

    def __init__(self, on_click: Callable[[str], None] | None = None) -> None:
        self.on_click = on_click
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self.click_count = 0
        self.last_click_method = ""
        self._reported_dialog_seen = False
        self._user32 = None
        self._enum_callback_type = None
        if os.name == "nt":
            self._user32 = ctypes.WinDLL("user32", use_last_error=True)
            self._enum_callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
            self._user32.EnumWindows.argtypes = [self._enum_callback_type, wintypes.LPARAM]
            self._user32.EnumChildWindows.argtypes = [wintypes.HWND, self._enum_callback_type, wintypes.LPARAM]
            self._user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
            self._user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
            self._user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
            self._user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            self._user32.IsWindow.argtypes = [wintypes.HWND]

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self.click_count = 0
        self.last_click_method = ""
        self._reported_dialog_seen = False
        self._thread = threading.Thread(target=self._watch, name="mathtype-inline-dialog-guard", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread and self._thread is not threading.current_thread():
            self._thread.join(timeout=1.0)
        self._thread = None

    def _window_text(self, hwnd: int) -> str:
        assert self._user32 is not None
        size = max(int(self._user32.GetWindowTextLengthW(hwnd)) + 1, 512)
        buffer = ctypes.create_unicode_buffer(size)
        self._user32.GetWindowTextW(hwnd, buffer, size)
        return buffer.value

    def _scan_once(self) -> bool:
        if self._user32 is None or self._enum_callback_type is None:
            return False
        clicked = False

        def dialog_still_open(hwnd: int) -> bool:
            return bool(self._user32.IsWindow(hwnd)) and is_inline_dialog_title(self._window_text(hwnd))

        @self._enum_callback_type
        def enum_window(hwnd, _):
            nonlocal clicked
            if not is_inline_dialog_title(self._window_text(hwnd)):
                return True
            if not self._reported_dialog_seen:
                self._reported_dialog_seen = True
                if self.on_click:
                    self.on_click("GUI 弹窗保护已检测到 MathType 行内/行间选择框，正在选择行内公式。\n")

            matching_child: int | None = None

            @self._enum_callback_type
            def enum_child(child, __):
                nonlocal matching_child
                caption = self._window_text(child)
                # MathType 6 can expose these controls through a custom or
                # owner-drawn window class.  The class name is therefore not a
                # reliable indication that this is a button.  The exact parent
                # dialog title plus exact control caption keeps the click
                # narrowly scoped without rejecting those legacy controls.
                if is_inline_button(caption):
                    matching_child = child
                    return False
                return True

            self._user32.EnumChildWindows(hwnd, enum_child, 0)
            if matching_child is not None:
                if self._user32.PostMessageW(matching_child, self.BM_CLICK, 0, 0):
                    self._stop.wait(0.12)
                    if not dialog_still_open(hwnd):
                        self.last_click_method = "BM_CLICK"
                        clicked = True
                        return False

            return not clicked

        self._user32.EnumWindows(enum_window, 0)
        return clicked

    def _watch(self) -> None:
        while not self._stop.is_set():
            try:
                if self._scan_once():
                    self.click_count += 1
                    if self.on_click:
                        self.on_click(
                            "GUI 弹窗保护已自动选择 Create Inline Style Equation"
                            f"（方式：{self.last_click_method}；累计 {self.click_count} 次）。\n"
                        )
                    # Let the modal close before scanning again, avoiding a
                    # second BM_CLICK against the same window.
                    self._stop.wait(0.4)
                    continue
            except Exception as exc:
                if self.on_click:
                    self.on_click(f"GUI 弹窗保护扫描异常：{exc}\n")
            self._stop.wait(0.1)
