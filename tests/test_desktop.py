from __future__ import annotations

import json
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core.pipeline import PipelineRunner, find_pandoc
from gui.app import load_frontend, window_geometry
from gui.bridge import DesktopApi, default_settings, make_config, normalize_settings


class DesktopTests(unittest.TestCase):
    """Exercise bridge and real Pandoc conversion without opening a webview."""
    def setUp(self):
        scratch = Path(__file__).resolve().parents[1] / '.test-tmp'
        scratch.mkdir(exist_ok=True)
        self.tmp = tempfile.TemporaryDirectory(dir=scratch)
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.api = DesktopApi(settings_path=self.root / 'settings.json',
                              runner_factory=lambda: PipelineRunner(log_directory=self.root / 'logs'))
        self.md = self.root / 'report.md'
        self.md.write_text('# Report\n\n$x$\n\n正文引用[1]', encoding='utf-8')
        self.data = {**self.api.bootstrap()['settings'], 'markdown': str(self.md),
                     'output': str(self.root / 'report.docx'), 'convert_formulas': False,
                     'python': sys.executable, 'pandoc': str(find_pandoc()), 'reference_superscript': True}

    def test_legacy_settings_migration_and_strict_booleans(self):
        settings = normalize_settings({'auto_word': False, 'number_format': '(1), (2), ...',
                                       'overwrite': 'true', 'stage_timeout': '120'}, default_settings())
        self.assertFalse(settings['auto_start_word'])
        self.assertEqual(settings['number_format'], 'Simple')
        self.assertFalse(settings['overwrite'])
        self.assertEqual(settings['stage_timeout'], 120)
        self.assertEqual(normalize_settings([], settings), settings)

    def test_settings_roundtrip_omits_document_paths(self):
        self.assertTrue(self.api.save_settings(self.data)['ok'])
        saved = json.loads((self.root / 'settings.json').read_text(encoding='utf-8'))
        self.assertNotIn('markdown', saved)
        self.assertNotIn('output', saved)
        self.assertTrue(saved['reference_superscript'])
        self.assertFalse((self.root / 'settings.tmp').exists())

    def test_config_native_mode_cannot_accidentally_enable_mathtype_numbers(self):
        config = make_config({**self.data, 'add_numbers': True}, default_settings())
        self.assertFalse(config.add_numbers)
        with self.assertRaises(ValueError):
            make_config({**self.data, 'stage_timeout': 'abc'}, default_settings())

    def test_validation_failure_does_not_start_worker(self):
        result = self.api.start_conversion({**self.data, 'markdown': ''})
        self.assertFalse(result['ok'])
        self.assertFalse(self.api.poll()['busy'])
        self.assertIsNone(self.api._worker)

    def test_real_conversion_through_bridge(self):
        self.assertTrue(self.api.start_conversion(self.data)['ok'])
        self.api._worker.join(20)
        self.assertFalse(self.api._worker.is_alive())
        state = self.api.poll()
        self.assertEqual(state['state'], 'success', state['error'])
        self.assertEqual(state['progress'], 100)
        self.assertTrue(Path(state['output']).is_file())
        self.assertTrue(state['logs'])
        self.assertEqual(self.api.poll(state['last_seq'])['logs'], [])

    def test_duplicate_start_and_close_wait_for_cleanup(self):
        started, cancelled, cleanup, destroyed = [threading.Event() for _ in range(4)]
        class Runner:
            warnings = []
            last_error = ''
            log_path = Path('unused')
            @property
            def cancelled(self): return cancelled.is_set()
            def cancel(self): cancelled.set()
            def run(self, config, log, event):
                started.set()
                cleanup.wait(5)
                return 130
        self.api._runner_factory = Runner
        self.api._window = Mock()
        self.api._window.destroy.side_effect = destroyed.set
        self.addCleanup(cleanup.set)
        self.assertTrue(self.api.start_conversion(self.data)['ok'])
        self.assertTrue(started.wait(2))
        self.assertFalse(self.api.start_conversion(self.data)['ok'])
        self.assertFalse(self.api._on_closing())
        self.assertTrue(self.api.poll()['close_requested'])
        self.api.request_close(True)
        self.assertTrue(cancelled.is_set())
        self.assertFalse(destroyed.is_set())
        cleanup.set()
        self.assertTrue(destroyed.wait(2))

    def test_completed_task_exposes_actual_fallback_output_path(self):
        alternate = self.root / 'report_转换结果.docx'
        self.api._runner.output_path = alternate
        self.api._runner.warnings = ['输出被占用，已另存成品']
        with patch.object(self.api._runner, 'run', return_value=0):
            self.api._convert(make_config(self.data, default_settings()))
        state = self.api.poll()
        self.assertEqual(state['state'], 'warning')
        self.assertEqual(state['output'], str(alternate))

    def test_logs_are_bounded_incremental_and_snapshot_safe(self):
        for index in range(2100):
            self.api._log(str(index))
        state = self.api.poll(1)
        self.assertEqual(len(state['logs']), 2000)
        self.assertTrue(state['truncated'])
        state['logs'][0]['text'] = 'changed by client'
        self.assertNotEqual(self.api.poll()['logs'][0]['text'], 'changed by client')

    def test_formula_progress_stays_within_overall_task_progress(self):
        self.api._event({'type': 'stage', 'current': 1, 'total': 2, 'label': '转换公式'})
        self.api._event({'type': 'progress', 'current': 1, 'total': 10})
        self.assertEqual(self.api.poll()['progress'], 55)
        self.api._event({'type': 'progress', 'current': 10, 'total': 10})
        self.assertEqual(self.api.poll()['progress'], 100)

    def test_copy_log_prefers_full_file_and_supports_unicode(self):
        self.api._log('屏幕截断的日志')
        path = self.api._runner.log_path
        path.parent.mkdir()
        path.write_text('完整日志：公式 π_k\n错误详情', encoding='utf-8')
        with patch('pyperclip.copy') as copy:
            self.assertEqual(self.api.copy_log(), {'ok': True, 'source': 'file'})
            copy.assert_called_once_with('完整日志：公式 π_k\n错误详情')

    def test_copy_log_buffer_empty_and_clipboard_failure(self):
        with patch('pyperclip.copy') as copy:
            self.assertFalse(self.api.copy_log()['ok'])
            copy.assert_not_called()
            self.api._log('当前日志')
            self.assertEqual(self.api.copy_log()['source'], 'buffer')
            copy.assert_called_once_with('当前日志')
        with patch('pyperclip.copy', side_effect=RuntimeError('clipboard unavailable')):
            result = self.api.copy_log()
            self.assertFalse(result['ok'])
            self.assertIn('clipboard unavailable', result['error'])

    def test_dialog_cancel_and_output_extension(self):
        import webview
        self.api._window = Mock()
        self.api._window.create_file_dialog.return_value = None
        self.assertIsNone(self.api.pick_file('markdown')['path'])
        self.api._window.create_file_dialog.return_value = [str(self.root / 'result.txt')]
        result = self.api.pick_file('output')
        self.assertTrue(result['path'].endswith('result.docx'))
        self.assertEqual(self.api._window.create_file_dialog.call_args.args[0], webview.FileDialog.SAVE)

    def test_diagnostics_replace_stale_log_without_starting_office(self):
        self.api._runner.log_path.parent.mkdir()
        self.api._runner.log_path.write_text('previous conversion', encoding='utf-8')
        with patch('gui.bridge.subprocess.run', return_value=Mock(returncode=0, stdout='{}', stderr='')) as run:
            self.assertTrue(self.api.diagnose(self.data)['ok'])
            self.api._worker.join(5)
        log = self.api._runner.log_path.read_text(encoding='utf-8')
        self.assertNotIn('previous conversion', log)
        self.assertIn('环境检测不启动 Office', log)
        self.assertIn('-CheckOnly', run.call_args.args[0])

    def test_frontend_is_bundled_without_external_requests(self):
        html = load_frontend()
        self.assertNotIn('<!-- APP_', html)
        self.assertIn("connect-src 'none'", html)
        self.assertNotRegex(html, r'(?:src|href)=[\"\']https?://')
        self.assertIn('pywebviewready', html)

    def test_window_fits_logical_work_area(self):
        for width, height in ((1366, 728), (960, 520), (800, 560), (1920, 1040)):
            w, h = window_geometry(width, height)
            self.assertLessEqual(w, width * .9)
            self.assertLessEqual(h, height * .9)

    def test_launcher_initializes_lazy_screens_before_renderer_check(self):
        from types import SimpleNamespace
        from gui.app import main
        fake = SimpleNamespace(renderer=None, create_window=Mock(), start=Mock())
        class LazyScreens:
            def __iter__(self):
                fake.renderer = 'edgechromium'
                return iter([SimpleNamespace(width=1920, height=1080, frame=None)])
        fake.screens = LazyScreens()
        window = Mock()
        window.events.closing = Mock()
        window.events.closing.__iadd__ = Mock(return_value=window.events.closing)
        fake.create_window.return_value = window
        with patch.dict(sys.modules, {'webview': fake}), patch('gui.app.DesktopApi'):
            main()
        self.assertEqual(fake.renderer, 'edgechromium')
        fake.create_window.assert_called_once()
        self.assertFalse(fake.create_window.call_args.kwargs['easy_drag'])
        self.assertFalse(fake.create_window.call_args.kwargs['frameless'])
        fake.start.assert_called_once()


if __name__ == '__main__':
    unittest.main()
