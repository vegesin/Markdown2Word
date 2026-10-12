from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from io import StringIO
from pathlib import Path
from unittest.mock import patch

from docx import Document
from backends.stage1 import build_stage1_docx as builder
from backends.stage1.markdown_ast import prepare_ast, formula_text, content_paragraphs
from core.pipeline import ConversionConfig, PipelineRunner, validate, find_pandoc


class ConversionTests(unittest.TestCase):
    def setUp(self):
        scratch = Path(__file__).resolve().parents[1] / ".test-tmp"
        scratch.mkdir(exist_ok=True)
        self.tmp = tempfile.TemporaryDirectory(dir=scratch)
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.md = self.root / "report.md"
        self.output = self.root / "report.docx"
        self.md.write_text("# Report\n\n$x$\n", encoding="utf-8")
        self.pandoc = find_pandoc()
        self.temp_patch = patch('tempfile.tempdir', str(self.root))
        self.temp_patch.start()
        self.addCleanup(self.temp_patch.stop)

    def config(self, **kwargs):
        return ConversionConfig(markdown=self.md, output=self.output,
            python=Path(sys.executable), pandoc=self.pandoc,
            convert_formulas=False, add_numbers=False, **kwargs)

    def build(self, *extra):
        command = [str(Path(builder.__file__)), str(self.md), str(self.output), "--pandoc", str(self.pandoc), *extra]
        with patch.object(sys, "argv", command), patch('sys.stdout', new_callable=StringIO):
            return builder.main()

    def test_ast_excludes_code_and_currency(self):
        self.md.write_text('`$x$`\n\n```python\ns = "$y$"\n```\n\n价格 \\$5 和 \\$10\n\n真正公式 $z$\n\n$$a+b$$\n', encoding="utf-8")
        _, formulas, _ = prepare_ast(self.md, str(self.pandoc))
        self.assertEqual([f["latex"] for f in formulas], ["z", "a+b"])
        self.build()
        doc = Document(self.output)
        text = "\n".join(p.text for p in content_paragraphs(doc))
        self.assertIn('s = "$y$"', text)
        self.assertIn("$5", text)
        self.assertEqual(sum(bool(formula_text(p)) for p in content_paragraphs(doc)), 2)

    def test_mermaid_tilde_and_nested(self):
        self.md.write_text('~~~mermaid\nflowchart TD\nA[开始] --> B[结束]\n~~~\n', encoding="utf-8")
        self.build("--write-visio-sources", "--visio-dir", str(self.root / "diagrams"))
        manifest = json.loads((self.root / "diagrams/manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(len(manifest), 1)
        self.assertIn("开始", (self.root / "diagrams/diagram-01.mmd").read_text(encoding="utf-8"))

    def test_display_math_blank_lines_compatible_without_source_changes(self):
        source = '$$\nH_r\n\n=\n\nE_\\theta(U_t)\n\n\\in\\mathbb R^{N_r\\times d_r}.\n$$\n'
        self.md.write_text(source, encoding='utf-8')
        before = self.md.read_bytes()
        for native in (False, True):
            notices = []
            _, formulas, _ = prepare_ast(self.md, str(self.pandoc), native_math=native, diagnostics=notices)
            self.assertEqual(len(formulas), 1)
            self.assertEqual(formulas[0]['kind'], 'display')
            self.assertIn(r'E_\theta(U_t)', formulas[0]['latex'])
            self.assertIn('第 1–9 行', notices[0])
        self.assertEqual(self.md.read_bytes(), before)
        self.build('--native-math')
        self.assertEqual(len(Document(self.output).element.xpath('.//m:oMath')), 1)

    def test_math_normalization_preserves_code_html_and_escaped_dollars(self):
        from backends.stage1.markdown_ast import normalize_display_math
        source = ('```python\n$$\nx\n\n$$\n```\n\n'
                  '~~~text\n$$\nnot closed\n~~~\n\n'
                  '    $$\n    literal\n\n'
                  '`$$` and ``$$\n\nliteral``\n\n'
                  '<pre>$$\nx\n\n$$</pre>\n\n'
                  '\\$\\$ and \\$$\n\n$$a+b$$\n')
        normalized, notices, _ = normalize_display_math(source)
        self.assertEqual(normalized, source)
        self.assertEqual(notices, [])
        self.md.write_text(source, encoding='utf-8')
        _, formulas, _ = prepare_ast(self.md, str(self.pandoc))
        self.assertEqual([f['latex'] for f in formulas], ['a+b'])

    def test_invalid_display_math_reports_original_line_number(self):
        for source, expected in [('正文\n\n$$\nx\n', '第 3 行'), ('正文\n\n$$\n\n$$\n', '第 3–5 行'),
                                 ('$$x\n\n+y$$\n', '第 1–3 行')]:
            self.md.write_text(source, encoding='utf-8')
            with self.subTest(source=source), self.assertRaisesRegex(ValueError, expected):
                prepare_ast(self.md, str(self.pandoc))

    def test_normalization_warning_reaches_desktop_pipeline(self):
        self.md.write_text('$$\nx\n\n+y\n$$\n', encoding='utf-8')
        runner = PipelineRunner(log_directory=self.root / 'logs')
        self.assertEqual(runner.run(self.config(), lambda _: None), 0, runner.last_error)
        self.assertTrue(any('第 1–5 行' in message for message in runner.warnings))

    def test_report_blank_line_formulas_are_now_detected(self):
        report = Path(__file__).parents[1] / 'test_doc'
        paths = list(report.glob('0916*.md'))
        if not paths:
            self.skipTest('Local example document unavailable')
        notices = []
        _, formulas, _ = prepare_ast(paths[0], str(self.pandoc), diagnostics=notices)
        self.assertEqual(sum(f['kind'] == 'display' for f in formulas), 23)
        self.assertEqual(len(notices), 9)
        for prefix in ('H_r', 'Z_r', 'H_m', r'H_{\mathrm{in}}'):
            self.assertTrue(any(f['kind'] == 'display' and f['latex'].startswith(prefix) for f in formulas))

    def test_literal_diagram_placeholder_in_code_is_not_targeted(self):
        self.md.write_text('`[VISIO:diagram-01]`\n\n~~~mermaid\nflowchart TD\nA --> B\n~~~\n', encoding='utf-8')
        self.build('--write-visio-sources')
        doc = Document(self.output)
        self.assertEqual(sum(builder.is_visio_placeholder(p) for p in content_paragraphs(doc)), 1)
        self.assertEqual(sum('[VISIO:diagram-01]' in p.text for p in content_paragraphs(doc)), 2)

    def test_mermaid_source_preserved_without_visio(self):
        self.md.write_text('```mermaid\nflowchart TD\nA --> B\n```', encoding='utf-8')
        self.build()
        text = '\n'.join(p.text for p in Document(self.output).paragraphs)
        self.assertIn('flowchart TD', text)
        self.assertNotIn('[VISIO:', text)

    def test_missing_image_is_reported_as_warning(self):
        self.md.write_text('![diagram](missing.png)\n\n$x$', encoding='utf-8')
        runner = PipelineRunner(log_directory=self.root / 'logs')
        self.assertEqual(runner.run(self.config(), lambda _: None), 0)
        self.assertTrue(runner.warnings)
        self.assertIn('missing.png', '\n'.join(runner.warnings))

    def test_caption_boundaries(self):
        self.assertIsNone(builder.caption_text_body("图像处理非常重要", "图"))
        self.assertIsNone(builder.caption_text_body("表明算法有效", "表"))
        self.assertEqual(builder.clean_caption_body("3D 结构", "图"), "3D 结构")
        self.assertEqual(builder.caption_text_body("图 1 系统框架", "图"), "系统框架")

    def test_tags_removed_only_from_mathtype_math(self):
        from backends.stage1.markdown_ast import strip_equation_tags
        self.assertEqual(strip_equation_tags(r'x+1\tag{A_{1}}+y\tag*{B}'), 'x+1+y')
        self.assertEqual(strip_equation_tags(r'\tagged{x}+\\tag{literal}'), r'\tagged{x}+\\tag{literal}')
        with self.assertRaises(ValueError):
            strip_equation_tags(r'x\tag{broken')
        self.md.write_text('$$x+1\\tag{12}$$\n\n`\\tag{keep}`\n', encoding='utf-8')
        _, formulas, _ = prepare_ast(self.md, str(self.pandoc))
        self.assertEqual(formulas[0]['latex'], 'x+1')
        _, native, _ = prepare_ast(self.md, str(self.pandoc), native_math=True)
        self.assertIn(r'\tag{12}', native[0]['latex'])
        self.build()
        doc = Document(self.output)
        self.assertIn('\\[x+1\\]', '\n'.join(p.text for p in doc.paragraphs))
        self.assertIn('\\tag{keep}', '\n'.join(p.text for p in doc.paragraphs))

    def test_numeric_citations_superscript_without_touching_code_or_references(self):
        self.md.write_text('结果见[1]、[1, 2]和[1-3]。链接[[4]](https://example.com)。\n\n`[5]` 和 $a[6]$\n\n```text\n[7]\n```\n\n# 8. 参考文献\n\n[1] Author [8]\n\n## Online\n\nOther reference [9]\n\n# 附录\n\n正文[10]\n', encoding='utf-8')
        self.build('--reference-superscript')
        doc = Document(self.output)
        raised = doc.element.xpath('.//w:r[w:rPr/w:vertAlign[@w:val="superscript"]]/w:t')
        self.assertEqual([node.text for node in raised], ['[1]', '[1, 2]', '[1-3]', '[4]', '[10]'])
        self.assertTrue(doc.element.xpath('.//w:hyperlink'))
        link = doc.element.xpath('.//w:hyperlink')[0]
        self.assertEqual(doc.part.rels[link.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')].target_ref, 'https://example.com')
        self.assertEqual(link.xpath('.//w:rPr/w:color/@w:val'), ['0563C1'])
        self.assertFalse(link.xpath('.//w:rPr/w:color/@w:themeColor'))

    def test_numeric_citations_disabled_by_default(self):
        self.md.write_text('引用[1]。', encoding='utf-8')
        self.build()
        self.assertFalse(Document(self.output).element.xpath('.//w:vertAlign[@w:val="superscript"]'))

    def test_code_blocks_preserve_text_and_have_readable_layout(self):
        source = 'def f(x):\n    text = "$x$ [1] \\tag{2}"\n    return x\n'
        self.md.write_text('```python\n' + source + '```\n', encoding='utf-8')
        self.build('--reference-superscript')
        doc = Document(self.output)
        code = [p for p in doc.paragraphs if p.style.name == 'Source Code']
        self.assertTrue(code)
        self.assertEqual('\n'.join(p.text for p in code).rstrip(), source.rstrip())
        self.assertTrue(all(p._p.xpath('./w:pPr/w:shd[@w:fill="F3F4F6"]') for p in code))
        self.assertFalse(doc.element.xpath('.//w:bookmarkStart[starts-with(@w:name,"M2WMath")]'))
        self.assertFalse(doc.element.xpath('.//w:vertAlign[@w:val="superscript"]'))
        self.assertTrue(all(not p.paragraph_format.keep_together for p in code))

    def test_local_images_embedded_from_multiple_path_forms(self):
        import base64
        from urllib.parse import quote
        directory = self.root / '图片 resources'
        directory.mkdir()
        image = directory / '示例 image.png'
        image.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aZ1sAAAAASUVORK5CYII='))
        relative = image.relative_to(self.root).as_posix()
        self.md.write_text('\n\n'.join([
            f'![相对路径](<{relative}>){{width=30in}}',
            f'![URL 编码]({quote(relative)})',
            f'![绝对路径](<{image.as_posix()}>)',
            f'![系统路径](<{str(image)}>)',
            f'![文件 URI]({image.as_uri()})',
        ]), encoding='utf-8')
        runner = PipelineRunner(log_directory=self.root / 'logs')
        self.assertEqual(runner.run(self.config(), lambda _: None), 0, runner.last_error)
        self.assertFalse(runner.warnings)
        doc = Document(self.output)
        self.assertEqual(len(doc.inline_shapes), 5)
        width = doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin
        self.assertTrue(all(shape.width <= width for shape in doc.inline_shapes))
        import zipfile
        with zipfile.ZipFile(self.output) as archive:
            self.assertTrue(any(name.startswith('word/media/') for name in archive.namelist()))

    def test_template_collision_rejected(self):
        template = self.output
        template.write_bytes(b"sentinel")
        config = ConversionConfig(markdown=self.md, output=self.output,
            template=template, python=Path(sys.executable), pandoc=self.pandoc)
        with self.assertRaises(ValueError):
            validate(config)
        self.assertEqual(template.read_bytes(), b"sentinel")

    def test_failed_audit_preserves_previous_output(self):
        self.output.write_bytes(b"previous output")
        with patch.object(builder, "audit_docx", return_value={"omml_nodes": 1, "ole_objects": 0, "equation_seq_fields": 0}):
            with self.assertRaises(RuntimeError):
                self.build("--force")
        self.assertEqual(self.output.read_bytes(), b"previous output")

    def test_existing_diagram_directory_preserved(self):
        directory = self.root / "diagrams"
        directory.mkdir()
        sentinel = directory / "manual.vsdx"
        sentinel.write_bytes(b"manual work")
        with self.assertRaises(FileExistsError):
            self.build("--write-visio-sources", "--visio-dir", str(directory))
        self.assertEqual(sentinel.read_bytes(), b"manual work")

    def test_runner_publishes_only_final_document(self):
        runner = PipelineRunner(log_directory=self.root / 'logs')
        events = []
        self.assertEqual(runner.run(self.config(), lambda line: None, events.append), 0)
        self.assertTrue(self.output.is_file())
        self.assertTrue(runner.log_path.is_file())
        self.assertEqual({p.name for p in self.root.iterdir()}, {'report.md', 'report.docx', 'logs'})
        self.assertEqual(len(Document(self.output).element.xpath('.//m:oMath')), 1)

    def test_runner_failure_preserves_previous_output(self):
        self.output.write_bytes(b"previous")
        runner = PipelineRunner(log_directory=self.root / 'logs')
        with patch("core.pipeline.build_commands", return_value=[
                [sys.executable, "-c", "raise SystemExit(7)"]]):
            self.assertEqual(runner.run(self.config(overwrite=True), lambda line: None), 7)
        self.assertEqual(self.output.read_bytes(), b"previous")
        self.assertEqual({p.name for p in self.root.iterdir()}, {'report.md', 'report.docx', 'logs'})
        self.assertTrue(runner.last_error)

    def test_locked_output_saves_verified_result_under_new_name(self):
        self.output.write_bytes(b'previous document')
        runner = PipelineRunner(log_directory=self.root / 'logs')
        with patch.object(Path, 'replace', side_effect=PermissionError(5, 'Access denied')):
            self.assertEqual(runner.run(self.config(overwrite=True), lambda _: None), 0, runner.last_error)
        self.assertEqual(self.output.read_bytes(), b'previous document')
        self.assertNotEqual(runner.output_path, self.output)
        self.assertEqual(runner.output_path.parent, self.output.parent)
        self.assertEqual(len(Document(runner.output_path).element.xpath('.//m:oMath')), 1)
        self.assertTrue(any(str(runner.output_path) in message for message in runner.warnings))
        self.assertFalse(list(self.root.glob('.markdown2word-*.tmp')))

    def test_temporary_output_lock_retries_original_path(self):
        self.output.write_bytes(b'previous document')
        runner = PipelineRunner(log_directory=self.root / 'logs')
        real_replace = Path.replace
        attempts = []
        def replace_after_one_failure(path, target):
            attempts.append(target)
            if len(attempts) == 1:
                raise PermissionError(5, 'Temporary lock')
            return real_replace(path, target)
        with patch.object(Path, 'replace', replace_after_one_failure):
            self.assertEqual(runner.run(self.config(overwrite=True), lambda _: None), 0)
        self.assertEqual(len(attempts), 2)
        self.assertEqual(runner.output_path, self.output)
        self.assertFalse(runner.warnings)
        self.assertEqual(len(Document(self.output).element.xpath('.//m:oMath')), 1)

    def test_unwritable_output_directory_uses_recovery_directory(self):
        runner = PipelineRunner(log_directory=self.root / 'logs')
        real_mkstemp = tempfile.mkstemp
        def deny_output_directory(*args, **kwargs):
            if kwargs.get('dir') == self.output.parent:
                raise PermissionError(5, 'Directory not writable')
            return real_mkstemp(*args, **kwargs)
        with patch('core.pipeline.tempfile.mkstemp', side_effect=deny_output_directory):
            self.assertEqual(runner.run(self.config(), lambda _: None), 0, runner.last_error)
        self.assertFalse(self.output.exists())
        self.assertEqual(runner.output_path.parent, self.root / 'logs' / 'recovered')
        self.assertEqual(len(Document(runner.output_path).element.xpath('.//m:oMath')), 1)
        self.assertTrue(runner.warnings)

    def test_cancel_during_publication_retry_preserves_old_document(self):
        self.output.write_bytes(b'previous document')
        runner = PipelineRunner(log_directory=self.root / 'logs')
        def cancel_on_retry(message):
            if '重试保存' in message:
                runner.cancel()
        with patch.object(Path, 'replace', side_effect=PermissionError(5, 'Access denied')):
            self.assertEqual(runner.run(self.config(overwrite=True), cancel_on_retry), 130)
        self.assertEqual(self.output.read_bytes(), b'previous document')
        self.assertIsNone(runner.output_path)
        self.assertFalse(list(self.root.glob('*转换结果*')))
        self.assertFalse(list(self.root.glob('.markdown2word-*.tmp')))

    def test_no_math_skips_office_even_with_mathtype_enabled(self):
        self.md.write_text('# Report\n\nNo formulas.', encoding='utf-8')
        config = ConversionConfig(markdown=self.md, output=self.output, python=Path(sys.executable), pandoc=self.pandoc)
        runner = PipelineRunner(log_directory=self.root / 'logs')
        self.assertEqual(runner.run(config, lambda _: None), 0)
        self.assertTrue(self.output.is_file())
        self.assertEqual({p.name for p in self.root.iterdir()}, {'report.md', 'report.docx', 'logs'})

    def test_cancel_returns_130(self):
        runner = PipelineRunner(log_directory=self.root / 'logs')
        def event(value):
            if value.get("type") == "stage":
                runner.cancel()
        code = runner.run(self.config(), lambda line: None, event)
        self.assertEqual(code, 130)
        self.assertFalse(self.output.exists())
        self.assertEqual({p.name for p in self.root.iterdir()}, {'report.md', 'logs'})

    def test_no_clobber(self):
        self.output.write_bytes(b"previous")
        with self.assertRaises(ValueError):
            PipelineRunner(log_directory=self.root / 'logs').run(self.config(), lambda line: None)
        self.assertEqual(self.output.read_bytes(), b"previous")

    def test_multiple_formulas_and_tex_dollar(self):
        self.md.write_text('$x$ 和 $y$，以及 $\\text{cost \\$5}$\n\n\\[a + b\\]\n', encoding='utf-8')
        self.build()
        from backends.stage1.markdown_ast import formula_literals
        literals = [t for p in content_paragraphs(Document(self.output)) for t in formula_literals(p)]
        self.assertEqual(len(literals), 4)
        self.assertIn('$\\text{cost \\$5}$', literals)

    def test_nested_table_math_counted_once(self):
        from backends.stage1.markdown_ast import materialize_formulas
        doc = Document()
        outer = doc.add_table(rows=1, cols=2)
        outer.cell(0, 0).merge(outer.cell(0, 1))
        nested = outer.cell(0, 0).add_table(rows=1, cols=1)
        nested.cell(0, 0).text = 'TOKEN'
        materialize_formulas(doc, [{'id': 'M2WMath000001', 'token': 'TOKEN', 'literal': '$x$'}])
        self.assertEqual(sum(bool(formula_text(p)) for p in content_paragraphs(doc)), 1)
        self.assertEqual(builder.format_tables(doc), 2)

    def test_formula_replacement_preserves_tabs_and_breaks(self):
        from backends.stage1.markdown_ast import materialize_formulas
        doc = Document()
        p = doc.add_paragraph()
        run = p.add_run('before')
        run.add_break()
        run.add_text('TOKEN1')
        run.add_tab()
        run.add_text('after TOKEN2')
        materialize_formulas(doc, [
            {'id': 'M2WMath000001', 'token': 'TOKEN1', 'literal': '$x$'},
            {'id': 'M2WMath000002', 'token': 'TOKEN2', 'literal': '$y$'}])
        self.assertEqual(p.text, 'before\n$x$\tafter $y$')

    def test_source_change_during_task_rejects_publication(self):
        runner = PipelineRunner(log_directory=self.root / 'logs')
        def event(value):
            if value.get('type') == 'stage':
                self.md.write_text('Changed while running', encoding='utf-8')
        self.assertEqual(runner.run(self.config(), lambda line: None, event), 1)
        self.assertFalse(self.output.exists())

    def test_rich_caption_retains_emphasis(self):
        doc = Document()
        builder.ensure_caption_style(doc)
        p = doc.add_paragraph('图 1 ')
        p.add_run('3D 结构').bold = True
        builder.write_caption(p, '图', '3D 结构', 1)
        self.assertIn('3D 结构', p.text)
        self.assertTrue(any(run.bold and '3D' in run.text for run in p.runs))

    def test_no_auto_captions_and_template_tables(self):
        self.md.write_text('| 名称 | 值 |\n| --- | --- |\n| 结果 | $x$ |\n', encoding='utf-8')
        self.build('--no-auto-captions', '--preserve-table-style')
        doc = Document(self.output)
        self.assertEqual(len(doc.tables), 1)
        self.assertFalse(any('请填写表题' in p.text for p in doc.paragraphs))

    def test_cancel_silent_backend_does_not_hang(self):
        runner = PipelineRunner(log_directory=self.root / 'logs')
        def commands(config):
            return [[sys.executable, '-c', "import os,time,pathlib; p=pathlib.Path(os.environ['M2W_CANCEL_FILE']); print('ready',flush=True)\nwhile not p.exists(): time.sleep(0.01)\nraise SystemExit(9)"]]
        def log(line):
            if 'ready' in line:
                runner.cancel()
        with patch('core.pipeline.build_commands', side_effect=commands):
            self.assertEqual(runner.run(self.config(), log), 130)
        self.assertIsNone(runner.process)

    def test_watchdog_handles_silent_backend(self):
        runner = PipelineRunner(log_directory=self.root / 'logs')
        # Call the process monitor directly so the test need not wait 30 seconds.
        from dataclasses import replace
        config = replace(self.config(), stage_timeout=0.02)
        runner._cancel_path = self.root / 'cancel.request'
        command = [sys.executable, '-c', "import os,time,pathlib; p=pathlib.Path(os.environ['M2W_CANCEL_FILE'])\nwhile not p.exists(): time.sleep(0.01)"]
        with (self.root / 'log').open('w', encoding='utf-8') as logfile:
            self.assertEqual(runner._execute(command, config, lambda line: None, lambda event: None, logfile), 124)

    def test_inline_ellipsis_extra_period_repair_is_strictly_local(self):
        import shutil
        powershell = shutil.which('powershell')
        if not powershell:
            self.skipTest('Windows PowerShell unavailable')
        backend = Path(builder.__file__).parents[1]
        script = r"""$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $args[0] 'mathtype/convert_word_latex_to_mathtype.ps1'),[ref]$tokens,[ref]$errors)
foreach ($name in @('Release-ComObject','Test-MathTypeInlineShape','Test-InlineConversionExtraPeriod','Repair-InlineConversionExtraPeriod')) {
  $fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
  Invoke-Expression $fn.Extent.Text
}
$script:deleted=0
$paragraph=[pscustomobject]@{Range=[pscustomobject]@{Start=0;End=20}}
$paragraphs=[pscustomobject]@{Entry=$paragraph}
$paragraphs | Add-Member ScriptMethod Item {param($i) return $this.Entry}
$shape=[pscustomobject]@{OLEFormat=[pscustomobject]@{ProgID='Equation.DSMT4'};Range=[pscustomobject]@{Start=5;End=6;Paragraphs=$paragraphs}}
$shapes=[pscustomobject]@{Count=1;Entry=$shape}
$shapes | Add-Member ScriptMethod Item {param($i) return $this.Entry}
$doc=[pscustomobject]@{Prefix='left:';Suffix='.rest';Shapes=$shapes}
$doc | Add-Member ScriptMethod Range {
  param($start,$end)
  if ($start -eq 5 -and $end -eq 6) {return [pscustomobject]@{InlineShapes=$this.Shapes}}
  if ($start -eq 0 -and $end -eq 5) {return [pscustomobject]@{Text=$this.Prefix}}
  if ($start -eq 6 -and $end -eq 20) {return [pscustomobject]@{Text=$this.Suffix}}
  if ($start -eq 6 -and $end -eq 7) {
    $extra=[pscustomobject]@{Text='.'}
    $extra | Add-Member ScriptMethod Delete {$script:deleted++}
    return $extra
  }
  throw "Unexpected range: $start..$end"
}
Repair-InlineConversionExtraPeriod -Document $doc -FormulaStart 5 -ExpectedPrefix 'left:' -ExpectedSuffix 'rest'
if ($script:deleted -ne 1) {throw 'Verified insertion was not repaired'}
# Preserve source punctuation, whitespace, and any genuinely changed prose.
foreach ($suffix in @('rest','..rest','.REST','.rest changed')) {
  $doc.Suffix=$suffix
  Repair-InlineConversionExtraPeriod -Document $doc -FormulaStart 5 -ExpectedPrefix 'left:' -ExpectedSuffix 'rest'
}
$doc.Suffix='.rest';$doc.Prefix='changed:'
Repair-InlineConversionExtraPeriod -Document $doc -FormulaStart 5 -ExpectedPrefix 'left:' -ExpectedSuffix 'rest'
if ($script:deleted -ne 1) {throw 'Unrelated text was modified'}
if (Test-InlineConversionExtraPeriod -ExpectedPrefix '' -ActualPrefix '' -ExpectedSuffix '.rest' -ActualSuffix '.rest') {throw 'Source period was removed'}
if (-not (Test-InlineConversionExtraPeriod -ExpectedPrefix '' -ActualPrefix '' -ExpectedSuffix '.rest' -ActualSuffix '..rest')) {throw 'Duplicate source period was not distinguished'}
"""
        script_path = self.root / 'check-period.ps1'
        script_path.write_text(script, encoding='utf-8-sig')
        result = subprocess.run([powershell, '-NoProfile', '-File', str(script_path), str(backend)], capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr.decode(errors='replace'))

    def test_powershell_parser_and_visio_labels(self):
        import shutil
        powershell = shutil.which('powershell')
        if not powershell:
            self.skipTest('Windows PowerShell unavailable')
        backend = Path(builder.__file__).parents[1]
        # Only parse scripts and invoke the pure Mermaid parser. No COM/Office.
        script = """$ErrorActionPreference='Stop'
Get-ChildItem -LiteralPath $args[0] -Filter *.ps1 -Recurse | ForEach-Object {
  $tokens=$null; $errors=$null
  [void][Management.Automation.Language.Parser]::ParseFile($_.FullName,[ref]$tokens,[ref]$errors)
  if (@($errors).Count) { throw ($errors | Out-String) }
}
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $args[0] 'visio/convert_mermaid_to_visio.ps1'),[ref]$tokens,[ref]$errors)
$fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Parse-MermaidFlowchart'},$true)
Invoke-Expression $fn.Extent.Text
$model=Parse-MermaidFlowchart -Source "flowchart TD`nA[Start] --> B[End]"
if ($model.Nodes[0].Text -ne 'Start' -or $model.Nodes[1].Text -ne 'End') { throw 'Node label lost' }
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $args[0] 'mathtype/convert_word_latex_to_mathtype.ps1'),[ref]$tokens,[ref]$errors)
foreach ($name in @('Release-ComObject','Get-TaskWordWindowHandle','Get-FormulaWordRanges','Get-DocumentTextWithoutMathTypeManagedFieldResults','Normalize-ConvertedText')) {
  $fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
  Invoke-Expression $fn.Extent.Text
}
Set-StrictMode -Version Latest
$word=[pscustomobject]@{ActiveWindow=[pscustomobject]@{Hwnd=12345}}
if ((Get-TaskWordWindowHandle -Word $word) -ne 12345) { throw 'Wrong Word window handle' }
$word.ActiveWindow.Hwnd=0
$rejected=$false
try { Get-TaskWordWindowHandle -Word $word } catch { $rejected=$true }
if (-not $rejected) { throw 'Invalid window handle accepted' }
Set-StrictMode -Off
$text='code $x$ actual $x$'
$range=[pscustomobject]@{Start=16;End=19;Text='$x$'}
$bookmark=[pscustomobject]@{Name='M2WMath000001';Range=[pscustomobject]@{Duplicate=$range}}
$bookmarks=[pscustomobject]@{Count=1;Entry=$bookmark}
$bookmarks | Add-Member ScriptMethod Item { param($i) return $this.Entry }
$doc=[pscustomobject]@{Bookmarks=$bookmarks;Fields=[pscustomobject]@{Count=0};Content=[pscustomobject]@{Start=0;End=$text.Length};Text=$text}
$doc | Add-Member ScriptMethod Range { param($start,$end) return [pscustomobject]@{Text=$this.Text.Substring($start,$end-$start)} }
$fingerprint=Get-DocumentTextWithoutMathTypeManagedFieldResults -Document $doc -ReplaceFormulas
if ($fingerprint -cne 'code $x$ actual [MATH]') { throw "Code was changed by fingerprint: $fingerprint" }
"""
        script_path = self.root / 'check.ps1'
        script_path.write_text(script, encoding='utf-8')
        result = subprocess.run([powershell, '-NoProfile', '-File', str(script_path), str(backend)], capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr.decode(errors='replace'))


if __name__ == "__main__":
    unittest.main()
