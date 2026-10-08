#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

try:
    from .markdown_ast import prepare_ast, materialize_formulas, formula_text, formula_literals, content_paragraphs
except ImportError:  # Direct execution by the compatibility PowerShell backend.
    from markdown_ast import prepare_ast, materialize_formulas, formula_text, formula_literals, content_paragraphs


FIGURE_LABEL = "\u56fe"
TABLE_LABEL = "\u8868"
CHINESE_NUMERAL_RE = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Build a template-based DOCX with LaTeX preserved for MathType.")
    p.add_argument("markdown", type=Path)
    p.add_argument("template_or_output", type=Path)
    p.add_argument("output", type=Path, nargs="?")
    p.add_argument("--pandoc", default="pandoc")
    p.add_argument("--visio-dir", type=Path)
    p.add_argument("--force", action="store_true")
    p.add_argument("--preserve-table-style", action="store_true")
    p.add_argument("--no-auto-captions", action="store_true")
    p.add_argument("--native-math", action="store_true", help="Generate native Word equations without Office automation")
    p.add_argument('--reference-superscript', action='store_true', help='Superscript numeric citations in prose')
    p.add_argument(
        "--write-visio-sources",
        action="store_true",
        help="Write Mermaid .mmd files and manifest.json for the optional Visio stage.",
    )
    p.add_argument(
        "--default-template",
        type=Path,
        default=Path(__file__).resolve().parent / "assets" / "default-template.docx",
    )
    return p.parse_args()


def clear_paragraph(paragraph) -> None:
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def first_existing_style(doc: Document, names: list[str]):
    by_name = {style.name: style for style in doc.styles}
    for name in names:
        if name in by_name:
            return by_name[name]
    return None


def ensure_caption_style(doc: Document):
    try:
        return doc.styles["Caption"]
    except KeyError:
        style = doc.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "宋体"
        style.font.size = Pt(10.5)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.space_before = Pt(3)
        style.paragraph_format.space_after = Pt(3)
        return style


def paragraph_has_drawing(paragraph) -> bool:
    return bool(paragraph._p.xpath(".//*[local-name()='drawing']"))


def drawing_alt_text(paragraph) -> str:
    texts: list[str] = []
    for node in paragraph._p.xpath(".//*[local-name()='docPr']"):
        for key in ("descr", "title"):
            value = node.get(key)
            if value:
                value = value.strip()
                if value and value not in texts:
                    texts.append(value)
    return " ".join(texts).strip()


def is_visio_placeholder(paragraph) -> bool:
    return any(node.get(qn('w:name'), '').startswith('M2WDiagram') for node in paragraph._p.xpath('.//w:bookmarkStart'))


def caption_text_body(text: str, label: str) -> str | None:
    pattern = rf"^\s*{re.escape(label)}(?:\s*(?:\d+(?:[-.]\d+)*|[{CHINESE_NUMERAL_RE}]+)(?=[\s:：、.．-])|\s+|[:：])\s*[、:：.．-]*\s*(.*?)\s*$"
    match = re.match(pattern, text)
    if not match:
        return None
    body = match.group(1).strip()
    return body or None


def clean_caption_body(text: str, label: str) -> str:
    body = re.sub(r"\s+", " ", text or "").strip()
    if not body:
        return ""
    # Remove an existing Markdown/alt-text caption prefix such as:
    # "图 1 系统框架", "表1: 指标", or the Pandoc image-alt variant
    # "1 系统框架". The Word SEQ field supplies the live number.
    parsed = caption_text_body(body, label)
    return parsed if parsed is not None else body


def insert_paragraph_after(paragraph):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    return paragraph._parent.paragraphs[[p._p for p in paragraph._parent.paragraphs].index(new_p)]


def insert_paragraph_before_table(table):
    new_p = OxmlElement("w:p")
    table._tbl.addprevious(new_p)
    body = table._parent
    return body.paragraphs[[p._p for p in body.paragraphs].index(new_p)]


def write_caption(paragraph, label: str, body: str, number: int) -> None:
    # For an explicit caption, remove only its textual prefix. Keep hyperlinks,
    # emphasis, and formula bookmarks in the remaining body.
    original = paragraph.text
    preserve = bool(body and original.endswith(body) and not paragraph._p.xpath('.//w:fldSimple | .//w:fldChar'))
    if preserve:
        remaining = len(original) - len(body)
        for node in paragraph._p.xpath('.//w:t'):
            value = node.text or ''
            cut = min(remaining, len(value))
            node.text = value[cut:]
            remaining -= cut
            if not remaining:
                break
        retained = [child for child in paragraph._p if child.tag != qn('w:pPr')]
        for child in retained:
            paragraph._p.remove(child)
    else:
        retained = []
        body = clean_caption_body(body, label)
        clear_paragraph(paragraph)
    paragraph.style = paragraph.part.document.styles["Caption"]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.keep_with_next = label == TABLE_LABEL
    paragraph.paragraph_format.keep_together = True
    paragraph.add_run(label)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), f"SEQ {label} \\* ARABIC")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = str(number)
    run.append(text)
    field.append(run)
    paragraph._p.append(field)
    default_body = "\u8bf7\u586b\u5199\u56fe\u9898" if label == FIGURE_LABEL else "\u8bf7\u586b\u5199\u8868\u9898"
    if preserve:
        paragraph.add_run(' ')
        for child in retained:
            paragraph._p.append(child)
    else:
        paragraph.add_run(" " + (body.strip() if body.strip() else default_body))


def iter_body_blocks(doc: Document):
    paragraphs_by_element = {p._p: p for p in doc.paragraphs}
    tables_by_element = {t._tbl: t for t in doc.tables}
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p") and child in paragraphs_by_element:
            yield ("paragraph", paragraphs_by_element[child])
        elif child.tag == qn("w:tbl") and child in tables_by_element:
            yield ("table", tables_by_element[child])


def apply_captions(doc: Document) -> dict[str, int]:
    ensure_caption_style(doc)
    blocks = list(iter_body_blocks(doc))
    figure_no = 0
    table_no = 0
    figure_captions = 0
    table_captions = 0
    inserted_figures = 0
    inserted_tables = 0

    for index, (kind, block) in enumerate(blocks):
        if kind != "paragraph" or not (paragraph_has_drawing(block) or is_visio_placeholder(block)):
            continue
        paragraph = block
        figure_no += 1
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.keep_with_next = True
        body = drawing_alt_text(paragraph)
        caption_paragraph = None
        if index + 1 < len(blocks) and blocks[index + 1][0] == "paragraph":
            candidate = blocks[index + 1][1]
            parsed = caption_text_body(candidate.text, FIGURE_LABEL)
            if parsed is not None:
                caption_paragraph = candidate
                body = parsed
        if caption_paragraph is None:
            caption_paragraph = insert_paragraph_after(paragraph)
            inserted_figures += 1
        write_caption(caption_paragraph, FIGURE_LABEL, body, figure_no)
        figure_captions += 1

    blocks = list(iter_body_blocks(doc))
    for index, (kind, block) in enumerate(blocks):
        if kind != "table":
            continue
        table = block
        table_no += 1
        body = ""
        caption_paragraph = None
        if index > 0 and blocks[index - 1][0] == "paragraph":
            candidate = blocks[index - 1][1]
            parsed = caption_text_body(candidate.text, TABLE_LABEL)
            if parsed is not None:
                caption_paragraph = candidate
                body = parsed
        if caption_paragraph is None:
            caption_paragraph = insert_paragraph_before_table(table)
            inserted_tables += 1
        write_caption(caption_paragraph, TABLE_LABEL, body, table_no)
        table_captions += 1

    return {
        "figure_captions": figure_captions,
        "table_captions": table_captions,
        "inserted_figure_caption_placeholders": inserted_figures,
        "inserted_table_caption_placeholders": inserted_tables,
    }


def set_line_spacing_at_least(paragraph, twips: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:line"), str(twips))
    spacing.set(qn("w:lineRule"), "atLeast")


def format_display_equations(doc: Document) -> int:
    count = 0
    for paragraph in content_paragraphs(doc):
        if not formula_text(paragraph):
            continue
        match = re.fullmatch(r"\s*\\\[(.*?)\\\]\s*", paragraph.text, re.DOTALL)
        if not match:
            continue
        count += 1
        display_style = first_existing_style(doc, ["MTDisplayEquation", "MTDisplayE"])
        if display_style is not None:
            paragraph.style = display_style
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.keep_together = True
        paragraph.paragraph_format.space_before = Pt(3)
        paragraph.paragraph_format.space_after = Pt(3)
        set_line_spacing_at_least(paragraph, 600)
    return count


def all_content_paragraphs(doc: Document):
    yield from content_paragraphs(doc)


def format_inline_formula_spacing(doc: Document) -> int:
    count = 0
    for paragraph in all_content_paragraphs(doc):
        if any(text.startswith('$') for text in formula_literals(paragraph)):
            count += 1
            set_line_spacing_at_least(paragraph, 500)
    return count


def set_cell_margins(cell, top=80, start=100, bottom=80, end=100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tbl_pr.append(borders)

    if table.rows:
        for cell in table.rows[0].cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            old_tc = tc_pr.find(qn("w:tcBorders"))
            if old_tc is not None:
                tc_pr.remove(old_tc)
            tc_borders = OxmlElement("w:tcBorders")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "8")
            bottom.set(qn("w:color"), "000000")
            tc_borders.append(bottom)
            tc_pr.append(tc_borders)


def format_tables(doc: Document) -> int:
    from docx.table import Table
    tables = [Table(node, doc._body) for node in doc.element.body.xpath(".//w:tbl")]
    for table in tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        set_table_borders(table)
        if table.rows:
            tr_pr = table.rows[0]._tr.get_or_add_trPr()
            repeat = OxmlElement("w:tblHeader")
            repeat.set(qn("w:val"), "true")
            tr_pr.append(repeat)
        for row_index, row in enumerate(table.rows):
            for col_index, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_margins(cell)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 or col_index > 0 else WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        run.font.size = Pt(12)
                        if row_index == 0:
                            run.bold = True
    return len(tables)


def format_visio_placeholders(doc: Document) -> int:
    count = 0
    paragraphs = list(content_paragraphs(doc))
    for index, paragraph in enumerate(paragraphs):
        if is_visio_placeholder(paragraph):
            count += 1
            if "pic" in [s.name for s in doc.styles]:
                paragraph.style = doc.styles["pic"]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.keep_with_next = True
            for run in paragraph.runs:
                run.bold = True
            if index + 1 < len(paragraphs) and re.match(r"^图\s*\d", paragraphs[index + 1].text.strip()):
                if "Caption" in [s.name for s in doc.styles]:
                    paragraphs[index + 1].style = doc.styles["Caption"]
    return count


def count_headings(doc: Document) -> dict[str, int]:
    result: dict[str, int] = {}
    for paragraph in doc.paragraphs:
        if paragraph.style and paragraph.style.name.startswith("Heading "):
            result[paragraph.style.name] = result.get(paragraph.style.name, 0) + 1
    return result


def audit_docx(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
        ns = {
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "o": "urn:schemas-microsoft-com:office:office",
        }
        return {
            "omml_nodes": len(root.xpath(".//m:oMath | .//m:oMathPara", namespaces=ns)),
            "ole_objects": len(root.xpath(".//w:object | .//o:OLEObject", namespaces=ns)),
            "equation_seq_fields": len(root.xpath(".//w:fldSimple[contains(@w:instr, 'SEQ Equation')]", namespaces=ns)),
        }


def audit_layout(doc, *, preserve_table_style: bool) -> None:
    for paragraph in content_paragraphs(doc):
        literals = list(formula_literals(paragraph))
        display = any(text.startswith('\\[') for text in literals) or bool(paragraph._p.xpath('.//m:oMathPara'))
        inline = bool(literals) or bool(paragraph._p.xpath('.//m:oMath'))
        if display or inline:
            spacing = paragraph._p.find(qn('w:pPr'))
            spacing = spacing.find(qn('w:spacing')) if spacing is not None else None
            minimum = 600 if display else 500
            if spacing is None or spacing.get(qn('w:lineRule')) != 'atLeast' or int(spacing.get(qn('w:line'), '0')) < minimum:
                raise RuntimeError('Equation paragraph spacing failed validation')
    if preserve_table_style:
        return
    for table in doc.element.xpath('.//w:tbl'):
        borders = table.xpath('./w:tblPr/w:tblBorders')
        values = {etree.QName(edge).localname: edge.get(qn('w:val')) for edge in borders[0]} if borders else {}
        header = table.xpath('./w:tr[1]/w:tc[1]/w:tcPr/w:tcBorders/w:bottom')
        if any(values.get(edge) != 'single' for edge in ('top', 'bottom')) or any(values.get(edge) != 'nil' for edge in ('left', 'right', 'insideH', 'insideV')) or not header or header[0].get(qn('w:val')) != 'single':
            raise RuntimeError('Table border formatting failed validation')


def format_code_blocks(doc: Document) -> int:
    count = 0
    for paragraph in content_paragraphs(doc):
        if not paragraph.style or paragraph.style.name != 'Source Code':
            continue
        count += 1
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt = paragraph.paragraph_format
        fmt.line_spacing = 1.0
        fmt.space_before = Pt(6)
        fmt.space_after = Pt(6)
        fmt.left_indent = Pt(8)
        fmt.right_indent = Pt(8)
        fmt.first_line_indent = Pt(0)
        fmt.keep_together = False  # Long blocks must be able to cross pages.
        fmt.keep_with_next = False
        properties = paragraph._p.get_or_add_pPr()
        shading = properties.find(qn('w:shd'))
        if shading is None:
            shading = OxmlElement('w:shd')
            properties.append(shading)
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:fill'), 'F3F4F6')
        for run in paragraph.runs:
            run.font.name = 'Consolas'
            run.font.size = Pt(9)
            run._r.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei UI')
            if run._r.rPr.find(qn('w:noProof')) is None:
                run._r.rPr.append(OxmlElement('w:noProof'))
    return count


def fit_images(doc: Document) -> int:
    # Match Pandoc's page-wide figures to the template's printable area.
    width = min(section.page_width - section.left_margin - section.right_margin for section in doc.sections)
    count = 0
    for shape in doc.inline_shapes:
        if shape.width > width:
            shape.height = round(shape.height * width / shape.width)
            shape.width = width
        count += 1
    return count


def main() -> int:
    args = parse_args()
    markdown = args.markdown.resolve()
    if args.output is None:
        template = args.default_template.resolve()
        output = args.template_or_output.resolve()
    else:
        template = args.template_or_output.resolve()
        output = args.output.resolve()
    if not markdown.is_file():
        raise FileNotFoundError(f"Markdown path does not exist: {markdown}")
    if not template.is_file():
        raise FileNotFoundError(f"Word template path does not exist: {template}")
    if output in {markdown, template}:
        raise ValueError("Output must not overwrite the input or template")
    if output.exists() and not args.force:
        raise FileExistsError(f"Output exists: {output}; use --force to replace")
    output.parent.mkdir(parents=True, exist_ok=True)
    visio_dir = (args.visio_dir or output.with_name(output.stem + "_visio_sources")).resolve()
    if args.write_visio_sources and visio_dir.exists():
        raise FileExistsError(f"Visio sources directory already exists; choose a new directory: {visio_dir}")
    if args.write_visio_sources and any(path == visio_dir or visio_dir in path.parents for path in (markdown, template, output)):
        raise ValueError("Visio directory must not contain input, template, or output")

    math_diagnostics = []
    ast, formulas, manifest = prepare_ast(markdown, args.pandoc, native_math=args.native_math, convert_diagrams=args.write_visio_sources, reference_superscript=args.reference_superscript, diagnostics=math_diagnostics)
    for message in math_diagnostics:
        print('M2W_EVENT ' + json.dumps({'type': 'warning', 'message': message}, ensure_ascii=False), flush=True)
    source_inline = sum(f["kind"] == "inline" for f in formulas)
    source_display = sum(f["kind"] == "display" for f in formulas)

    with tempfile.TemporaryDirectory(prefix="md-template-docx-", dir=output.parent) as tmp:
        tmp_path = Path(tmp)
        temp_md = tmp_path / "source.json"
        temp_docx = tmp_path / "pandoc.docx"
        audited_docx = tmp_path / "audited.docx"
        temp_md.write_text(json.dumps(ast, ensure_ascii=False), encoding="utf-8")
        cmd = [
            args.pandoc,
            "--from=json",
            "--to=docx",
            f"--reference-doc={template}",
            f"--resource-path={markdown.parent}",
            "--wrap=none",
            f"--output={temp_docx}",
            str(temp_md),
        ]
        result = subprocess.run(cmd, capture_output=True, encoding='utf-8', errors='replace')
        if result.returncode:
            raise RuntimeError(f'Pandoc 转换失败：{result.stderr.strip()}')
        if result.stderr.strip():
            detail = result.stderr.strip()
            message = ('图片资源缺失，文档中已用描述替代。请补齐 Markdown 引用的图片后重新转换。\n' if 'Could not fetch resource' in detail else 'Pandoc 转换警告：\n') + detail
            print('M2W_EVENT ' + json.dumps({'type': 'warning', 'message': message}, ensure_ascii=False), flush=True)
        doc = Document(temp_docx)
        code_blocks = format_code_blocks(doc)
        images = fit_images(doc)
        materialize_formulas(doc, ([] if args.native_math else formulas) + [{"id": d["bookmark"], "token": d["token"], "literal": d["placeholder"]} for d in manifest])
        display_count = format_display_equations(doc)
        table_count = 0 if args.preserve_table_style else format_tables(doc)
        visio_count = format_visio_placeholders(doc)
        caption_summary = {} if args.no_auto_captions else apply_captions(doc)
        inline_formula_paragraphs = format_inline_formula_spacing(doc)
        doc.core_properties.title = markdown.stem
        doc.core_properties.subject = "Markdown technical report"
        if args.native_math:
            for paragraph in content_paragraphs(doc):
                if paragraph._p.xpath('.//m:oMathPara'):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_line_spacing_at_least(paragraph, 600)
                elif paragraph._p.xpath('.//m:oMath'):
                    set_line_spacing_at_least(paragraph, 500)
        doc.save(audited_docx)

        reopened = Document(audited_docx)
        audit_layout(reopened, preserve_table_style=args.preserve_table_style)
        math_texts = [text for p in content_paragraphs(reopened) for text in formula_literals(p)]
        inline_count = sum(text.startswith('$') for text in math_texts)
        display_text_count = sum(text.startswith('\\[') for text in math_texts)
        package_audit = audit_docx(audited_docx)
        if not args.native_math and (source_inline != inline_count or source_display != display_text_count):
            raise RuntimeError("Formula counts changed during conversion")
        if len(manifest) != visio_count:
            raise RuntimeError("Mermaid placeholder count mismatch")
        native_count = len(reopened.element.xpath('.//m:oMath'))
        if args.native_math and native_count != len(formulas):
            raise RuntimeError(f"Native equation count mismatch: {native_count}/{len(formulas)}")
        if package_audit['ole_objects'] or package_audit['equation_seq_fields'] or (not args.native_math and package_audit['omml_nodes']):
            raise RuntimeError("Prepared DOCX contains unexpected formula objects or numbering fields")
        if args.write_visio_sources:
            visio_dir.mkdir(parents=True, exist_ok=False)
            for diagram in manifest:
                (visio_dir / diagram["source"]).write_text(diagram["text"], encoding="utf-8")
            (visio_dir / "manifest.json").write_text(json.dumps(
                [{k: v for k, v in d.items() if k != "text"} for d in manifest], ensure_ascii=False, indent=2), encoding="utf-8")
        # A failed audit never replaces a previous successful document.
        if args.force:
            audited_docx.replace(output)
        else:
            import os
            os.link(audited_docx, output)  # Atomic no-clobber publication.
    summary = {
        "output": str(output),
        "visio_dir": str(visio_dir) if args.write_visio_sources else None,
        "visio_sources_written": args.write_visio_sources,
        "source_inline_math": source_inline,
        "source_display_math": source_display,
        "docx_inline_latex": inline_count,
        "docx_display_latex": display_text_count,
        "display_paragraphs_formatted": display_count,
        "inline_formula_paragraphs_at_least_25pt": inline_formula_paragraphs,
        "tables_formatted": table_count,
        "code_paragraphs_formatted": code_blocks,
        "embedded_images": images,
        **caption_summary,
        "mermaid_sources": len(manifest),
        "visio_placeholders": visio_count,
        "headings": count_headings(reopened),
        **package_audit,
    }
    print('M2W_EVENT ' + json.dumps({"type": "prepared", "inline": source_inline, "display": source_display, **summary}, ensure_ascii=False))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
