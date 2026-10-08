"""Use Pandoc's parser as the single source of truth for Markdown structure."""
from __future__ import annotations

import json
import re
import subprocess
import uuid
from copy import deepcopy
from pathlib import Path
from urllib.parse import unquote, urlsplit

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

READER = "markdown+tex_math_dollars+tex_math_single_backslash"
BOOKMARK_PREFIX = "M2WMath"


def strip_equation_tags(latex: str) -> str:
    """Remove tag/tag* commands with balanced groups, respecting escaped braces."""
    result = []
    cursor = 0
    while cursor < len(latex):
        if latex[cursor] != '\\':
            result.append(latex[cursor])
            cursor += 1
            continue
        match = re.match(r'\\tag\*?(?![A-Za-z])\s*\{', latex[cursor:])
        if match:
            end = cursor + match.end()
            depth = 1
            while end < len(latex) and depth:
                if latex[end] == '\\':
                    end += 2
                    continue
                depth += (latex[end] == '{') - (latex[end] == '}')
                end += 1
            if depth:
                raise ValueError('公式中的 \\tag 参数缺少闭合花括号。')
            cursor = end
        else:
            # Consume an entire command or escaped character, including \\\\.
            command = re.match(r'\\(?:[A-Za-z]+|.)', latex[cursor:], re.DOTALL)
            end = cursor + (command.end() if command else 1)
            result.append(latex[cursor:end])
            cursor = end
    return ''.join(result).strip()


REFERENCE_RE = re.compile(r'\[[1-9]\d*(?:\s*[,，;；–—-]\s*[1-9]\d*)*\]')


def windows_image_targets(source: str) -> dict[str, str]:
    """Recover literal Windows separators consumed by Markdown punctuation escapes.

    This map is consulted only for parsed Image nodes; code is never rewritten.
    Ambiguous targets are left to Pandoc instead of guessing a source file.
    """
    patterns = [r'!\[[^\]\n]*\]\(\s*(?:<([^>\n]+)>|([^\s)]+))',
                r'^\s{0,3}\[[^\]\n]+\]:\s*(?:<([^>\n]+)>|([^\s]+))']
    candidates: dict[str, set[str]] = {}
    punctuation = r'''\!"#$%&'()*+,-./:;<=>?@[]^_`{|}~'''
    for pattern in patterns:
        for match in re.finditer(pattern, source, re.MULTILINE):
            raw = match.group(1) or match.group(2)
            if not re.match(r'^[A-Za-z]:\\|^\\\\', raw):
                continue
            parsed = re.sub(r'\\([' + re.escape(punctuation) + r'])', r'\1', raw)
            key = unquote(parsed).replace('\\', '/')
            candidates.setdefault(key, set()).add(unquote(raw))
    return {key: next(iter(values)) for key, values in candidates.items() if len(values) == 1}


def superscript_references(ast: dict) -> None:
    """Change textual numeric citations, excluding bibliography, math and code."""
    def inline_walk(nodes):
        result = []
        index = 0
        while index < len(nodes):
            node = nodes[index]
            if node.get('t') in {'Str', 'Space'}:
                end = index
                while end < len(nodes) and nodes[end].get('t') in {'Str', 'Space'}:
                    end += 1
                text = ''.join(' ' if n['t'] == 'Space' else n['c'] for n in nodes[index:end])
                matches = list(REFERENCE_RE.finditer(text))
                if not matches:
                    result.extend(nodes[index:end])
                else:
                    cursor = 0
                    for match in matches:
                        if match.start() > cursor:
                            result.append({'t': 'Str', 'c': text[cursor:match.start()]})
                        result.append({'t': 'Superscript', 'c': [{'t': 'Str', 'c': match.group()}]})
                        cursor = match.end()
                    if cursor < len(text):
                        result.append({'t': 'Str', 'c': text[cursor:]})
                index = end
                continue
            if node.get('t') in {'Emph', 'Strong', 'Strikeout', 'SmallCaps', 'Underline'}:
                node['c'] = inline_walk(node['c'])
            elif node.get('t') in {'Link', 'Span'}:
                node['c'][1] = inline_walk(node['c'][1])
            result.append(node)
            index += 1
        return result

    def blocks_walk(value, reference_level=None):
        if isinstance(value, list):
            for item in value:
                if isinstance(item, dict) and item.get('t') == 'Header':
                    title = ''.join(n.get('c', '') for n in item['c'][2] if n.get('t') == 'Str')
                    level = item['c'][0]
                    if re.fullmatch(r'(?:\d+(?:\.\d+)*[.、]?\s*)?(?:参考文献|references|bibliography)[:：]?', title.strip(), re.IGNORECASE):
                        reference_level = level
                    elif reference_level is not None and level <= reference_level:
                        reference_level = None
                elif isinstance(item, dict) and item.get('t') in {'Para', 'Plain'}:
                    nodes = item['c']
                    first = next((n for n in nodes if n.get('t') != 'Space'), {})
                    # A leading reference label is a bibliography entry, not a citation.
                    leading_label = first.get('t') == 'Str' and re.match(r'^\[\d+\]', first.get('c', ''))
                    if reference_level is None and not leading_label:
                        item['c'] = inline_walk(nodes)
                else:
                    blocks_walk(item, reference_level)
        elif isinstance(value, dict) and value.get('t') not in {'Code', 'CodeBlock', 'Math', 'RawInline', 'RawBlock'}:
            for item in value.values():
                if isinstance(item, (list, dict)):
                    blocks_walk(item, reference_level)
    blocks_walk(ast.get('blocks', []))


def normalize_display_math(source: str) -> tuple[str, list[str], list[int]]:
    """Normalize explicit display blocks, leaving code and source files intact."""
    lines = source.splitlines(keepends=True)
    masked = []
    fence = None
    for line in lines:
        # Recognize fences in lists/quotes too; normalize only top-level math.
        probe = re.sub(r'^\s*(?:>\s*)+', '', line).lstrip()
        match = re.match(r'(?:[-+*]\s+|\d+[.)]\s+)?(`{3,}|~{3,})(.*)', probe)
        protected = fence is not None or line.startswith(('    ', '\t'))
        if fence:
            if re.fullmatch(re.escape(fence[0]) + '{' + str(fence[1]) + r',}\s*', probe):
                fence = None
        elif match and not (match[1][0] == '`' and '`' in match[2]):
            fence = (match[1][0], len(match[1]))
            protected = True
        masked.append(re.sub(r'[^\r\n]', ' ', line) if protected else line)
    visible = ''.join(masked)
    # Raw HTML and inline/multiline code spans are literal content.
    def hide(match):
        return re.sub(r'[^\r\n]', ' ', match.group())
    visible = re.sub(r'<!--.*?-->|<(pre|code|script|style)\b[^>]*>.*?</\1\s*>', hide, visible, flags=re.DOTALL | re.IGNORECASE)
    visible = re.sub(r'(?<![\\`])(`+)(?!`)(.*?)(?<!`)\1(?!`)', hide, visible, flags=re.DOTALL)
    positions = [m.start() for m in re.finditer(r'(?<!\$)\$\$(?!\$)', visible)
                 if (len(visible[:m.start()]) - len(visible[:m.start()].rstrip('\\'))) % 2 == 0]
    line_numbers = [visible.count('\n', 0, pos) + 1 for pos in positions]
    if len(positions) % 2:
        raise ValueError(f'Markdown 第 {line_numbers[-1]} 行附近的 $$ 公式分隔符未配对，请检查开头和结尾；字面文本请放入代码或使用转义。')
    removed = set()
    diagnostics = []
    for pair, (start, end) in enumerate(zip(line_numbers[::2], line_numbers[1::2])):
        if not source[positions[2 * pair] + 2:positions[2 * pair + 1]].strip():
            raise ValueError(f'Markdown 第 {start}–{end} 行的 $$ 公式内容为空。')
        if start == end:
            continue
        blanks = [i for i in range(start, end - 1) if not lines[i].strip()]
        if not blanks:
            continue
        if not (re.fullmatch(r' {0,3}\$\$\s*', lines[start - 1]) and re.fullmatch(r' {0,3}\$\$\s*', lines[end - 1])):
            raise ValueError(f'Markdown 第 {start}–{end} 行的 $$ 公式内含空行，请将分隔符独占一行或删除公式内空行。')
        removed.update(blanks)
        diagnostics.append(f'已兼容 Markdown 第 {start}–{end} 行的公式：忽略公式内部 {len(blanks)} 个空白行，源文件未修改。')
    return ''.join(line for i, line in enumerate(lines) if i not in removed), diagnostics, line_numbers


def prepare_ast(markdown: Path, pandoc: str, *, native_math: bool = False, convert_diagrams: bool = True, reference_superscript: bool = False, diagnostics: list[str] | None = None) -> tuple[dict, list[dict], list[dict]]:
    source = markdown.read_text(encoding='utf-8-sig')
    normalized, notices, math_lines = normalize_display_math(source)
    result = subprocess.run(
        [pandoc, f"--from={READER}", "--to=json", f"--resource-path={markdown.parent}"],
        input=normalized,
        capture_output=True, encoding="utf-8", errors='replace',
    )
    if result.returncode:
        raise RuntimeError(f'Pandoc Markdown 解析失败：{result.stderr.strip()}')
    ast = json.loads(result.stdout)
    def display_count(value):
        if isinstance(value, list):
            return sum(display_count(item) for item in value)
        if not isinstance(value, dict):
            return 0
        if value.get('t') == 'Math':
            return int(value['c'][0]['t'] == 'DisplayMath')
        return sum(display_count(item) for item in value.values())
    if display_count(ast) < len(math_lines) // 2:
        raise ValueError('Markdown 中存在未识别的 $$ 公式，请检查第 ' + '、'.join(map(str, math_lines)) + ' 行附近的分隔符和段落格式。')
    if diagnostics is not None:
        diagnostics.extend(notices)
    windows_targets = windows_image_targets(source)
    if reference_superscript:
        superscript_references(ast)
    formulas: list[dict] = []
    diagrams: list[dict] = []
    nonce = uuid.uuid4().hex

    def walk(value):
        if isinstance(value, list):
            return [walk(item) for item in value]
        if not isinstance(value, dict):
            return value
        if value.get("t") == "Math":
            kind, latex = value["c"]
            display = kind["t"] == "DisplayMath"
            # Preserve spaces inside TeX commands; only remove physical newlines.
            latex = latex.replace("\r", " ").replace("\n", " ").strip()
            if not native_math:
                latex = strip_equation_tags(latex)
            index = len(formulas) + 1
            token = f"M2WTOKEN{nonce}F{index:06d}END"
            formulas.append({
                "id": f"{BOOKMARK_PREFIX}{index:06d}", "token": token,
                "kind": "display" if display else "inline",
                "latex": latex,
                "literal": f"\\[{latex}\\]" if display else f"${latex}$",
            })
            return {"t": "Math", "c": [kind, latex]} if native_math else {"t": "Str", "c": token}
        if value.get('t') == 'Image':
            target = value['c'][2][0]
            parsed = urlsplit(target)
            if parsed.scheme.lower() == 'file':
                target = unquote(parsed.path)
                if parsed.netloc:
                    target = '//' + parsed.netloc + target
                elif re.match(r'^/[A-Za-z]:/', target):
                    target = target[1:]
            elif not parsed.scheme or re.match(r'^[A-Za-z]:[\\/]', target):
                target = unquote(target)
            else:
                return value  # HTTP and data URLs are handled by Pandoc.
            target = windows_targets.get(target.replace('\\', '/'), target)
            path = Path(target.replace('\\', '/')).expanduser()
            if not path.is_absolute():
                path = markdown.parent / path
            value['c'][2][0] = path.resolve().as_posix()
            return value
        if convert_diagrams and value.get("t") == "CodeBlock" and "mermaid" in value["c"][0][1]:
            source = value["c"][1].strip() + "\n"
            diagram_id = f"diagram-{len(diagrams) + 1:02d}"
            placeholder = f"[VISIO:{diagram_id}]"
            token = f"M2WTOKEN{nonce}D{len(diagrams) + 1:06d}END"
            diagrams.append({"id": diagram_id, "placeholder": placeholder,
                             "bookmark": f"M2WDiagram{len(diagrams) + 1:06d}", "token": token,
                             "source": f"{diagram_id}.mmd", "text": source,
                             "diagram_type": source.splitlines()[0] if source.strip() else ""})
            return {"t": "Para", "c": [{"t": "Str", "c": token}]}
        return {key: walk(item) for key, item in value.items()}

    # Split standalone display nodes at the block level, including list/table
    # blocks, without touching code, escaped dollars, or literal prose.
    def split_blocks(value):
        if isinstance(value, list):
            result = []
            for item in value:
                if isinstance(item, dict) and item.get("t") in {"Para", "Plain"}:
                    pending = []
                    for inline in item["c"]:
                        if inline.get("t") == "Math" and inline["c"][0]["t"] == "DisplayMath":
                            if any(i.get("t") not in {"Space", "SoftBreak", "LineBreak"} for i in pending):
                                result.append({"t": item["t"], "c": pending})
                            result.append({"t": "Para", "c": [inline]})
                            pending = []
                        else:
                            pending.append(inline)
                    if any(i.get("t") not in {"Space", "SoftBreak", "LineBreak"} for i in pending):
                        result.append({"t": item["t"], "c": pending})
                else:
                    result.append(split_blocks(item))
            return result
        if isinstance(value, dict):
            return {key: split_blocks(item) for key, item in value.items()}
        return value

    return walk(split_blocks(ast)), formulas, diagrams


def materialize_formulas(doc, formulas: list[dict]) -> None:
    """Replace unique AST tokens with bookmarked LaTeX, preserving other runs."""
    by_token = {formula["token"]: formula for formula in formulas}
    if not by_token:
        return
    pattern = re.compile("|".join(re.escape(token) for token in by_token))
    ids = [int(node.get(qn("w:id"), "0")) for node in doc.element.xpath(".//w:bookmarkStart")]
    bookmark_id = max(ids, default=0) + 1
    found = set()
    for node in list(doc.element.xpath(".//w:t")):
        text = node.text or ""
        matches = list(pattern.finditer(text))
        if not matches:
            continue
        run = node.getparent()
        if run.tag != qn("w:r"):
            raise RuntimeError("Unexpected Pandoc formula token location")
        # A run may also contain tabs, line breaks or several text nodes. Keep
        # both sides of the target node instead of replacing the entire run.
        children = list(run)
        following = children[children.index(node) + 1:]
        for child in following:
            run.remove(child)
        run.remove(node)
        position = run
        cursor = 0

        def append_text(content):
            nonlocal position
            if not content:
                return
            new_run = OxmlElement("w:r")
            if run.rPr is not None:
                new_run.append(deepcopy(run.rPr))
            text_node = OxmlElement("w:t")
            text_node.set(qn("xml:space"), "preserve")
            text_node.text = content
            new_run.append(text_node)
            position.addnext(new_run)
            position = new_run

        for match in matches:
            append_text(text[cursor:match.start()])
            formula = by_token[match.group()]
            if formula["id"] in found:
                raise RuntimeError("Duplicate formula token")
            found.add(formula["id"])
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(bookmark_id))
            start.set(qn("w:name"), formula["id"])
            position.addnext(start)
            position = start
            append_text(formula["literal"])
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(bookmark_id))
            position.addnext(end)
            position = end
            bookmark_id += 1
            cursor = match.end()
        append_text(text[cursor:])
        if following:
            suffix = OxmlElement('w:r')
            if run.rPr is not None:
                suffix.append(deepcopy(run.rPr))
            for child in following:
                suffix.append(child)
            position.addnext(suffix)
        if not any(child.tag != qn('w:rPr') for child in run):
            run.getparent().remove(run)
    if len(found) != len(formulas):
        raise RuntimeError(f"Formula token count mismatch: {len(found)}/{len(formulas)}")


def formula_literals(paragraph) -> list[str]:
    """Only return AST-approved math; literal delimiters in code stay literal."""
    texts = {}
    active = set()
    for node in paragraph._p.iter():
        if node.tag == qn("w:bookmarkStart") and node.get(qn("w:name"), "").startswith(BOOKMARK_PREFIX):
            active.add(node.get(qn("w:id")))
            texts[node.get(qn("w:id"))] = []
        elif node.tag == qn("w:bookmarkEnd"):
            active.discard(node.get(qn("w:id")))
        elif active and node.tag == qn("w:t"):
            for bookmark in active:
                texts[bookmark].append(node.text or "")
    return ["".join(parts) for parts in texts.values()]


def formula_text(paragraph) -> str:
    return " ".join(formula_literals(paragraph))


def content_paragraphs(doc):
    from docx.text.paragraph import Paragraph
    for node in doc.element.body.xpath(".//w:p"):
        yield Paragraph(node, doc._body)
