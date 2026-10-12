from __future__ import annotations

import copy
import json
import math
from pathlib import Path

from jsonschema import Draft202012Validator

ROOT = Path(__file__).parent
STYLE_NAMES = {
    'body': 'Normal', 'first_paragraph': 'First Paragraph',
    'figure_caption': 'Figure Caption', 'table_caption': 'Table Caption',
    'table_body': 'Table Body', 'table_header': 'Table Header',
    'code': 'Source Code', 'quote': 'Block Text', 'equation': 'M2W Equation',
    **{f'heading_{i}': f'Heading {i}' for i in range(1, 10)},
}


def merge(base: dict, override: dict) -> dict:
    result = copy.deepcopy(base)
    for key, value in override.items():
        replace_mode = isinstance(value,dict) and 'mode' in value
        result[key] = merge(result[key], value) if isinstance(value, dict) and isinstance(result.get(key), dict) and not replace_mode else copy.deepcopy(value)
    return result


def validate_spec(data: dict) -> None:
    schema = json.loads((ROOT / 'schema.json').read_text(encoding='utf-8'))
    _finite_numbers(data)
    errors = sorted(Draft202012Validator(schema).iter_errors(data), key=lambda error: str(list(error.path)))
    if errors:
        error = errors[0]
        raise ValueError('/' + '/'.join(map(str,error.path)) + ': ' + error.message)


def _finite_numbers(value):
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError('协议不允许 NaN 或 Infinity。')
    if isinstance(value, dict):
        for item in value.values():
            _finite_numbers(item)
    elif isinstance(value, list):
        for item in value:
            _finite_numbers(item)


def resolve_spec(data: dict) -> dict:
    validate_spec(data)
    result = merge(json.loads((ROOT / 'defaults.json').read_text(encoding='utf-8')), data)
    styles = result['styles']
    resolved = {}

    def resolve(name, trail=()):
        if name in resolved:
            return resolved[name]
        if name in trail:
            raise ValueError('样式循环继承：' + ' → '.join((*trail, name)))
        if name not in styles:
            raise ValueError('未知基础样式：' + name)
        style = copy.deepcopy(styles[name])
        parent = style.pop('based_on', None)
        resolved[name] = merge(resolve(parent, (*trail, name)), style) if parent else style
        return resolved[name]

    for name in styles:
        resolve(name)
    result['styles'] = resolved
    margins = result['page']['margins_mm']
    width, height = (210, 297) if result['page']['size'] == 'A4' else (215.9, 279.4)
    if result['page']['orientation'] == 'landscape':
        width, height = height, width
    if margins['left'] + margins['right'] >= width - 20 or margins['top'] + margins['bottom'] >= height - 20:
        raise ValueError('页边距没有留下足够的正文区域。')
    return result


def load_spec(path: Path) -> dict:
    if path.stat().st_size > 1_000_000:
        raise ValueError('格式协议超过 1 MB 限制。')
    return resolve_spec(json.loads(path.read_text(encoding='utf-8-sig')))


def dump_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
