from docx.oxml.ns import qn
from .spec import STYLE_NAMES
from .word import ALIGN


def audit_format(doc, spec):
    """Check persisted style and layout properties, including Office output."""
    for role, definition in spec['styles'].items():
        name = STYLE_NAMES[role]
        if name not in doc.styles:
            raise ValueError('缺少样式：' + name)
        style = doc.styles[name]
        if abs(style.font.size.pt - definition['font']['size_pt']) > 0.01:
            raise ValueError('样式字号不符合协议：' + name)
        if style.element.rPr.rFonts.get(qn('w:eastAsia')) != definition['font']['east_asia']:
            raise ValueError('中文字体不符合协议：' + name)
        font = definition['font']
        if style.font.name != font['latin'] or style.font.bold != font['bold'] or style.font.italic != font['italic'] or str(style.font.color.rgb) != font['color'].upper():
            raise ValueError('字体属性不符合协议：' + name)
        paragraph = definition['paragraph']
        if style.paragraph_format.alignment != ALIGN[paragraph['alignment']]:
            raise ValueError('样式对齐不符合协议：' + name)
        for key in ('space_before','space_after'):
            if abs(getattr(style.paragraph_format,key).pt - paragraph[key+'_pt']) > 0.05:
                raise ValueError('段间距不符合协议：' + name)
    for section in doc.sections:
        for side, value in spec['page']['margins_mm'].items():
            if abs(getattr(section, side + '_margin').mm - value) > 0.05:
                raise ValueError('页边距不符合协议：' + side)
    expected = spec['table']['borders']
    for table in doc.tables:
        borders = table._tbl.tblPr.find(qn('w:tblBorders'))
        if borders is None:
            raise ValueError('表格缺少边框规则。')
        for edge in ('top','bottom','left','right','insideH','insideV'):
            node = borders.find(qn('w:' + edge))
            value = 'single' if expected == 'grid' or (expected == 'three_line' and edge in ('top','bottom')) else 'nil'
            if node is None or node.get(qn('w:val')) != value:
                raise ValueError('表格边框不符合协议：' + edge)
    from backends.stage1.markdown_ast import content_paragraphs, formula_literals
    for p in content_paragraphs(doc):
        literals = list(formula_literals(p))
        display = any(t.startswith('\\[') for t in literals) or bool(p._p.xpath('.//m:oMathPara'))
        math_ole = bool(p._p.xpath('.//*[local-name()="OLEObject" and namespace-uri()="urn:schemas-microsoft-com:office:office" and @ProgID="Equation.DSMT4"]'))
        display = display or (p.style.name == STYLE_NAMES['equation'] and math_ole)
        inline = bool(literals) or bool(p._p.xpath('.//m:oMath')) or math_ole
        if display or inline:
            spacing = p._p.xpath('./w:pPr/w:spacing')
            minimum = round(spec['equations']['display_min_pt' if display else 'inline_min_pt'] * 20)
            if not spacing or spacing[0].get(qn('w:lineRule')) != 'atLeast' or int(spacing[0].get(qn('w:line'),'0')) < minimum:
                raise ValueError('公式段落行距不符合协议。')
