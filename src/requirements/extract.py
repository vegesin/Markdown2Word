from pathlib import Path
import zipfile

from lxml import etree


def extract(path: Path) -> list[dict]:
    path = path.expanduser().resolve()
    if path.stat().st_size > 10_000_000:
        raise ValueError('要求文件超过 10 MB 限制。')
    suffix = path.suffix.lower()
    if suffix in ('.md','.markdown','.txt'):
        blocks = [{'source':f'行 {i}', 'text':line.strip()} for i,line in enumerate(path.read_text(encoding='utf-8-sig').splitlines(),1) if line.strip()]
    elif suffix == '.docx':
        # Read XML only. Never launch Word, macros, or embedded relationships.
        try:
            with zipfile.ZipFile(path) as package:
                if len(package.infolist()) > 2000 or sum(i.file_size for i in package.infolist()) > 50_000_000:
                    raise ValueError('Word 要求文件解压后过大。')
                xml = package.read('word/document.xml')
            root = etree.fromstring(xml,etree.XMLParser(resolve_entities=False,no_network=True))
        except (zipfile.BadZipFile,KeyError,etree.XMLSyntaxError):
            raise ValueError('无法读取 DOCX，文件可能损坏或不是标准 Word 文档。') from None
        ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        blocks = []
        body=root.find('w:body',ns)
        if body is None:
            raise ValueError('DOCX 中没有可读取的正文。')
        for i,node in enumerate(body,1):
            if node.tag.endswith('}tbl'):
                for j,row in enumerate(node.findall('w:tr',ns),1):
                    cells = [''.join(cell.xpath('.//w:t/text()',namespaces=ns)) for cell in row.findall('w:tc',ns)]
                    blocks.append({'source':f'块 {i} 表格行 {j}','text':'；'.join(cells)})
            elif node.tag.endswith('}p'):
                text = ''.join(node.xpath('.//w:t/text()',namespaces=ns))
                if text.strip(): blocks.append({'source':f'段落块 {i}','text':text})
    else:
        raise ValueError('本版支持 TXT、Markdown 和 DOCX 格式要求；PDF、扫描件和旧 DOC 请先转成文本或 DOCX。')
    if not blocks or sum(len(b['text']) for b in blocks) > 80000:
        raise ValueError('要求内容为空或超过 80000 字符，请拆分后导入。')
    return blocks
