from __future__ import annotations

import json
import re
from pathlib import Path

from formatting.spec import resolve_spec

SIZES = {'初号':42,'小初':36,'一号':26,'小一':24,'二号':22,'小二':18,'三号':16,'小三':15,'四号':14,'小四':12,'五号':10.5,'小五':9,'六号':7.5}
ROLES = {'正文':'body','一级标题':'heading_1','二级标题':'heading_2','三级标题':'heading_3','表正文':'table_body','表头':'table_header','图题':'figure_caption','图注':'figure_caption','表题':'table_caption','表注':'table_caption','代码':'code'}


def local_parse(blocks, name='导入的格式规范'):
    spec = {'schema_version':'1.0','name':name,'styles':{}}
    evidence, issues, assigned = [], [], {}

    def assign(pointer,value,block):
        if pointer in assigned and assigned[pointer] != value:
            issues.append({'kind':'conflict','message':f'{pointer} 出现不同要求，请编辑后解决。','source':block['source']})
            return
        assigned[pointer]=value
        target=spec
        parts=pointer.strip('/').split('/')
        for key in parts[:-1]: target=target.setdefault(key,{})
        target[parts[-1]]=value
        evidence.append({'path':pointer,'source':block['source'],'quote':block['text']})

    for block in blocks:
        line=block['text']
        rest=line
        role = next((value for label,value in ROLES.items() if line.startswith(label)),None)
        if role:
            label=next(label for label,value in ROLES.items() if value==role and line.startswith(label))
            rest=rest[len(label):]
            base='/styles/'+role
            font=re.search(r'(仿宋|黑体|宋体|楷体|微软雅黑)',rest)
            if font:
                assign(base+'/font/east_asia',font[0],block); rest=rest.replace(font[0],'',1)
            size=re.search('|'.join(sorted(SIZES,key=len,reverse=True)),rest)
            if size:
                assign(base+'/font/size_pt',SIZES[size[0]],block); rest=rest.replace(size[0],'',1)
            latin=re.search(r'Times New Roman|Arial|Consolas',rest,re.I)
            if latin:
                assign(base+'/font/latin',latin[0],block); rest=rest.replace(latin[0],'',1)
            for pattern,key in [(r'字号\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)','/font/size_pt'),(r'段前\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)','/paragraph/space_before_pt'),(r'段后\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)','/paragraph/space_after_pt')]:
                m=re.search(pattern,rest,re.I)
                if m: assign(base+key,float(m[1]),block); rest=rest.replace(m[0],'',1)
            m=re.search(r'(\d+(?:\.\d+)?)\s*倍行距',rest)
            if m:
                assign(base+'/paragraph/line_spacing',{'mode':'multiple','value':float(m[1])},block); rest=rest.replace(m[0],'',1)
            m=re.search(r'固定(?:行距|值)\s*(\d+(?:\.\d+)?)\s*(?:pt|磅)',rest)
            if m:
                assign(base+'/paragraph/line_spacing',{'mode':'exact','value':float(m[1])},block); rest=rest.replace(m[0],'',1)
            m=re.search(r'首行缩进\s*([0-9]+|两|二)\s*(?:个)?字符',rest)
            if m:
                assign(base+'/paragraph/first_line_indent',{'value':2 if m[1] in ('两','二') else int(m[1]),'unit':'char'},block); rest=rest.replace(m[0],'',1)
            for token,key,value in [('不加粗','/font/bold',False),('加粗','/font/bold',True),('居中','/paragraph/alignment','center'),('两端对齐','/paragraph/alignment','justify'),('左对齐','/paragraph/alignment','left')]:
                if token in rest: assign(base+key,value,block); rest=rest.replace(token,'',1)
        elif re.fullmatch(r'(?:纸张[：: ]*)?A4[。.]?',line,re.I):
            assign('/page/size','A4',block); rest=''
        elif line.startswith('页边距'):
            rest=line[len('页边距'):]
            for side,key in [('上','top'),('下','bottom'),('左','left'),('右','right')]:
                m=re.search(side+r'\s*(\d+(?:\.\d+)?)\s*(mm|毫米|cm|厘米)',rest,re.I)
                if m:
                    assign('/page/margins_mm/'+key,float(m[1])*(10 if m[2].lower() in ('cm','厘米') else 1),block); rest=rest.replace(m[0],'',1)
        elif line.startswith('表格'):
            for text,value in [('三线表','three_line'),('网格表','grid'),('无边框','none')]:
                if text in line:
                    assign('/table/borders',value,block); rest=line.replace('表格','',1).replace(text,'',1); break
        if re.sub(r'[\s，,。.;；:：、]+','',rest):
            issues.append({'kind':'unresolved','message':'未完整识别：'+rest.strip(),'source':block['source']})
    resolve_spec(spec)
    return {'spec':spec,'evidence':evidence,'issues':issues}


def validate_draft(data, blocks=None):
    if not isinstance(data,dict) or set(data) != {'spec','evidence','issues'}:
        raise ValueError('草稿必须且只能包含 spec、evidence、issues。')
    resolve_spec(data['spec'])
    if not isinstance(data['evidence'],list) or not isinstance(data['issues'],list):
        raise ValueError('evidence 和 issues 必须是数组。')
    if len(data['evidence']) > 1000 or len(data['issues']) > 1000:
        raise ValueError('草稿条目过多。')
    for item in data['issues']:
        if not isinstance(item,dict) or set(item) != {'kind','message','source'} or item['kind'] not in ('conflict','unresolved','unsupported') or not all(isinstance(v,str) for v in item.values()):
            raise ValueError('问题条目格式不正确。')
    sources = {b['source']:b['text'] for b in blocks} if blocks else None
    for item in data['evidence']:
        if not isinstance(item,dict) or set(item) != {'path','source','quote'} or not all(isinstance(v,str) and v for v in item.values()):
            raise ValueError('来源条目格式不正确。')
        if not item['path'].startswith('/'):
            raise ValueError('来源 path 必须是 JSON Pointer。')
        node=data['spec']
        try:
            for key in item['path'][1:].split('/'):
                node=node[key.replace('~1','/').replace('~0','~')]
        except (KeyError,TypeError):
            raise ValueError('来源指向不存在的协议字段。') from None
        if sources is not None and (item['source'] not in sources or item['quote'] not in sources[item['source']]):
            raise ValueError('AI 返回了无法在原文找到的来源引用。')
    return data


def decode_draft(text, blocks=None):
    text=text.strip()
    if text.startswith('```'):
        match=re.fullmatch(r'```(?:json)?\s*\n(.*?)\n```',text,re.S)
        if not match: raise ValueError('AI 返回的代码围栏不完整。')
        text=match[1]
    return validate_draft(json.loads(text),blocks)
