from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from .parser import decode_draft


def system_prompt():
    from formatting.spec import ROOT
    return Path(__file__).with_name('prompt.md').read_text(encoding='utf-8').replace('{{SCHEMA}}',(ROOT/'schema.json').read_text(encoding='utf-8'))


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise ValueError('AI 服务重定向被拒绝，请填写最终 API 地址。')


def chat(config, blocks, message, history=(), api_key='', opener=None):
    """Chat Completions-compatible adapter; no agent tools or local execution."""
    endpoint=str(config.get('base_url','')).strip().rstrip('/')
    parsed=urllib.parse.urlsplit(endpoint)
    if parsed.scheme != 'https' and not (parsed.scheme=='http' and parsed.hostname in ('localhost','127.0.0.1','::1')):
        raise ValueError('AI 地址须使用 HTTPS；本机服务可使用 HTTP。')
    if not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise ValueError('AI 地址不得包含凭据、查询参数或片段。')
    model=str(config.get('model','')).strip()
    if not model or len(model)>200: raise ValueError('请填写模型名称。')
    if not isinstance(message,str) or not message.strip() or len(message)>80000:
        raise ValueError('对话上下文须为 1～80000 字符。')
    timeout=int(config.get('timeout',60))
    if not 5 <= timeout <= 180: raise ValueError('AI 超时应为 5～180 秒。')
    key=api_key or os.environ.get(str(config.get('api_key_env','M2W_AI_API_KEY')),'')
    if not isinstance(key,str) or len(key)>4096 or any(ord(c)<33 or ord(c)>126 for c in key):
        raise ValueError('API Key 必须是无空白的 ASCII 文本。')
    messages=[{'role':'system','content':system_prompt()}, {'role':'user','content':'格式要求原文块：\n'+json.dumps(blocks,ensure_ascii=False)}]
    for item in history:
        if item.get('role') not in ('user','assistant') or not isinstance(item.get('content'),str):
            raise ValueError('对话历史格式不正确。')
        messages.append({'role':item['role'],'content':item['content']})
    messages.append({'role':'user','content':message})
    if len(json.dumps(messages,ensure_ascii=False))>180000:
        raise ValueError('会话过长，请重新导入要求开始新会话。')
    payload={'model':model,'messages':messages,'stream':False}
    if config.get('json_mode',False): payload['response_format']={'type':'json_object'}
    headers={'Content-Type':'application/json'}
    if key: headers['Authorization']='Bearer '+key
    request=urllib.request.Request(endpoint+'/chat/completions',data=json.dumps(payload).encode(),headers=headers,method='POST')
    opener=opener or urllib.request.build_opener(NoRedirect())
    try:
        with opener.open(request,timeout=timeout) as response:
            raw=response.read(2_000_001)
        if len(raw)>2_000_000: raise ValueError('AI 返回内容超过 2 MB。')
        result=json.loads(raw)
        choice=result['choices'][0]
        if choice.get('finish_reason') not in (None,'stop'):
            raise ValueError('AI 输出未正常结束，请缩短要求或调整服务输出限制。')
        content=choice['message']['content']
        draft=decode_draft(content,blocks)
        return draft, [*history,{'role':'user','content':message},{'role':'assistant','content':content}]
    except urllib.error.HTTPError as exc:
        raise ValueError(f'AI 请求失败（HTTP {exc.code}），请检查地址、模型、凭据及服务额度。') from None
    except (urllib.error.URLError, TimeoutError, OSError):
        raise ValueError('AI 连接失败或超时，请检查服务地址和网络。') from None
    except (KeyError,IndexError,TypeError,json.JSONDecodeError):
        raise ValueError('AI 返回格式无效，需要 Chat Completions 响应及标准 JSON 草稿。') from None
