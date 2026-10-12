from __future__ import annotations

import copy
import json
import threading
from pathlib import Path

from formatting.spec import resolve_spec
from formatting.word import build_project
from requirements.extract import extract
from requirements.parser import local_parse, validate_draft


class TemplateApi:
    """Mixin sharing the desktop task lock; AI state is per-window and ephemeral."""
    def _init_templates(self):
        self._requirement_blocks=[]
        self._chat_history=[]
        self._template_result={'revision':0,'draft':'','messages':[],'project':'','error':''}
        self._template_cancel=threading.Event()

    def template_snapshot(self):
        with self._lock:
            return copy.deepcopy(self._template_result)

    def _template_task(self, label, operation):
        with self._lock:
            if self._busy() or self._closing:
                return {'ok':False,'error':'请等待当前任务结束。'}
            self._template_cancel.clear()
            self._template_result['error']=''
            self._template_result['project']=''
            self._state.update(state='templating',status=label,error='',output='',warnings=[],progress=0)
            def work():
                try:
                    result=operation()
                    with self._lock:
                        if self._template_cancel.is_set() and not result.get('project'):
                            self._state.update(state='idle',status='已取消模板任务',progress=0)
                        else:
                            self._template_result.update(result)
                            self._template_result['revision']+=1
                            self._state.update(state='idle',status='模板任务完成，请检查规范草稿和问题清单',progress=100)
                except Exception as exc:
                    with self._lock:
                        if self._template_cancel.is_set():
                            self._state.update(state='idle',status='已取消模板任务',progress=0)
                        else:
                            self._template_result.update(error=str(exc),revision=self._template_result['revision']+1)
                            self._state.update(state='error',status='模板任务失败',error=str(exc),progress=0)
            self._worker=threading.Thread(target=work,name='template-task',daemon=False)
            self._worker.start()
            return {'ok':True}

    def import_requirements(self, path):
        def operation():
            source=Path(path)
            if source.suffix.lower()=='.json':
                if source.stat().st_size>1_000_000: raise ValueError('JSON 文件超过 1 MB。')
                data=json.loads(source.read_text(encoding='utf-8-sig'))
                draft=validate_draft(data) if 'spec' in data else {'spec':data,'evidence':[],'issues':[]}
                resolve_spec(draft['spec'])
                self._requirement_blocks=[]
            else:
                self._requirement_blocks=extract(source)
                draft=local_parse(self._requirement_blocks,source.stem)
            self._chat_history=[]
            return {'draft':json.dumps(draft,ensure_ascii=False,indent=2),'messages':[]}
        return self._template_task('正在读取格式要求…',operation)

    def template_chat(self, data):
        if not isinstance(data,dict): return {'ok':False,'error':'AI 参数必须是对象。'}
        def operation():
            from requirements.ai import chat
            current=validate_draft(json.loads(data['draft']))
            message=data.get('message','')
            if len(message)>10000: raise ValueError('对话消息过长。')
            # Include edits from the JSON editor in the next turn.
            context='当前人工编辑的草稿：\n'+json.dumps(current,ensure_ascii=False)+'\n用户消息：'+message
            if len(context)>80000: raise ValueError('当前草稿与消息超过 80000 字符，请精简草稿。')
            draft,history=chat(data['config'],self._requirement_blocks,context,self._chat_history,data.get('api_key',''))
            if self._template_cancel.is_set(): return {}
            self._chat_history=history
            messages=[*self._template_result['messages'], '你：'+message, 'AI：规范已更新，待处理问题 '+str(len(draft['issues']))+' 项。']
            return {'draft':json.dumps(draft,ensure_ascii=False,indent=2),'messages':messages}
        return self._template_task('正在向配置的 AI 服务发送格式要求…',operation)

    def validate_template(self, text):
        try:
            draft=validate_draft(json.loads(text))
            resolved=resolve_spec(draft['spec'])
            return {'ok':True,'issues':draft['issues'],'resolved':json.dumps(resolved,ensure_ascii=False,indent=2)}
        except (ValueError,TypeError) as exc:
            return {'ok':False,'error':str(exc)}

    def save_template_draft(self, text, output):
        with self._lock:
            if self._busy(): return {'ok':False,'error':'请等待当前任务结束。'}
            try:
                draft=validate_draft(json.loads(text))
                if not output: raise ValueError('请选择草稿保存路径。')
                path=Path(output)
                if path.suffix.lower()!='.json': raise ValueError('草稿文件须使用 .json 扩展名。')
                with path.open('x',encoding='utf-8') as stream:
                    json.dump(draft,stream,ensure_ascii=False,indent=2)
                return {'ok':True,'path':str(path)}
            except (ValueError,OSError) as exc:
                return {'ok':False,'error':str(exc)}

    def generate_template(self, text, output, reviewed=False):
        if not reviewed:
            return {'ok':False,'error':'请核对协议及默认值，并勾选确认。'}
        def operation():
            if not isinstance(output,str) or not output.strip():
                raise ValueError('请填写新的模板目录。')
            draft=validate_draft(json.loads(text))
            if draft['issues']:
                raise ValueError('仍有未解决的问题。请修正规范并从 issues 中移除已经处理的条目。')
            project=build_project(draft['spec'],Path(output),report=draft,cancelled=self._template_cancel.is_set)
            return {'project':str(project),'draft':text}
        return self._template_task('正在生成模板和排版样张…',operation)

    def cancel_template(self):
        with self._lock:
            if self._state['state']!='templating': return {'ok':False,'error':'没有运行中的模板任务。'}
            # urllib has no cooperative in-flight cancellation; wait at most the configured timeout.
            self._template_cancel.set()
            self._state['status']='已请求取消，等待当前请求返回或超时…'
            return {'ok':True}
