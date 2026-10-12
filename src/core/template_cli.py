"""Independent template operations; normal conversion never calls AI."""
import argparse
import json
from pathlib import Path

from formatting.spec import dump_json, load_spec
from formatting.word import build_project
from requirements.extract import extract
from requirements.parser import local_parse, validate_draft


def main():
    parser=argparse.ArgumentParser(description='Create and validate Word formatting templates')
    sub=parser.add_subparsers(dest='command',required=True)
    parse=sub.add_parser('parse')
    parse.add_argument('source',type=Path)
    parse.add_argument('output',type=Path)
    parse.add_argument('--ai-config',type=Path)
    build=sub.add_parser('build')
    build.add_argument('spec',type=Path)
    build.add_argument('output',type=Path)
    check=sub.add_parser('validate')
    check.add_argument('spec',type=Path)
    args=parser.parse_args()
    try:
        if args.command=='parse':
            if args.output.exists(): raise ValueError('输出草稿已存在。')
            blocks=extract(args.source)
            if args.ai_config:
                from requirements.ai import chat
                config=json.loads(args.ai_config.read_text(encoding='utf-8-sig'))
                draft,_=chat(config,blocks,'请解析全部格式要求并生成规范草稿。')
            else:
                draft=local_parse(blocks,args.source.stem)
            with args.output.open('x',encoding='utf-8') as output:
                json.dump(draft,output,ensure_ascii=False,indent=2)
            print(f'草稿已保存；待处理问题 {len(draft["issues"])} 项。')
        else:
            raw=json.loads(args.spec.read_text(encoding='utf-8-sig'))
            report=None
            if 'spec' in raw:
                validate_draft(raw)
                if raw['issues']: raise ValueError('草稿仍有未解决问题，请编辑并解决 issues 后再生成。')
                report=raw
                raw=raw['spec']
            from formatting.spec import resolve_spec
            resolve_spec(raw)
            if args.command=='build': print(build_project(raw,args.output,report=report))
            else: print('协议校验通过。')
        return 0
    except (ValueError,OSError) as exc:
        print('错误：'+str(exc))
        return 1


if __name__=='__main__':
    raise SystemExit(main())
