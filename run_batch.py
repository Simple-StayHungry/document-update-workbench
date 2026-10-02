from pathlib import Path
import argparse,json,tempfile
from explain_core import analyze,unpack_targets,export_parallel,zip_output

def main():
    ap=argparse.ArgumentParser(description='说明性文件更新批处理')
    ap.add_argument('prospectus');ap.add_argument('targets_zip');ap.add_argument('--out',default='说明性文件更新结果.zip');ap.add_argument('--author',default='柒')
    args=ap.parse_args()
    with tempfile.TemporaryDirectory(prefix='explain-update-') as td:
        root=Path(td);targets=unpack_targets(args.targets_zip,root/'targets');result=analyze(args.prospectus,targets)
        # 高置信度正式来源与纯口吻更新默认执行；待定位项默认保留。
        export_parallel(result,{},root/'result',args.author)
        (root/'result'/'分析结果.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        zip_output(root/'result',args.out)
        print(json.dumps(result['summary'],ensure_ascii=False));print(args.out)
if __name__=='__main__':main()
