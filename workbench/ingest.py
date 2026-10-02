"""Safe, recursive upload ingestion with Unicode ZIP names and content-based roles."""
from __future__ import annotations
import copy,io,zipfile,re,json,hashlib
from pathlib import Path,PurePosixPath
from .model import Document
from .docxio import sha

SKIP={'.venv','venv','node_modules','__pycache__','__MACOSX','.git','.DS_Store'}
MAX_TOTAL=600*1024**2;MAX_FILE=180*1024**2;MAX_ENTRIES=30000;MAX_DEPTH=4

def source_exclusion_reason(record):
    """Conservative provenance gate, including ZIP folders and deduplicated aliases."""
    names=[record.get('name',''),record.get('origin',''),*record.get('aliases',[])]
    for value in names:
        value=str(value).replace('\\','/')
        if re.search(r'老师|教师对照|teacher(?:s)?(?:[ _./-]|$)',value,re.I):
            return '老师或教师对照材料不能作为事实来源。'
        if (any(x in value.lower() for x in (
                '/out/','/exports/','自动更新修订版','引擎v','盲测修订版',
                '核查分析文件_修订版','check_v','历史问题输出','历史自动输出',
                '错误输出','程序输出','v19开发残留','历史参考'))
                or re.match(r'^\d+(?:-\d+)+_修订版',PurePosixPath(value).name)):
            return '历史程序输出或开发对照材料不能作为正式来源。'
    if record.get('profile',{}).get('kind')=='workpaper':
        return '核查底稿只能作为待更新或对照材料，不能作为新事实来源。'
    return ''

def validate_source_record(record):
    reason=source_exclusion_reason(record)
    if reason:raise ValueError(record.get('name','材料')+'：'+reason)

def validate_material_hash(record):
    path=Path(record['path'])
    if not path.is_file() or sha(path.read_bytes())!=record['id']:
        raise ValueError(record.get('name','材料')+'：材料内容已经变化或丢失，请重新导入后分析。')

def decode_name(info):
    n=info.filename
    if not info.flag_bits&0x800:
        for enc in ('utf-8','gb18030'):
            try:
                candidate=n.encode('cp437').decode(enc)
                if enc=='utf-8' or re.search('[\u4e00-\u9fff]',candidate):n=candidate;break
            except UnicodeError:pass
    return n.replace('\\','/')

class Importer:
    def __init__(self,directory,existing=()):
        self.directory=Path(directory);self.directory.mkdir(parents=True,exist_ok=True)
        self.files=copy.deepcopy(list(existing));self.byhash={f['id']:f for f in self.files};self.inventory=[];self.total=0;self.count=0
    def add(self,name,data,forced_role=None,depth=0):
        name=name.replace('\\','/');parts=PurePosixPath(name).parts
        self.count+=1
        if self.count>MAX_ENTRIES:raise ValueError('压缩包条目过多。')
        if '..' in parts or name.startswith('/') or re.match(r'^[A-Za-z]:',name):raise ValueError('检测到不安全的压缩包路径，已停止导入。')
        if any(p in SKIP or p.startswith(('._','.~','~$')) for p in parts):self.inventory.append({'name':name,'role':'忽略','reason':'运行环境、系统缓存或临时文件'});return
        ext=Path(name).suffix.lower()
        if ext=='.zip':
            if depth>=MAX_DEPTH:raise ValueError('压缩包嵌套层级超过安全上限。')
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for info in z.infolist():
                        if info.is_dir():continue
                        sub=decode_name(info)
                        if '..' in PurePosixPath(sub).parts or sub.startswith('/') or re.match(r'^[A-Za-z]:',sub):raise ValueError('检测到不安全的压缩包路径，已停止导入。')
                        if any(p in SKIP or p.startswith(('._','.~','~$')) for p in PurePosixPath(sub).parts):
                            self.count+=1
                            if self.count>MAX_ENTRIES:raise ValueError('压缩包条目过多。')
                            continue
                        if info.file_size>MAX_FILE:raise ValueError('单个解压文件过大：'+sub)
                        if info.file_size>50*1024**2 and info.file_size/max(1,info.compress_size)>500:raise ValueError('检测到异常压缩比，已停止导入。')
                        self.add(name+'/'+sub,z.read(info),forced_role,depth+1)
            except zipfile.BadZipFile:raise ValueError('压缩包损坏或不是有效 ZIP。')
            return
        if ext!='.docx':
            self.inventory.append({'name':name,'role':'参考或忽略','reason':'PDF 不参与自动回填，请提供同版 Word。' if ext=='.pdf' else '非核查 Word；不执行上传的代码，也不读取其运行环境。'})
            return
        self.total+=len(data)
        if self.total>MAX_TOTAL:raise ValueError('本次材料体积超过安全上限。')
        fid=sha(data);basename=PurePosixPath(name).name
        if fid in self.byhash:
            f=self.byhash[fid];f.setdefault('aliases',[]).append(name)
            reason=source_exclusion_reason(f)
            if reason and f['role']=='source':f['role']='reference';f['reason']=reason
            self.inventory.append({'name':name,'role':'去重','reason':'与 '+f['name']+' 字节完全相同'});return
        path=self.directory/(fid+'.docx');path.write_bytes(data)
        try:
            d=Document(path,basename);head='\n'.join(b.text for b in d.blocks[:30])
            provenance_exclusion=source_exclusion_reason({'name':basename,'origin':name})
            exclusion=provenance_exclusion or source_exclusion_reason({'profile':d.profile})
            if provenance_exclusion:role='reference';reason=provenance_exclusion
            elif forced_role=='source' and exclusion:role='reference';reason=exclusion
            elif forced_role in ('source','target','reference'):role=forced_role;reason='按上传区域指定用途'
            elif '募集说明书' in head[:1300] and '核查记录' not in head[:500] and '核查意见' not in head[:500]:role='source';reason='封面识别为募集说明书'
            elif any(x in name for x in ['/out/',' / out/','自动更新修订版','引擎v','盲测修订版','核查分析文件_修订版','check_v'] ) or re.match(r'^\d+(?:-\d+)+_修订版',basename):role='ignored';reason='历史程序输出，不作为最新版原始证据'
            else:role='candidate';reason='核查文稿，待按发行人和章节选择工作底稿'
            file={'id':fid,'name':basename,'path':str(path),'role':role,'reason':reason,'origin':name,'forced_role':bool(forced_role),'role_locked':bool(forced_role),'size':len(data),'profile':d.profile,'aliases':[]}
            self.files.append(file);self.byhash[fid]=file
            self.inventory.append({'name':name,'role':role,'reason':reason,'sha256':fid})
        except Exception as exc:
            self.inventory.append({'name':name,'role':'读取失败','reason':str(exc)})
    def recommend(self):
        sources=[f for f in self.files if f['role']=='source'];issuers={f['profile']['issuer'] for f in sources if f['profile']['issuer']}
        if len(issuers)!=1:return
        issuer=next(iter(issuers));groups={}
        has_workpapers=any(f['profile'].get('kind')=='workpaper' and f['role'] not in ('reference','ignored') and f['profile']['issuer']==issuer for f in self.files)
        if has_workpapers:
            for f in self.files:
                if f['role'] in ('candidate','target') and not f.get('forced_role') and not f.get('role_locked') and not source_exclusion_reason(f) and f['profile'].get('kind')=='opinion' and f['profile']['issuer']==issuer:
                    f['role']='source';f['reason']='同项目已有核查分析底稿，主承销商核查意见作为补充依据，不另作待更新稿'
        selected={(f['profile']['issuer'],f['profile']['chapter']) for f in self.files if f['role']=='target'}
        for f in self.files:
            if f['role']!='candidate' or f.get('role_locked') or f['profile']['issuer']!=issuer:continue
            key=(issuer,f['profile']['chapter'] or f['name'])
            groups.setdefault(key,[]).append(f)
        for key,fs in groups.items():
            if key in selected:continue
            def rank(f):
                score=0
                if '原版' in f['name']:score+=100
                if '【待更新】' in f['origin']:score+=30
                if '修订版' in f['name']:score-=40
                if not f['profile']['historical_revisions']:score+=15
                return score
            f=max(fs,key=rank);f['role']='target';f['reason']='同发行人、同章节候选中优先推荐明确原版；可自行更换，未猜测最新版'
