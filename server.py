from __future__ import annotations
import email,email.policy,hashlib,io,json,mimetypes,os,re,secrets,tempfile,threading,urllib.parse,webbrowser,zipfile
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from pathlib import Path
from explain_core import analyze,unpack_targets,export_parallel,zip_output,BUILD
from workbench.model import Document
ROOT=Path(__file__).resolve().parent
STATE={'result':None,'decisions':{},'work':None,'zip':None};TOKEN=secrets.token_hex(16);LOCK=threading.RLock()

def clean_work():
    old=STATE.get('work')
    if old:
        try:__import__('shutil').rmtree(old,ignore_errors=True)
        except Exception:pass
    td=tempfile.mkdtemp(prefix='explain-workbench-');STATE.update(result=None,decisions={},work=td,zip=None);return Path(td)

def _source_cover(doc:Document,name:str)->bool:
    """Narrow source identification: a real prospectus cover, not a body reference."""
    if '募集说明书' in Path(name).name and not re.search(r'说明性文件|情况说明|说明文件|核查',Path(name).name):
        return True
    head=[(b.text or '').strip() for b in doc.blocks[:12] if (b.text or '').strip()]
    if any(x=='募集说明书' for x in head[:8]):return True
    return any(len(x)<180 and x.endswith('募集说明书') and re.search(r'20\d{2}年|公司债券|企业债券|中期票据',x) for x in head[:8])

def prepare_materials(work:Path,inputs:list[tuple[str,bytes]]):
    """Flatten DOCX/ZIP uploads, dedupe bytes, identify exactly one prospectus, keep the rest as targets."""
    intake=work/'intake';intake.mkdir(parents=True,exist_ok=True)
    docs=[];ignored=[];seen=set();seq=0
    def add_doc(name,data,origin):
        nonlocal seq
        base=Path(name.replace('\\','/')).name
        if not base.lower().endswith('.docx') or base.startswith('~$'):return
        digest=hashlib.sha256(data).hexdigest()
        if digest in seen:return
        seen.add(digest);seq+=1;p=intake/f'{seq:03d}.docx';p.write_bytes(data);docs.append({'path':p,'name':base,'origin':origin,'hash':digest})
    for name,data in inputs:
        low=name.lower()
        if low.endswith('.docx'):add_doc(name,data,'loose')
        elif low.endswith('.zip'):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for zi in z.infolist():
                        if zi.is_dir() or not zi.filename.lower().endswith('.docx'):continue
                        if zi.file_size>80*1024*1024:raise ValueError('ZIP 内单个 DOCX 过大：'+Path(zi.filename).name)
                        add_doc(zi.filename,z.read(zi),Path(name).name)
            except zipfile.BadZipFile:raise ValueError('无法读取 ZIP：'+Path(name).name)
        else:ignored.append(Path(name).name)
    if not docs:raise ValueError('没有找到可处理的 DOCX。')
    sources=[]
    for rec in docs:
        try:d=Document(rec['path'],rec['name'])
        except Exception as e:raise ValueError(f"无法读取 DOCX：{rec['name']}（{e}）")
        rec['doc']=d
        if _source_cover(d,rec['name']):sources.append(rec)
    if not sources:
        raise ValueError('没有识别到募集说明书。请把当前募集说明书 DOCX 一起放入上传框。')
    if len(sources)>1:
        names='、'.join(x['name'] for x in sources[:4])
        raise ValueError('识别到多个募集说明书：'+names+'。请只保留当前版本。')
    source=sources[0];targets=[x for x in docs if x is not source]
    if not targets:raise ValueError('没有识别到说明性文件。请同时加入说明性文件 DOCX、ZIP 或文件夹。')
    return source,[(str(x['path']),x['name']) for x in targets],{'source_name':source['name'],'targets':len(targets),'uploaded_items':len(inputs),'docx_found':len(docs),'ignored':ignored}

class H(BaseHTTPRequestHandler):
    def log_message(self,fmt,*args):pass
    def send(self,data,status=200,kind='application/json; charset=utf-8'):
        if isinstance(data,(dict,list)):data=json.dumps(data,ensure_ascii=False).encode()
        elif isinstance(data,str):data=data.encode()
        self.send_response(status);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data)
    def jsonbody(self):
        n=int(self.headers.get('Content-Length','0'));return json.loads(self.rfile.read(n) or b'{}')
    def auth(self):return self.headers.get('X-Workbench-Token')==TOKEN
    def do_GET(self):
        u=urllib.parse.urlsplit(self.path);p=u.path
        if p=='/api/state':
            r=STATE['result'];return self.send({'build':BUILD,'result':r,'decisions':STATE['decisions'],'download':'/api/download' if STATE.get('zip') else None})
        if p=='/api/download':
            fp=STATE.get('zip')
            if not fp or not Path(fp).is_file():return self.send({'error':'尚未导出'},404)
            raw=Path(fp).read_bytes();self.send_response(200);self.send_header('Content-Type','application/zip');self.send_header('Content-Length',str(len(raw)));self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+urllib.parse.quote('说明性文件更新结果.zip'));self.end_headers();return self.wfile.write(raw)
        if p in ('/','/index.html','/app.js','/style.css'):
            name='index.html' if p=='/' else p[1:];fp=ROOT/'web'/name;kind=mimetypes.guess_type(name)[0] or 'text/plain'
            if name.endswith(('.html','.js','.css')):kind+='; charset=utf-8'
            return self.send(fp.read_bytes(),kind=kind)
        return self.send({'error':'not found'},404)
    def do_POST(self):
        try:
            if not self.auth():return self.send({'error':'仅允许本机页面操作'},403)
            if self.path=='/api/analyze':return self.upload_analyze()
            if self.path=='/api/decision':
                b=self.jsonbody();pid=b.get('id');d=b.get('decision')
                if d not in ('update','keep'):raise ValueError('决定无效')
                if not STATE['result']:raise ValueError('请先分析')
                proposal=next((p for p in STATE['result']['proposals'] if p['id']==pid),None)
                if proposal is None:raise ValueError('改动项不存在')
                if proposal.get('status')=='retained' and d=='update':
                    raise ValueError('该项没有可安全完整替换的来源，已按规则自动保留，无需人工处理。')
                rec={'decision':d}
                if 'source_index' in b and b.get('source_index') is not None:rec['source_index']=int(b['source_index'])
                STATE['decisions'][pid]=rec;return self.send({'ok':True,'decisions':STATE['decisions']})
            if self.path=='/api/export':
                if not STATE['result']:raise ValueError('请先分析')
                work=Path(STATE['work']);out=work/'output';out.mkdir(exist_ok=True)
                rep=export_parallel(STATE['result'],STATE['decisions'],out,'柒');zip_path=work/'说明性文件更新结果.zip';zip_output(out,zip_path);STATE['zip']=str(zip_path)
                return self.send({'ok':True,'report':rep,'download':'/api/download'})
            return self.send({'error':'not found'},404)
        except Exception as e:return self.send({'error':str(e)},400)
    def upload_analyze(self):
        n=int(self.headers.get('Content-Length','0'))
        if not 0<n<320*1024*1024:raise ValueError('上传体积无效')
        ctype=self.headers.get('Content-Type','')
        if 'multipart/form-data' not in ctype:raise ValueError('上传格式错误')
        raw=self.rfile.read(n);msg=email.message_from_bytes(('Content-Type: '+ctype+'\r\nMIME-Version: 1.0\r\n\r\n').encode()+raw,policy=email.policy.default)
        inputs=[];legacy={}
        for part in msg.iter_parts():
            name=part.get_filename();field=part.get_param('name',header='content-disposition')
            if not name:continue
            item=(name,part.get_payload(decode=True));inputs.append(item)
            if field in ('prospectus','targets'):legacy[field]=item
        work=clean_work()
        # Backward-compatible old two-field form remains accepted, but all new UI uploads use one mixed material box.
        if 'prospectus' in legacy and 'targets' in legacy and not any(part.get_param('name',header='content-disposition')=='files' for part in msg.iter_parts()):
            pros=work/'prospectus.docx';pros.write_bytes(legacy['prospectus'][1]);tz=work/'targets.zip';tz.write_bytes(legacy['targets'][1]);targets=unpack_targets(tz,work/'targets');intake={'source_name':legacy['prospectus'][0],'targets':len(targets),'uploaded_items':2,'docx_found':len(targets)+1,'ignored':[]}
        else:
            source,targets,intake=prepare_materials(work,inputs);pros=source['path']
        STATE['result']=analyze(pros,targets);STATE['result']['source']['name']=intake['source_name'];STATE['result']['intake']=intake
        return self.send({'ok':True,'result':STATE['result']})

def main():
    import argparse
    ap=argparse.ArgumentParser();ap.add_argument('--port',type=int,default=0);ap.add_argument('--no-browser',action='store_true');a=ap.parse_args()
    srv=ThreadingHTTPServer(('127.0.0.1',a.port),H);port=srv.server_address[1];url=f'http://127.0.0.1:{port}/?token={TOKEN}'
    print(url,flush=True)
    if not a.no_browser:threading.Timer(.5,lambda:webbrowser.open(url)).start()
    try:srv.serve_forever()
    finally:clean_work()
if __name__=='__main__':main()
