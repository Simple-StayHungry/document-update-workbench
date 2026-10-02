"""Document structure, independent of issuer names, matter IDs and chapter order."""
from __future__ import annotations
from dataclasses import dataclass,field
from pathlib import Path
import re,copy,collections,html
from .docxio import Package,NS,w,text,xml,local,sha

MATTER=re.compile(r'^\s*(\d{1,2}(?:\s*[-－—–]\s*\d{1,2}){1,6})\s*(.{3,220})$')
COMPANY=re.compile(r'[\u4e00-\u9fffA-Za-z（）()·]{3,55}(?:有限责任公司|股份有限公司|有限公司)')
SECTION=re.compile(r'^(?:[一二三四五六七八九十\d]+[、.．]|[（(][一二三四五六七八九十\d]+[）)])?\s*((?:核查|调查|分析)(?:程序|方法|情况|过程|结论|意见|结果)(?:及分析)?)\s*[:：]?$')
PROTECTED=re.compile(r'^(?:\d+[、.．]\s*)?(?:经核查|经分析|主承销商认为|项目组认为|本项目组|我们认为|经企查查|企查查查询|信用中国查询|通过查询|经查询|根据.*?(?:查询结果|访谈|核查))')
EVIDENCE=re.compile(r'企查查|信用中国|执行信息|裁判文书|公示系统|查询结果|网站查询|搜索结果|征信报告')

@dataclass
class Block:
    index:int;el:object;kind:str;text:str;level:int|None=None;heading:bool=False
    path:list=field(default_factory=list);matter:str='';section:str='other';protected:bool=False;reason:str='';images:int=0;is_toc:bool=False
    @property
    def title(self):return self.text[:180]

@dataclass
class Unit:
    id:str;start:int;end:int;blocks:list;heading:str;path:list;matter:str;title:str='';kind:str='material'
    @property
    def text(self):return '\n'.join(b.text for b in self.blocks if b.text)
    @property
    def schema(self):
        out=[]
        for b in self.blocks:
            if b.kind=='table':
                for tr in b.el.findall(w('tr'))[:2]:out.append('|'.join(text(tc) for tc in tr.findall(w('tc'))))
        return '\n'.join(out)
    @property
    def has_table(self):return any(b.kind=='table' for b in self.blocks)
    @property
    def locator(self):return ' > '.join(self.path[-4:]) or self.heading or ('正文片段 '+str(self.start+1))

def strip_lead(s):
    s=re.sub(r'^\s*(?:第[一二三四五六七八九十百\d]+[章节部分]|[（(][一二三四五六七八九十百\d]+[）)]|[一二三四五六七八九十百]+[、.．]|\d+(?:[.．]\d+)*[、.．)）])\s*','',s)
    return s.strip()

def is_short_numbered_title(value):
    # Retain existing compact labels; extend only to nominal analysis labels.
    # A short punctuation-free statement is not automatically a heading.
    match=re.fullmatch(r'[①-⑳]\s*([\u4e00-\u9fff]{3,20})',value.strip())
    if not match:return False
    label=match[1]
    return len(label)<=5 or (label.endswith('分析') and not re.search(r'已|将|应当|必须|需要|正在|可以',label))

def is_bold_numbered_title(value,el):
    """A numbered/bulleted, explicitly bold standalone nominal label."""
    label=value.strip()
    if not re.fullmatch(r'[\u4e00-\u9fff]{3,32}',label):return False
    if re.search(r'已|将|应当|必须|需要|正在|可以|建立|开展|提供|负责|进行|包括|属于',label):return False
    if not re.search(r'(?:分析|情况|方案|流程|方式|方法|规则|原则|制度|机制|指标|范围|措施|目标|公式|减排量|节能量)$',label):return False
    num=el.find('./'+w('pPr')+'/'+w('numPr')+'/'+w('numId'))
    if num is None or num.get(w('val')) in (None,'','0'):return False
    runs=[r for r in el.iter(w('r')) if any((t.text or '').strip() for t in r.iter(w('t')))]
    if not runs:return False
    for run in runs:
        bold=run.find('./'+w('rPr')+'/'+w('b'))
        if bold is None or bold.get(w('val'),'1').lower() in ('0','false','off'):return False
    return True

def heading_level(s,el,styles):
    if len(s)>160 or not s or s.endswith(('。','；',';')):return None
    if re.match(r'^\s*(?:单位\s*[:：]|表\s*[:：]|注[一二三四五\d]*[:：])',s):return None
    # Numbered headings have document-independent lexical boundaries.
    if re.match(r'^第[一二三四五六七八九十百\d]+[章节]',s):return 0
    if re.match(r'^[一二三四五六七八九十百]+[、．.]',s) and len(s)<90:return 1
    if re.match(r'^[（(][一二三四五六七八九十百]+[）)]',s) and len(s)<90:return 2
    if re.match(r'^\d+[、．.]\s*[^\d]',s) and len(s)<80:return 3
    if re.match(r'^[（(]\d+[）)]',s) and len(s)<80:return 4
    if re.match(r'^\d+[)）]',s) and len(s)<70:return 5
    if is_short_numbered_title(s):return 6
    if is_bold_numbered_title(s,el):return 7
    pr=el.find(w('pPr'));ol=pr.find(w('outlineLvl')) if pr is not None else None
    if ol is not None:
        val=int(ol.get(w('val'),'9'))
        if val<9:return val
    ps=pr.find(w('pStyle')) if pr is not None else None
    sid=ps.get(w('val')) if ps is not None else None
    seen=set()
    while sid in styles and sid not in seen:
        seen.add(sid);style=styles[sid];v=style.find('./'+w('pPr')+'/'+w('outlineLvl'))
        if v is not None and int(v.get(w('val'),'9'))<9:return int(v.get(w('val')))
        nm=style.find(w('name'));name=nm.get(w('val'),'') if nm is not None else ''
        m=re.search(r'(?:heading|标题)\s*([1-9])',name,re.I)
        if m:return int(m.group(1))-1
        based=style.find(w('basedOn'));sid=based.get(w('val')) if based is not None else None
    if re.match(r'^第[一二三四五六七八九十百\d]+[章节]',s):return 0
    if re.match(r'^[一二三四五六七八九十百]+[、．.]',s) and len(s)<90:return 1
    if re.match(r'^[（(][一二三四五六七八九十百]+[）)]',s) and len(s)<90:return 2
    if re.match(r'^\d+[、．.]\s*[^\d]',s) and len(s)<80:return 3
    if re.match(r'^[（(]\d+[）)]',s) and len(s)<80:return 4
    if re.match(r'^\d+[)）]',s) and len(s)<70:return 4
    return None

class Document:
    def __init__(self,path,name=None):
        self.path=Path(path);self.pkg=Package(path).baseline();self.name=name or Path(path).name;self.hash=self.pkg.hash
        self.blocks=[];self.matters=[];self.units=[];self.source_units=[];self.warnings=[]
        self.parents={c:p for p in self.pkg.root.iter() for c in p}
        st=xml(self.pkg.entries['word/styles.xml']) if 'word/styles.xml' in self.pkg.entries else []
        styles={s.get(w('styleId')):s for s in st}
        def walk(parent):
            for e in parent:
                if e.tag==w('p'):yield e
                elif e.tag==w('tbl'):
                    # Some institutions place the whole report in layout tables.
                    tt=text(e)
                    if ('核查程序' in tt or '调查程序' in tt) and len(re.findall(r'\d+(?:-\d+){1,5}\s*[^\d]',tt))>=2:
                        for row in e.findall(w('tr')):
                            for cell in row.findall(w('tc')):yield from walk(cell)
                    else:yield e
                elif e.tag in (w('sdt'),w('sdtContent'),w('customXml')):
                    yield from walk(e)
        for i,e in enumerate(walk(self.pkg.body)):
            tx=text(e).strip();kind='table' if e.tag==w('tbl') else 'paragraph'
            if kind=='table':tx='\n'.join(' | '.join(text(tc).strip() for tc in tr.findall(w('tc'))) for tr in e.findall(w('tr')))
            level=heading_level(tx,e,styles) if kind=='paragraph' else None
            image_count=len(e.findall('.//'+w('drawing')))+len(e.findall('.//'+w('pict')))
            toc=any('TOC' in (x.text or '') for x in e.iter(w('instrText'))) or any((x.get(w('val')) or '').startswith('TOC') for x in e.iter(w('pStyle')))
            if kind=='table' and '调查事项' in tx and not ('核查程序' in tx or '核查情况' in tx):toc=True
            # Typical TOC field cached result lines end in a tab + page number.
            if re.search(r'\t\s*\d+\s*$',tx) and len(tx)<200:toc=True
            self.blocks.append(Block(i,e,kind,tx,level,level is not None,images=image_count,is_toc=toc))
        quote_stack=[]
        pairs={'“':'”','「':'」','『':'』'}
        for b in self.blocks:
            inherited=bool(quote_stack)
            for ch in b.text:
                if ch in pairs:quote_stack.append(pairs[ch])
                elif quote_stack and ch==quote_stack[-1]:quote_stack.pop()
            b.quoted=inherited or bool(quote_stack)
        self.profile=self._profile()
        self._structure()
        from .docxio import comment_covered_elements,complex_reference_reason
        covered=comment_covered_elements(self.pkg.root)
        self.safety_issues=[]
        for b in self.blocks:
            reason=('原有人工批注范围保留，不移动或删除人工批注。' if id(b.el) in covered else complex_reference_reason([b.el]))
            if not reason and any(str(e.tag).startswith('{http://schemas.openxmlformats.org/officeDocument/2006/math}') for e in b.el.iter()):
                b.protected=True;b.reason='原生公式由完整对象来源通道处理，不作为普通文字切片。'
            if reason:
                if '批注' in reason or '公式' in reason:b.protected=True;b.reason=reason
                self.safety_issues.append({'rule':'protected-object','severity':'warning','title':'复杂对象原样保留','block':b.index,'locator':' > '.join(b.path),'detail':reason})
        self._units()
        from .assurance import glossary_periods
        self.profile['reporting_periods']=glossary_periods(self)
        from .precision import augment_sources
        augment_sources(self)
    def _profile(self):
        head='\n'.join(b.text for b in self.blocks[:35])
        names=COMPANY.findall(head)
        issuer=''
        # Prefer the issuer directly before the bond year. Never infer from underwriter name.
        m=re.search(r'([\u4e00-\u9fffA-Za-z（）()·]{3,55}(?:有限责任公司|股份有限公司|有限公司))\s*\n?\s*20\d{2}年',head)
        if m:issuer=m.group(1)
        if not issuer:
            issuer=next((n for n in names if not re.search(r'证券|会计|律师|信用评级|资信评估',n)),names[0] if names else '')
        issuer=re.sub(r'^(?:关于|发行人|公司名称[:：]?)','',issuer)
        bond=''
        for i,b in enumerate(self.blocks[:20]):
            s=b.text.replace('\n','')
            m=re.search(r'20\d{2}年[^。\n]{0,110}?(?:公司债券|企业债券|中期票据|短期融资券)',s)
            if m:bond=m.group(0);break
        if not bond:
            cover=''.join(b.text.replace('\n','') for b in self.blocks[:10])
            mt=re.search(r'20\d{2}年[^。]{0,110}?(?:公司债券|企业债券|中期票据|短期融资券)',cover)
            if mt:bond=mt.group()
        period=next((b.text for b in self.blocks[:25] if '报告期：' in b.text or '报告期:' in b.text),'')
        chapter=re.search(r'第[一二三四五六七八九十百\d]+章',self.name+' '+head[:500])
        return {'issuer':issuer,'bond':bond,'period':period,'chapter':chapter.group() if chapter else '', 'historical_revisions':self.pkg.revision_count,'blocks':len(self.blocks),'tables':sum(b.kind=='table' for b in self.blocks),'images':sum(b.images for b in self.blocks),'kind':'opinion' if '核查意见' in head[:700] else 'workpaper' if re.search(r'核查记录|核查分析文件',head[:700]) else 'prospectus' if '募集说明书' in head[:700] else 'unknown'}
    def _structure(self):
        starts=[]
        for i,b in enumerate(self.blocks):
            if b.kind!='paragraph' or b.is_toc:continue
            m=MATTER.match(b.text)
            if m and re.search(r'核查|调查|分析|核验',m.group(2)):
                upcoming=[x.text for x in self.blocks[i+1:i+12] if x.kind=='paragraph']
                if any(SECTION.match(t.replace(' ','')) for t in upcoming):
                    starts.append((i,re.sub(r'[－—–]','-',re.sub(r'\s+','',m.group(1))),m.group(2)))
        for k,(i,mid,title) in enumerate(starts):
            end=starts[k+1][0] if k+1<len(starts) else len(self.blocks)
            self.matters.append({'id':mid,'title':title,'start':i,'end':end})
            cur='other'
            for b in self.blocks[i:end]:
                b.matter=mid
                mt=SECTION.match(b.text.replace(' ',''))
                if mt:
                    s=mt.group(1);cur='proc' if any(x in s for x in ['程序','方法']) else 'concl' if any(x in s for x in ['结论','意见']) else 'situ'
                    b.protected=True;b.reason='核查结构';b.heading=True;b.level=1
                b.section=cur
                if b.index==i:b.protected=True;b.reason='事项标题';b.heading=True;b.level=0
                elif cur in ('proc','concl','other'):b.protected=True;b.reason='核查程序或判断'
                if PROTECTED.search(b.text) or (not b.heading and re.match(r'^(?:综上|经.{0,14}核查|根据.{0,35}(?:访谈|核查)|项目组|本项目组|本公司(?:经|认为|通过|核查)|律师认为)',b.text)):
                    b.protected=True;b.reason='独立核查或分析判断'
        if not starts:
            # Generic reports can still be compared; no implicit "everything is a fact" rule.
            cur='other'
            first_chapter=next((b.index for b in self.blocks if re.match(r'^第.+[章节]',b.text) and b.heading),len(self.blocks))
            for b in self.blocks:
                if self.profile['kind']=='opinion' and b.text.strip()=='核查意见' and b.index<first_chapter:
                    b.section='other';continue
                mt=SECTION.match(b.text.replace(' ',''))
                if mt:
                    s=mt.group(1);cur='proc' if '程序' in s or '方法' in s else 'concl' if '结论' in s or '意见' in s else 'situ'
                    b.protected=True;b.reason='核查结构'
                b.section=cur
                if PROTECTED.search(b.text) or cur in ('proc','concl'):b.protected=True;b.reason='独立核查或分析判断'
            self.warnings.append('未检测到标准调查事项编号；按正文标题提供对照，不按固定章节编号猜测。')
        stack=[]
        for b in self.blocks:
            if b.is_toc:continue
            if b.heading:
                while stack and stack[-1][0]>=b.level:stack.pop()
                stack.append((b.level,b.text))
            b.path=[x[1] for x in stack]
            if b.images:
                nearby=' '.join(x.text for x in self.blocks[max(0,b.index-3):b.index+1])
                # Raster evidence is never declared checked from document text alone.
                b.protected=True;b.reason='保留原图，未核验图内内容' if EVIDENCE.search(nearby) else '保留原图，图片不参与自动文字匹配'
        # Opinion documents have no literal “核查情况” container. Enable only
        # clearly factual chapter families, never the entire opinion by default.
        if not starts and self.profile['kind']=='opinion':
            factual=re.compile(r'发行人基本情况|发行人概况|主要业务情况|主要发行条款|发行人存在主要风险')
            for b in self.blocks:
                scope=' '.join(b.path)
                if not b.protected and factual.search(scope) and not re.search(r'核查结论|内核意见|承诺',scope):
                    b.section='situ'
                if re.match(r'^(?:综上|由此可见|主承销商|经.{0,18}核查)',b.text) and not b.heading:
                    b.protected=True;b.reason='独立核查或分析判断'
        if self.matters:
            for b in self.blocks:
                tx=b.text.strip();scope=' '.join(b.path)
                if not tx:continue
                if re.match(r'^(?:式中[：:]?|[（(]以下无正文[）)]|节能量|图[：:])',tx) or (len(tx)<12 and re.fullmatch(r'[①②③④⑤⑥⑦⑧⑨⑩]?(?:式中|节能量|计算公式)[：:]?',tx)):
                    b.protected=True;b.reason='公式或图示结构保留，未作自动事实核验'
                if MATTER.match(tx) and re.search(r'核查记录.*不适用',tx):
                    b.protected=True;b.reason='不适用事项标题保留'
                if re.match(r'^(?:经项目组(?:查阅|核对|访谈)|经示例证券对.{0,85}(?:核查|认为)|经对.{0,85}核查|经(?:查阅|审阅).{0,80}(?:律师|审计|执业|许可证)|通过核查.{0,80}(?:审计报告|法律意见书)|审阅.{0,40}审计报告|全面阅读.{0,50}专业意见|经.{1,45}(?:律师|会计师).{0,8}(?:审阅|核查)|基于上述.{0,20}(?:律师|项目组).{0,8}认为|阅读.{0,55}(?:法律意见书|审计报告)|查阅.{0,75}(?:执业资格|从业资格|执业证|许可证)|对.{0,30}(?:合法有效性|主体资格|执业资格).{0,15}进行核查|针对.{0,70}履行.{0,10}(?:程序|义务)|取得.{0,35}承诺|项目组(?:将|已|通过|进行|核查))',tx):
                    b.protected=True;b.reason='独立核查程序与结论保留，未代替重新执行'
                if re.search(r'可比(?:公司|企业)|同行业.*比较|同业.*比较',scope) or (b.kind=='table' and re.search(r'公司名称.*(?:平均|行业)|行业平均|公司名称.*毛利率|可比公司',tx)):
                    b.protected=True;b.reason='项目组同业比较保留，需另行核验比较依据'
                if re.match(r'^(?:根据|结合).{0,70}(?:可比公司|同业|行业平均)|^经.{0,35}同行业比较',tx):
                    b.protected=True;b.reason='项目组比较分析保留，未重新验证分析结论'
                if len(tx)<80 and re.search(r'(?:图如下|图如下所示|简历如下|制度主要内容如下|专项信息披露情况如下)',tx):
                    b.protected=True;b.reason='结构引导语保留，不单独替换或推定其下方事实已核验'
                nearby=' '.join(x.text for x in self.blocks[max(0,b.index-2):b.index])
                if b.kind=='table' and re.search(r'与其他.{0,15}(?:同类|可比)|同行业|同行可比',nearby):
                    b.protected=True;b.reason='项目组同业比较表保留，需另行核验比较依据'
                if tx.startswith(('表：','表:')) and b.index+1<len(self.blocks):
                    nxt=self.blocks[b.index+1]
                    if nxt.images:
                        b.protected=True;b.reason='图文混合表的表题保留，未自动替换图内信息'
        # Add a precise warning instead of treating unparsed narrative as updated.
    def _units(self):
        # Source leaf sections. Body paragraphs before the first real heading are excluded
        # from automatic source reuse unless they have sufficient factual content.
        headings=[b.index for b in self.blocks if b.heading and not b.is_toc]
        for n,i in enumerate(headings):
            end=headings[n+1] if n+1<len(headings) else len(self.blocks)
            bs=[b for b in self.blocks[i+1:end] if b.text and not b.is_toc and not b.images]
            if bs:self.source_units.append(Unit('s'+str(i),bs[0].index,bs[-1].index+1,bs,self.blocks[i].text,self.blocks[i].path,self.blocks[i].matter))
        # Individual blocks and table bundles provide fallback for split/reordered sections.
        for i,b in enumerate(self.blocks):
            if b.is_toc or b.heading or b.images or not b.text:continue
            if b.kind=='paragraph' and (len(b.text)<20 or b.text.startswith(('表：','图：','单位：'))):continue
            lo=i
            if b.kind=='table':
                for j in range(i-1,max(-1,i-4),-1):
                    prev=self.blocks[j]
                    if prev.kind=='paragraph' and not prev.heading and (prev.text.startswith(('表：','表:','单位：','单位:')) or (len(prev.text)<140 and re.search(r'如下|下表|列示',prev.text))):lo=j
                    else:break
            bs=self.blocks[lo:i+1];self.source_units.append(Unit('b'+str(i),lo,i+1,bs,b.path[-1] if b.path else '',b.path,b.matter))
        # Target units: contiguous factual material, never spanning a protected block or heading.
        start=None;heading='';matter=''
        def flush(end):
            nonlocal start
            if start is None:return
            bs=self.blocks[start:end]
            while bs and not (bs[0].text or bs[0].images):bs=bs[1:]
            while bs and not (bs[-1].text or bs[-1].images):bs=bs[:-1]
            if bs:self.units.append(Unit('u'+str(bs[0].index),bs[0].index,bs[-1].index+1,bs,heading,bs[0].path,bs[0].matter))
            start=None
        for b in self.blocks:
            eligible=not b.protected and not b.is_toc and not b.heading and (b.section=='situ' if self.matters else b.section=='situ')
            if not eligible or (start is not None and self.parents.get(self.blocks[start].el) is not self.parents.get(b.el)):flush(b.index)
            if b.heading:heading=b.text
            if eligible and start is None:start=b.index
        flush(len(self.blocks))
    def split_unit(self,u):
        units=[];i=u.start
        while i<u.end:
            b=self.blocks[i]
            if not b.text:i+=1;continue
            end=i+1
            if b.kind=='paragraph' and (b.text.startswith(('表：','表:','单位：','单位:')) or (len(b.text)<140 and re.search(r'如下|下表|列示',b.text))):
                j=i+1
                while j<u.end and self.blocks[j].kind=='paragraph' and self.blocks[j].text.startswith(('表：','表:','单位：','单位:')):j+=1
                if j<u.end and self.blocks[j].kind=='table':end=j+1
            bs=self.blocks[i:end];units.append(Unit(u.id+'b'+str(i),i,end,bs,u.heading,u.path,u.matter));i=end
        return units
    def protected_summary(self):
        counts=collections.Counter(b.reason for b in self.blocks if b.protected and b.text or b.images)
        return dict(counts)
    def images(self,b):
        out=[];rels={r.get('Id'):r for r in self.pkg.rels()}
        for el in b.el.iter():
            for k,v in el.attrib.items():
                if k in ('{'+NS['r']+'}embed','{'+NS['r']+'}id') and v in rels:
                    r=rels[v]
                    if '/image' in r.get('Type','') and r.get('TargetMode')!='External':
                        import posixpath
                        part=posixpath.normpath(posixpath.join('word',r.get('Target')))
                        if part in self.pkg.entries:out.append((part,self.pkg.entries[part]))
        return out
