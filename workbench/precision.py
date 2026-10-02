"""Scope-aware comparison units and verifiable source spans.

No issuer names, expected figures, chapter numbers or answer labels are stored here.
"""
from __future__ import annotations
import copy,re,hashlib,json,unicodedata
from .docxio import w,text,local,NS,complex_reference_reason
from .assurance import periods,compact
from .model import Unit,strip_lead,is_short_numbered_title

CAPTION=re.compile(r'^\s*(?:表\s*[:：]|表\s*\d+\s*[:：、.．]?)')
UNIT=re.compile(r'^\s*单位\s*[:：]')
NOTE=re.compile(r'^\s*(?:注[\d一二三四五]*[:：]|注\s*\d+[、.．]|数据来源[:：]|资料来源[:：])')
GENERIC=re.compile(r'^(?:[一二三四五六七八九十\d]+[、．.]|[（(][一二三四五六七八九十\d]+[）)])?\s*(?:核查|调查|分析)(?:情况|记录|结果|过程|程序|意见|结论)?$')


def is_caption(b):return b.kind=='paragraph' and bool(CAPTION.match(b.text) or UNIT.match(b.text))
def same_parent(doc,a,b):return doc.parents.get(a.el) is doc.parents.get(b.el)


def useful_title(u):
    for b in reversed(u.blocks):
        if CAPTION.match(b.text):return b.text
    for s in reversed(u.path):
        if not GENERIC.match(s) and len(strip_lead(s))>2:return s
    return u.heading


def nearest_scope(path,prefix=''):
    """A leaf 'consolidated' overrides a parent 'consolidated and parent statements'."""
    if re.search(r'发行人(?:本部|母公司)|母公司(?:口径|本部)',prefix):return '发行人本部'
    for s in reversed(path):
        if re.search(r'会计政策变更|会计估计变更|调整影响',s):return '会计调整'
        if re.search(r'母公司|本部',s) and re.search(r'合并[及和与]|合并.*母公司',s):continue
        if re.search(r'^(?:[一二三四五六七八九十\d]+[、.．]|[（(]\d+[）)])?\s*(?:保证人|担保人|增信主体)(?:的|基本|主要|财务|资信|信用|情况|$)',s) or re.search(r'^(?:第.{1,6}[节章]\s*)?增信机制$',s):return '保证人'
        if '控股股东' in s and '实际控制人' not in s:return '控股股东'
        if '实际控制人' in s and '控股股东' not in s:return '实际控制人'
        if re.search(r'母公司|发行人本部|本部财务',s):return '发行人本部'
        if re.search(r'合并(?:财务|资产|利润|现金|报表)',s):return '发行人'
    return '发行人'


def table_identity(u):
    tables=[b for b in u.blocks if b.kind=='table']
    caption=' '.join(b.text for b in u.blocks if CAPTION.match(b.text))
    title=caption or useful_title(u)
    header_rows=[]
    for b in tables:
        for ri,row in enumerate(b.el.findall(w('tr'))[:3]):
            cells=[text(c).strip() for c in row.findall(w('tc'))]
            explicit=row.find('./'+w('trPr')+'/'+w('tblHeader')) is not None
            # Never treat the first data row's incorporation/maturity dates as
            # reporting periods. Multi-level date-only header rows are allowed.
            header_like=bool(cells) and all(not x or re.fullmatch(r'(?:项目|科目|指标|金额|比例|占比|单位|名称|本期|上期|余额|年末|年度|期间|[\s\d年\-/月日度末（）()%％、.]+)',x) for x in cells)
            if ri and not explicit and not header_like:break
            header_rows.append(text(row))
    header='\n'.join(header_rows)
    p=sorted(set(periods(caption)+periods(header)))
    if not p:p=periods(title)
    units=[]
    for b in u.blocks:
        if UNIT.match(b.text):units.extend(re.split(r'[、,，/／\s]+',UNIT.sub('',b.text.strip())))
    if not units:
        units=re.findall(r'[（(]([^（）()]{1,12})[）)]',header)
        units=[v for v in units if re.search(r'元|%|％|倍|平方|瓦|头|吨',v)]
    def category_of(v):
        customer=bool(re.search(r'客户|销售对象',v));supplier=bool(re.search(r'供应商|采购对象',v))
        return '客户' if customer and not supplier else '供应商' if supplier and not customer else ''
    tc,hc=category_of(title),category_of(header[:120])
    category=tc or hc
    category_conflict=bool(tc and hc and tc!=hc)
    businesses=[x for x in ('电力','食品','物业','供水','屠宰','房地产','住宅','商业','工程建设','担保业务') if x in title]
    metric_text=title+' '+header[:130]
    metrics=[label for label,pat in [('资产负债表',r'资产负债表'),('现金流量表',r'现金流量表'),('利润表',r'(?<!毛)利润表'),('营业收入',r'营业收入|主营业务收入'),('营业成本',r'营业成本|主营业务成本'),('毛利润',r'毛利润|毛利情况'),('毛利率',r'毛利率')] if re.search(pat,metric_text)]
    return {'category_conflict':category_conflict,'metrics':metrics,'title':title,'caption':caption,'periods':p,'units':sorted(set(x.replace('％','%') for x in units if x)), 'scope':nearest_scope(u.path),'category':category,'business':businesses,'tables':len(tables),'header':header}


def metric_signature(s):
    # The leading measured fact, not every business word mentioned in the paragraph.
    t=compact(s)
    m=re.search(r'分别(?:为|是)|(?:账面价值|余额|金额)(?:分别)?(?:为|是)',t)
    if not m:return set()
    lead=t[max(0,m.start()-95):m.start()]
    keys=[('其他业务收入',r'其他业务收入'),('其他业务成本',r'其他业务成本'),('其他业务毛利',r'其他业务毛利'),
          ('营业外收入',r'营业外收入'),('其他收益',r'其他收益'),('资产处置收益',r'资产处置收益'),
          ('公允价值变动收益',r'公允价值变动收益'),('投资收益',r'投资收益'),
          ('营业收入',r'营业收入'),('营业成本',r'营业成本'),('净利润',r'净利润'),('营业利润',r'营业利润'),
          ('毛利润',r'毛利润'),('毛利率',r'毛利率'),('资产总计',r'总资产|资产总计'),
          ('负债总计',r'总负债|负债总计|负债合计'),('净资产',r'净资产|所有者权益|股东权益')]
    return {k for k,pat in keys if re.search(pat,lead)}


def numeric_evidence_equivalent(a,b):
    """Allow explicit currency unit conversion and rounding, never missing quantities."""
    from decimal import Decimal, ROUND_HALF_UP
    def values(s):
        s=unicodedata.normalize('NFKC',compact(s)).replace(',','').replace('，','')
        result=[]
        for m in re.finditer(r'[-+]?\d+(?:\.\d+)?(?:亿元|万元|元|%|％)?',s):
            token=m.group();unit=next((u for u in ('亿元','万元','元','%','％') if token.endswith(u)),'')
            raw=token[:-len(unit)] if unit else token
            value=Decimal(raw);dec=len(raw.partition('.')[2]);step=Decimal(10)**(-dec)/2
            mult=Decimal(100000000 if unit=='亿元' else 10000 if unit=='万元' else 1)
            result.append(('money' if unit in ('亿元','万元','元') else 'percent' if unit in ('%','％') else 'raw',value*mult,step*mult))
        return result
    aa,bb=values(a),values(b)
    if len(aa)!=len(bb):return False
    def equal(x,y):
        if x[0]!=y[0]:return False
        if x[1]==y[1]:return True
        if x[0]=='raw' or x[2]==y[2]:return False
        coarse,fine=(x,y) if x[2]>y[2] else (y,x)
        quantum=coarse[2]*2
        return (fine[1]/quantum).to_integral_value(rounding=ROUND_HALF_UP)*quantum==coarse[1]
    return all(equal(x,y) for x,y in zip(aa,bb))

def descriptor_guard(t,s):
    ts=nearest_scope(t.path,t.text[:110] if not t.has_table else '')
    ss=nearest_scope(s.path,s.text[:110] if not s.has_table else '')
    if ts!=ss and {'发行人本部','保证人'} & {ts,ss}:return '主体口径不同：'+ts+' / '+ss
    if t.has_table:
        a,b=table_identity(t),table_identity(s)
        if a['category_conflict'] or b['category_conflict']:return '表题与表头的客户/供应商类别相互冲突。'
        if a['metrics'] and b['metrics'] and set(a['metrics']).isdisjoint(b['metrics']):return '表格指标不同：'+'、'.join(a['metrics'])+' / '+'、'.join(b['metrics'])
        if a['tables']!=b['tables']:return '表格数量不同，不能以多张表替换一张表'
        if a['periods'] and b['periods'] and set(a['periods'])!=set(b['periods']):return '表格期间不同：'+ '、'.join(a['periods'])+' / '+'、'.join(b['periods'])
        if a['category'] and b['category'] and a['category']!=b['category']:return '客户表与供应商表不能混用'
        if a['business'] and b['business'] and set(a['business']).isdisjoint(b['business']):return '表格业务板块不同'
        if a['units'] and b['units'] and a['units']!=b['units']:return '计量单位不同：'+str(a['units'])+' / '+str(b['units'])
    # Protect names when comparing discrete corporate profiles, not counterparties
    # inside a whole customer/supplier table whose composition can legitimately change.
    if not t.has_table and not s.has_table:
        def measured_business(v):
            lead=compact(v).split('分别')[0][-85:]
            return {b for b in ('电力','食品','物业','供水','屠宰','房地产','住宅','工程建设','其他') if re.search(re.escape(b)+r'[^，。]{0,12}业务(?:收入|成本|毛利润|毛利率)',lead)}
        tb,sb=measured_business(t.text),measured_business(s.text)
        if tb and sb and tb.isdisjoint(sb):return '段落的业务板块不同，不能以同句式金额替换'
        tm,sm=metric_signature(t.text),metric_signature(s.text)
        if tm and sm and tm.isdisjoint(sm):return '段落描述的财务指标不同，不能互换'
        pat=r'^([\u4e00-\u9fffA-Za-z（）()]{4,55}(?:有限公司|有限责任公司|股份有限公司))'
        a,b=re.match(pat,t.text),re.match(pat,s.text)
        if a and b and a[1]!=b[1] and re.search(r'成立|注册资本|经营范围',t.text[:180]+s.text[:180]):return '公司简介的具名主体不同'
    return ''


def source_table_change(t,s):
    a,b=table_identity(t),table_identity(s)
    if not a['caption'] or not b['caption'] or not a['periods'] or not b['periods']:return False
    def key(c):return re.sub(r'[\s:：、（）()，,。.]','',c).replace('发行人','').replace('公司','')
    return key(a['caption'])==key(b['caption']) and a['periods']==b['periods'] and a['units']==b['units'] and not descriptor_guard(t,s)


def evidence_numbers(s):
    s=unicodedata.normalize('NFKC',compact(s)).replace(',','').replace('，','')
    return re.findall(r'(?<![A-Za-z])[-+]?\d+(?:\.\d+)?(?:%|％)?',s)


def atomic_units(doc,u):
    """Table caption + unit + full table + footnotes; other paragraphs separately."""
    out=[];i=u.start
    while i<u.end:
        b=doc.blocks[i]
        if not b.text:i+=1;continue
        j=i+1
        if is_caption(b):
            while j<u.end and is_caption(doc.blocks[j]):j+=1
            if j<u.end and doc.blocks[j].kind=='table':j+=1
            else:j=i+1
        if any(x.kind=='table' for x in doc.blocks[i:j]):
            while j<u.end and NOTE.match(doc.blocks[j].text):j+=1
        bs=doc.blocks[i:j]
        unit=Unit('a'+str(i),i,j,bs,u.heading,bs[0].path,u.matter)
        unit.heading=useful_title(unit);out.append(unit);i=j
    return out


def sentence_boundaries(value):
    """Only cut at complete sentences, including their closing quotation marks.

    A full stop inside a quotation is not a source boundary. Splitting before
    its closing mark would turn a direct quote into the reviewer's own voice.
    """
    pairs={'“':'”','‘':'’','「':'」','『':'』','《':'》'}
    stack=[];bounds=[];last_stop=False
    for i,char in enumerate(value):
        if char in pairs:stack.append(pairs[char])
        elif stack and char==stack[-1]:stack.pop()
        elif char=='"':
            if stack and stack[-1]=='"':stack.pop()
            else:stack.append('"')
        if char=='。':last_stop=True
        elif char not in ('”','’','」','』','"') and not char.isspace():last_stop=False
        if last_stop and not stack and (char=='。' or char in ('”','’','」','』','"')):
            end=i+1
            while end<len(value) and value[end].isspace():end+=1
            if end<len(value):bounds.append(end)
            last_stop=False
    return sorted(set(bounds))


def augment_sources(doc):
    """Add honest source paragraph spans when templates divide an existing paragraph."""
    # Replace old table bundles by deterministic full-table bundles. No adjacent
    # prose with a different assertion is absorbed merely because it says '如下'.
    items=[];seen=set()
    for u in doc.source_units:
        if not u.has_table and not any(b.heading or is_short_numbered_title(b.text) for b in u.blocks):items.append(u)
    for b in doc.blocks:
        if b.kind=='table' and not b.is_toc:
            lo=b.index;hi=lo+1
            while lo>0 and is_caption(doc.blocks[lo-1]) and same_parent(doc,doc.blocks[lo-1],b):lo-=1
            while hi<len(doc.blocks) and NOTE.match(doc.blocks[hi].text) and same_parent(doc,b,doc.blocks[hi]):hi+=1
            u=Unit('t'+str(b.index),lo,hi,doc.blocks[lo:hi],b.path[-1] if b.path else '',b.path,b.matter)
            u.heading=useful_title(u);items.append(u)
        elif b.kind=='paragraph' and not (b.heading or b.is_toc or b.images or is_short_numbered_title(b.text)) and len(b.text)>=7 and not is_caption(b):
            u=Unit('p'+str(b.index),b.index,b.index+1,[b],b.path[-1] if b.path else '',b.path,b.matter);items.append(u)
            # Natural information boundary: corporate profile followed by dated
            # financial figures, separated within the source by a full stop.
            cuts=sentence_boundaries(b.text)
            if cuts:
                bounds=[0]+cuts+[len(b.text)]
                for k,j in ((k,j) for k in range(len(bounds)-1) for j in range(k+1,min(k+4,len(bounds)))):
                    start,end=bounds[k],bounds[j]
                    if end-start<25 or (start==0 and end==len(b.text)):continue
                    block=copy.copy(b)
                    try:block.el=slice_paragraph(b.el,start,end)
                    except ValueError:continue
                    block.text=text(block.el)
                    v=Unit('p'+str(b.index)+'_'+str(start),b.index,b.index+1,[block],u.heading,u.path,u.matter)
                    v.span={'block':b.index,'start':start,'end':end};items.append(v)
    result=[]
    for u in items:
        key=(u.start,u.end,getattr(u,'span',{}).get('start'),getattr(u,'span',{}).get('end'))
        if key in seen:continue
        seen.add(key);result.append(u)
    doc.source_units=result


def slice_paragraph(el,start,end):
    clone=copy.deepcopy(el);raw=text(clone)
    if not 0<=start<end<=len(raw):raise ValueError('来源文字区间已变化。')
    if clone.tag!=w('p') or clone.findall('.//'+w('drawing')) or clone.findall('.//'+w('pict')) or complex_reference_reason([clone]) or any(str(e.tag).startswith('{http://schemas.openxmlformats.org/officeDocument/2006/math}') for e in clone.iter()):raise ValueError('此复杂段落不能作文字切片。')
    cursor=0
    for node in clone.iter():
        if node.tag==w('t'):
            val=node.text or '';a=max(0,start-cursor);b=min(len(val),end-cursor)
            node.text=val[a:b] if b>a else '';cursor+=len(val)
            node.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        elif node.tag in (w('tab'),w('br')):
            if cursor<start or cursor>=end:node.tag=w('t');node.text=''
            cursor+=1
    if text(clone)!=raw[start:end]:raise ValueError('来源切片无法无损提取，已停止。')
    return clone



def _section_key(value):
    """Normalize a visible heading for block-level synchronization."""
    value=strip_lead(value or '')
    value=re.sub(r'\s+','',value)
    value=value.replace('公司','').replace('发行人','')
    value=re.sub(r'[：:，,。；;（）()【】\[\]“”"\'、]','',value)
    return value

def _meaningful_parent_keys(path):
    out=[]
    for part in path[:-1]:
        raw=strip_lead(part or '')
        if not raw or re.search(r'核查(?:情况|记录|分析文件)|调查(?:情况|记录)|分析(?:情况|记录)',raw):
            continue
        k=_section_key(raw)
        if len(k)>=2:out.append(k)
    return out[-3:]

def complete_section_candidates(doc,target,source_docs,source_flags=None):
    """Find complete source sections bounded by the same visible heading.

    This is deliberately stricter than paragraph retrieval.  It is used only when
    the target Unit itself fills the complete area under a heading up to the next
    heading.  That lets the exporter synchronize the whole factual block and avoids
    the visually confusing mixture of updated and retained paragraphs inside one
    copied disclosure section.
    """
    source_flags=source_flags or {}
    if any(getattr(b,'formal_source_correspondence',None) for b in target.blocks):return []
    # Locate the heading immediately owning this unit.
    hidx=None
    for i in range(target.start-1,-1,-1):
        b=doc.blocks[i]
        if b.heading and not b.is_toc:
            hidx=i;break
    if hidx is None:return []
    hb=doc.blocks[hidx]
    # Only a complete, uninterrupted leaf range is eligible.
    next_heading=next((b.index for b in doc.blocks[hidx+1:] if b.heading and not b.is_toc),len(doc.blocks))
    if target.start!=hidx+1 or target.end!=next_heading:return []
    if any(b.protected or b.images or b.is_toc for b in doc.blocks[target.start:target.end]):return []
    target_key=_section_key(hb.text)
    if len(target_key)<2:return []
    target_parents=_meaningful_parent_keys(hb.path)
    result=[]
    for sd in source_docs:
        for sh in sd.blocks:
            if not sh.heading or sh.is_toc:continue
            source_key=_section_key(sh.text)
            # Exact leaf title is the automatic path.  Near-title matching is
            # allowed only as a review candidate and still needs parent context.
            exact=source_key==target_key
            if exact:leaf_score=1.0
            else:
                # Local import avoids a module cycle at import time.
                from .matching import ratio
                leaf_score=ratio(target_key,source_key)
                if leaf_score<.88:continue
            source_parents=_meaningful_parent_keys(sh.path)
            parent_score=0.0;immediate_parent_score=0.0
            if target_parents and source_parents:
                from .matching import ratio
                parent_score=max(ratio(a,b) for a in target_parents for b in source_parents)
                immediate_parent_score=ratio(target_parents[-1],source_parents[-1])
            elif not target_parents and not source_parents:
                parent_score=1.0;immediate_parent_score=1.0
            # Generic leaf headings such as “基本情况/业务概况/经营情况” are
            # ubiquitous.  Their immediate parent must identify the same business
            # or corporate subject; sharing only a grandparent is not enough.
            if len(target_key)<8 and immediate_parent_score<.90:continue
            if parent_score<.68 or immediate_parent_score<.72:continue
            send=next((b.index for b in sd.blocks[sh.index+1:] if b.heading and not b.is_toc),len(sd.blocks))
            blocks=[b for b in sd.blocks[sh.index+1:send] if not b.is_toc]
            if not blocks:continue
            # Do not silently import raster evidence/formulas or other complex OOXML.
            unsupported=complex_reference_reason([b.el for b in blocks])
            if any(b.images for b in blocks):unsupported=unsupported or '来源完整事项含图片；为避免图片关系错位，保留原稿并提示人工核对。'
            body='\n'.join(b.text for b in blocks if b.text)
            if not body.strip():continue
            warnings=sorted({v for b in blocks for v in source_flags.get((sd.hash,b.index),[])})
            score=round(.72*leaf_score+.28*parent_score,4)
            result.append({'span':None,'identity':None,'index':-1,'doc_hash':sd.hash,'source_name':sd.name,'unit_id':'section'+str(sh.index),'start':blocks[0].index,'end':blocks[-1].index+1,'indices':[b.index for b in blocks],'score':score,'content_score':0.0,'literal_score':0.0,'heading_score':round(leaf_score,4),'table_score':0.0,'locator':' > '.join(sh.path),'text':body,'exact':normal_section_text(target.text)==normal_section_text(body),'unsupported':unsupported,'source_warnings':warnings,'scope_verified':bool(exact and parent_score>=.92 and immediate_parent_score>=.86),'whole_section':True,'parent_score':round(parent_score,4),'immediate_parent_score':round(immediate_parent_score,4),'source_kind':sd.profile.get('kind','')})
    result.sort(key=lambda c:(-int(c['scope_verified']),-c.get('immediate_parent_score',0),-c['score'],0 if next((d for d in source_docs if d.hash==c['doc_hash']),None).profile.get('kind')=='prospectus' else 1,c['source_name'],c['start']))
    # Deduplicate identical evidence across source documents/locations.
    unique=[];seen=set()
    for c in result:
        k=(c['doc_hash'],normal_section_text(c['text']))
        if k in seen:continue
        seen.add(k);unique.append(c)
    return unique[:4]

def normal_section_text(value):
    # Used only to judge whether a located section is already the same.  Treat the
    # frozen archive date wording and prospectus issuer voice as editorial aliases,
    # but do not collapse generic `本募集说明书` wording or other substantive text.
    s=value or ''
    s=re.sub(r'(?:本)?募集说明书(?:签署|出具)(?:之)?日|本核查(?:分析文件|文件|意见)出具(?:之)?日','本核查分析文件出具日',s)
    s=s.replace('本公司','发行人')
    return re.sub(r'\s+','',s)

def assess_complete_section(doc,target,source_docs,source_flags=None):
    c=complete_section_candidates(doc,target,source_docs,source_flags)
    if not c:return None
    best=c[0]
    if best.get('unsupported'):
        return 'review',c,best['unsupported']
    strong=[x for x in c if x.get('scope_verified')]
    if not strong:
        return 'review',c,'标题接近但不是完全一致；请确认整个事项是否对应后再采用。'
    # For factual workpaper synchronization the prospectus is the primary disclosure
    # source.  A companion underwriting opinion may intentionally contain a shorter
    # or differently framed summary; it must not block an exact prospectus section.
    primary=[x for x in strong if x.get('source_kind')=='prospectus'] or strong
    variants={normal_section_text(x['text']) for x in primary}
    if len(variants)>1:
        return 'review',c,'同一完整事项在同类最新版来源中存在内容差异；请先确定适用版本。'
    if best not in primary:
        best=primary[0]
        c=[best]+[x for x in c if x is not best]
    other=next((x for x in strong if x.get('source_kind')=='opinion' and normal_section_text(x['text'])!=normal_section_text(best['text'])),None)
    if best.get('source_kind')=='prospectus' and other:
        best['source_precedence']={'rule':'conflicting-formal-sources-prefer-prospectus',
            'preferred_source':best['doc_hash'],'other_source':other['doc_hash'],
            'other_indices':other['indices'],'other_text':other['text']}

    if best.get('source_warnings'):
        return 'review',c,'对应完整事项来源存在一致性提示：'+'；'.join(best['source_warnings'])+'。'
    if normal_section_text(target.text)==normal_section_text(best['text']):
        return 'auto',c,'已按完整事项定位来源；即使内容相同也整块复制来源并保留修订。'
    return 'auto',c,'已按同名标题及父级结构锁定完整事项；从本标题后到下一个同级/后续标题前整块同步，段落、表格和注释一起更新，不再逐段拼接。'

def source_elements(doc,c):
    if c.get('span'):
        x=c['span'];return [slice_paragraph(doc.blocks[x['block']].el,x['start'],x['end'])]
    return [copy.deepcopy(doc.blocks[i].el) for i in c.get('indices',range(c['start'],c['end']))]


def fingerprint(elements):
    # Prefix-independent tree identity, excludes no numeric or structural content.
    def tree(e):return (str(e.tag),tuple(sorted(e.attrib.items())),e.text or '',tuple(tree(c) for c in e),e.tail or '')
    return hashlib.sha256(json.dumps([tree(e) for e in elements],ensure_ascii=False,separators=(',',':')).encode()).hexdigest()


def table_cells(elements):
    return [[[{'text':text(tc),'grid':next((e.get(w('val'),'1') for e in tc.findall('./'+w('tcPr')+'/'+w('gridSpan'))),'1'),'merge':next((e.get(w('val'),'continue') for e in tc.findall('./'+w('tcPr')+'/'+w('vMerge'))),'')} for tc in tr.findall(w('tc'))] for tr in tbl.findall(w('tr'))] for e in elements for tbl in e.iter(w('tbl'))]
