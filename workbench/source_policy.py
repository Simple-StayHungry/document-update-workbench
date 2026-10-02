"""Conservative ownership, complete paragraph alignment and full object accounting.

Retrieval scores never establish ownership. Workpaper procedures/conclusions and
external evidence stay independent. Corresponding formal underwriter opinions can
update opinion-derived narrative without blanket-unlocking the workpaper.
"""
from __future__ import annotations
import copy, hashlib, json, re
from pathlib import Path
from .docxio import text, w, complex_reference_reason
from .matching import excerpt_identity, title_key

JUDGMENT = re.compile(r'项目组|本项目组|主承销商(?:认为|经|通过)|经.{0,30}(?:核查|查询|访谈)|根据.{0,45}(?:查询结果|核查结果)|通过.{0,30}(?:核查|查询)|(?:企查查|天眼查|信用中国|国家企业信用信息公示系统|失信被执行人名单)|可能造成.{0,20}股权结构不稳定|严重依赖于少数')
NO_CORRESPONDENCE = '当前正式来源未找到明确对应内容，按规则保留底稿原样；不删除、不改写、不新增修订。'
_SHORT_SOURCE_ONLY = '来源明显短于目标，可能遗漏原段落中的事实或子事项。'
_WEAK_CANDIDATE_ONLY = '此候选只有弱相关性，不支持直接覆盖。'

def _weak_short_source_without_correspondence(p):
    """A generic length veto cannot establish a match below retrieval's floor.

    Do not relax concrete coverage, object, layout or source-consistency vetoes.
    The single length warning can precede the score check in Index.assess, leaving
    unrelated institution/contact snippets mislabeled as unresolved replacements.
    """
    if (p.get('status') != 'review' or p.get('reason') != _SHORT_SOURCE_ONLY or
            p.get('has_table') or p.get('figure') or p.get('object_kind') == 'figure'):
        return False
    candidates = p.get('candidates', [])
    return bool(candidates) and all(
        isinstance(c.get('score'), (int, float)) and c['score'] < .52 and
        not any(c.get(k) for k in ('exact', 'scope_verified', 'correspondence',
                                   'span', 'projection', 'unsupported', 'source_warnings', 'layout_risks')) and
        c.get('blocked_reason', '') in ('', _WEAK_CANDIDATE_ONLY)
        for c in candidates)

def retain_unmatched(p):
    """Absence is a completed retention decision, never a request to delete.

    Only explicit missing matches take this path. Copy vetoes, unresolved source
    versions and unsupported objects remain distinct from not finding a source.
    Weak retrieval candidates are diagnostic leads, not actionable copy plans.
    """
    if p.get('content_class')=='no_correspondence_retained':return p
    if p.get('decision') not in ('pending','keep'):
        return p
    if p.get('status') != 'missing' and not _weak_short_source_without_correspondence(p):
        return p
    p['correspondence_search']={'reason':p.get('reason',''),
        'candidates':[{'source_name':c.get('source_name'), 'locator':c.get('locator'),
                       'score':c.get('score')} for c in p.get('candidates',[])]}
    p.update(status='missing',decision='keep',selected=None,candidates=[],
             content_class='no_correspondence_retained',reason=NO_CORRESPONDENCE)
    return p

def code_fingerprint():
    root=Path(__file__).resolve().parents[1]
    files=sorted([*root.glob('*.py'),*root.glob('workbench/*.py'),*root.glob('web/*')])
    return hashlib.sha256(json.dumps([(str(p.relative_to(root)),hashlib.sha256(p.read_bytes()).hexdigest()) for p in files if p.is_file()],separators=(',',':')).encode()).hexdigest()

def independent_reason(block):
    if getattr(block,'formal_source_correspondence',None):return ''
    if block.section in ('proc','concl'):
        return '核查程序或项目组结论保留；相似的来源文字不构成覆盖授权。'
    if JUDGMENT.search(block.text):
        return '含独立核查判断、外部查询或分析责任；来源匹配不改变其保护属性。'
    return ''

def _opinion_correspondence(doc, sources):
    """Locate opinion-derived narrative before relaxing a responsibility boundary.

    Exact/full-paragraph evidence or a same-topic near-identical passage is needed.
    A source's existence alone cannot unlock project procedures or external evidence.
    """
    from .matching import Index,normal
    from .model import Unit
    from .assurance import replacement_risks
    opinions=[s for s in sources if s.profile.get('kind')=='opinion']
    if not opinions:return
    index=Index(opinions)
    external=re.compile(r'企查查|天眼查|信用中国|国家企业信用信息公示系统|失信被执行人名单|项目组|同业比较|同行业.*比较|可比公司')
    reasons={'独立核查或分析判断','核查程序或判断','独立核查程序与结论保留，未代替重新执行'}
    for b in doc.blocks:
        if (b.heading or b.is_toc or b.images or b.kind!='paragraph' or
                b.section=='proc' or len(b.text)<25 or not b.path or
                external.search(b.text+' '.join(b.path)) or
                complex_reference_reason([b.el]) or b.el.find('.//'+w('sectPr')) is not None):continue
        # No broad enabling of cover text, formula definitions or figure captions.
        if not (independent_reason(b) or b.reason in reasons):continue
        u=Unit('opinion'+str(b.index),b.index,b.index+1,[b],b.path[-1],b.path,b.matter)
        status,candidates,_=index.assess(u)
        if not candidates:continue
        best=candidates[0]
        if (best.get('span') or best.get('unsupported') or best.get('source_warnings') or
                len(best['indices'])!=1):continue
        item=index.items[best['index']]
        if any(x.heading or x.is_toc or x.images or x.section=='proc' for x in item['unit'].blocks):continue
        if external.search(best['text']):continue
        others=[c for c in candidates[1:] if normal(c['text'])!=normal(best['text'])]
        margin=best['literal_score']-max((c['literal_score'] for c in others),default=0)
        exact=best['exact'] and best.get('context_score',0)>=.25
        same_topic=status=='auto' and best.get('context_score',0)>=.85 and best['literal_score']>=.70
        near=(best['literal_score']>=.90 and best.get('context_score',0)>=.40 and
              margin>=.06 and not replacement_risks(u,item['unit']))
        if len(normal(best['text']))<.85*len(normal(b.text)):continue
        if not (exact or same_topic or near):continue
        b.formal_source_correspondence={'doc_hash':best['doc_hash'],'indices':best['indices'],
            'method':'complete-opinion-paragraph','text':best['text'],
            'reason':'正式核查意见同一事项已有完整对应正文；按来源复制，不重新推断项目组结论。'}
        b.protected=False;b.reason='';b.section='situ'

def prepare_target(doc, formal_sources):
    """Keep responsibility boundaries; enable only proven formal-opinion counterparts."""
    from .model import Unit
    # Engine caches parsed documents: changing the selected sources must revoke an
    # earlier opinion permission rather than silently reusing that permission.
    for b in doc.blocks:
        if not hasattr(b,'_source_policy_original'):
            b._source_policy_original=(b.protected,b.reason,b.section)
        b.protected,b.reason,b.section=b._source_policy_original
        if hasattr(b,'formal_source_correspondence'):del b.formal_source_correspondence
    _opinion_correspondence(doc,formal_sources)
    for b in doc.blocks:
        reason=independent_reason(b)
        if reason:b.protected=True;b.reason=reason
        if b.el.find('.//'+w('sectPr')) is not None:
            b.protected=True;b.reason='目标段落含节分界；保留正文与节设置，当前不能安全整块迁移，待核对。'
    doc.units=[];doc.source_units=[];doc._units()
    units=[]
    for u in doc.units:
        start=u.start
        for b in u.blocks:
            if not getattr(b,'formal_source_correspondence',None):continue
            if start<b.index:
                bs=doc.blocks[start:b.index]
                units.append(Unit('u'+str(start),start,b.index,bs,u.heading,bs[0].path,u.matter))
            units.append(Unit('u'+str(b.index),b.index,b.index+1,[b],u.heading,b.path,u.matter))
            start=b.index+1
        if start<u.end:
            bs=doc.blocks[start:u.end]
            units.append(Unit('u'+str(start),start,u.end,bs,u.heading,bs[0].path,u.matter))
    doc.units=units

def classify_proposal(p,doc):
    blocks=doc.blocks[p['start']:p['end']]
    reason=next((independent_reason(b) for b in blocks if independent_reason(b)), '')
    if reason:
        p.update(status='protected',decision='protected',selected=None,candidates=[],reason=reason,content_class='independent_protected')
    elif p.get('decision') in ('accept','same'):
        p['content_class']='prospectus_source'
    else:
        p['content_class']='source_missing_or_ambiguous'
    return p

def paragraph_merges(doc,sources,props,source_flags=None):
    """Recover a full current paragraph using TWO literal, ordered anchors.

    A stable opening sentence plus a later complete paragraph may identify a
    current paragraph that merged two old ones. Intervening objects are retained;
    only the consumed text paragraph receives a separate tracked deletion.
    """
    result=list(props)
    by_start={p['start']:p for p in props if p['end']==p['start']+1}
    used=set()
    for i,p in sorted(by_start.items()):
        b=doc.blocks[i]
        if i in used or b.protected or b.kind!='paragraph' or b.images or independent_reason(b):continue
        first=re.match(r'^(.{40,}?[。；])',b.text)
        if not first:continue
        opening=first.group(1)
        if len(b.text)<=len(opening):continue
        heading=title_key(b.path[-1]) if b.path else ''
        if len(heading)<4:continue
        matches=[]
        for sd in sources:
            for s in sd.blocks:
                if s.kind!='paragraph' or s.images or s.heading or s.is_toc or complex_reference_reason([s.el]):continue
                if not s.path or title_key(s.path[-1])!=heading:continue
                if not s.text.startswith(opening):continue
                tail=s.text[len(opening):]
                if len(tail)<40:continue
                for j in range(i+1,min(i+6,len(doc.blocks))):
                    other=doc.blocks[j];op=by_start.get(j)
                    if other.heading or other.matter!=b.matter:break
                    if not op or other.protected or other.images or other.kind!='paragraph':continue
                    if excerpt_identity(other.text)!=excerpt_identity(tail):continue
                    if any(x.text and not x.images and not re.match(r'^图[:：]',x.text) for x in doc.blocks[i+1:j]):continue
                    matches.append((sd,s,j,op))
        unique={(sd.hash,s.index,j):(sd,s,j,op) for sd,s,j,op in matches}
        if len(unique)!=1:continue
        sd,s,j,op=next(iter(unique.values()))
        warnings=sorted(set((source_flags or {}).get((sd.hash,s.index),[])) | {v for candidate in p.get('candidates',[]) + op.get('candidates',[]) if candidate.get('doc_hash')==sd.hash and s.index in candidate.get('indices',[]) for v in candidate.get('source_warnings',[])})
        c={'doc_hash':sd.hash,'source_name':sd.name,'unit_id':'merged'+str(s.index),'start':s.index,'end':s.index+1,'indices':[s.index],
           'score':1.0,'literal_score':0,'content_score':0,'heading_score':1,'context_score':1,'table_score':0,'locator':' > '.join(s.path),
           'text':s.text,'exact':False,'unsupported':'','source_warnings':warnings,'scope_verified':True,
           'correspondence':{'method':'ordered-opening-and-complete-tail','opening':opening,'consumed_target':j}}
        p.update(candidates=[c],selected=0,status='auto',decision='accept',content_class='prospectus_source',
                 reason='同一标题内由完整起句与后续完整段落两个有序锚点确认合段；复制当前完整来源段，夹在两段间的对象另行保留核对。')
        op.update(candidates=[copy.deepcopy(c)],selected=0,status='auto',decision='accept',action='source_merge',merge_into=p['id'],
                  content_class='prospectus_source',reason='本段全文已被当前来源合入前段；随已登记的完整来源复制删除重复旧段，保留删除修订。')
        if warnings:
            for item in (p,op):item.update(selected=None,status='review',decision='pending',reason='已定位合段但对应来源存在一致性提示，未自动复制或删除：'+'；'.join(warnings))
        used.update((i,j))
    return result

def inventory(doc,props):
    entries=[]
    first_matter=min((m['start'] for m in doc.matters),default=0)
    for b in doc.blocks:
        owners=[p for p in props if p['start']<=b.index<p['end']]
        if len(owners)>1:raise ValueError('目标对象被多个计划重复占用：'+str(b.index))
        if owners:
            p=owners[0];state='copied' if p['decision'] in ('accept','same') else 'protected' if p['decision'] in ('protected','keep') else 'pending'
            reason=p['reason'];cls=p.get('content_class','source_missing_or_ambiguous');pid=p['id']
        else:
            cover=b.index<first_matter
            unsupported=('目标段落含节分界，正文与节设置保留待核对。' if b.el.find('.//'+w('sectPr')) is not None else complex_reference_reason([b.el]))
            state='pending' if unsupported and not independent_reason(b) else 'protected' if cover or b.protected or b.heading or b.is_toc or not b.text else 'pending'
            reason=unsupported if state=='pending' and unsupported else b.reason or ('封面、签字或项目参数保留；没有用户新参数时不改写' if cover else '章节框架或空段保留' if state=='protected' else '该正文对象未建立可靠来源范围，保留待核对')
            cls='independent_protected' if independent_reason(b) else 'framework' if cover or b.heading or b.is_toc or not b.text else 'unclassified_retained';pid=None
            if state=='pending' and not unsupported:
                state='protected';cls='no_correspondence_retained';reason=NO_CORRESPONDENCE
        entries.append({'block':b.index,'kind':'image' if b.images else b.kind,'matter':b.matter,'locator':' > '.join(b.path),
                        'status':state,'content_class':cls,'reason':reason,'proposal_id':pid,'text':b.text})
    return entries
