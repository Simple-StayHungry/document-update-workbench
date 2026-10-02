"""Recover missed source body content; workpaper headings remain untouched.

Exact text needs subject and adjacent-copy evidence. Procedures, independent
judgments, external evidence, comments and every original heading remain closed.
"""
from __future__ import annotations
import re
from .docxio import w, complex_reference_reason
from .matching import title_key
from .precision import fingerprint
from .source_policy import independent_reason


def _literal(value):
    return re.sub(r'\s+', '', value or '')


def _eligible(b):
    if b.section!='situ' or b.is_toc or b.heading or b.kind!='paragraph' or b.images or not b.text.strip():return False
    if independent_reason(b) or complex_reference_reason([b.el]) or b.el.find('.//'+w('sectPr')) is not None:return False
    if any(x in b.reason for x in ('批注','核查结构','事项标题','不适用','同业','比较分析','核查程序')):return False
    return (b.reason in ('公式或图示结构保留，未作自动事实核验','结构引导语保留，不单独替换或推定其下方事实已核验') or
            (b.reason=='独立核查或分析判断' and bool(re.match(r'^(综上|由此可见)',b.text))))


def _subject(b,s):
    if b.heading or s.heading:return False
    a=[title_key(x) for x in b.path if not re.search(r'核查(?:情况|记录|程序|结论)',x)]
    z=[title_key(x) for x in s.path if not re.search(r'核查(?:情况|记录|程序|结论)',x)]
    if not a or not z or a[-1]!=z[-1]:return False
    if len(a[-1])>=4:return True
    return bool(len(a)>=2 and len(z)>=2 and a[-2:]==z[-2:] and len(a[-2])>=4)


def _retained_heading_bridge(doc,sd,b,s,p,c,source_flags):
    """One unchanged heading may locate body text without becoming a copy.

    Both sides must be the same consecutive triplet: a complete, identical
    already-copied paragraph, an identical heading, then the candidate body.
    """
    if b.index<2 or s.index<2 or p['end']!=b.index-1 or c['end']!=s.index-1:return False
    if p['end']-p['start']!=1 or c['end']-c['start']!=1 or c.get('blocked_reason'):return False
    h,sh=doc.blocks[b.index-1],sd.blocks[s.index-1]
    if (not h.heading or not sh.heading or h.level!=sh.level or h.section!='situ' or
            h.matter!=b.matter or _literal(h.text)!=_literal(sh.text) or len(title_key(h.text))<8):return False
    if (h.text not in b.path or sh.text not in s.path or source_flags.get((sd.hash,sh.index)) or
            any(x.is_toc or x.images or complex_reference_reason([x.el]) or
                x.el.find('.//'+w('sectPr')) is not None or independent_reason(x) or
                any(reason in x.reason for reason in ('批注','核查结构','事项标题','核查程序')) for x in (h,sh))):return False
    before,source_before=doc.blocks[p['start']],sd.blocks[c['start']]
    if (any(x.heading or x.kind!='paragraph' or x.images or independent_reason(x) for x in (before,source_before)) or
            len(_literal(before.text))<20 or _literal(before.text)!=_literal(source_before.text)):return False
    return (doc.parents.get(before.el) is doc.parents.get(h.el) is doc.parents.get(b.el) and
            sd.parents.get(source_before.el) is sd.parents.get(sh.el) is sd.parents.get(s.el))


def recover_source_coverage(doc,sources,proposals,source_flags=None):
    result=list(proposals);source_flags=source_flags or {};catalog={}
    for sd in sources:
        if sd.profile.get('kind') not in ('prospectus','opinion') or sd.profile.get('issuer')!=doc.profile.get('issuer'):continue
        for b in sd.blocks:
            if b.heading or b.is_toc or b.kind!='paragraph' or b.images or not b.text.strip() or complex_reference_reason([b.el]):continue
            catalog.setdefault(_literal(b.text),[]).append((sd,b))
    changed=True
    while changed:
        changed=False;occupied={i for p in result for i in range(p['start'],p['end'])};anchors=[]
        for p in result:
            sel=p.get('selected')
            if p.get('decision') not in ('accept','same') or p.get('action')=='source_merge' or not isinstance(sel,int):continue
            c=p['candidates'][sel]
            if c.get('span') or c.get('projection') or c.get('source_warnings') or c.get('unsupported'):continue
            if c.get('indices',list(range(c['start'],c['end'])))!=list(range(c['start'],c['end'])):continue
            anchors.append((p,c))
        for b in doc.blocks:
            if b.index in occupied or not _eligible(b):continue
            matches=[]
            for sd,s in catalog.get(_literal(b.text),[]):
                if not _subject(b,s):continue
                evidence=[]
                for p,c in anchors:
                    if c['doc_hash']!=sd.hash or p.get('matter')!=(b.matter or '正文'):continue
                    if p['end']==b.index and c['end']==s.index:evidence.append({'side':'before','proposal':p['id']})
                    if p['start']==b.index+1 and c['start']==s.index+1:evidence.append({'side':'after','proposal':p['id']})
                    if _retained_heading_bridge(doc,sd,b,s,p,c,source_flags):
                        evidence.append({'side':'before','proposal':p['id'],'method':'identical-retained-heading',
                                         'target_heading':b.index-1,'source_heading':s.index-1})
                if evidence:matches.append((sd,s,evidence))
            if len(matches)!=1:continue
            sd,s,evidence=matches[0];warnings=sorted(set(source_flags.get((sd.hash,s.index),[])))
            c={'doc_hash':sd.hash,'source_name':sd.name,'source_kind':sd.profile.get('kind'),'unit_id':'coverage'+str(s.index),
               'start':s.index,'end':s.index+1,'indices':[s.index],'score':1.0,'literal_score':1.0,'content_score':1.0,
               'heading_score':1.0,'context_score':1.0,'table_score':0.0,'locator':' > '.join(s.path),'text':s.text,
               'exact':True,'unsupported':'','source_warnings':warnings,'scope_verified':True,
               'source_fingerprint':fingerprint([s.el]),
               'correspondence':{'method':'exact-object-adjacent-source-copy','anchors':evidence}}
            p={'id':doc.hash+':coverage:'+str(b.index),'target_id':doc.hash,'target_name':doc.name,
               'start':b.index,'end':b.index+1,'matter':b.matter or '正文','heading':b.path[-1] if b.path else b.text,
               'locator':' > '.join(b.path),'old_text':b.text,'status':'review' if warnings else 'auto',
               'decision':'pending' if warnings else 'accept','selected':None if warnings else 0,'candidates':[c],
               'reason':'正文原文、所属主题和紧邻来源复制范围共同定位；实际复制来源并保留修订。' if not warnings else '已定位原文但来源存在一致性提示：'+'；'.join(warnings),
               'content_class':'source_missing_or_ambiguous' if warnings else 'prospectus_source','kind':'material',
               'has_table':False,'table_mode':'copy','sync_mode':'block','identity':None,'note':''}
            result.append(p);occupied.add(b.index);changed=True
    return result


def recover_quoted_clause(doc,sources,proposals,source_flags=None):
    """Restore a missed quote opener only when the complete quoted tail was copied.

    This is not general punctuation stripping: a numbered clause, its source
    opener and every following block through the old closing quote must form one
    consecutive, already sourced range. Copy the real opener, never retype it.
    """
    result=list(proposals);source_flags=source_flags or {}
    close={'“':'”','「':'」','『':'』'}
    by_start={p['start']:p for p in result if p['end']==p['start']+1}
    for p in list(result):
        if p.get('content_class')!='no_correspondence_retained' or p['end']!=p['start']+1:continue
        b=doc.blocks[p['start']];value=b.text.strip()
        if len(value)<40 or not re.match(r'^[“「『]\d+(?:\.\d+)+',value) or value[-1:] not in ('：',':'):continue
        if b.section!='situ' or b.images or independent_reason(b) or complex_reference_reason([b.el]) or '批注' in b.reason:continue
        matches=[]
        for sd in sources:
            if sd.profile.get('kind') not in ('prospectus','opinion') or sd.profile.get('issuer')!=doc.profile.get('issuer'):continue
            for s in sd.blocks:
                if _literal(s.text)!=_literal(value[1:]) or s.images or complex_reference_reason([s.el]):continue
                chain=[];closed=False
                for offset in range(1,21):
                    if b.index+offset>=len(doc.blocks) or s.index+offset>=len(sd.blocks):break
                    tail=doc.blocks[b.index+offset];sp=sd.blocks[s.index+offset];owner=by_start.get(tail.index)
                    if tail.heading or tail.matter!=b.matter or not owner or owner.get('decision') not in ('accept','same'):break
                    sel=owner.get('selected');c=owner['candidates'][sel] if isinstance(sel,int) else None
                    if not c or c['doc_hash']!=sd.hash or c.get('indices')!=[sp.index] or c.get('span') or c.get('source_warnings') or c.get('unsupported'):break
                    chain.append(owner['id'])
                    if tail.text.rstrip().endswith(close[value[0]]):closed=True;break
                if closed and len(chain)>=2:matches.append((sd,s,chain))
        if len(matches)!=1:continue
        sd,s,chain=matches[0];warnings=sorted(set(source_flags.get((sd.hash,s.index),[])))
        c={'doc_hash':sd.hash,'source_name':sd.name,'source_kind':sd.profile.get('kind'),'unit_id':'quoted'+str(s.index),
           'start':s.index,'end':s.index+1,'indices':[s.index],'span':None,'score':1.0,'literal_score':1.0,'content_score':1.0,
           'heading_score':1.0,'context_score':1.0,'table_score':0.0,'locator':' > '.join(s.path),'text':s.text,
           'exact':False,'unsupported':'','source_warnings':warnings,'scope_verified':True,'source_fingerprint':fingerprint([s.el]),
           'correspondence':{'method':'numbered-quote-complete-copied-tail','tail_proposals':chain}}
        updated=dict(p,candidates=[c],status='review' if warnings else 'auto',decision='pending' if warnings else 'accept',
                     selected=None if warnings else 0,content_class='source_missing_or_ambiguous' if warnings else 'prospectus_source',
                     reason='条款编号、完整引导句及截至旧闭引号的连续来源尾段已对应，补复制完整来源引导段。' if not warnings else '对应来源存在一致性提示：'+'；'.join(warnings))
        updated.pop('correspondence_search',None)
        result[result.index(p)]=updated
    return result
