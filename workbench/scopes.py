"""Bounded source-section reconciliation, separate from sentence similarity.

A whole leaf may change its words and grow. We still need a unique structural
address, complete source boundaries and compatible subject/table identities.
Uncertain candidates never become auto replacements by lowering a score.
"""
from __future__ import annotations
import re,copy,collections
from .model import Unit,strip_lead,SECTION,MATTER
from .docxio import w,text,complex_reference_reason
from .precision import descriptor_guard,nearest_scope,atomic_units
from .matching import normal,title_key,ratio,skeleton,subject_guard

GENERIC=re.compile(r'^(?:基本情况|经营情况|业务概况|主要业务|核查情况|情况|其他|说明|相关情况|基本信息|核查记录|分析)$')

def key(t):return title_key(t)
def informative(t):
    v=key(t)
    return len(v)>=4 and not GENERIC.fullmatch(strip_lead(t)) and not SECTION.match(t.replace(' ',''))

def scope_end(doc,heading):
    """Same-level lexical numbering already overrides corrupt Word style levels."""
    i=heading.index+1
    while i<len(doc.blocks):
        b=doc.blocks[i]
        if b.is_toc:return i
        if b.matter and heading.matter and b.matter!=heading.matter:return i
        if b.heading and b.level is not None and b.level<=heading.level:return i
        if doc.parents.get(b.el) is not doc.parents.get(heading.el):return i
        i+=1
    return i

def get_source_scopes(doc):
    result=[]
    for h in doc.blocks:
        if not h.heading or h.is_toc:continue
        end=scope_end(doc,h)
        bs=doc.blocks[h.index+1:end]
        while bs and not (bs[-1].text or bs[-1].images or complex_reference_reason([bs[-1].el])):bs=bs[:-1]
        if not bs or not any(b.text for b in bs):continue
        result.append((h,Unit('scope'+str(h.index),h.index+1,bs[-1].index+1,bs,h.text,h.path,h.matter)))
    return result

def compatible_tables(target,source):
    # Old tables must each have an ordered corresponding complete table. Added
    # source tables inside a verified section are included, never silently lost.
    def singles(u):
        out=[]
        for n,b in enumerate(u.blocks):
            if b.kind!='table':continue
            lo=n
            while lo and re.match(r'^(?:表[:：]|单位[:：])',u.blocks[lo-1].text):lo-=1
            bs=u.blocks[lo:n+1];out.append(Unit('t',bs[0].index,b.index+1,bs,u.heading,b.path,u.matter))
        return out
    tt,ss=singles(target),singles(source)
    if len(tt)>len(ss):return False
    cursor=0
    for t in tt:
        found=False
        while cursor<len(ss):
            s=ss[cursor];cursor+=1
            if not descriptor_guard(t,s) and ratio(skeleton(t.schema),skeleton(s.schema))>=.70:
                found=True;break
        if not found:return False
    return True

def _outside_context(h):
    return [key(v) for v in h.path[:-1] if informative(v)]

def reconcile(doc,index,props):
    """Consolidate only complete, unprotected factual leaves into source scopes.

    Exact old matches act as anchors, NOT evidence that the current source has
    no extra paragraphs. A source scope is always read again in its entirety.
    """
    catalog=[(sd,h,u) for sd in index.documents for h,u in get_source_scopes(sd)]
    newprops=list(props);notes=[]
    for group in doc.units:
        original=[p for p in newprops if group.start<=p['start'] and p['end']<=group.end and p['kind']=='material']
        if not original:continue
        # Only a whole leaf is eligible. A protected judgment in the middle
        # splits a group and must never be swallowed by a section replacement.
        hi=next((b for b in reversed(doc.blocks[:group.start]) if b.heading),None)
        if not hi or not informative(hi.text) or MATTER.match(hi.text):continue
        before=doc.blocks[hi.index+1:group.start]
        after=doc.blocks[group.end:scope_end(doc,hi)]
        if any(b.text or b.images for b in before+after):continue
        if any(b.heading or b.protected or b.images for b in group.blocks):continue
        # An opinion may establish permission for one previously independent
        # paragraph.  That narrow correspondence is not permission to replace
        # the rest of its section with an entire formal-source scope.
        if any(getattr(b,'formal_source_correspondence',None) for b in group.blocks):continue
        if any(p['decision']=='protected' for p in original):continue
        # A quoted subclause is not a request to replace its entire chapter.
        if group.text.lstrip().startswith(('“','「','『','"')) or any(getattr(b,'quoted',False) for b in group.blocks):continue
        matter_title=next((m['title'] for m in doc.matters if m['id']==group.matter),'')
        narrow_counterparties=bool(re.search(r'主要客户|主要供应商|前五大|客户和供应商|客户及供应商',matter_title))
        strong_anchors=[p for p in original if p['candidates'] and p['candidates'][0].get('score',0)>=.82]
        candidates=[]
        for sd,sh,su in catalog:
            exact_heading=key(hi.text)==key(sh.text)
            # A renamed heading is recoverable only through TWO independent,
            # ordered, unchanged body anchors. Numbers aren't counted as anchors.
            anchor_hits=[]
            for p in strong_anchors:
                for c in p['candidates']:
                    if c['doc_hash']==sd.hash and su.start<=c['start'] and c['end']<=su.end and c.get('exact') and not c.get('span') and not p.get('has_table') and c['end']-c['start']==1 and len(normal(p['old_text']))>=20:
                        anchor_hits.append(c['start'])
            renamed=not exact_heading and len(set(anchor_hits))>=2 and anchor_hits==sorted(anchor_hits)
            if not (exact_heading or renamed):continue
            if narrow_counterparties and not re.search(r'客户|供应商',sh.text):continue
            # A source section already assigned to sibling template ranges is
            # deliberately split, not a renamed complete section. Do not paste
            # the same corporate profile/financial table into both headings.
            siblings=[p for p in props if p['matter']==group.matter and not (group.start<=p['start'] and p['end']<=group.end)]
            if any(c['doc_hash']==sd.hash and c.get('score',0)>=.9 and su.start<=c['start'] and c['end']<=su.end for p in siblings for c in p['candidates'][:1]):continue
            context=max((ratio(a,b) for a in _outside_context(hi) for b in _outside_context(sh)),default=0)
            anchor=any(c['doc_hash']==sd.hash and su.start<=c['start'] and c['end']<=su.end for p in strong_anchors for c in p['candidates'] if c.get('score',0)>=.78)
            if not anchor and not (exact_heading and context>=.82):continue
            if subject_guard(group,su):continue
            # Headings only locate corresponding content; scope expansion must
            # not import source subheadings into the preserved workpaper outline.
            if any(b.is_toc or b.images or b.heading or b.protected for b in su.blocks):continue
            unsupported=complex_reference_reason([b.el for b in su.blocks])
            if unsupported:continue
            if not compatible_tables(group,su):continue
            # A matching heading elsewhere in a large document isn't sufficient.
            if not anchor and context<.90 and ratio(skeleton(group.text),skeleton(su.text))<.30:continue
            warnings=sorted({v for b in su.blocks for v in index.source_flags.get((sd.hash,b.index),[])})
            candidates.append({'doc_hash':sd.hash,'source_name':sd.name,'unit_id':su.id,'start':su.start,'end':su.end,'indices':[b.index for b in su.blocks],
                'score':.98,'content_score':round(ratio(skeleton(group.text),skeleton(su.text)),4),'literal_score':round(ratio(normal(group.text),normal(su.text)),4),
                'heading_score':1.0 if exact_heading else 0.0,'table_score':1.0,'locator':su.locator,'text':su.text,'exact':normal(group.text)==normal(su.text),
                'unsupported':'','source_warnings':warnings,'table_count':sum(b.kind=='table' for b in su.blocks),'scope_verified':True,'scope_heading':sh.text,'scope_heading_index':sh.index,'renamed_heading':renamed})
        if not candidates:continue
        if any(not c['renamed_heading'] for c in candidates):
            candidates=[c for c in candidates if not c['renamed_heading']]
        else:
            # Ancestors also contain both anchors; use the smallest enclosing
            # scope per source rather than pretending an ancestor was renamed.
            candidates=[c for c in candidates if not any(d['doc_hash']==c['doc_hash'] and c['start']<=d['start'] and d['end']<=c['end'] and (d['start'],d['end'])!=(c['start'],c['end']) for d in candidates)]
        # Correspondence and copying are separate.  Once the structural address is
        # found, the source scope is copied wholesale even when every visible word
        # is unchanged.  Do not keep black target paragraphs just because they
        # already look the same.
        kind_by_hash={sd.hash:sd.profile.get('kind','unknown') for sd in index.documents}
        for c in candidates:c['source_kind']=kind_by_hash.get(c['doc_hash'],'unknown')

        # The source index can expose a full leaf and shorter fragments under the
        # same visible heading.  A shorter contained fragment is not a competing
        # version; keep the complete source scope.
        filtered=[]
        for c in candidates:
            nc=normal(c['text'])
            contained=False
            for d in candidates:
                if c is d or c['doc_hash']!=d['doc_hash']:continue
                if key(c.get('scope_heading',''))!=key(d.get('scope_heading','')):continue
                nd=normal(d['text'])
                if len(nd)>len(nc) and nc and nc in nd:
                    contained=True;break
            if not contained:filtered.append(c)
        candidates=filtered or candidates

        # Group identical visible source contents across documents.  The same text
        # repeated in a prospectus and an underwriter opinion is confirmation of
        # one correspondence, not an ambiguity.
        groups=collections.defaultdict(list)
        # Corresponding formal sources have equal retrieval standing.  When the
        # established topic contains different wording, the user's authority rule
        # gives the prospectus precedence, irrespective of which old draft is
        # textually closer.  Keep alternatives as provenance, not silent choices.
        prospective=[c for c in candidates if c.get('source_kind')=='prospectus']
        conflicts=bool(prospective and any(c.get('source_kind')=='opinion' and
                       all(normal(c['text'])!=normal(p['text']) for p in prospective) for c in candidates))
        preferred=prospective if conflicts else candidates
        for c in preferred:groups[normal(c['text'])].append(c)
        grouped=[]
        for body,items in groups.items():
            grouped.append({
                'body':body,'items':items,
                'docs':len({x['doc_hash'] for x in items}),
                'exact':max(int(x.get('exact',False)) for x in items),
                'literal':max(x.get('literal_score',0) for x in items),
                'content':max(x.get('content_score',0) for x in items),
                'heading':max(x.get('heading_score',0) for x in items),
                'score':max(x.get('score',0) for x in items),
            })
        grouped.sort(key=lambda g:(-g['exact'],-g['docs'],-g['literal'],-g['content'],-g['heading'],-g['score'],-len(g['body'])))

        chosen=None
        if len(grouped)==1:
            chosen=grouped[0]
        elif grouped[0]['exact']>grouped[1]['exact']:
            # The old workpaper itself identifies which source lineage it came
            # from; use that lineage, without judging substantive wording.
            chosen=grouped[0]
        elif grouped[0]['docs']>=2 and grouped[0]['docs']>grouped[1]['docs']:
            # The same complete content occurs independently in more source files.
            chosen=grouped[0]
        elif grouped[0]['heading']>=.985 and grouped[0]['literal']>=.90 and grouped[0]['literal']-grouped[1]['literal']>=.05:
            chosen=grouped[0]

        resolved=chosen is not None
        if resolved:
            # Pick one physical source only for OOXML copy/relationships.  For a
            # workpaper, an opinion is the closest document family when content is
            # otherwise identical; exact lineage still wins first.
            items=chosen['items']
            items.sort(key=lambda c:(-int(c.get('exact',False)),-c.get('literal_score',0),-c.get('content_score',0),
                0 if (doc.profile.get('kind')=='workpaper' and c.get('source_kind')=='opinion') else 1,
                0 if c.get('source_kind')=='prospectus' else 1,c['source_name'],c['start']))
            top=items[0]
            # Present the selected correspondence first but keep alternatives in
            # the UI for traceability.
            rest=[c for c in candidates if c is not top]
            candidates=[top]+rest
        else:
            candidates.sort(key=lambda c:(-int(c.get('exact',False)),-c.get('literal_score',0),-c.get('content_score',0),-c.get('heading_score',0),-c.get('score',0),c['source_name'],c['start']))
            top=candidates[0]

        selected_indices={i for p in original for c in p['candidates'][:1] if c['doc_hash']==top['doc_hash'] for i in c['indices']}
        extra=[i for i in top['indices'] if i not in selected_indices]
        reason='已按标题、编号和上下文定位完整对应来源；不比较新旧是否相同，直接复制完整来源范围。'
        if top['renamed_heading']:reason='标题表述有变化，但正文锚点已唯一定位对应来源；直接复制完整来源范围。'
        if not resolved:
            reason='同一结构位置仍对应多个不同来源版本，无法仅靠标题和上下文确定唯一来源；此处不擅自选择。'
        if top.get('source_warnings'):
            resolved=False;reason='完整来源事项含一致性提示，先核实该提示；未自动覆盖。'
        st='auto' if resolved else 'review'
        base=copy.deepcopy(original[0]);base.update(id=doc.hash+':scope'+str(group.start),start=group.start,end=group.end,heading=hi.text,locator=group.locator,
            old_text=group.text,candidates=candidates[:4],status=st,reason=reason,decision='accept' if st=='auto' else 'pending',
            selected=0 if resolved else None,has_table=bool(top['table_count']) or group.has_table,table_count=top['table_count'],
            scope_mode='完整事项',scope_added_blocks=len(extra),table_mode='copy',identity=None)
        if conflicts:
            base['source_precedence']={'rule':'同一事项内容冲突时以募集说明书为准','preferred':top['doc_hash'],
                'other':[{'doc_hash':c['doc_hash'],'start':c['start'],'end':c['end']} for c in candidates if c.get('source_kind')=='opinion']}
        newprops=[p for p in newprops if p not in original]+[base]
        notes.append({'rule':'source-scope','title':hi.text,'mode':'renamed' if top['renamed_heading'] else 'bounded','status':st,'source_blocks':len(top['indices']),'old_blocks':group.end-group.start})
    return sorted(newprops,key=lambda p:p['start']),notes


def source_gaps(doc,index,props):
    """Report uncopied content in structurally matched sections, not 'new facts'.

    Without a previous prospectus we cannot assert that any such content was
    recently added. The report states only that these current source blocks are
    not covered by the current replacement plan.
    """
    output=[];seen=set();source_by_hash={sd.hash:sd for sd in index.documents}
    catalog=[(sd,h,u) for sd in index.documents for h,u in get_source_scopes(sd)]
    for group in doc.units:
        hi=next((b for b in reversed(doc.blocks[:group.start]) if b.heading),None)
        if not hi or not informative(hi.text) or MATTER.match(hi.text):continue
        relevant=[p for p in props if p['matter']==group.matter]
        current=[p for p in props if p['start']<=group.start and group.end<=p['end'] and p.get('scope_mode')]
        if current:continue  # Entire scope is already copied OR explicitly pending.
        for sd,sh,su in catalog:
            exact_heading=key(hi.text)==key(sh.text)
            selected=[c for p in relevant if p.get('decision') in ('accept','same') and isinstance(p.get('selected'),int)
                      for c in [p['candidates'][p['selected']]] if c['doc_hash']==sd.hash and c['score']>=.87 and
                      su.start<=c['start'] and c['end']<=su.end]
            if not exact_heading:
                # Changed headings with imperfect body anchors are NOT silently
                # treated as complete. Report, don't auto-expand, when two
                # different target paragraphs point into the same parent scope.
                context=max((ratio(a,b) for a in _outside_context(hi) for b in _outside_context(sh)),default=0)
                if context<.90:continue
                hints={p['start'] for p in relevant if group.start<=p['start']<group.end and not p.get('has_table') and len(normal(p['old_text']))>=20 for c in p.get('candidates',[]) if c['doc_hash']==sd.hash and c.get('score',0)>=.82 and su.start<=c['start'] and c['end']<=su.end}
                if context<.90 or len(hints)<2 or subject_guard(group,su):continue
            elif not selected:continue
            covered={k for c in selected for k in c['indices']}
            # Workpaper headings and numbering are deliberately retained. They
            # are locators, not required source-copy coverage or missing content.
            # The same source paragraph can be repeated elsewhere in the current
            # prospectus.  Copying its identical body in this target topic already
            # fulfils coverage; do not report that duplicate as a new disclosure.
            copied_text={normal(source_by_hash[c['doc_hash']].blocks[i].text) for p in relevant if p.get('decision') in ('accept','same')
                         and isinstance(p.get('selected'),int) for c in [p['candidates'][p['selected']]]
                         if c['doc_hash'] in source_by_hash and not c.get('span') and not c.get('projection')
                         for i in c.get('indices',range(c['start'],c['end'])) if source_by_hash[c['doc_hash']].blocks[i].text}
            gaps=[b for b in su.blocks if not b.heading and b.index not in covered and
                  (b.text or b.images or complex_reference_reason([b.el])) and
                  not (b.kind in ('paragraph','table') and b.text and not b.images and not complex_reference_reason([b.el]) and
                       normal(b.text) in copied_text)]
            if not gaps:continue
            marker=(group.matter,sd.hash,sh.index)
            if marker in seen:continue
            seen.add(marker)
            excerpt='\n'.join((b.text or b.reason or '图形或公式对象')[:600] for b in gaps[:3])
            output.append({'rule':'source-coverage-gap','severity':'warning','title':'对应来源仍有未纳入的内容','locator':group.locator,'block':group.start,
                'detail':('标题已有变化，正文线索指向可能对应的范围，但不足以自动扩大替换。' if not exact_heading else '')+'来源的对应标题下还有 '+str(len(gaps))+' 个段落或对象未进入本次更新范围。可能是新增披露，也可能不属于原稿摘录范围；不能仅因旧句已匹配就认定完整。',
                'evidence':[{'file_id':sd.hash,'file':sd.name,'block':gaps[0].index,'locator':su.locator,'text':excerpt}]})
    return output
