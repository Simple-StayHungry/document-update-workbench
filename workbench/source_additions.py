"""Copy missing body ranges only inside a proved, unchanged workpaper topic.

An addition is a zero-width source insertion.  The source heading is a locator,
never inserted content.  Similar words alone cannot create an insertion: two
consecutive copy anchors, or a copied leaf plus its exact heading boundary, must
prove both location and reading order.  Rejected ranges remain in diagnostics.
"""
from __future__ import annotations
import collections
from .docxio import w, complex_reference_reason
from .matching import title_key, normal
from .precision import fingerprint
from .scopes import scope_end, informative


def _visible(block):
    return bool(block.text.strip() or block.images or complex_reference_reason([block.el]))


def _heading(doc, index):
    return next((b for b in reversed(doc.blocks[:index]) if b.heading and not b.is_toc), None)


def _leaf_end(doc, heading):
    end=scope_end(doc,heading)
    while end>heading.index+1 and not _visible(doc.blocks[end-1]):end-=1
    return end


def _same_topic(target_heading, source_heading):
    return bool(target_heading and source_heading and informative(target_heading.text) and
                title_key(target_heading.text)==title_key(source_heading.text))


def _anchors(doc,sources,proposals):
    by_hash={s.hash:s for s in sources if s.profile.get('kind') in ('prospectus','opinion') and
             s.profile.get('issuer')==doc.profile.get('issuer')}
    result=[]
    for p in sorted(proposals,key=lambda p:(p['start'],p['end'],p['id'])):
        sel=p.get('selected')
        if (p.get('kind')!='material' or p.get('action') in ('source_merge','source_insert') or
                p.get('decision') not in ('accept','same') or not isinstance(sel,int) or
                p['start']>=p['end'] or not 0<=sel<len(p.get('candidates',[]))):continue
        c=p['candidates'][sel];sd=by_hash.get(c['doc_hash'])
        if not sd or any(c.get(k) for k in ('span','projection','unsupported','source_warnings','blocked_reason')):continue
        if c.get('indices',list(range(c['start'],c['end'])))!=list(range(c['start'],c['end'])):continue
        if any(b.heading or b.is_toc for b in doc.blocks[p['start']:p['end']]):continue
        if any(b.heading or b.is_toc for b in sd.blocks[c['start']:c['end']]):continue
        th,sh=_heading(doc,p['start']),_heading(sd,c['start'])
        if not th or not sh:continue
        if any(_heading(doc,i) is not th for i in range(p['start'],p['end'])):continue
        if any(_heading(sd,i) is not sh for i in range(c['start'],c['end'])):continue
        result.append((p,c,sd,th,sh))
    return result


def recover_source_additions(doc,sources,proposals,source_flags=None):
    """Return existing plans plus additions, and a complete proposed-gap ledger.

    Each new plan has ``action=source_insert`` and ``start==end``.  The insertion
    is immediately before that original block (or after the final block at EOF).
    ``insert_before``/``insert_after`` and their fingerprints freeze both sides.
    Export must retain these original objects and track only the new source copy.
    """
    result=list(proposals);flags=source_flags or {};anchors=_anchors(doc,sources,proposals)
    gaps=[];seen=set()
    def add(p,c,sd,th,sh,lo,hi,position,method,evidence):
        while lo<hi and not _visible(sd.blocks[lo]):lo+=1
        while hi>lo and not _visible(sd.blocks[hi-1]):hi-=1
        if lo>=hi:return
        key=(position,sd.hash,lo,hi)
        if key in seen:return
        seen.add(key)
        gaps.append({'position':position,'source':sd,'heading':th,'source_heading':sh,
                     'start':lo,'end':hi,'matter':p.get('matter','正文'),
                     'method':method,'anchors':evidence})

    for i,(p,c,sd,th,sh) in enumerate(anchors):
        if not _same_topic(th,sh):continue
        # Direct neighbors in BOTH the target topic and selected source lineage.
        for q,qc,qd,qh,qsh in anchors[i+1:]:
            if q['start']>p['end']:break
            if (p['end']==q['start'] and qd.hash==sd.hash and qh is th and qsh is sh and
                    p.get('matter')==q.get('matter') and c['end']<qc['start']):
                add(p,c,sd,th,sh,c['end'],qc['start'],p['end'],'two-adjacent-source-anchors',[p['id'],q['id']])
        # For a single-ended gap, every visible target body in the leaf must
        # already be covered by a source-copy anchor in this same source lineage.
        te,se=_leaf_end(doc,th),_leaf_end(sd,sh)
        target_body=doc.blocks[th.index+1:te];source_body=sd.blocks[sh.index+1:se]
        if any(b.heading for b in target_body+source_body):continue
        covered={k for q,qc,qd,qh,qsh in anchors if qd.hash==sd.hash and qh is th and qsh is sh
                 for k in range(q['start'],q['end'])}
        if any(_visible(b) and b.index not in covered for b in target_body):continue
        if p['start']==th.index+1 and c['start']>sh.index+1:
            add(p,c,sd,th,sh,sh.index+1,c['start'],p['start'],'exact-leaf-prefix',[p['id']])
        if p['end']==te and c['end']<se:
            add(p,c,sd,th,sh,c['end'],se,p['end'],'exact-leaf-suffix',[p['id']])

    # Some templates omit source subheadings while already copying the complete
    # corresponding body.  A missing leading table is still addressable without
    # importing that heading: the preceding copied section ends at its boundary,
    # and TWO consecutive identical paragraphs identify its entire remaining leaf.
    # This is deliberately narrower than a heading/similarity-based expansion.
    for i in range(1,len(anchors)-1):
        previous,pc,pd,ph,psh=anchors[i-1]
        p,c,sd,th,sh=anchors[i]
        following,fc,fd,fh,fsh=anchors[i+1]
        if not (pd.hash==sd.hash==fd.hash and ph is th is fh and sh is fsh and psh is not sh):continue
        if _same_topic(th,sh):continue  # The ordinary exact-leaf path handles it.
        if previous.get('matter')!=p.get('matter') or following.get('matter')!=p.get('matter'):continue
        if not (previous['end']==p['start'] and p['end']==following['start'] and
                pc['end']==sh.index and c['end']==fc['start'] and fc['end']==_leaf_end(sd,sh)):continue
        if c['start']<=sh.index+1:continue
        if any(len(normal(q['old_text']))<40 or normal(q['old_text'])!=normal(qc['text'])
               for q,qc in ((p,c),(following,fc))):continue
        prefix=sd.blocks[sh.index+1:c['start']]
        if not any(b.kind=='table' for b in prefix):continue
        if any(b.heading or b.is_toc for b in prefix):continue
        add(p,c,sd,th,sh,sh.index+1,c['start'],p['start'],
            'complete-leaf-table-prefix-three-anchors',[previous['id'],p['id'],following['id']])

    ledger=[];eligible=[]
    for g in gaps:
        sd=g['source'];bs=sd.blocks[g['start']:g['end']];pos=g['position'];why=[]
        if any(b.heading or b.is_toc for b in bs):why.append('来源范围包含标题或目录，不能带入原稿标题框架')
        if any(b.images for b in bs):why.append('新增范围包含图像，需单独图文对应与对象兼容验证')
        unsupported=complex_reference_reason([b.el for b in bs])
        if unsupported:why.append(unsupported)
        if any(b.el.find('.//'+w('sectPr')) is not None for b in bs):why.append('来源范围包含节分界')
        if len({id(sd.parents.get(b.el)) for b in bs})!=1:why.append('来源对象不在同一父节点')
        sides=[doc.blocks[k] for k in (pos-1,pos) if 0<=k<len(doc.blocks)]
        if len({id(doc.parents.get(b.el)) for b in sides})>1:why.append('插入边界跨越原稿父节点')
        warnings=sorted({x for b in bs for x in flags.get((sd.hash,b.index),[])})
        if warnings:why.append('来源存在一致性提示：'+'；'.join(warnings))
        # Repeated units such as “单位：万元” accompany different complete tables;
        # they do not establish that the new table/body was already copied.
        texts={normal(b.text) for b in bs if b.text.strip() and (b.kind=='table' or len(normal(b.text))>=15)}
        originals={normal(b.text) for b in doc.blocks if (b.matter or '正文')==g['matter'] and b.text.strip()}
        used={i for p in proposals if p.get('matter')==g['matter'] and p.get('decision') in ('accept','same')
              and isinstance(p.get('selected'),int) for c in [p['candidates'][p['selected']]]
              if c['doc_hash']==sd.hash for i in c.get('indices',range(c['start'],c['end']))}
        if texts & originals:why.append('原稿同一事项已有相同正文，不能重复新增')
        if used & set(range(g['start'],g['end'])):why.append('该来源范围已进入同一事项的其他复制计划')
        if any(p.get('action')=='source_insert' and p['start']==pos for p in proposals):why.append('该插入边界已有来源新增计划')
        record={'rule':'source-addition','target_block':pos,'target_heading':g['heading'].text,
                'source_name':sd.name,'source_hash':sd.hash,'source_kind':sd.profile.get('kind'),
                'source_heading':g['source_heading'].text,'source_indices':list(range(g['start'],g['end'])),
                'source_text':'\n'.join(b.text for b in bs),'method':g['method'],'anchor_proposals':g['anchors'],
                'status':'retained' if why else 'candidate','reason':'；'.join(why)}
        ledger.append(record);g['record']=record
        if not why:eligible.append(g)

    by_position=collections.defaultdict(list)
    for g in eligible:by_position[g['position']].append(g)
    for pos,items in by_position.items():
        # User-directed authority rule: corresponding prospectus wording wins.
        preferred=[g for g in items if g['source'].profile.get('kind')=='prospectus'] or items
        bodies={normal(g['record']['source_text']) for g in preferred}
        if len(bodies)>1:
            for g in items:g['record'].update(status='retained',reason='同一插入边界存在不同来源范围，未唯一定位')
            continue
        chosen=sorted(preferred,key=lambda g:(g['source'].name,g['start'],g['end']))[0]
        for g in items:
            if g is not chosen:g['record'].update(status='covered-by-selected-source',reason='同一位置只复制选定正式来源一次；冲突时以募集说明书为准')
        sd=chosen['source'];bs=sd.blocks[chosen['start']:chosen['end']];h=chosen['heading'];sh=chosen['source_heading']
        c={'doc_hash':sd.hash,'source_name':sd.name,'source_kind':sd.profile.get('kind'),
           'unit_id':'addition'+str(chosen['start']),'start':chosen['start'],'end':chosen['end'],
           'indices':[b.index for b in bs],'score':1.0,'literal_score':1.0,'content_score':1.0,
           'heading_score':1.0,'table_score':1.0 if any(b.kind=='table' for b in bs) else 0.0,
           'locator':' > '.join(bs[0].path),'text':'\n'.join(b.text for b in bs),'exact':False,
           'unsupported':'','source_warnings':[],'scope_verified':True,'source_fingerprint':fingerprint([b.el for b in bs]),
           'correspondence':{'method':chosen['method'],'anchors':chosen['anchors'],
                             'target_heading':h.index,'source_heading':sh.index}}
        anchor_sources={}
        for anchor_id in chosen['anchors']:
            anchor=next(p for p in proposals if p['id']==anchor_id)
            ac=anchor['candidates'][anchor['selected']]
            anchor_sources[anchor_id]={'doc_hash':ac['doc_hash'],'start':ac['start'],'end':ac['end'],
                                       'indices':ac.get('indices',list(range(ac['start'],ac['end'])))}
        proposal={'id':doc.hash+':addition:'+str(pos)+':'+sd.hash[:12]+':'+str(chosen['start']),
                  'target_id':doc.hash,'target_name':doc.name,'start':pos,'end':pos,
                  'insert_before':pos if pos<len(doc.blocks) else None,'insert_after':pos-1 if pos else None,
                  'insert_before_fingerprint':fingerprint([doc.blocks[pos].el]) if pos<len(doc.blocks) else None,
                  'insert_after_fingerprint':fingerprint([doc.blocks[pos-1].el]) if pos else None,
                  'matter':chosen['matter'],'heading':h.text,'locator':' > '.join(h.path),
                  'old_text':'','status':'auto','decision':'accept','selected':0,'candidates':[c],
                  'action':'source_insert','anchor_sources':anchor_sources,'kind':'material','content_class':'formal_source_addition',
                  'reason':'对应原标题及相邻来源正文已确认；将同一科目内缺少的完整来源正文以新增修订插入，原标题不变。',
                  'has_table':any(b.kind=='table' for b in bs),'table_mode':'copy','sync_mode':'block','identity':None,'note':''}
        result.append(proposal);chosen['record'].update(status='added',reason=proposal['reason'],proposal_id=proposal['id'])
    return sorted(result,key=lambda p:(p['start'],p['end'],p['id'])),ledger
