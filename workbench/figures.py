"""Conservative, source-backed proposals for captioned raster figure groups.

Only Documents supplied as current sources are searched. A caption is a locator,
not ownership evidence: issuer, local heading, adjacent prose and real package
relationships must agree before a prospectus figure can be copied. Media equality
never causes a skip. External workpaper evidence is retained before matching.
"""
from __future__ import annotations

import hashlib
import json
import posixpath
import re
import unicodedata
from difflib import SequenceMatcher

from .docxio import R, w, local, complex_reference_reason
from .model import strip_lead, SECTION, MATTER
from .precision import fingerprint

_CAPTION = re.compile(r'^\s*图(?:\s*[0-9一二三四五六七八九十]+(?:[-－.．][0-9]+)*)?\s*[:：、.．]\s*\S')
_NOTE = re.compile(r'^\s*(?:数据来源|资料来源|来源|注(?:[0-9一二三四五六七八九十]+)?|说明)\s*[:：]')
_EXTERNAL = re.compile(r'企查查|天眼查|信用中国|裁判文书|执行信息|公示系统|征信报告|(?:官网|网站|系统|网页).{0,12}(?:查询|截图|核查)|(?:查询|截图).{0,12}(?:官网|网站|系统|网页)|查询结果|核查截图|外部证据')
_MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
_SIGNATURE = re.compile(r'^\s*(?:调查人|核查人|调查人员|核查人员|项目负责人|复核人)\s*签字\s*[:：]?\s*$')


def _key(value):
    return re.sub(r'[\s:：,，.。；;（）()“”\"、]', '', unicodedata.normalize('NFKC', value or ''))


def _caption_key(value):
    return _key(re.sub(r'^\s*图(?:\s*[0-9一二三四五六七八九十]+(?:[-－.．][0-9]+)*)?\s*[:：、.．]', '', value or ''))


def _is_object(block):
    return bool(block.images or any(local(e) in ('chart', 'oMath', 'oMathPara', 'wgp', 'grpSp', 'group', 'relIds', 'object', 'OLEObject') for e in block.el.iter()))


def _has_native_math(block):
    return any(e.tag in ('{'+_MATH+'}oMath', '{'+_MATH+'}oMathPara') for e in block.el.iter())


def _math_signature(block):
    """Compare the formula tree, not a flattened string or its visual styling."""
    def semantic(node):
        tag=local(node)
        if tag in ('rPr', 'ctrlPr'):return None
        children=[value for child in node if (value:=semantic(child)) is not None]
        return (node.tag, tuple(sorted(node.attrib.items())), node.text or '', children)
    roots=[semantic(e) for e in block.el.iter() if e.tag=='{'+_MATH+'}oMath']
    return hashlib.sha256(json.dumps(roots,ensure_ascii=False,separators=(',',':')).encode()).hexdigest()


def _formula_context(doc, block):
    headings=[]
    for value in block.path:
        if SECTION.match(value.replace(' ','')) or MATTER.match(value):continue
        clean=_key(strip_lead(value))
        if len(clean)>=4:headings.append(clean)
    neighbors={}
    for name,step in (('before',-1),('after',1)):
        for index in range(block.index+step,block.index+step*9,step):
            if not 0<=index<len(doc.blocks):break
            other=doc.blocks[index]
            if other.heading or other.matter!=block.matter:break
            value=_key(other.text)
            if len(value)>=12:
                neighbors[name]={'block':index,'text':other.text,'key':value};break
    return {'headings':headings[-2:], 'neighbors':neighbors}


def _native_formula_candidates(doc, block, source_docs, source_flags):
    """Locate one complete formula paragraph by scope and two-sided anchors.

    Formula equality alone is not ownership. A changed formula can still be
    copied when both independent adjacent texts identify one source position.
    At a section edge we additionally require the same explanatory text and
    formula tree, so an absent neighbor never weakens a changed-formula match.
    """
    from .docxio import native_math_copy_reason
    target_context=_formula_context(doc,block)
    target_math=_math_signature(block)
    candidates=[]
    if len(target_context['headings'])<2:return candidates
    for sd in source_docs:
        if (sd.profile.get('kind')!='prospectus' or sd.hash==doc.hash or
                not doc.profile.get('issuer') or sd.profile.get('issuer')!=doc.profile['issuer']):continue
        for sb in sd.blocks:
            if sb.kind!='paragraph' or not _has_native_math(sb):continue
            context=_formula_context(sd,sb)
            if context['headings']!=target_context['headings']:continue
            agreements={side:target_context['neighbors'][side]['key']==context['neighbors'][side]['key']
                        for side in ('before','after')
                        if side in target_context['neighbors'] and side in context['neighbors']}
            same_formula=target_math==_math_signature(sb)
            same_text=bool(len(_key(block.text))>=12 and _key(block.text)==_key(sb.text))
            if not (all(agreements.get(side,False) for side in ('before','after')) or
                    (same_formula and same_text and any(agreements.values()))):continue
            warnings=source_flags.get((sd.hash,sb.index),[])
            unsupported=(complex_reference_reason([block.el]) or
                         native_math_copy_reason(doc.pkg,sd.pkg,[sb.el]))
            if block.images or sb.images:
                unsupported=unsupported or '公式与其他图示混合，尚未验证完整对象迁移。'
            evidence={'object_type':'native_formula','source_xml_path':_physical_locator(sd,sb.el),
                      'source_formula_sha256':_math_signature(sb),'target_formula_sha256':target_math,
                      'formula_equal':same_formula,'target_context':target_context,'source_context':context,
                      'neighbor_agreements':agreements,'whole_source_paragraph':True}
            candidates.append({'doc_hash':sd.hash,'source_name':sd.name,'source_kind':'prospectus',
                'source_role':'current_prospectus','unit_id':'formula'+str(sb.index),
                'start':sb.index,'end':sb.index+1,'indices':[sb.index],'score':1.0,
                'literal_score':1.0 if same_text else 0.0,'context_score':1.0,'heading_score':1.0,
                'content_score':1.0 if same_formula else 0.0,'table_score':0.0,
                'locator':' > '.join(sb.path[-4:]),'text':sb.text,'span':None,'exact':False,
                'unsupported':unsupported,'source_warnings':warnings,'scope_verified':True,
                'table_count':0,'source_fingerprint':fingerprint([sb.el]),'correspondence':evidence})
    return candidates


def _physical_locator(doc, element):
    parts=[]
    current=element
    while current is not None:
        parent=doc.parents.get(current)
        label=local(current)
        if parent is not None:
            siblings=[e for e in parent if e.tag==current.tag]
            label+='['+str(siblings.index(current)+1)+']'
        parts.append(label)
        current=parent
    return '/'+'/'.join(reversed(parts))


def _object_evidence(doc, blocks):
    """Resolve embedded image relationships to immutable package bytes."""
    media=[]
    unsupported=[]
    rels={r.get('Id'):r for r in doc.pkg.rels()}
    for block in blocks:
        if block.kind!='paragraph':
            unsupported.append('图示位于表格内，保留完整表格等待人工核对。')
        reason=complex_reference_reason([block.el])
        if reason:unsupported.append(reason)
        for node in block.el.iter():
            tag=local(node)
            if tag in ('object','OLEObject'):
                unsupported.append('含 OLE 嵌入对象；尚未验证其预览、原生数据和形状引用的完整迁移，保留待核对，不视为已更新。')
            if str(node.tag).startswith('{'+_MATH+'}'):
                unsupported.append('原生公式保留，当前未实现公式依赖的完整迁移。')
            if tag in ('chart', 'relIds'):
                unsupported.append('原生图表或 SmartArt 保留，当前未实现其数据及依赖的完整迁移。')
            if tag in ('wgp', 'grpSp', 'group', 'txbx', 'txbxContent'):
                unsupported.append('组合形状或文本框保留，当前未实现其完整迁移。')
            if tag not in ('blip', 'imagedata'):continue
            rid=node.get('{'+R+'}embed') or node.get('{'+R+'}id')
            if node.get('{'+R+'}link'):
                unsupported.append('图片使用外部链接，不能作为已取得的来源图复制。')
            relationship=rels.get(rid)
            if relationship is None:
                unsupported.append('图片关系无法解析，未取得可复制的来源媒体。')
                continue
            if relationship.get('TargetMode')=='External' or not (relationship.get('Type') or '').endswith('/image'):
                unsupported.append('图片关系不是包内图像资源，保留待核对。')
                continue
            target=relationship.get('Target','')
            part=posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('word',target))
            payload=doc.pkg.entries.get(part)
            if payload is None:
                unsupported.append('图片关系指向的媒体不存在，保留待核对。')
                continue
            if posixpath.splitext(part)[1].lower() not in ('.png','.jpg','.jpeg','.gif','.bmp','.tif','.tiff'):
                unsupported.append('图像资源不是已支持的位图格式，保留真实对象等待人工核对。')
            media.append({'block':block.index,'xml_path':_physical_locator(doc,block.el),
                          'relationship_id':rid,'part':part,'sha256':hashlib.sha256(payload).hexdigest(),
                          'bytes':len(payload),'object_type':'raster-image'})
    if not media and not unsupported:unsupported.append('对象没有可验证的包内图片关系，保留待核对。')
    return media,'；'.join(dict.fromkeys(unsupported))


def _context(doc, start, path):
    meaningful=[]
    for value in path:
        if SECTION.match(value.replace(' ','')) or MATTER.match(value):continue
        clean=strip_lead(value)
        if len(_key(clean))>=4:meaningful.append(_key(clean))
    lead=''
    for b in reversed(doc.blocks[max(0,start-6):start]):
        if b.heading or _is_object(b):break
        if b.text and len(_key(b.text))>=30 and not _CAPTION.match(b.text) and not _NOTE.match(b.text):
            lead=b.text;break
    return {'headings':meaningful,'lead':lead}


def _groups(doc):
    groups=[]
    consumed=set()
    for block in doc.blocks:
        if block.index in consumed or not _is_object(block):continue
        start=block.index;end=start+1;caption=''
        if start and doc.parents.get(doc.blocks[start-1].el) is doc.parents.get(block.el) and _CAPTION.match(doc.blocks[start-1].text):
            start-=1;caption=doc.blocks[start].text
        # Captionless or text-mixed objects are still represented as pending or
        # protected proposals; none disappear behind a generic image counter.
        while end<len(doc.blocks) and end-start<8:
            nxt=doc.blocks[end]
            if doc.parents.get(nxt.el) is not doc.parents.get(block.el):break
            if nxt.heading or _is_object(nxt) or not _NOTE.match(nxt.text):break
            end+=1
        blocks=doc.blocks[start:end]
        media,unsupported=_object_evidence(doc,blocks)
        if block.text.strip():unsupported=unsupported or '图文混合段落无法安全分离，保留整段待核对。'
        surrounding=[]
        for b in reversed(doc.blocks[max(0,start-3):start]):
            if b.heading:break
            surrounding.append(b.text)
        context_text=' '.join(block.path+surrounding+[b.text for b in blocks])
        signature=bool(_SIGNATURE.fullmatch(block.text))
        independent=block.section in ('proc','concl') or bool(_EXTERNAL.search(context_text)) or signature
        real_protection=any('人工批注' in b.reason for b in blocks)
        groups.append({'start':start,'end':end,'blocks':blocks,'caption':caption,'path':block.path,
                       'matter':block.matter,'media':media,'unsupported':unsupported,
                       'external':independent,'signature':signature,'manual_protection':real_protection,
                       'context':_context(doc,start,block.path)})
        consumed.update(range(start,end))
    return groups


def _context_match(target, source):
    th,sh=target['context']['headings'],source['context']['headings']
    leaf=bool(th and sh and th[-1]==sh[-1])
    # Generic headings alone do not establish provenance. The same neighbouring
    # narrative must also identify the chart's subject and measurement context.
    left,right=_key(target['context']['lead']),_key(source['context']['lead'])
    neighbor=SequenceMatcher(None,left,right,autojunk=False).ratio() if left and right else 0.0
    return bool(leaf and neighbor>=.60),neighbor


def _captionless_context(doc, group):
    """One raster paragraph expressly introduced by its immediate prose.

    A missing printed caption is not evidence that the source lacks the figure.
    Require two concrete enclosing headings, the exact preceding introduction and
    the exact following heading; image similarity alone never authorizes a copy.
    """
    if group['caption'] or len(group['blocks'])!=1:return None
    block=group['blocks'][0]
    if block.kind!='paragraph' or block.text.strip() or not block.images:return None
    if group['external'] or group['manual_protection']:return None
    if not 0<block.index<len(doc.blocks)-1:return None
    before,after=doc.blocks[block.index-1],doc.blocks[block.index+1]
    if (doc.parents.get(before.el) is not doc.parents.get(block.el) or
            doc.parents.get(after.el) is not doc.parents.get(block.el)):return None
    if (before.kind!='paragraph' or before.heading or _is_object(before) or
            len(_key(before.text))<30 or not re.search(r'(?:如下图|下图|图所示|图示)',before.text)):
        return None
    if not after.heading or MATTER.match(after.text) or SECTION.match(after.text.replace(' ','')):return None
    if _EXTERNAL.search(before.text+' '+after.text):return None
    headings=group['context']['headings'][-2:]
    if len(headings)!=2:return None
    return {'headings':headings,'before_block':before.index,'before_text':before.text,
            'before_key':_key(before.text),'after_block':after.index,'after_heading':after.text,
            'after_key':_key(strip_lead(after.text))}


def _captionless_candidates(doc,group,catalog,source_flags):
    target=_captionless_context(doc,group)
    if target is None:return []
    candidates=[]
    for sd,sg in catalog:
        context=_captionless_context(sd,sg)
        if context is None or any(context[k]!=target[k] for k in ('headings','before_key','after_key')):continue
        warnings=sorted({v for b in sg['blocks'] for v in source_flags.get((sd.hash,b.index),[])})
        candidates.append({'doc_hash':sd.hash,'source_name':sd.name,'source_kind':'prospectus',
            'source_role':'current_prospectus','unit_id':'figure'+str(sg['start']),
            'start':sg['start'],'end':sg['end'],'indices':list(range(sg['start'],sg['end'])),
            'score':1.0,'literal_score':1.0,'context_score':1.0,'heading_score':1.0,
            'content_score':0.0,'table_score':0.0,'locator':' > '.join(sg['path'][-4:]),
            'text':'\n'.join(b.text for b in sg['blocks'] if b.text),'span':None,'exact':False,
            'unsupported':sg['unsupported'],'source_warnings':warnings,'scope_verified':True,
            'figure':True,'table_count':0,'source_fingerprint':fingerprint([b.el for b in sg['blocks']]),
            'figure_evidence':{'source_media':sg['media'],'context':sg['context'],
                'caption_match':False,'captionless_context_match':True,
                'target_boundary_context':target,'source_boundary_context':context}})
    return candidates


def figure_proposals(doc, source_docs, source_flags=None):
    """Return engine-compatible proposals; input source_docs must be role=source.

    Each auto candidate contains figure_evidence with physical source relationships
    and hashes. Unsupported targets never expose a force-copy candidate. Unknown
    ownership and ambiguous/current-source gaps remain pending, with no deletion.
    """
    source_flags=source_flags or {}
    catalog=[]
    issuer=doc.profile.get('issuer')
    for sd in source_docs:
        if sd.profile.get('kind')!='prospectus' or sd.hash==doc.hash:continue
        if not issuer or sd.profile.get('issuer')!=issuer:continue
        catalog.extend((sd,group) for group in _groups(sd))
    proposals=[]
    for group in _groups(doc):
        start,end=group['start'],group['end']
        p={'id':doc.hash+':figure:'+str(start),'target_id':doc.hash,'target_name':doc.name,
           'start':start,'end':end,'matter':group['matter'] or '正文',
           'heading':group['caption'] or (group['path'][-1] if group['path'] else '图示对象'),
           'locator':' > '.join(group['path'][-4:]),'old_text':'\n'.join(b.text for b in group['blocks'] if b.text),
           'status':'review','decision':'pending','selected':None,'candidates':[],'note':'',
           'kind':'material','object_kind':'figure','figure':True,'has_table':False,
           'table_mode':'copy','identity':None,'sync_mode':'block',
           'figure_evidence':{'target_media':group['media'],'context':group['context']}}
        if group['external'] or group['manual_protection']:
            p.update(status='protected',decision='protected',reason=(
                '调查或复核人员签字属于底稿责任记录，保留原始签字和所在段落，不属于募集复制范围。'
                if group['signature'] else '项目组外部核查证据、独立程序/结论或人工批注范围保留，不以募集图片覆盖。'))
        elif len(group['blocks'])==1 and _has_native_math(group['blocks'][0]):
            p.update(figure=False,object_kind='native_formula')
            candidates=_native_formula_candidates(doc,group['blocks'][0],source_docs,source_flags)
            p['candidates']=candidates[:4]
            if len(candidates)>1:
                p['reason']='原生公式的同层标题和相邻正文指向多个来源位置，保留待核对。'
            elif candidates and candidates[0]['unsupported']:
                p['reason']=candidates[0]['unsupported']
            elif candidates and candidates[0]['source_warnings']:
                p['reason']='对应来源公式存在一致性提示，保留待核对：'+'；'.join(candidates[0]['source_warnings'])
            elif candidates:
                p.update(status='auto',decision='accept',selected=0,
                         reason='同发行人当前募集中的同层标题、相邻正文和公式段落已唯一对应，整段复制来源原生公式及文字并保留修订。')
            elif complex_reference_reason([group['blocks'][0].el]):
                p['reason']=complex_reference_reason([group['blocks'][0].el])
            else:
                p.update(status='missing',reason='未找到同发行人当前募集内标题和相邻正文唯一对应的原生公式段落。')
        elif group['unsupported']:
            p['reason']=group['unsupported']
        elif not group['caption']:
            candidates=_captionless_candidates(doc,group,catalog,source_flags)
            p['candidates']=candidates[:4]
            if len(candidates)>1:
                p['reason']='无独立图题图片的同层标题、直接引图正文和后邻标题指向多个来源位置，保留待核对。'
            elif candidates and candidates[0]['unsupported']:
                p['reason']=candidates[0]['unsupported']
            elif candidates and candidates[0]['source_warnings']:
                p['reason']='对应来源图组存在一致性提示，保留待核对：'+'；'.join(candidates[0]['source_warnings'])
            elif candidates:
                p.update(status='auto',decision='accept',selected=0,
                         reason='无独立图题，但同发行人当前募集的同层标题、直接引图正文和后邻标题已唯一定位图片；复制完整来源对象并保留修订，相同图片也实际复制。')
            else:
                p.update(status='missing',reason='未通过紧邻图题或完整引图上下文唯一定位当前募集对应图组。')
        else:
            candidates=[]
            for sd,sg in catalog:
                if not sg['caption'] or _caption_key(sg['caption'])!=_caption_key(group['caption']):continue
                matched,neighbor=_context_match(group,sg)
                if not matched or sg['external'] or sg['manual_protection']:continue
                warnings=sorted({v for b in sg['blocks'] for v in source_flags.get((sd.hash,b.index),[])})
                candidate={'doc_hash':sd.hash,'source_name':sd.name,'source_kind':'prospectus',
                           'source_role':'current_prospectus','unit_id':'figure'+str(sg['start']),
                           'start':sg['start'],'end':sg['end'],'indices':list(range(sg['start'],sg['end'])),
                           'score':round(.85+.15*neighbor,4),'literal_score':neighbor,'context_score':neighbor,
                           'heading_score':1.0,'content_score':0.0,'table_score':0.0,
                           'locator':' > '.join(sg['path'][-4:])+' > '+sg['caption'],
                           'text':'\n'.join(b.text for b in sg['blocks'] if b.text),'span':None,
                           'exact':False,'unsupported':sg['unsupported'],'source_warnings':warnings,
                           'scope_verified':True,'figure':True,'table_count':0,
                           'source_fingerprint':fingerprint([b.el for b in sg['blocks']]),
                           'figure_evidence':{'source_media':sg['media'],'context':sg['context'],
                                              'caption_match':True,'neighbor_score':round(neighbor,4)}}
                candidates.append(candidate)
            candidates.sort(key=lambda c:(-c['context_score'],c['source_name'],c['start']))
            p['candidates']=candidates[:4]
            variants={(tuple(x['sha256'] for x in c['figure_evidence']['source_media']),_key(c['text'])) for c in candidates}
            if len(variants)>1:
                p['reason']='同一图题及上下文指向不同当前来源图，未选择版本，保留待核对。'
            elif candidates and candidates[0]['unsupported']:
                p['reason']=candidates[0]['unsupported']
            elif candidates and candidates[0]['source_warnings']:
                p['reason']='对应来源图组存在一致性提示，保留待核对：'+'；'.join(candidates[0]['source_warnings'])
            elif candidates:
                p.update(status='auto',decision='accept',selected=0,
                         reason='同发行人当前募集中的图题、所属标题及相邻正文已定位图组，媒体关系和 SHA-256 已核实；整组复制来源并保留修订，内容相同也不跳过。')
            else:
                p.update(status='missing',reason='未取得同发行人当前募集内图题、所属标题和相邻正文同时对应的图组。')
        from .source_policy import retain_unmatched
        retain_unmatched(p)
        proposals.append(p)
    return proposals
