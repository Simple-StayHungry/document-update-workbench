"""One source-to-output builder shared by preview, decisions and DOCX export."""
from __future__ import annotations
import copy,re,json,hashlib
from .docxio import w,text,transformed,transform_paragraph
from .precision import source_elements,table_cells,fingerprint

from .wording import adapt, POLICY


def content_digest(elements):
    obj=[{'kind':'table' if e.tag==w('tbl') else 'paragraph','text':text(e),'tables':table_cells([e])} for e in elements]
    return hashlib.sha256(json.dumps(obj,ensure_ascii=False,separators=(',',':')).encode()).hexdigest()


def drop_imported_paragraph_style(el):
    """Drop an imported paragraph style reference but keep direct formatting.

    Every ordinary source copy in this program carries no pStyle, so the copied
    paragraph inherits the target's own Normal. An approved missing item copied
    from a paragraph that does carry one would instead pull the source style
    table in: importing ``zhengwen`` also clones its ``basedOn`` chain, including
    the document-level Normal and Default Paragraph Font. A second Normal is
    resolved by name across the whole file and silently moves every page.
    """
    for para in el.iter(w('p')):
        props=para.find(w('pPr'))
        if props is None:continue
        ref=props.find(w('pStyle'))
        if ref is not None:props.remove(ref)


def build(engine,p,files,selected=None):
    target=next((f for f in files if f['id']==p['target_id']),None)
    if not target:raise ValueError('待更新文件已移除，请重新分析。')
    doc=engine.load(target);old=[copy.deepcopy(b.el) for b in doc.blocks[p['start']:p['end']]]
    if not old and p.get('action')!='source_insert':raise ValueError('替换范围无效。')
    if p.get('action')=='source_insert':
        for side in ('before','after'):
            index=p.get('insert_'+side)
            if index is not None and (not 0<=index<len(doc.blocks) or fingerprint([doc.blocks[index].el])!=p.get('insert_'+side+'_fingerprint')):
                raise ValueError('新增正文的原始边界已改变，请重新分析。')
    if any(el.find('.//'+w('sectPr')) is not None for el in old):raise ValueError('目标范围含节分界，不能整块覆盖；请保留并核对。')
    if p['kind'] in ('metadata','field','voice'):
        new=[copy.deepcopy(old[0])];transform_paragraph(new[0],p.get('replacements',[]) if p['kind'] in ('field','voice') else [(text(old[0]),p['new_text'])])
        if text(new[0])!=p['new_text']:raise ValueError('字段预览与冻结计划不一致。')
        c=p['candidates'][0] if p['candidates'] else None
        origin=(c['source_name']+' > '+c['locator']) if c else '归档文种规范' if p['kind']=='voice' else '项目参数'
        return {'new':new,'source':None,'origin':origin,'source_hash':c['doc_hash'] if c else '', 'source_kind':'', 'digest':content_digest(new)}
    ci=(p.get('selected') if selected is None else selected)
    if ci is None:ci=0
    if type(ci) is not int or not 0<=ci<len(p['candidates']):raise ValueError('没有可靠来源；此处保留原文。')
    c=p['candidates'][ci]
    if c.get('unsupported'):raise ValueError(c['unsupported'])
    src=next((f for f in files if f['id']==c['doc_hash'] and f.get('role')=='source'),None)
    if not src:raise ValueError('来源文件已移除，请重新分析。')
    sd=engine.load(src);raw=source_elements(sd,c)
    if c.get('source_fingerprint') and fingerprint(raw)!=c['source_fingerprint']:raise ValueError('来源片段指纹不一致，已阻止导出。')
    if p.get('target_fingerprint') and fingerprint(old)!=p['target_fingerprint']:
        raise ValueError('目标范围与冻结计划不一致，已阻止导出。')
    if p.get('action')=='source_merge':
        if not p.get('merge_into') or not c.get('correspondence'):
            raise ValueError('合段删除缺少对应的来源复制证据。')
        return {'new':[],'source':sd.pkg,'origin':c['source_name']+' > '+c['locator'],'source_hash':sd.hash,'source_kind':sd.profile.get('kind','unknown'),'digest':content_digest([])}
    if c.get('projection'):
        # Projection/synthesis is intentionally forbidden. The archive workflow only
        # permits literal source copying after a deterministic location match.
        raise ValueError('该候选需要抽取或重建来源内容，已阻止自动写入；请定位可直接复制的完整来源对象。')
    quotation=False  # Source-copy is literal; do not silently add quotation marks.
    # Copy real source XML first, then apply only the frozen editorial wording.
    new=copy.deepcopy(raw)
    wording_changes=adapt(new,sd.profile.get('kind','unknown'))
    if quotation:
        oldtext='\n'.join(text(e) for e in old).strip()
        newtext='\n'.join(text(e) for e in new).strip()
        ts=[t for el in new for t in el.iter(w('t')) if t.text]
        if ts:
            if oldtext[:1] in ('“','「','『') and not newtext.startswith(oldtext[:1]):ts[0].text=oldtext[:1]+ts[0].text
            if oldtext[-1:] in ('”','」','』') and not newtext.endswith(oldtext[-1:]):ts[-1].text+=oldtext[-1:]
    if p.get('missing_heading') is not None:
        # Only the approved missing-item adapter sets this key. The heading is
        # created on the target side from the target's own style shell and the
        # copied source range stays heading-free, so the source-heading export
        # guard is neither bypassed nor relaxed.
        new=[copy.deepcopy(p['missing_heading'])]+new
        # The copied body must not drag the source style table in. Ordinary
        # source copies here carry no pStyle, so the paragraph inherits the
        # target's own Normal. Copying one that does carry a style would import
        # its basedOn chain too — including the document-level Normal and
        # Default Paragraph Font — and a second Normal is resolved by name
        # across the whole file, silently moving every page.
        for el in new[1:]:
            drop_imported_paragraph_style(el)
    digest=content_digest(new)
    if c.get('replacement_digest') and digest!=c['replacement_digest']:raise ValueError('替换后预览指纹不一致，已阻止导出。')
    origin=c['source_name']+' > '+c['locator']
    if c.get('projection'):origin+='；从同一来源表按期间和同名科目抽取完整行，未重算数字；保留原表题与单位'
    if c.get('span'):origin+='；该来源段落字符 '+str(c['span']['start']+1)+'—'+str(c['span']['end'])
    return {'new':new,'source':sd.pkg,'origin':origin,'source_hash':sd.hash,'source_kind':sd.profile.get('kind','unknown'),'digest':digest,'wording_policy':POLICY,'wording_changes':wording_changes}


def freeze(engine,p,files):
    if p['kind'] in ('metadata','field','voice'):
        p['replacement_digest']=build(engine,p,files)['digest'];return
    td=engine.load(next(f for f in files if f['id']==p['target_id']))
    p['target_fingerprint']=fingerprint([b.el for b in td.blocks[p['start']:p['end']]])
    from .model import Unit
    from .assurance import replacement_risks
    from .precision import source_table_change
    from .layout_policy import table_layout_risks
    selected_before=p.get('selected') if isinstance(p.get('selected'),int) else 0
    target=Unit('target',p['start'],p['end'],td.blocks[p['start']:p['end']],p['heading'],td.blocks[min(p['start'],len(td.blocks)-1)].path,p['matter'])
    for i,c in enumerate(p['candidates']):
        if c.get('unsupported'):continue
        sd=engine.load(next(f for f in files if f['id']==c['doc_hash']))
        # Record hard copy vetoes alongside each candidate; UI and export both enforce them.
        blocks=[copy.copy(sd.blocks[k]) for k in c.get('indices',range(c['start'],c['end']))]
        if c.get('span'):
            blocks[0].el=source_elements(sd,c)[0];blocks[0].text=text(blocks[0].el)
        source=Unit('source',c['start'],c['end'],blocks,c['locator'],blocks[0].path,'')
        risks=replacement_risks(target,source)
        if c.get('projection') or c.get('scope_verified'):risks=[]
        hard=[r for r in risks if '计量单位' in r or '表题' in r or '未覆盖目标中明确列示' in r or ('缺少原稿部分行项目' in r and not source_table_change(target,source))]
        if not target.has_table and len(target.text)<220 and re.search(r'(?:持股比例|持有发行人)[^。；]{0,30}\d+(?:\.\d+)?%',target.text) and not re.search(r'(?:持股|持有|股权)[^。；]{0,45}\d+(?:\.\d+)?%',source.text):
            hard.append('来源缺少原稿的持股比例。请补充直接依据，不能删除该事实后确认。')
        if c['score']<.52:hard.append('此候选只有弱相关性，不支持直接覆盖。')
        anchor=td.blocks[min(p['start'],len(td.blocks)-1)].el if p.get('action')=='source_insert' else None
        c['layout_risks']=table_layout_risks(td.pkg,[b.el for b in target.blocks],sd.pkg,[b.el for b in source.blocks],insertion_anchor=anchor)
        hard.extend(r['reason'] for r in c['layout_risks'])
        c['blocked_reason']='；'.join(hard)
        c['source_fingerprint']=fingerprint(source_elements(sd,c))
        planned=build(engine,p,files,i);c['replacement_digest']=planned['digest'];c['planned_text']='\n'.join(text(e) for e in planned['new'])
    if p.get('candidates') and 0<=selected_before<len(p['candidates']):
        layout=p['candidates'][selected_before].get('layout_risks',[])
        if layout:
            reason='；'.join(dict.fromkeys(r['reason'] for r in layout))
            p.update(status='review',decision='pending',selected=None,content_class='layout_pending',reason=reason)


def font_normalize(pkg,author):
    """Normalize fonts INSIDE inserted content only, without extra format redlines.

    Existing workpapers are already overwhelmingly Songti/Times New Roman. Creating
    independent rPrChange revisions on retained text made WPS stay red after the user
    rejected the content change, which looked like an unrejected replacement. Font
    cleanup is therefore folded into the inserted content itself.
    """
    from xml.etree import ElementTree as E
    parents={c:p for p in pkg.root.iter() for c in p};count=0
    for run in list(pkg.root.iter(w('r'))):
        if not any((t.text or '') for t in run.iter(w('t'))):continue
        cur=run;tags=[];chain=[]
        while cur in parents:
            cur=parents[cur];tags.append(cur.tag);chain.append(cur)
        # Paragraph insertions are wrapped in w:ins. Whole-table insertions use
        # grouped w:trPr/w:ins markers sharing one logical revision id, so their cell
        # runs need the same font normalization even though the run is not under w:ins.
        row_inserted=any(a.tag==w('tr') and a.find('./'+w('trPr')+'/'+w('ins')) is not None for a in chain)
        row_deleted=any(a.tag==w('tr') and a.find('./'+w('trPr')+'/'+w('del')) is not None for a in chain)
        if (w('ins') not in tags and not row_inserted) or w('del') in tags or w('moveFrom') in tags or row_deleted:continue
        if any('officeDocument/2006/math}' in tag for tag in tags):continue
        pr=run.find(w('rPr'))
        if pr is None:pr=E.Element(w('rPr'));run.insert(0,pr)
        rf=pr.find(w('rFonts'));want={w('ascii'):'Times New Roman',w('hAnsi'):'Times New Roman',w('eastAsia'):'宋体',w('cs'):'Times New Roman'}
        if rf is None:rf=E.Element(w('rFonts'));pr.insert(0,rf)
        before=dict(rf.attrib);rf.attrib.clear();rf.attrib.update(want)
        if before!=rf.attrib:count+=1
    return count


def retained_plan(engine,p,files):
    """Pending/kept material stays byte-for-byte unchanged.

    Wording changes belong inside an accepted source insertion only; pending
    factual decisions must not create independent editorial revisions.
    """
    doc=engine.load(next(f for f in files if f['id']==p['target_id']))
    old=[b.el for b in doc.blocks[p['start']:p['end']]]
    return {'new':copy.deepcopy(old),'changed':False,'digest':content_digest(old)}

def effective_context(engine,proposal,result,files,start,end):
    """Compose neighbouring blocks using CURRENT decisions, not stale target text."""
    doc=engine.load(next(f for f in files if f['id']==proposal['target_id']))
    plans=sorted((p for p in result['proposals'] if p['target_id']==proposal['target_id'] and p['id']!=proposal['id']),key=lambda p:p['start'])
    relevant=[p for p in plans if (p['start']<end and p['end']>start) or (p.get('action')=='source_insert' and start<=p['start']<=end)]
    if relevant:start=min(start,min(p['start'] for p in relevant));end=max(end,max(p['end'] for p in relevant))
    at={p['start']:p for p in relevant if p.get('action')!='source_insert'}
    additions={p['start']:p for p in relevant if p.get('action')=='source_insert'}
    out=[];i=start
    while i<=end:
        addition=additions.get(i)
        if addition and addition['decision'] in ('accept','same'):out.extend(build(engine,addition,files)['new'])
        if i==end:break
        p=at.get(i)
        if p:
            if p['decision'] in ('accept','same'):out.extend(build(engine,p,files)['new'])
            elif p['decision']=='pending':out.extend(retained_plan(engine,p,files)['new'])
            else:out.extend(copy.deepcopy(b.el) for b in doc.blocks[p['start']:p['end']])
            i=p['end']
        else:out.append(copy.deepcopy(doc.blocks[i].el));i+=1
    return out
