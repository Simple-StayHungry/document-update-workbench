from __future__ import annotations
import copy, difflib, hashlib, json, math, os, re, shutil, zipfile
try:
    from rapidfuzz.fuzz import ratio as _rapid_ratio
    def _ratio(a,b): return _rapid_ratio(a,b)/100.0
except Exception:
    def _ratio(a,b): return difflib.SequenceMatcher(None,a,b,autojunk=False).ratio()
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as E

from workbench.model import Document
from workbench.docxio import Package, w, text, local, transform_paragraph, validate_package
from workbench.plans import font_normalize

BUILD='explain-workbench-1.20'

DATE_TERMS=tuple(sorted((
 '本募集说明书签署之日','本募集说明书签署日','募集说明书签署之日','募集说明书签署日',
 '本募集说明书出具之日','本募集说明书出具日','募集说明书出具之日','募集说明书出具日',
 '本说明性文件出具之日','本说明性文件出具日','说明性文件出具之日','说明性文件出具日',
 '本说明出具之日','本说明出具日','说明出具之日','说明出具日'
), key=len, reverse=True))
_DATE_RE=re.compile('|'.join(re.escape(x) for x in DATE_TERMS))
_ADAPT_RE=re.compile('|'.join(re.escape(x) for x in list(DATE_TERMS)+['本募集说明书','本公司','发行人']))
_LEAD_RE=re.compile(r'^\s*(?:[（(][一二三四五六七八九十百\d]+[）)]|[一二三四五六七八九十百]+[、.]|\d+[、.)）])\s*')

def _split_lead(s:str):
    m=_LEAD_RE.match(s or '')
    return (m.group(0), (s or '')[m.end():]) if m else ('',s or '')


def _compact(s:str)->str:
    return re.sub(r'\s+','',s or '').replace('（','(').replace('）',')').replace('，',',').replace('。','.').replace('：',':').replace('；',';')

def neutral(s:str)->str:
    s=s or ''
    s=_DATE_RE.sub('<DATE>',s)
    s=s.replace('本募集说明书','募集说明书')
    # issuer voice is semantically the same between prospectus and standalone issuer statement.
    s=re.sub(r'发行人|本公司|公司','<C>',s)
    return _compact(s)

def adapt_text(s:str)->str:
    s=s or ''
    s=_DATE_RE.sub('本说明性文件出具日',s)
    s=s.replace('本募集说明书','募集说明书')
    s=s.replace('发行人','公司').replace('本公司','公司')
    return s

def adapt_text_for_target(source_text:str,target_text:str)->str:
    """Apply issuer-voice rules while preserving the standalone file's numbering shell."""
    s=adapt_text(source_text)
    sp,rest=_split_lead(s);tp,_=_split_lead(target_text)
    if sp!=tp and sp:
        s=tp+rest
    return s

def _looks_substantive_heading_text(s:str)->bool:
    """Some WPS body lead-ins are styled as headings; do not mistake them for immutable structure."""
    s=(s or '').strip()
    if not s:return False
    # A sentence-like enumerated lead-in or factual statement should still refresh from the prospectus.
    if any(x in s for x in ('截至','报告期','具体情况如下','情况如下表','不存在已发行','存在已注册','累计发行')):return True
    # Numbered list clauses are factual body content even when WPS assigns them a
    # heading-like style or their text is short.  A terminal sentence/list mark is
    # the reliable distinction: real structural headings normally do not end in
    # 。/；/:, while governance powers and other enumerated facts do.
    if s.endswith(('。','；',';', '：',':')):return True
    return False

def _is_structural_heading(b)->bool:
    return bool((b.heading or _headingish(b.text)) and not _looks_substantive_heading_text(b.text))

def _adapt_mapping(before:str):
    """Build non-overlapping wording replacements, longest term first."""
    mapping=[];seen=set()
    for m in _ADAPT_RE.finditer(before or ''):
        old=m.group(0)
        if old in seen:continue
        seen.add(old)
        if old in DATE_TERMS:new='本说明性文件出具日'
        elif old=='本募集说明书':new='募集说明书'
        else:new='公司'
        mapping.append((old,new))
    return mapping

def _copy_target_paragraph_shell(clone,target_el):
    """Keep the standalone document's paragraph geometry/style shell.

    Facts come from the prospectus, but spacing/indent/alignment/numbering stay
    with the explanatory document, matching the workpaper workflow.
    """
    if clone.tag!=w('p') or target_el is None or target_el.tag!=w('p'):
        return clone
    tp=target_el.find(w('pPr'));cp=clone.find(w('pPr'))
    if cp is not None:clone.remove(cp)
    if tp is not None:clone.insert(0,copy.deepcopy(tp))
    return clone

def adapt_element(el,target_text:str|None=None,target_el=None):
    clone=copy.deepcopy(el)
    _copy_target_paragraph_shell(clone,target_el)
    first_para=True
    for p in clone.iter(w('p')):
        before=''.join(t.text or '' for t in p.iter(w('t')))
        mapping=_adapt_mapping(before)
        if mapping:transform_paragraph(p,mapping)
        # Prospectus body paragraphs may carry local numbering that the standalone
        # explanatory file deliberately omits or styles differently. Keep target numbering.
        if first_para and target_text is not None and clone.tag==w('p'):
            after=''.join(t.text or '' for t in p.iter(w('t')))
            sp,_=_split_lead(after);tp,_=_split_lead(target_text)
            if sp and sp!=tp:transform_paragraph(p,[(sp,tp)])
        first_para=False
    return clone

def _has_drawing(block):
    return any(local(e) in ('drawing','pict','object') for e in block.el.iter())

def _figure_candidate(source_doc,prev_source,lookahead=5):
    if prev_source is None:return None
    hits=[]
    for idx in range(prev_source+1,min(len(source_doc.blocks),prev_source+1+lookahead)):
        b=source_doc.blocks[idx]
        if b.heading and b.text.strip():break
        if _has_drawing(b):hits.append(b)
    return hits[0] if len(hits)==1 else None

_PROJ_DATE_RE=re.compile(r'截至(?:本)?(?:募集说明书|说明性文件|说明)(?:签署|出具)?(?:之)?日')
_PROJ_NUM_RE=re.compile(r'(?<!\d)(?:\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d+)%')

def _projection_normal(s:str)->str:
    s=_PROJ_DATE_RE.sub('',s or '')
    s=re.sub(r'发行人|本公司|公司','',s)
    return re.sub(r'[\s，。；：、“”‘’（）()《》]+','',s)

def _projection_candidate(target,index):
    """Find a longer prospectus paragraph that contains the target fact clause.

    This is deliberately narrow: at least one percentage/decimal anchor must be
    preserved and >=82% of the target's 4-grams must occur in one source paragraph.
    It supports standalone summary sentences without copying a much longer source
    paragraph into the explanatory file.
    """
    if target.kind!='paragraph' or len(target.text.strip())<18:return None
    nums=set(_PROJ_NUM_RE.findall(target.text))
    if not nums:return None
    tn=_projection_normal(target.text);tg=_grams(tn)
    if len(tg)<8:return None
    out=[]
    for row in index.rows:
        if row['kind']!='paragraph':continue
        st=row['block'].text
        if not all(n in st for n in nums):continue
        sn=_projection_normal(st);sg=_grams(sn)
        coverage=len(tg & sg)/max(1,len(tg))
        if coverage<.82:continue
        # Avoid tiny incidental matches in a huge unrelated paragraph.
        phrase_bonus=.04 if any(x in tn and x in sn for x in ('表决权','控股股东及实际控制人','持股比例','注册资本','资产负债率','营业收入','净利润')) else 0
        score=min(1.0,coverage+phrase_bonus)
        out.append((score,len(sn),row['block']))
    if not out:return None
    out.sort(key=lambda x:(-x[0],x[1],x[2].index))
    score,_,b=out[0]
    return {'source_index':b.index,'score':round(score,4),'literal':round(score,4),'kind':b.kind,
            'source_text':b.text,'source_path':block_path(b),'projection':True}

def _projection_mask_percentage(s:str)->str:
    s=_PROJ_DATE_RE.sub('',s or '')
    s=re.sub(r'发行人|本公司|公司','',s)
    s=_PROJ_NUM_RE.sub('<P>',s)
    return re.sub(r'[\s，。；：、“”‘’（）()《》]+','',s)


def _projection_numeric_drift_candidate(target,index):
    """Allow one strongly identified percentage in a target-shaped summary to refresh.

    This is intentionally much narrower than ordinary fuzzy matching.  It is used only
    after full-block matching and the strict projection above both fail.  The target
    must contain exactly one percentage, and one source sentence/semicolon clause must
    have the same semantic skeleton with exactly one percentage.  Only that percentage
    is projected; wording, entity naming and the standalone date shell stay target-owned.
    """
    if target.kind!='paragraph' or len(target.text.strip())<18:return None
    tvals=_PROJ_NUM_RE.findall(target.text)
    if len(tvals)!=1:return None
    tn=_projection_mask_percentage(target.text);tg=_grams(tn)
    if len(tg)<8:return None
    hits=[]
    strong=('表决权','控股股东及实际控制人','持股比例','注册资本','资产负债率','营业收入','净利润')
    for row in index.rows:
        if row['kind']!='paragraph':continue
        st=row['block'].text
        for seg in re.findall(r'[^。；;]+[。；;]?',st):
            svals=_PROJ_NUM_RE.findall(seg)
            if len(svals)!=1:continue
            sn=_projection_mask_percentage(seg);sg=_grams(sn)
            coverage=len(tg & sg)/max(1,len(tg));ratio=_ratio(tn,sn)
            if coverage<.78 or ratio<.72:continue
            phrase_bonus=.04 if any(x in tn and x in sn for x in strong) else 0
            score=min(1.0,.65*coverage+.35*ratio+phrase_bonus)
            hits.append((score,coverage,ratio,len(sn),row['block'],seg,svals[0]))
    if not hits:return None
    hits.sort(key=lambda x:(-x[0],-x[1],-x[2],x[3],x[4].index))
    top=hits[0]
    # Do not auto-project across two equally plausible clauses/chapters.
    if len(hits)>1 and top[0]-hits[1][0]<.035 and hits[1][4].index!=top[4].index:
        return None
    score,coverage,ratio,_,b,seg,sval=top
    return {'source_index':b.index,'score':round(score,4),'literal':round(ratio,4),'kind':b.kind,
            'source_text':b.text,'source_path':block_path(b),'projection':True,'projection_numeric_drift':True,
            'projection_segment':seg,'projection_target_value':tvals[0],'projection_source_value':sval}


def _materialize_projection(target_el,target_value=None,source_value=None,segment_text=None,target_text=None):
    clone=copy.deepcopy(target_el)
    if segment_text is not None and clone.tag==w('p'):
        before=text(clone)
        replacement=adapt_text_for_target(segment_text,target_text if target_text is not None else before)
        if before and replacement!=before:
            transform_paragraph(clone,[(before,replacement)])
    for p in clone.iter(w('p')):
        before=''.join(t.text or '' for t in p.iter(w('t')))
        mapping=_adapt_mapping(before)
        if mapping:transform_paragraph(p,mapping)
        if target_value and source_value and target_value!=source_value:
            transform_paragraph(p,[(target_value,source_value)])
    return clone

_TRAILING_HEADING_RE=re.compile(r'([。；;]\s*)([（(][一二三四五六七八九十百\d]+[）)].{2,100})$')
_TARGET_OWNED_TABLE_RE=re.compile(r'同行业|同业对比|同行业对比|可比公司|平均水平|项目组|核查结论')
_UNIT_LINE_RE=re.compile(r'^\s*单位\s*[:：]')

def _trailing_heading_tail(target_text:str,source_text:str):
    m=_TRAILING_HEADING_RE.search(target_text or '')
    if not m:return None
    tail=m.group(2)
    if neutral(tail) and neutral(tail) not in neutral(source_text or ''):
        body=(target_text or '')[:m.start(2)]
        return body,tail
    return None

def _mixed_target_table(target_text:str,source_text:str)->bool:
    """Protect analysis/comparison rows that are not prospectus-owned facts."""
    t=target_text or '';s=source_text or ''
    return bool(_TARGET_OWNED_TABLE_RE.search(t) and not _TARGET_OWNED_TABLE_RE.search(s))

def _simple_table_schema(el):
    """Return a rectangular, unmerged table schema or ``None``.

    A target-shaped projection is only safe when every row has the same number
    of ordinary cells and the header labels are unique.  Merged/nested tables
    stay in review instead of being guessed into a new structure.
    """
    if el is None or el.tag!=w('tbl') or el.find('.//'+w('tbl')) is not None:
        return None
    rows=el.findall(w('tr'))
    if len(rows)<2:return None
    if el.findall('.//'+w('vMerge')) or el.findall('.//'+w('gridSpan')):
        return None
    grids=[]
    for tr in rows:
        cells=tr.findall(w('tc'))
        if not cells:return None
        grids.append(cells)
    width=len(grids[0])
    if any(len(row)!=width for row in grids):return None
    headers=[text(tc).strip() for tc in grids[0]]
    keys=[neutral(x) for x in headers]
    if any(not k for k in keys) or len(set(keys))!=len(keys):return None
    return {'rows':grids,'headers':headers,'keys':keys,'width':width}


def _table_projection_columns(target_el,source_el):
    """Return source column positions for a deterministic target-schema projection.

    The explanatory file owns the table schema; the prospectus owns the facts.
    If the explanatory headers are an ordered strict subset of a simple source
    table, project the latest source rows into that existing schema and replace
    the *whole* target table as one tracked object.  Any ambiguity is review-only.
    """
    tg=_simple_table_schema(target_el);sg=_simple_table_schema(source_el)
    if not tg or not sg:return None
    if tg['keys']==sg['keys']:return []  # identical schema; ordinary whole copy
    pos=[];last=-1
    for key in tg['keys']:
        hits=[i for i,k in enumerate(sg['keys']) if k==key]
        if len(hits)!=1:return None
        i=hits[0]
        if i<=last:return None
        pos.append(i);last=i
    if len(pos)>=sg['width']:return None
    return pos


def _project_source_table(source_el,columns):
    """Deep-copy a source table while retaining only selected ordinary columns."""
    clone=copy.deepcopy(source_el)
    schema=_simple_table_schema(clone)
    if not schema or not columns:return clone
    for tr in clone.findall(w('tr')):
        cells=tr.findall(w('tc'))
        kept=[copy.deepcopy(cells[i]) for i in columns]
        for c in list(cells):tr.remove(c)
        for c in kept:tr.append(c)
    grid=clone.find(w('tblGrid'))
    if grid is not None:
        cols=grid.findall(w('gridCol'))
        if len(cols)==schema['width']:
            for c in list(cols):grid.remove(c)
            for i in columns:grid.append(copy.deepcopy(cols[i]))
    return clone


def _table_structure_mismatch(target_el,source_el):
    tg=_simple_table_schema(target_el);sg=_simple_table_schema(source_el)
    if not tg or not sg:return False
    return tg['keys']!=sg['keys']

def _merge_source_body_preserve_tail(target_el,target_text,source_text,body,tail):
    clone=copy.deepcopy(target_el)
    new_body=adapt_text_for_target(source_text,body)
    transform_paragraph(clone,[(body,new_body)])
    for p in clone.iter(w('p')):
        mapping=_adapt_mapping(''.join(t.text or '' for t in p.iter(w('t'))))
        if mapping:transform_paragraph(p,mapping)
    return clone

def visible_equal(a,b):
    return _compact(a)==_compact(b)

def _grams(s,n=4):
    if not s:return set()
    if len(s)<n:return {s}
    return {s[i:i+n] for i in range(len(s)-n+1)}

def content_bounds(doc:Document):
    start=1;end=len(doc.blocks)
    for b in doc.blocks:
        if '说明如下' in b.text:
            start=b.index+1;break
    for b in doc.blocks[start:]:
        t=b.text.strip()
        if t.startswith('特此说明') or '本页无正文' in t:
            end=b.index;break
    return start,end

def block_path(b):
    return ' > '.join(getattr(b,'path',[]) or [])

def _headingish(s):
    s=(s or '').strip()
    return len(s)<90 and bool(re.match(r'^(?:第[一二三四五六七八九十百\d]+[章节]|[一二三四五六七八九十百]+[、.]|[（(][一二三四五六七八九十百\d]+[）)]|\d+[、.)）])',s))

class ProspectusIndex:
    def __init__(self,doc:Document):
        self.doc=doc;self.rows=[];self.exact=defaultdict(list);self.inv=defaultdict(list)
        for b in doc.blocks:
            if not b.text.strip() or b.is_toc or b.kind not in ('paragraph','table'):continue
            n=neutral(b.text)
            if len(n)<4:continue
            j=len(self.rows);g=_grams(n)
            row={'j':j,'index':b.index,'kind':b.kind,'n':n,'grams':g,'block':b,'path':block_path(b)}
            self.rows.append(row);self.exact[(b.kind,n)].append(j)
            for gram in list(g)[:4000]:self.inv[(b.kind,gram)].append(j)
    def candidates(self,target,limit=8):
        tn=neutral(target.text)
        if not tn:return []
        exact=self.exact.get((target.kind,tn),[])
        if exact:
            return [self._pack(self.rows[j],1.0,1.0) for j in exact[:limit]]
        tg=_grams(tn);hits=Counter()
        for g in tg:
            for j in self.inv.get((target.kind,g),[]):hits[j]+=1
        out=[]
        for j,_ in hits.most_common(45):
            row=self.rows[j];sn=row['n']
            size=max(len(tn),len(sn))/max(1,min(len(tn),len(sn)))
            if size>4.0:continue
            sim=_ratio(tn,sn)
            jacc=hits[j]/max(1,len(tg|row['grams']))
            score=.88*sim+.12*jacc
            # Short headings must be near-exact. Tables may legitimately change many numeric cells.
            if _headingish(target.text) and len(tn)<55 and sim<.76:continue
            if target.kind=='table':
                # Table headers are stronger than total-cell similarity.
                th=(target.text.split('\n',1)[0] if target.text else '')
                sh=(row['block'].text.split('\n',1)[0] if row['block'].text else '')
                hsim=_ratio(neutral(th),neutral(sh))
                if hsim<.58 and sim<.78:continue
                score=.72*score+.28*hsim
            if score<.30:continue
            out.append(self._pack(row,score,sim))
        out.sort(key=lambda x:(-x['score'],-x['literal'],x['source_index']))
        return out[:limit]
    def _pack(self,row,score,literal):
        b=row['block']
        return {'source_index':b.index,'score':round(score,4),'literal':round(literal,4),'kind':b.kind,
                'source_text':b.text,'source_path':row['path']}

def _bind_unit_lines_to_adjacent_tables(target_doc,source_doc,proposals):
    """Bind a standalone ``单位：`` line to the table it actually annotates.

    Unit labels are extremely non-unique in a prospectus (``单位：万元`` may occur
    dozens of times).  They must never be source-owned merely because identical text
    exists somewhere else.  The following target table is the identity anchor:

    * ready table -> use the unit line immediately preceding that exact source table;
    * review/target-owned table -> keep the unit line with that table in review.

    This prevents an analyst-owned comparison table from acquiring a fake prospectus
    provenance through a generic unit label.
    """
    by_idx={p['target_index']:p for p in proposals}
    for p in list(proposals):
        if p.get('kind')!='paragraph' or not _UNIT_LINE_RE.match(p.get('old_text','')):
            continue
        ti=p['target_index']
        # The unit label must directly annotate the next top-level table.  Do not
        # search across prose/headings and accidentally bind a distant table.
        if ti+1>=len(target_doc.blocks) or target_doc.blocks[ti+1].kind!='table':
            continue
        tp=by_idx.get(ti+1)
        if tp is None:
            continue
        if tp.get('status')!='ready' or tp.get('selected_source_index') is None:
            p.update(status='review',decision='keep',selected_source_index=None,selected_score=0,
                     source_owned=False,projection_mode=False,same_visible=False,
                     reason='该单位行从属于后续表格，而该表未确认唯一募集说明书来源；单位行随表默认保留，禁止从全文同名单位行猜来源。')
            continue
        table_si=tp['selected_source_index'];unit_block=None
        # In the source, unit text is normally the paragraph immediately before the
        # table.  Allow one extra plain paragraph only when it is also a unit line;
        # never cross another table or a structural heading.
        for si in range(table_si-1,max(-1,table_si-3),-1):
            if si<0:break
            sb=source_doc.blocks[si]
            if sb.kind=='table' or _is_structural_heading(sb):break
            if _UNIT_LINE_RE.match(sb.text or ''):
                unit_block=sb;break
            if sb.text.strip():break
        if unit_block is None:
            p.update(status='review',decision='keep',selected_source_index=None,selected_score=0,
                     source_owned=False,projection_mode=False,same_visible=False,
                     reason='后续表格已定位，但其募集说明书来源表前未找到唯一单位行；为避免跨章节误取同名单位，默认保留。')
            continue
        sim=_ratio(neutral(p.get('old_text','')),neutral(unit_block.text))
        if sim<.92:
            p.update(status='review',decision='keep',selected_source_index=None,selected_score=round(sim,4),
                     source_owned=False,projection_mode=False,same_visible=False,
                     reason='后续表格来源已定位，但目标单位行与该来源表的单位行不一致；默认保留并待人工确认。')
            continue
        cand={'source_index':unit_block.index,'score':round(sim,4),'literal':round(sim,4),'kind':unit_block.kind,
              'source_text':unit_block.text,'source_path':block_path(unit_block)}
        p.update(status='ready',decision='update',selected_source_index=unit_block.index,selected_score=round(sim,4),
                 candidates=[cand],source_owned=True,projection_mode=False,
                 same_visible=visible_equal(adapt_text_for_target(unit_block.text,p.get('old_text','')),p.get('old_text','')),
                 reason='单位行按其后已确认来源表绑定到同一募集说明书表前单位行；不再使用全文同名单位行匹配。')


def _lead_token(s:str)->str:
    lead,_=_split_lead(s or '')
    return _compact(lead)


def _resolve_structural_gap(target_doc,source_doc,proposals):
    """Resolve a changed body block only when surrounding structure proves identity.

    This handles cases such as a renamed department where text similarity is low but
    the target/source sequences are bracketed by already-confirmed blocks and carry
    the same numbered heading shells.  No free-form interpolation is allowed.
    """
    ordered=sorted(proposals,key=lambda p:p['target_index'])
    for p in ordered:
        if p.get('status')!='review' or p.get('kind')!='paragraph':
            continue
        ti=p['target_index']
        prev=next((q for q in reversed(ordered) if q['target_index']<ti and q.get('status')=='ready' and q.get('selected_source_index') is not None),None)
        nxt=next((q for q in ordered if q['target_index']>ti and q.get('status')=='ready' and q.get('selected_source_index') is not None),None)
        if not prev or not nxt:continue
        pt,nt=prev['target_index'],nxt['target_index'];ps,ns=prev['selected_source_index'],nxt['selected_source_index']
        if not (pt<ti<nt and ps<ns):continue
        if nt-pt!=ns-ps or nt-pt>8:continue
        tseq=target_doc.blocks[pt+1:nt]
        sseq=source_doc.blocks[ps+1:ns]
        if len(tseq)!=len(sseq) or not tseq:continue
        heading_pairs=0;ok=True
        for tb,sb in zip(tseq,sseq):
            if tb.kind!=sb.kind:ok=False;break
            th=_is_structural_heading(tb);sh=_is_structural_heading(sb)
            if th!=sh:ok=False;break
            if th:
                tl=_lead_token(tb.text);sl=_lead_token(sb.text)
                if not tl or tl!=sl:ok=False;break
                heading_pairs+=1
        if not ok or heading_pairs<1:continue
        offset=ti-pt-1
        if offset<0 or offset>=len(sseq):continue
        sb=sseq[offset];tb=tseq[offset]
        if _is_structural_heading(tb) or _is_structural_heading(sb) or sb.kind!='paragraph':continue
        cand={'source_index':sb.index,'score':1.0,'literal':round(_ratio(neutral(tb.text),neutral(sb.text)),4),
              'kind':sb.kind,'source_text':sb.text,'source_path':block_path(sb),'structural_gap':True}
        p.update(status='ready',decision='update',selected_source_index=sb.index,selected_score=1.0,
                 candidates=[cand]+[c for c in p.get('candidates',[]) if c.get('source_index')!=sb.index],
                 source_owned=True,projection_mode=False,same_visible=visible_equal(adapt_text_for_target(sb.text,tb.text),tb.text),
                 reason='前后已确认来源锚点之间的目标/募集结构等长，且编号标题壳逐项一致；按结构位置确认该正文来源并从最新募集说明书重物化。')


def _segment_key(s:str)->str:
    return neutral(re.sub(r'[。；;]+\\s*$','',s or ''))


def _resolve_clause_before_next_anchor(source_doc,proposals):
    """Resolve an explanatory sentence that is one exact clause of a longer source paragraph.

    The search is restricted to the small source window immediately before the next
    already-confirmed anchor.  It therefore cannot jump to a same-looking clause in
    another prospectus chapter.  The whole target sentence must match one source
    clause at >=95%; partial facts remain review-only.
    """
    ordered=sorted(proposals,key=lambda p:p['target_index'])
    for p in ordered:
        if p.get('status')!='review' or p.get('kind')!='paragraph':continue
        ti=p['target_index'];tk=_segment_key(p.get('old_text',''))
        if len(tk)<20:continue
        nxt=next((q for q in ordered if q['target_index']>ti and q.get('status')=='ready' and q.get('selected_source_index') is not None),None)
        if not nxt:continue
        ns=nxt['selected_source_index'];lo=max(0,ns-10)
        hits=[]
        for si in range(lo,ns):
            sb=source_doc.blocks[si]
            if sb.kind!='paragraph' or _is_structural_heading(sb):continue
            for seg in re.findall(r'[^。；;]+[。；;]?',sb.text or ''):
                sk=_segment_key(seg)
                if len(sk)<20:continue
                ratio=_ratio(tk,sk)
                lr=max(len(tk),len(sk))/max(1,min(len(tk),len(sk)))
                if ratio>=.95 and lr<=1.15:
                    hits.append((ratio,lr,si,sb,seg))
        if not hits:continue
        hits.sort(key=lambda x:(-x[0],x[1],-x[2]))
        top=hits[0]
        if len(hits)>1 and top[0]-hits[1][0]<.02 and _segment_key(top[4])!=_segment_key(hits[1][4]):continue
        ratio,_,si,sb,seg=top
        cand={'source_index':si,'score':round(ratio,4),'literal':round(ratio,4),'kind':'paragraph',
              'source_text':sb.text,'source_path':block_path(sb),'projection':True,'projection_clause':True,
              'projection_segment':seg}
        p.update(status='ready',decision='update',selected_source_index=si,selected_score=round(ratio,4),
                 candidates=[cand]+[c for c in p.get('candidates',[]) if c.get('source_index')!=si],
                 source_owned=True,projection_mode=True,same_visible=visible_equal(adapt_text_for_target(seg,p.get('old_text','')),p.get('old_text','')),
                 projection_segment_text=seg,
                 reason='本说明性文件完整句与下一已确认来源锚点之前的募集说明书单一分句高度一致；按该分句重物化，不复制同段其他事项。')




def _review_candidate_actionable(proposal, candidate):
    """Only expose a human choice when one candidate can safely replace the whole target.

    Mirrors the mature workbench rule: weak/partial retrieval is diagnostic evidence,
    not a task for the user. A review candidate must cover the target as a whole;
    otherwise the completed decision is to retain the target unchanged.
    """
    if not candidate or proposal.get('kind') != candidate.get('kind'):
        return False
    if proposal.get('kind') == 'table':
        # Mixed/schema-mismatched tables need a dedicated projection, not a manual
        # whole-table source copy that could delete target-owned rows.
        return False
    old=neutral(proposal.get('old_text','')); new=neutral(candidate.get('source_text',''))
    if len(old) < 10 or len(new) < 10:
        return False
    grams=_grams(old); source_grams=_grams(new)
    coverage=len(grams & source_grams)/max(1,len(grams))
    if coverage < .72:
        return False
    if len(new) < len(old)*.78:
        return False
    if candidate.get('score',0) < .60:
        return False
    return True

def _settle_nonactionable_reviews(proposals):
    """Convert non-actionable `review` rows into completed retention decisions.

    This follows the核查工作台 architecture: no correspondence and weak/partial
    candidates are not human work. Only a safe whole-object candidate can remain
    pending for a human choice. Diagnostic retrieval evidence is retained separately.
    """
    for p in proposals:
        if p.get('status') != 'review':
            continue
        reason=p.get('reason','')
        raw=list(p.get('candidates') or [])
        actionable=[c for c in raw if _review_candidate_actionable(p,c)]
        if actionable:
            p['candidates']=actionable
            p['review_required']=True
            p['content_class']='ambiguous_safe_source'
            p['decision']='keep'
            p['reason']='存在可完整替换的可靠来源，但自动规则无法唯一确定；仅此类事项需要人工确认。'
            continue
        p['correspondence_search']={'reason':reason,'candidates':raw}
        p.update(status='retained',decision='keep',selected_source_index=None,candidates=[],
                 review_required=False,content_class='no_safe_correspondence_retained',
                 reason='未找到可安全完整替换的唯一来源，按规则自动保留原文；无需人工检查。')
    return proposals

def _stable_id(target_name,target_index,source_index,old_text):
    raw=f'{target_name}|{target_index}|{source_index}|{_compact(old_text)[:180]}'
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()[:20]

def analyze(prospectus_path, targets:list[tuple[str,str]]):
    source=Document(prospectus_path,Path(prospectus_path).name);index=ProspectusIndex(source)
    result={'build':BUILD,'source':{'name':source.name,'path':str(prospectus_path),'issuer':source.profile.get('issuer',''),'hash':source.hash},'documents':[],'proposals':[],'summary':{}}
    for target_path,target_name in targets:
        doc=Document(target_path,target_name);lo,hi=content_bounds(doc);prev_source=None;proposals=[]
        for b in doc.blocks:
            if not(lo<=b.index<hi):continue
            # Image-only paragraphs are source-owned when they immediately follow a
            # proved local source anchor. This covers organization/shareholding charts
            # without guessing by OCR or visual similarity.
            if not b.text.strip() and _has_drawing(b):
                fb=_figure_candidate(source,prev_source)
                if fb is None:continue
                cand={'source_index':fb.index,'score':1.0,'literal':1.0,'kind':fb.kind,
                      'source_text':fb.text,'source_path':block_path(fb)}
                pid=_stable_id(target_name,b.index,fb.index,'<FIGURE>')
                p={'id':pid,'target_name':target_name,'target_path':str(target_path),'target_index':b.index,'kind':b.kind,
                   'target_path_text':block_path(b),'old_text':b.text,'status':'ready','decision':'update',
                   'reason':'与前一已确认来源锚点相邻的图形对象；按最新募集说明书重新物化并保留修订。',
                   'selected_source_index':fb.index,'selected_score':1.0,'candidates':[cand],'source_owned':True,'figure':True}
                proposals.append(p);result['proposals'].append(p);prev_source=fb.index;continue
            if not b.text.strip():continue
            cands=index.candidates(b)
            contextual=False
            if prev_source is not None and cands:
                top=cands[0]['score']
                nearby=[c for c in cands if prev_source<=c['source_index']<=prev_source+36 and c['score']>=max(.45,top-.18)]
                if nearby:
                    nearby.sort(key=lambda c:(c['source_index']-prev_source,-c['score']))
                    chosen=nearby[0];dist=chosen['source_index']-prev_source
                    chosen['score']=round(min(1.0,chosen['score']+(.18 if dist<=12 else .10)),4)
                    cands=[chosen]+[c for c in cands if c is not chosen]
                    contextual=True
                else:
                    for c in cands:
                        if c['source_index']<prev_source-15:c['score']=round(max(0,c['score']-.03),4)
                    cands.sort(key=lambda x:(-x['score'],-x['literal'],x['source_index']))
            best=cands[0] if cands else None
            # Standalone headings are target-owned structure. They can receive narrow
            # wording normalization, but source headings are never copied into them.
            if _is_structural_heading(b):
                if best and best['score']>=.70:prev_source=best['source_index']
                if adapt_text(b.text)==b.text:continue
                pid=_stable_id(target_name,b.index,-1,b.text)
                p={'id':pid,'target_name':target_name,'target_path':str(target_path),'target_index':b.index,'kind':b.kind,
                   'target_path_text':block_path(b),'old_text':b.text,'status':'wording','decision':'update',
                   'reason':'保留说明性文件自身标题层级和编号，仅统一独立出具口吻。','selected_source_index':None,'selected_score':1.0,'candidates':[],'source_owned':False}
                proposals.append(p);result['proposals'].append(p);continue
            threshold=.80 if b.kind=='paragraph' else .67
            if b.kind=='paragraph' and len(neutral(b.text))<110:threshold=.88
            if contextual and b.kind=='paragraph':threshold=min(threshold,.82)
            accepted=bool(best and best['score']>=threshold)
            mixed_table=bool(accepted and b.kind=='table' and _mixed_target_table(b.text,best['source_text']))
            table_projection=None
            table_schema_blocked=False
            if accepted and b.kind=='table' and not mixed_table:
                source_el=source.blocks[best['source_index']].el
                cols=_table_projection_columns(b.el,source_el)
                if cols is not None and cols:
                    table_projection=cols
                elif _table_structure_mismatch(b.el,source_el):
                    # A structurally different table is never blindly pasted into
                    # the standalone file.  Either project an exact header subset
                    # or require a human decision.
                    table_schema_blocked=True
            tail_keep=_trailing_heading_tail(b.text,best['source_text']) if accepted and b.kind=='paragraph' else None
            if mixed_table or table_schema_blocked:accepted=False
            projection=None
            if not accepted and not mixed_table:
                projection=_projection_candidate(b,index)
                if projection is None:
                    projection=_projection_numeric_drift_candidate(b,index)
            if accepted:prev_source=best['source_index']
            wording_needed=adapt_text(b.text)!=b.text
            # Core archive rule: once source ownership is proved, rematerialize from
            # the latest prospectus even when visible text is currently identical.
            if accepted:
                status='ready';decision='update'
                if table_projection:
                    reason='已可靠定位到募集说明书对应表；说明性文件表头是来源表的确定子集，按说明性文件既有列结构投影最新来源后整表重物化。'
                else:
                    reason='已可靠定位到募集说明书对应完整对象；无论内容是否相同，均删除旧对象并从最新来源重新物化。'
                selected=best['source_index'];source_owned=True;projection_mode=False
            elif projection is not None:
                status='ready';decision='update'
                reason='募集说明书较长段落完整包含本说明性文件摘要事实及关键数值；保留说明性文件摘要写法并强制重物化。'
                selected=projection['source_index'];source_owned=True;projection_mode=True
                cands=[projection]+[c for c in cands if c.get('source_index')!=projection['source_index']]
                best=projection
            elif mixed_table:
                status='review';decision='keep'
                reason='该表同时包含募集说明书事实与说明性文件自有同行业/比较分析行，禁止整表覆盖；默认保留并待人工确认。'
                selected=None;source_owned=False;projection_mode=False
            elif table_schema_blocked:
                status='review';decision='keep'
                reason='募集说明书候选表与说明性文件既有列结构不同，且无法按唯一表头子集安全投影；禁止直接改变说明性文件表结构。'
                selected=None;source_owned=False;projection_mode=False
            elif wording_needed:
                status='wording';decision='update'
                reason='仅统一说明性文件自身口吻，不改变事实内容。'
                selected=None;source_owned=False;projection_mode=False
            else:
                status='review';decision='keep'
                reason='来源定位不够稳定，默认保留原文；如需更新必须人工明确选择具体来源候选。'
                selected=None;source_owned=False;projection_mode=False
            same_visible=False
            if status=='ready' and not projection_mode and best is not None and not _has_drawing(b):
                same_visible=visible_equal(adapt_text_for_target(best.get('source_text',''),b.text),b.text)
            pid=_stable_id(target_name,b.index,selected if selected is not None else -1,b.text)
            p={'id':pid,'target_name':target_name,'target_path':str(target_path),'target_index':b.index,'kind':b.kind,
               'target_path_text':block_path(b),'old_text':b.text,'status':status,'decision':decision,'reason':reason,
               'selected_source_index':selected,'selected_score':best['score'] if best else 0,'candidates':cands,
               'source_owned':source_owned,'projection_mode':projection_mode,'same_visible':same_visible,'preserve_tail':tail_keep,
               'table_projection':table_projection,
               'projection_target_value':projection.get('projection_target_value') if projection else None,
               'projection_source_value':projection.get('projection_source_value') if projection else None,
               'projection_segment':projection.get('projection_segment') if projection else None,
               'projection_segment_text':projection.get('projection_segment') if projection and projection.get('projection_clause') else None}
            proposals.append(p);result['proposals'].append(p)
        _bind_unit_lines_to_adjacent_tables(doc,source,proposals)
        _resolve_structural_gap(doc,source,proposals)
        _resolve_clause_before_next_anchor(source,proposals)
        _settle_nonactionable_reviews(proposals)
        # Independent issuer-voice sweep is deliberately limited to the body.
        # The standalone document title and seal-page title reference are frozen
        # document identity, not issuer-voice prose. Source-backed body blocks already
        # have a proposal; do not double-redline them.
        covered={p['target_index'] for p in proposals}
        for b in doc.blocks:
            if not(lo<=b.index<hi):continue
            if b.index in covered or not b.text.strip():continue
            adapted=adapt_text(b.text)
            if adapted==b.text:continue
            pid=_stable_id(target_name,b.index,-1,b.text)
            p={'id':pid,'target_name':target_name,'target_path':str(target_path),'target_index':b.index,'kind':b.kind,
               'target_path_text':block_path(b),'old_text':b.text,'status':'wording','decision':'update',
               'reason':'统一独立说明性文件口吻，不改变事实内容。','selected_source_index':None,'selected_score':1.0,'candidates':[],'source_owned':False}
            proposals.append(p);result['proposals'].append(p)
        result['documents'].append({'name':target_name,'path':str(target_path),'hash':doc.hash,'blocks':len(doc.blocks),'content_range':[lo,hi],
                                    'proposals':len(proposals),'ready':sum(p['status']=='ready' for p in proposals),
                                    'wording':sum(p['status']=='wording' for p in proposals),'review':sum(p['status']=='review' for p in proposals),'retained':sum(p['status']=='retained' for p in proposals)})
    ps=result['proposals'];result['summary']={'documents':len(result['documents']),'changes':sum(p['decision']=='update' for p in ps),
        'source_updates':sum(p['status']=='ready' for p in ps),'wording_updates':sum(p['status']=='wording' for p in ps),'review':sum(p['status']=='review' for p in ps),'retained':sum(p['status']=='retained' for p in ps)}
    return result

def _lookup_source_block(source_doc,index):
    return source_doc.blocks[index].el

def _lookup_target_block(target_doc,index):
    return target_doc.blocks[index].el

def _decision_record(decisions,pid):
    raw=decisions.get(pid)
    if raw is None:return {}
    if isinstance(raw,str):return {'decision':raw}
    return dict(raw)

def export(result, decisions:dict, outdir, author='柒', _source_doc=None):
    outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True)
    src=_source_doc if _source_doc is not None else Document(result['source']['path'],result['source']['name'])
    report={'build':BUILD,'author':author,'files':[],'errors':[],'blocked':[]}
    expected_source_hash=result.get('source',{}).get('hash')
    if expected_source_hash and src.hash!=expected_source_hash:
        report['blocked'].append({'file':result['source'].get('name',''),'reason':'募集说明书在分析后发生变化；旧分析结果已失效，请重新分析。'})
        (outdir/'更新记录.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        return report
    bydoc=defaultdict(list)
    for p in result['proposals']:bydoc[p['target_name']].append(p)

    def effective(p):
        dr=_decision_record(decisions,p['id'])
        return dr,dr.get('decision',p['decision'])

    for docrec in result['documents']:
        name=docrec['name'];path=docrec['path'];tdoc=Document(path,name);pkg=tdoc.pkg
        expected_target_hash=docrec.get('hash')
        if expected_target_hash and tdoc.hash!=expected_target_hash:
            report['blocked'].append({'file':name,'reason':'说明性文件在分析后发生变化；旧分析结果已失效，请重新分析。'})
            continue
        applied=[];skipped=[];failed=False
        ps=sorted(bydoc[name],key=lambda x:x['target_index'])
        active=[]
        for p in ps:
            dr,decision=effective(p)
            if decision!='update':skipped.append(p['id']);continue
            active.append((p,dr))

        # Collapse adjacent source-owned updates into range operations.  Visible
        # equality is only an audit label, never an execution boundary: both changed
        # and unchanged source-owned blocks are pasted from the frozen prospectus.
        # Grouping across that label mirrors one continuous WPS delete/paste and
        # avoids hundreds of package/import passes on long explanatory files.
        ops=[];i=0
        while i<len(active):
            p,dr=active[i]
            ordinary=(p.get('status')=='ready' and not p.get('projection_mode') and not p.get('figure') and not p.get('preserve_tail') and not p.get('table_projection'))
            if ordinary:
                group=[(p,dr)];j=i+1
                while j<len(active):
                    q,qdr=active[j];prev=group[-1][0]
                    qordinary=(q.get('status')=='ready' and not q.get('projection_mode') and not q.get('figure') and not q.get('preserve_tail') and not q.get('table_projection'))
                    if not qordinary or q['target_index']!=prev['target_index']+1:break
                    if q.get('selected_source_index')!=prev.get('selected_source_index')+1:break
                    group.append((q,qdr));j+=1
                # Only group blocks in the same target XML container.
                els=[_lookup_target_block(tdoc,x[0]['target_index']) for x in group]
                parents={id(tdoc.parents.get(e)) for e in els}
                if len(parents)==1 and len(group)>1:
                    ops.append(('source_range_copy',group));i=j;continue
            ops.append(('single',[(p,dr)]));i+=1

        # Later ranges first keeps all original target anchors stable.
        ops.sort(key=lambda op:op[1][0][0]['target_index'],reverse=True)
        for opmode,group in ops:
            try:
                if opmode=='source_range_copy':
                    old=[_lookup_target_block(tdoc,p['target_index']) for p,_ in group]
                    new=[]
                    for (p,_),o in zip(group,old):
                        si=p.get('selected_source_index')
                        if si is None:raise ValueError('ready 项缺少冻结来源索引')
                        new.append(adapt_element(_lookup_source_block(src,si),p['old_text'],o))
                    pkg.replace(old,new,author=author,source=src.pkg,whole=True)
                    for p,_ in group:
                        applied.append({'id':p['id'],'status':'ready',
                                        'mode':'source_confirmed_refresh' if p.get('same_visible') else 'source_copy',
                                        'source_index':p.get('selected_source_index'),
                                        'score':p.get('selected_score',0),'chosen_by_user':False})
                    continue

                p,dr=group[0];old=[_lookup_target_block(tdoc,p['target_index'])];status=p.get('status')
                if status=='wording':
                    new=[adapt_element(old[0],target_el=old[0])]
                    if text(new[0])==text(old[0]):skipped.append(p['id']);continue
                    pkg.replace(old,new,author=author,source=None,whole=True)
                    applied.append({'id':p['id'],'status':status,'mode':'wording'});continue
                if status=='ready':
                    si=p.get('selected_source_index')
                    if si is None:raise ValueError('ready 项缺少冻结来源索引')
                    if p.get('same_visible') and not p.get('table_projection') and not p.get('projection_mode') and not p.get('preserve_tail') and not p.get('figure'):
                        new=[adapt_element(_lookup_source_block(src,si),p['old_text'],old[0])];mode='source_confirmed_refresh';source_pkg=src.pkg
                    elif p.get('table_projection'):
                        projected=_project_source_table(_lookup_source_block(src,si),p['table_projection'])
                        new=[projected];mode='source_table_projection';source_pkg=src.pkg
                    elif p.get('preserve_tail'):
                        body,tail=p['preserve_tail'];source_text=src.blocks[si].text
                        new=[_merge_source_body_preserve_tail(old[0],p['old_text'],source_text,body,tail)];mode='source_copy_preserve_tail';source_pkg=None
                    elif p.get('projection_mode'):
                        new=[_materialize_projection(old[0],p.get('projection_target_value'),p.get('projection_source_value'),p.get('projection_segment_text'),p.get('old_text'))];mode='source_projection';source_pkg=None
                    else:
                        new=[adapt_element(_lookup_source_block(src,si),p['old_text'],old[0])];mode='source_copy';source_pkg=src.pkg
                    pkg.replace(old,new,author=author,source=source_pkg,whole=True)
                    applied.append({'id':p['id'],'status':status,'mode':mode,'source_index':si,'score':p.get('selected_score',0),'chosen_by_user':False});continue
                if status=='review':
                    si=dr.get('source_index');valid={c.get('source_index') for c in p.get('candidates',[])}
                    if si is None or si not in valid:raise ValueError('待定位项必须明确选择当前来源候选，程序不会自动取第一候选。')
                    chosen=next(c for c in p['candidates'] if c.get('source_index')==si)
                    new=[adapt_element(_lookup_source_block(src,si),p['old_text'],old[0])]
                    pkg.replace(old,new,author=author,source=src.pkg,whole=True)
                    applied.append({'id':p['id'],'status':status,'mode':'source_copy','source_index':si,'score':chosen.get('score',0),'chosen_by_user':True});continue
                raise ValueError('未知 proposal 状态')
            except Exception as exc:
                report['blocked'].append({'file':name,'id':group[0][0]['id'],'reason':str(exc)});failed=True;break
        if failed:continue
        if any(a['status']=='wording' and a['mode']=='source_copy' for a in applied):
            report['blocked'].append({'file':name,'reason':'wording proposal 禁止 source_copy'});continue
        font_changes=font_normalize(pkg,author)
        dst=outdir/name;pkg.save(dst);issues=validate_package(dst)
        if issues:
            report['errors'].append({'file':name,'issues':issues});dst.unlink(missing_ok=True);continue
        report['files'].append({'file':name,'applied':len(applied),'skipped':len(skipped),'font_runs_standardized':font_changes,'applied_records':applied})
    (outdir/'更新记录.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report

def _subset_result(result,names):
    wanted=set(names)
    out=dict(result)
    out['documents']=[d for d in result.get('documents',[]) if d.get('name') in wanted]
    out['proposals']=[p for p in result.get('proposals',[]) if p.get('target_name') in wanted]
    ps=out['proposals']
    out['summary']={'documents':len(out['documents']),'changes':sum(p.get('decision')=='update' for p in ps),
                    'source_updates':sum(p.get('status')=='ready' for p in ps),
                    'wording_updates':sum(p.get('status')=='wording' for p in ps),
                    'review':sum(p.get('status')=='review' for p in ps),
                    'retained':sum(p.get('status')=='retained' for p in ps)}
    return out


_EXPORT_WORKER_SOURCE=None


def _init_export_worker(source_path,source_name,expected_hash=None):
    """Load the immutable prospectus once per process for dynamic document tasks."""
    global _EXPORT_WORKER_SOURCE
    src=Document(source_path,source_name)
    if expected_hash and src.hash!=expected_hash:
        raise RuntimeError('募集说明书在并行导出前发生变化；请重新分析。')
    _EXPORT_WORKER_SOURCE=src


def _export_document_task(payload):
    """Export one explanatory document using the worker-local prospectus cache."""
    result,decisions,outdir,author=payload
    if _EXPORT_WORKER_SOURCE is None:
        raise RuntimeError('并行导出 worker 未加载募集说明书')
    return export(result,decisions,outdir,author,_source_doc=_EXPORT_WORKER_SOURCE)


def export_parallel(result,decisions:dict,outdir,author='柒',workers=None):
    """Dynamically schedule independent DOCX exports across a small process pool.

    Each worker loads the prospectus exactly once, while documents are submitted as
    independent tasks.  This avoids the previous fixed-batch tail problem (one heavy
    explanatory file could hold the entire export open) without touching the mature
    Word writer.  Set EXPLAIN_NO_PARALLEL=1 to force the deterministic sequential path.
    """
    docs=list(result.get('documents',[]));outdir=Path(outdir)
    if len(docs)<=1 or os.environ.get('EXPLAIN_NO_PARALLEL')=='1':
        return export(result,decisions,outdir,author)
    if workers is None:
        workers=min(4,max(2,(os.cpu_count() or 2)),len(docs))
    workers=max(1,min(int(workers),len(docs)))
    if workers==1:return export(result,decisions,outdir,author)
    from concurrent.futures import ProcessPoolExecutor,as_completed

    workroot=outdir.parent/(outdir.name+'.parts')
    shutil.rmtree(workroot,ignore_errors=True);shutil.rmtree(outdir,ignore_errors=True)
    workroot.mkdir(parents=True,exist_ok=True);outdir.mkdir(parents=True,exist_ok=True)
    merged={'build':BUILD,'author':author,'files':[],'errors':[],'blocked':[],
            'parallel_workers':workers,'parallel_mode':'dynamic_document_tasks'}
    expected_source_hash=result.get('source',{}).get('hash')
    futures={}
    try:
        with ProcessPoolExecutor(
            max_workers=workers,
            initializer=_init_export_worker,
            initargs=(result['source']['path'],result['source']['name'],expected_source_hash),
        ) as pool:
            # Submit heavier documents first, but let idle workers pull the next task
            # dynamically.  Proposal count is only a priority hint, never a partition.
            counts=Counter(p.get('target_name') for p in result.get('proposals',[]))
            ordered=sorted(docs,key=lambda d:(-counts.get(d.get('name'),0),d.get('name','')))
            for task_id,d in enumerate(ordered):
                name=d['name'];part=workroot/f'{task_id:03d}';part.mkdir(parents=True,exist_ok=True)
                sub=_subset_result(result,[name])
                fut=pool.submit(_export_document_task,(sub,decisions,str(part),author))
                futures[fut]=(name,part)
            for fut in as_completed(futures):
                name,part=futures[fut]
                try:rep=fut.result()
                except Exception as exc:
                    merged['blocked'].append({'file':name,'reason':'并行导出进程失败：'+str(exc)})
                    continue
                merged['files'].extend(rep.get('files',[]));merged['errors'].extend(rep.get('errors',[]));merged['blocked'].extend(rep.get('blocked',[]))
                for fp in part.iterdir():
                    if fp.is_file() and fp.name!='更新记录.json':
                        shutil.move(str(fp),str(outdir/fp.name))
    finally:
        shutil.rmtree(workroot,ignore_errors=True)
    order={d['name']:i for i,d in enumerate(docs)}
    merged['files'].sort(key=lambda x:order.get(x.get('file'),10**9))
    (outdir/'更新记录.json').write_text(json.dumps(merged,ensure_ascii=False,indent=2),encoding='utf-8')
    return merged


def unpack_targets(zip_path,folder):
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True);items=[]
    with zipfile.ZipFile(zip_path) as z:
        for i,n in enumerate(z.namelist()):
            if n.endswith('/') or not n.lower().endswith('.docx'):continue
            original=Path(n).name;short=f'{i:03d}.docx';p=folder/short;p.write_bytes(z.read(n));items.append((str(p),original))
    return items

def zip_output(folder,zip_path):
    folder=Path(folder);zip_path=Path(zip_path)
    with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
        for p in sorted(folder.rglob('*')):
            if not p.is_file():continue
            arc=p.relative_to(folder).as_posix()
            # z.write preserves the generated file's actual mtime instead of
            # ZipInfo()'s 1980-01-01 default. Unicode arc names are emitted with
            # the UTF-8 flag by Python's zipfile implementation.
            arc.encode('utf-8',errors='strict')
            z.write(p,arcname=arc)
    return zip_path

