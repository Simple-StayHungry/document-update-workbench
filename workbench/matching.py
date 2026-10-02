"""Transparent retrieval and matching. Teacher files are deliberately not indexed.

Scores are structural similarity, not a probability or legal conclusion.
No company name, matter number or expected answer is hardcoded.
"""
from __future__ import annotations
import re,math,collections,functools,unicodedata

def _lcs_length(a,b):
    """Bit-parallel LCS used to reproduce RapidFuzz fuzz.ratio without a binary dependency."""
    if not a or not b:return 0
    if len(a)>len(b):a,b=b,a
    block={};bit=1
    for ch in a:
        block[ch]=block.get(ch,0)|bit;bit<<=1
    mask=(1<<len(a))-1;state=mask
    for ch in b:
        matches=block.get(ch,0);u=state&matches;state=(state+u)|(state-u)
    return ((~state)&mask).bit_count()

@functools.lru_cache(maxsize=50000)
def _python_ratio(a,b):
    if not a or not b:return 0.0
    if a==b:return 1.0
    # Common edges are guaranteed members of an LCS; trimming them makes
    # near-identical long document fragments substantially cheaper.
    prefix=0;limit=min(len(a),len(b))
    while prefix<limit and a[prefix]==b[prefix]:prefix+=1
    aa=a[prefix:];bb=b[prefix:]
    suffix=0;limit=min(len(aa),len(bb))
    while suffix<limit and aa[-1-suffix]==bb[-1-suffix]:suffix+=1
    if suffix:
        aa=aa[:-suffix];bb=bb[:-suffix]
    lcs=prefix+suffix+_lcs_length(aa,bb)
    return (2.0*lcs)/(len(a)+len(b))

def ratio(a,b):
    # Deterministic stdlib implementation: production and tests use the same matcher.
    return _python_ratio(a,b)
from .model import Unit,strip_lead
from .docxio import complex_reference_reason
from .assurance import replacement_risks
from .precision import descriptor_guard,nearest_scope,table_identity,useful_title,source_table_change,evidence_numbers,numeric_evidence_equivalent,is_short_numbered_title,GENERIC

def normal(s):
    s=re.sub(r'(?:本)?募集说明书(?:签署|出具)(?:之)?日|本核查(?:分析文件|文件|意见)出具(?:之)?日','本核查分析文件出具日',s)
    for a,b in [('本核查意见出具日','本核查分析文件出具日'),('本募集说明书签署之日','本核查分析文件出具日'),('本募集说明书签署日','本核查分析文件出具日'),('本核查文件出具之日','本核查分析文件出具日'),('本核查文件出具日','本核查分析文件出具日'),('本公司','发行人')]:s=s.replace(a,b)
    return re.sub(r'\s+','',s).replace('，',',').replace('：',':').replace('（','(').replace('）',')')

def excerpt_identity(s):
    # Exact excerpt permission does not inherit retrieval's date/voice aliases.
    # Different disclosure dates or document self-references remain differences.
    return re.sub(r'[\s，,。.;；：:（）()“”"、]','',unicodedata.normalize('NFKC',s))

def skeleton(s):
    s=normal(s);s=re.sub(r'\d[\d,，.．%％\-/－]*','#',s)
    return re.sub(r'[^\u4e00-\u9fffA-Za-z#]','',s)

def title_key(s):
    s=strip_lead(s)
    for v in ['发行人的','发行人','公司','的核查记录','核查记录','调查记录','的分析文件','分析文件','关于','情况','具体','基本']:s=s.replace(v,'')
    # Common heading variants that name the same disclosure location. This is
    # locator normalization only; it never changes copied source content.
    s=s.replace('所处行业地位','行业地位').replace('行业地位及竞争优势','行业地位竞争优势')
    s=s.replace('及其','和').replace('以及','和').replace('及','和').replace('与','和')
    s=s.replace('发行挂牌文件','发行上市文件').replace('挂牌申请文件','上市申请文件')
    return skeleton(s)

def grams(s):
    s=skeleton(s)
    return {s[i:i+2] for i in range(max(0,len(s)-1)) if '#' not in s[i:i+2]}

def subject_guard(target,source):
    a=' '.join(target.path);b=' '.join(source.path)
    veto=descriptor_guard(target,source)
    if veto:return veto
    # Symmetric issuer/parent/guarantor scope separation.
    parent=lambda u: nearest_scope(u.path,u.text[:110] if not u.has_table else '')=='发行人本部'
    if parent(target)!=parent(source):return '母公司或本部口径与合并口径不能混用'
    def owner(u):
        if len(u.text)<600:
            m=re.search(r'发行人(?:的)?(控股股东|实际控制人)',u.text)
            if m:return m[1]
        for part in reversed(u.path):
            roles=[r for r in ('控股股东','实际控制人') if r in part]
            if len(roles)==1:return roles[0]
        return ''
    ta,sa=owner(target),owner(source)
    if ta and sa and ta!=sa:return '实际控制人与控股股东是不同核查主体'
    # A current ownership/capital snapshot does not answer whether control changed
    # during the reporting period. Keep both named responsibility scopes intact.
    control_roles={r for r in ('控股股东','实际控制人') if r in target.text}
    if control_roles and re.search(r'(?:未|没有).{0,4}(?:变更|变化)',target.text):
        source_roles={r for r in ('控股股东','实际控制人') if r in source.text}
        if not control_roles.issubset(source_roles) or not re.search(r'变更|变化',source.text):
            return '控制主体在报告期是否变更与当前资本或持有人概况不是同一事实，不能替换'
    # Accounting error corrections are not policy/estimate changes. Similar
    # boilerplate must not surface an actionable candidate for another event.
    te=bool(re.search(r'会计差错|差错更正',target.text))
    se=bool(re.search(r'会计差错|差错更正',source.text))
    if te!=se and re.search(r'会计政策|会计估计',(source.text if te else target.text)):
        return '会计差错更正与会计政策、会计估计变更是不同事项，不能互相替换'
    # Short negative statements can look almost identical but concern different
    # propositions. A heading is never allowed to override the sentence itself.
    def objects(u):
        if len(u.text)>600:return set()
        if not re.search(r'不存在|未发生|没有|不涉及|无重大',u.text):return set()
        families={'重大承诺':r'重大承诺','诉讼仲裁':r'诉讼|仲裁|行政处罚','违规担保':r'违规担保','对外担保':r'对外担保','关联担保':r'关联.{0,5}担保|担保.{0,5}关联','股权质押':r'股权.{0,8}质押'}
        return {k for k,v in families.items() if re.search(v,u.text)}
    at,bt=objects(target),objects(source)
    if at and bt and at!=bt:return '否定命题的对象不同，不能用相似句式替换'

    guarantor=lambda u:nearest_scope(u.path)=='保证人'
    if guarantor(source)!=guarantor(target):return '来源属于保证人或增信主体，不是发行人'
    if ('风险' in b and '风险' not in a) and not (target.text.startswith('风险') or len(target.blocks)==1 and '风险' in target.text):return '来源位于风险提示，不能直接替代事实明细'
    if re.search(r'募集资金用途|债券基本',b) and re.search(r'子公司|参股|关联交易|营业收入|毛利率',a):return '来源主题与目标不一致'
    return ''

def prefer_prospectus(candidates):
    """Apply the user's priority only to competing copies of the same passage.

    Similarity does not create a source match here. Candidates have already passed
    subject/table boundaries; this only resolves conflicting current formal sources.
    """
    if not candidates:return candidates
    best=candidates[0]
    if best.get('source_kind')!='opinion':return candidates
    from .wording import mappings
    def editorial_identity(c):
        value=c['text']
        for old,new in mappings(value,c.get('source_kind','')):value=value.replace(old,new)
        return normal(value)
    primary=[c for c in candidates if c.get('source_kind')=='prospectus' and
             ratio(skeleton(c['text']),skeleton(best['text']))>=.92 and
             (c.get('heading_score',0)>=.85 or c.get('context_score',0)>=.85 or
              c.get('literal_score',0)>=.90 or editorial_identity(c)==editorial_identity(best))]
    if not primary:return candidates
    chosen=primary[0]
    if normal(chosen['text'])!=normal(best['text']):
        chosen['source_precedence']={'rule':'conflicting-formal-sources-prefer-prospectus',
            'preferred_source':chosen['doc_hash'],'other_source':best['doc_hash'],
            'other_indices':best['indices'],'other_text':best['text']}
    return [chosen]+[c for c in candidates if c is not chosen]

class Index:
    def __init__(self,documents,source_issues=()):
        self.source_flags=collections.defaultdict(list)
        for issue in source_issues:
            for block in issue.get('blocks',[]):self.source_flags[(issue['file_id'],block)].append(issue['title'])
        self.assess_cache={}
        self.items=[];self.inv=collections.defaultdict(set);self.exact=collections.defaultdict(list);self.weights={};self.documents=documents
        for doc in documents:
            for u in doc.source_units:
                # Do not use source query procedures, cover TOCs or signatures as facts.
                if u.blocks and all(b.section=='proc' for b in u.blocks):continue
                idx=len(self.items);body=normal(u.text);sk=skeleton(u.text)
                self.items.append({'doc':doc,'unit':u,'body':body,'skeleton':sk,'heading':title_key(useful_title(u)),'grams':grams(u.text+' '+u.heading+' '+u.locator)})
                self.exact[body].append(idx)
                for g in self.items[-1]['grams']:self.inv[g].add(idx)
        n=len(self.items)
        self.weights={g:math.log(1+n/(1+len(ids))) for g,ids in self.inv.items()}
    def candidates(self,target,limit=4):
        body=normal(target.text);sk=skeleton(target.text);head=title_key(useful_title(target))
        scores=collections.defaultdict(float)
        for g in sorted(grams(target.text+' '+target.heading+' '+target.locator)):
            for i in self.inv.get(g,[]):scores[i]+=self.weights.get(g,1)
        top=[i for i,_ in sorted(scores.items(),key=lambda x:(-x[1],x[0]))[:55]]
        for i in self.exact.get(body,[]):
            if i not in top:top.insert(0,i)
        ranked=[]
        for i in top:
            item=self.items[i];u=item['unit'];guard=subject_guard(target,u)
            if is_short_numbered_title(target.text) or getattr(u,'short_title_exact_only',False) or is_short_numbered_title(u.text):
                # A short label conveys too little for fuzzy matching. Its full
                # text and the two nearest informative parents must coincide.
                parents=lambda path:[re.sub(r'\s+','',strip_lead(part)) for part in path
                                     if not GENERIC.match(part) and len(strip_lead(part))>=3]
                tp,sp=parents(target.path),parents(u.path)
                if target.text.strip()!=u.text.strip() or len(tp)<2 or len(sp)<2 or tp[-2:]!=sp[-2:]:continue
                if any(b.section=='proc' or b.protected or b.is_toc or b.images for b in u.blocks) or complex_reference_reason([b.el for b in u.blocks]):continue
            # A whole target paragraph is never replaced by the most similar
            # sentence sliced out of a longer current source paragraph.
            if getattr(u,'span',None) and (target.has_table or excerpt_identity(target.text)!=excerpt_identity(u.text)):continue
            if guard:continue
            # Tables must be copied with their data and units, not replaced with prose.
            if target.has_table!=u.has_table:continue
            content=ratio(sk[:20000],item['skeleton'][:20000]);literal=ratio(body[:20000],item['body'][:20000]);h=ratio(head,item['heading'])
            context=max((ratio(title_key(a),title_key(b)) for a in target.path[-3:] for b in u.path[-3:]),default=0)
            schema=ratio(skeleton(target.schema),skeleton(u.schema)) if target.has_table else 0
            first=ratio(skeleton(target.blocks[0].text),skeleton(u.blocks[0].text))
            score=max(.86*content+.14*context,.35*h+.30*content+.20*first+.15*(schema if target.has_table else literal))
            if body==item['body']:score=1.0
            # An exact, informative heading plus compatible content/schema is useful across projects.
            if h>.96 and len(head)>=6 and (content>.55 or schema>.83) and context>.93:score=max(score,.84)
            # Prevent a small excerpt replacing an entire larger section, or vice versa.
            size=max(len(sk),len(item['skeleton']))/max(1,min(len(sk),len(item['skeleton'])))
            if size>4 and body!=item['body']:score-=min(.25,(size-4)*.035)
            if score<.32:continue
            ranked.append({'span':getattr(u,'span',None),'identity':table_identity(u) if u.has_table else None,'index':i,'doc_hash':item['doc'].hash,'source_name':item['doc'].name,'source_kind':item['doc'].profile.get('kind',''),'unit_id':u.id,'start':u.start,'end':u.end,'indices':[b.index for b in u.blocks],'score':round(score,4),'content_score':round(content,4),'literal_score':round(literal,4),'heading_score':round(h,4),'context_score':round(context,4),'table_score':round(schema,4),'locator':u.locator,'text':u.text,'exact':body==item['body'],'unsupported':complex_reference_reason([b.el for b in u.blocks]),'source_warnings':sorted({v for b in u.blocks for v in self.source_flags.get((item['doc'].hash,b.index),[])})})
        # A genuinely matching informative heading is a stronger locator than
        # prose similarity.  Similarity is only used inside that structural
        # address, never to choose a different titled disclosure.
        heading_exact=lambda x: int(x.get('heading_score',0)>=.985 and len(head)>=4)
        # Exact visible source content is the strongest locator.  Previously an exact
        # leaf heading was sorted ahead of an exact body match, which could select a
        # generic business-overview paragraph instead of the literal customer-table
        # lead sentence under a differently named subheading.  That is a target/source
        # range error, not a harmless score preference.
        # When the same visible sentence/table is repeated in several source
        # locations, exact text alone must not decide provenance. Prefer the
        # candidate whose surrounding headings best match the workpaper path.
        # The full-paragraph near-identity accepted by assess() must survive
        # retrieval sorting; a short opinion summary sharing a broad parent
        # heading cannot displace that more complete source object.
        ranked.sort(key=lambda x:(-int(x['exact']),-int(x['score']>=.52),-int(x['literal_score']>=.97 and x['score']>=.80),-heading_exact(x),-x.get('context_score',0),-x['score'],-x['literal_score'],0 if self.items[x['index']]['doc'].profile['kind']=='prospectus' else 1,x['source_name'],x['start'],x['end'],(x.get('span') or {}).get('start',-1),(x.get('span') or {}).get('end',-1),x['doc_hash'],x['unit_id']))
        locked=[getattr(b,'formal_source_correspondence',None) for b in target.blocks]
        if any(locked):
            if len(locked)!=1 or not locked[0]:return []
            permission=locked[0]
            kept=[]
            for candidate in ranked:
                original=(candidate['doc_hash']==permission['doc_hash'] and candidate['indices']==permission['indices'] and not candidate.get('span'))
                primary=(candidate.get('source_kind')=='prospectus' and not candidate.get('span') and
                         ratio(skeleton(candidate['text']),skeleton(permission['text']))>=.92 and
                         (candidate.get('context_score',0)>=.85 or candidate.get('literal_score',0)>=.90))
                if original or primary:
                    candidate['formal_correspondence']=True;kept.append(candidate)
            ranked=kept
        ranked=prefer_prospectus(ranked)
        # Same visible evidence repeated in several sections is not an ambiguity by itself.
        unique=[];seen=set()
        for r in ranked:
            key=(r['doc_hash'],normal(r['text']))
            if key in seen:continue
            seen.add(key);unique.append(r)
        return unique[:limit]
    def assess(self,u):
        c=self.candidates(u)
        own=complex_reference_reason([b.el for b in u.blocks])
        if own:return 'protected',c,own
        if not c:return 'missing',c,'最新版材料中没有找到足够接近的对应内容。'
        best=c[0];margin=best['score']-max((v['score'] for v in c[1:] if normal(v['text'])!=normal(best['text']) and not (evidence_numbers(v['text']) and ratio(skeleton(v['text']),skeleton(best['text']))>.985 and numeric_evidence_equivalent(v['text'],best['text']))),default=0)
        for other in c[1:]:
            overlap=best['doc_hash']==other['doc_hash'] and bool(set(best['indices']) & set(other['indices']))
            same_shape=ratio(skeleton(best['text']),skeleton(other['text']))>.94
            containment=normal(best['text']) in normal(other['text']) or normal(other['text']) in normal(best['text'])
            priority=best.get('source_precedence',{})
            if priority.get('other_source')==other['doc_hash'] and priority.get('other_indices')==other['indices']:continue
            if not overlap and not containment and best['score']-other['score']<.04 and same_shape and not numeric_evidence_equivalent(best['text'],other['text']):
                return 'review',c,'不同位置或版本给出高度相似但文字、数字不同的来源；请先确定适用版本，未自动采用。'

        if best.get('source_warnings'):return 'review',c,'对应来源自身存在一致性提示：'+'；'.join(best['source_warnings'])+'。即使文字相同，也未标记核对完成。'
        if best['exact']:return 'auto',c,'已定位对应来源对象；即使内容相同也直接复制来源，不跳过修订。'
        if best.get('unsupported'):return 'review',c,best['unsupported']+' 可保留原文；不支持强制自动写入该来源。'
        unit=self.items[best['index']]['unit']
        # Accuracy gate: a locator can be correct while its source span is too
        # narrow. Never auto-delete facts merely because a leaf heading matches.
        # Run every range-completeness veto before any non-exact auto decision.
        strong_table=source_table_change(u,unit)
        risks=replacement_risks(u,unit)
        # A verified same-caption/same-period table may legitimately add/remove
        # row items in the latest source. Other completeness risks still block.
        if strong_table:
            risks=[v for v in risks if not v.startswith('来源表格缺少原稿部分行项目')]
        if not u.has_table and len(u.text)<220:
            oldperc=re.findall(r'(?:持股比例|持有.{0,6}股权|持有发行人|占比)[^。；]{0,30}\d+(?:\.\d+)?%',u.text)
            if oldperc and not re.search(r'(?:持股|持有|股权|占比)[^。；]{0,45}\d+(?:\.\d+)?%',unit.text):
                risks.append('来源段落未包含原稿中的持股比例；不能删去该事实。需定位持股表或股权图。')
        if risks:
            return 'review',c,'；'.join(risks)
        if best.get('formal_correspondence'):
            return 'auto',c,'正式核查意见同一事项已通过完整段落与主题对应校验；复制完整来源正文并保留修订。'
        # Literal near-identity is used only as a locator.  It does NOT decide
        # whether an update is needed: once the corresponding source object is
        # uniquely located, the entire object is copied even when the words are
        # already the same.
        divergent=[v for v in c[1:] if normal(v['text'])!=normal(best['text'])]
        second_literal=max((v.get('literal_score',0) for v in divergent),default=0)
        if best.get('literal_score',0)>=.97 and best.get('score',0)>=.80 and best.get('literal_score',0)-second_literal>=.08:
            return 'auto',c,'正文关键词和顺序已唯一定位对应来源对象；完整性校验通过，直接复制完整来源对象。'
        # An informative same heading is a strong locator, but not authority to
        # shrink the target range.  The completeness veto above must pass first.
        exact_heading = best.get('heading_score',0)>=.985 and len(title_key(useful_title(u)))>=4
        if exact_heading and best['score']>=.75:
            return 'auto',c,'同名标题并结合正文与结构信息定位对应来源对象；完整性校验通过，直接复制完整来源对象。'
        if strong_table and best['table_score']>=.94 and best['heading_score']>=.96 and margin>=.05:
            return 'auto',c,'表题、期间、单位、主体及表结构一致；直接复制完整来源表，不逐格拼接。'
        # A changed paragraph is auto-copyable only when its OWN structural heading
        # is the same informative heading in the source. Near-identical prose by
        # itself is only a locator hint; it is never authority to overwrite.
        if exact_heading and best.get('literal_score',0)>=.90 and margin>=.05:
            return 'auto',c,'同名标题已唯一定位对应来源对象；直接复制完整来源对象，不按句内差异推断。'
        # All remaining fuzzy matches are review-only. Similarity is used to locate
        # a candidate, never to decide that the source should overwrite the draft.
        if best['score']<.52:return 'missing',c,'只有弱相关片段，保留原文，不按关键词强行替换。'
        return 'review',c,'存在候选，但相似度或候选差距不足；请核对一次后决定。'
