"""Evidence-first consistency checks; never imports teacher/reference contents.

All values originate in the selected documents. Decimal arithmetic checks are
limited to explicitly labelled financial statements and bond simulations.
Findings are not a legal opinion and cannot authorize an unsourced rewrite.
"""
from __future__ import annotations
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from .docxio import text, w, xml

NUM = r'[-−－]?(?:\d{1,3}(?:[,，]\d{3})+|\d+)(?:\.\d+)?'
MONEY = re.compile(r'('+NUM+r')\s*(亿元|万元|元)')
SCALE = {'亿元':Decimal(10000), '万元':Decimal(1), '元':Decimal('0.0001'), '%':Decimal(1), '倍':Decimal(1)}
METRICS = {
 '投资现金净额': r'投资活动(?:产生的)?(?:现金流量|净现金流|现金流)(?:量)?净额',
 '经营现金净额': r'经营活动(?:产生的)?(?:现金流量|现金流)净额',
 '筹资现金净额': r'筹资活动(?:产生的)?(?:现金流量|现金流)净额',
 '资产总额': r'(?:资产总计|总资产|(?<!流动)资产合计)',
 '负债总额': r'(?:(?<!流动)负债合计|总负债)',
 '净资产': r'(?:所有者权益(?:合计)?|净资产)',
 '净利润': r'(?<!扣除非经常性损益后)(?<!归属于母公司的)(?<!归属于母公司所有者的)(?<!归母)(?<!扣非)净利润',
 '营业收入': r'(?:营业总收入|营业收入)',
 '营业成本': r'营业成本',
 '毛利润': r'(?:综合毛利润|营业毛利润|毛利润合计)',
 '资产负债率': r'资产负债率',
 '流动比率': r'流动比率',
}
TABLE_NAMES = {
 '资产总计':'资产总额','资产总额':'资产总额','总资产':'资产总额',
 '负债合计':'负债总额','总负债':'负债总额','所有者权益合计':'净资产','所有者权益':'净资产','净资产':'净资产',
 '营业收入':'营业收入','营业收入合计':'营业收入','营业总收入':'营业收入',
 '营业成本':'营业成本','营业成本合计':'营业成本','净利润':'净利润','净利润（净亏损以“－”号填列）':'净利润',
 '毛利润合计':'毛利润','投资活动产生的现金流量净额':'投资现金净额',
 '经营活动产生的现金流量净额':'经营现金净额','筹资活动产生的现金流量净额':'筹资现金净额',
 '资产负债率':'资产负债率','流动比率':'流动比率',
}

def compact(s):return re.sub(r'\s+','',s).replace('−','-').replace('－','-')
def dec(s):
    try:return Decimal(compact(s).replace(',','').replace('，','').replace('%','').replace('％',''))
    except InvalidOperation:return None

def periods(s):
    """Explicit year/range dates in their document order; no wall-clock inference."""
    s=compact(s)
    s=re.sub(r'(20\d{2})[年]?[~～—–至-](20\d{2})年?',lambda m:'、'.join(str(y)+'年' for y in range(int(m[1]),min(int(m[2]),int(m[1])+5)+1)),s)
    found=[]
    for m in re.finditer(r'(20\d{2})年(?:(\d{1,2})(?:[~～—–至-](\d{1,2}))?月)?',s):
        y=m[1];end=m[3] or m[2]
        p=y+'-'+str(int(end)).zfill(2) if end else y
        if p not in found:found.append(p)
    return found

def glossary_periods(doc):
    for b in doc.blocks[:250]:
        if b.kind=='table':
            for row in b.el.findall(w('tr')):
                vals=[text(tc) for tc in row.findall(w('tc'))]
                if vals and ('报告期' in vals[0] or '最近两年及一期' in vals[0]) and not '末' in vals[0]:
                    p=periods(' '.join(vals[1:]))
                    if p:return p
        elif re.search(r'报告期[:：]',b.text):
            p=periods(b.text)
            if p:return p
    return []

def scope_for(b, prefix=''):
    """Avoid matching parent, guarantor, sector business and issuer totals."""
    from .precision import nearest_scope
    ctx=' '.join(b.path)
    scope=nearest_scope(b.path,prefix)
    if scope!='发行人':return scope
    if re.search(r'募集资金运用对.*财务|本次发行对.*财务|财务结构的影响|模拟',ctx):return '发行模拟'
    if re.search(r'主要子公司|参股公司|关联方',ctx):return '其他主体'
    # A named subsidiary preceding the metric is not an issuer total.
    tail=re.split(r'[，,；;。]',prefix)[-1]
    if re.search(r'(?:业务|板块|子公司).{0,12}$',tail):return '业务分项'
    if tail and not re.search(r'发行人|公司|集团|截至|年度|年末|月末|报告期',tail) and len(tail)>2:return '其他主体'
    return '发行人'

@dataclass
class Fact:
    metric:str; period:str; scope:str; value:Decimal; tolerance:Decimal; display:str; unit:str; block:int; file:str; hash:str; locator:str; origin:str
    @property
    def key(self):return (self.metric,self.period,self.scope,'ratio' if self.unit in ('%','倍') else 'money')
    def evidence(self):return {'file':self.file,'file_id':self.hash,'block':self.block,'locator':self.locator,'text':self.origin,'value':self.display,'unit':self.unit,'period':self.period,'metric':self.metric,'scope':self.scope}

def fact(doc,b,metric,p,raw,unit,scope,origin):
    value=dec(raw)
    if value is None:return None
    precision=len(raw.rsplit('.',1)[1]) if '.' in raw else 0
    return Fact(metric,p,scope,value*SCALE[unit],Decimal(5).scaleb(-precision-1)*SCALE[unit],raw+unit,unit,b.index,doc.name,doc.hash,' > '.join(b.path[-4:]) or '正文块 '+str(b.index+1),origin)

def table_rows(block):
    rows=[]
    for tr in block.el.findall(w('tr')):
        row=[]
        for tc in tr.findall(w('tc')):
            g=tc.find('./'+w('tcPr')+'/'+w('gridSpan'));span=int(g.get(w('val'),'1')) if g is not None else 1
            row.extend([compact(text(tc))]*span)
        rows.append(row)
    return rows

def surrounding_unit(doc,b):
    for prev in reversed(doc.blocks[max(0,b.index-5):b.index]):
        m=re.search(r'单位[:：]\s*(亿元|万元|元)',compact(prev.text))
        if m:return m[1]
        if prev.heading:break
    return ''

def extract_facts(doc):
    facts=[];default=glossary_periods(doc)
    for b in doc.blocks:
        if b.is_toc or b.images:continue
        if b.kind=='table':
            rows=table_rows(b)
            if not rows:continue
            first=rows[0];unit=surrounding_unit(doc,b)
            # Before/after simulation columns are separate scenarios, not two
            # incompatible historical values for the same issuer and period.
            if any(re.search(r'发行后|模拟|调整后',v) for v in first):continue
            for ri,row in enumerate(rows):
                if not row:continue
                label=re.sub(r'^[一二三四五六七八九十\d]+[、.．]','',row[0]).replace('（%）','').replace('(%)','').replace('（倍）','').replace('(倍)','')
                key=TABLE_NAMES.get(label)
                if not key:continue
                if not unit and key not in ('资产负债率','流动比率'):continue
                for ci,cell in enumerate(row[1:],1):
                    # Only use unambiguously labelled year columns. Multi-tier header
                    # percentages are excluded from money totals.
                    if ci>=len(first):continue
                    per=periods(first[ci])
                    if len(per)!=1:continue
                    if any(ci<len(r) and re.search(r'占比|比例|比重',r[ci]) for r in rows[1:min(ri,3)]):continue
                    if not re.fullmatch(NUM+r'[%％]?',cell):continue
                    u='%' if key=='资产负债率' else '倍' if key=='流动比率' else unit
                    f=fact(doc,b,key,per[0],cell.replace('%','').replace('％',''),u,scope_for(b),label+' | '+first[ci]+' | '+cell)
                    if f:facts.append(f)
            continue
        s=compact(b.text)
        for metric,pattern in METRICS.items():
            for m in re.finditer(pattern+r'(?:余额|账面价值)?(?:分别)?(?:为|是)([^。；;]{0,130})',s):
                prefix=s[:m.start()];scope=scope_for(b,prefix)
                if scope in ('业务分项','其他主体'):continue
                # Restrict total metrics to direct issuer statements (not attribution).
                if metric=='净利润' and re.search(r'扣除|归属于|归母',prefix[-18:]):continue
                values=[]
                valuepart=m[1]
                if metric in ('资产负债率','流动比率'):
                    # A ratio sequence ends before the following descriptive clause.
                    for v in re.finditer('('+NUM+r')(%|％|倍)?',valuepart.split('，')[0]):values.append((v[1],'%' if metric=='资产负债率' else '倍'))
                else:
                    values=[(v[1],v[2]) for v in MONEY.finditer(valuepart.split('，')[0])]
                if not values:continue
                ps=periods(prefix)
                if not ps and re.search(r'近[两二三]年|报告期|各期',prefix):ps=default
                if len(values)==1 and ps:ps=ps[-1:]
                if len(values)!=len(ps) or not ps:continue
                for p,(v,u) in zip(ps,values):
                    f=fact(doc,b,metric,p,v,u,scope,s[max(0,m.start()-80):m.end()])
                    if f:facts.append(f)
    return facts

def equal(a,b):return abs(a.value-b.value)<=max(a.tolerance,b.tolerance)

def finding(doc,block,rule,title,detail,evidence=(),severity='warning',source=False,blocks=None):
    b=doc.blocks[block] if block is not None else None
    return {'id':doc.hash[:12]+':'+rule+':'+str(block)+':'+str(len(detail)), 'rule':rule,'severity':severity,'file':doc.name,'file_id':doc.hash,'block':block,'blocks':blocks or ([block] if block is not None else []),'matter':b.matter if b else '', 'locator':' > '.join(b.path[-4:]) if b else '全文','title':title,'detail':detail,'evidence':list(evidence),'origin':'source' if source else 'target'}

def bond_amount(doc):
    hits=[]
    for b in doc.blocks:
        if b.is_toc:continue
        s=compact(b.text)
        for pattern in [r'本次债券(?:的)?(?:发行规模|发行总额|发行金额|募集资金规模)(?:为|拟为)?[^。；\d]{0,20}('+NUM+r')(亿元|万元)',r'发行金额[:：]本次债券[^。；\d]{0,30}('+NUM+r')(亿元|万元)']:
            m=re.search(pattern,s)
            if m:hits.append((dec(m[1])*SCALE[m[2]],b,m[1],m[2]));break
    vals={a for a,_,_,_ in hits}
    return hits[0] if len(vals)==1 else None

def internal_checks(doc):
    issues=[];facts=extract_facts(doc);groups=defaultdict(list)
    for f in facts:groups[f.key].append(f)
    for key,fs in groups.items():
        distinct=[]
        for f in fs:
            if not any(equal(f,g) for g in distinct):distinct.append(f)
        if len(distinct)>1:
            bs=sorted({f.block for f in fs})
            issues.append(finding(doc,bs[0],'numeric-conflict','同指标披露数值不同：'+key[0],key[1]+'，'+key[2]+'。不同位置分别披露 '+ '、'.join(f.display for f in distinct[:6])+'；可能涉及口径差异，需要回到原始依据确认，不能按出现次数选值。',[f.evidence() for f in fs[:8]],source=True,blocks=bs))
    # Definition contradictions require no inference about current exchange rules.
    for b in doc.blocks:
        if b.kind!='table':continue
        for tr in b.el.findall(w('tr')):
            cols=[compact(text(tc)) for tc in tr.findall(w('tc'))]
            if len(cols)<2:continue
            a=cols[0];v=' '.join(cols[1:])
            if ('深交所' in a and '上海证券交易所' in v) or ('上交所' in a and '深圳证券交易所' in v):
                issues.append(finding(doc,b.index,'definition-conflict','交易所简称与释义不一致',a+' 对应 '+v+'。保留原始资料；请确认应采用的简称，不自动替来源改口径。',source=True))
    controls=[]
    for b in doc.blocks:
        if b.kind!='table':continue
        for tr in b.el.findall(w('tr')):
            cs=[compact(text(tc)) for tc in tr.findall(w('tc'))]
            if len(cs)>1 and '控股股东' in cs[0] and len(cs[0])<45:
                controls.append((b.index,cs[0],''.join(cs[1:]).removeprefix('指')))
    vals={v for _,_,v in controls if v}
    if len(vals)>1:
        bs=sorted({i for i,_,_ in controls})
        issues.append(finding(doc,bs[0],'control-definition','控股股东释义存在不同指向',
            '同份材料的释义分别为：'+'；'.join(k+'＝'+v for _,k,v in controls)+'。不自动把实际控制人与直接控股股东合并。',source=True,blocks=bs))
    # Explicit balance-sheet identities, only one coherent table / period.
    for b in doc.blocks:
        if b.kind!='table':continue
        bf=[f for f in facts if f.block==b.index]
        by=defaultdict(dict)
        for f in bf:by[(f.period,f.scope)][f.metric]=f
        for (period,scope),m in by.items():
            if {'资产总额','负债总额','净资产'}<=m.keys():
                a,l,e=[m[k] for k in ('资产总额','负债总额','净资产')]
                if abs(a.value-l.value-e.value)>a.tolerance+l.tolerance+e.tolerance:
                    issues.append(finding(doc,b.index,'balance-equation','资产负债表勾稽未通过',period+'，'+scope+'。资产 '+a.display+' 不等于负债 '+l.display+' 加权益 '+e.display+'。',[v.evidence() for v in (a,l,e)],source=True))
    issues.extend(asset_percentage_checks(doc,facts))
    # Simulation uses explicitly declared issue proceeds, not an estimated figure.
    amt=bond_amount(doc)
    if amt:
        amount,ab,raw,unit=amt
        for b in doc.blocks:
            if b.kind!='table':continue
            rows=table_rows(b)
            if not rows or not any('模拟变动额' in x for x in rows[0]):continue
            context=''.join(x.text for x in doc.blocks[max(0,b.index-13):b.index])
            if not ('计入所有者权益' in context or '分类为权益' in context or '计入权益' in context):continue
            if not re.search(r'全部用于偿还',context):continue
            tu=surrounding_unit(doc,b)
            if not tu:continue
            ci=next((i for i,x in enumerate(rows[0]) if '模拟变动额' in x),None)
            eq=next((r for r in rows if r and r[0] in ('所有者权益','所有者权益合计','净资产')),None)
            if eq and ci is not None and ci<len(eq) and dec(eq[ci]) is not None:
                val=dec(eq[ci])*SCALE[tu]
                if abs(amount-val)>Decimal('0.02'):
                    issues.append(finding(doc,b.index,'simulation-proceeds','发行规模与模拟权益增量不一致','文字假设募集 '+raw+unit+'、全部偿债并计入权益；模拟表权益只增加 '+eq[ci]+tu+'。请核实是否仍沿用其他发行规模的模拟表。',[{'file':doc.name,'file_id':doc.hash,'block':ab.index,'text':ab.text,'locator':'发行规模'}],source=True,blocks=[b.index]))
    return issues,facts

def asset_percentage_checks(doc,facts):
    """Recompute explicitly paired asset amounts and percentages, same scope/year.
    Accept display-rounding intervals, never infer amounts from a missing unit.
    """
    assets=defaultdict(list)
    for f in facts:
        if f.metric=='资产总额':assets[(f.period,f.scope)].append(f)
    issues=[]
    pattern=r'(应收账款|其他应收款|投资性房地产|存货|固定资产|在建工程|无形资产)(?:账面价值|余额)?(?:分别)?为([^。；]{0,180}?)，(?:在总资产中占比|占总资产(?:的)?(?:比例|比重|占比)?)(?:分别)?为?([^。；，]{0,70})'
    for b in doc.blocks:
        if b.kind!='paragraph' or b.is_toc:continue
        tx=compact(b.text)
        for m in re.finditer(pattern,tx):
            amounts=[(x[1],x[2]) for x in MONEY.finditer(m[2])]
            percents=re.findall('('+NUM+r')[%％]',m[3]);ps=periods(tx[:m.start()]);scope=scope_for(b,tx[:m.start()])
            if scope not in ('发行人','发行人本部'):continue
            if not (len(amounts)==len(percents)==len(ps)) or not amounts:continue
            for period,(raw,unit),rate in zip(ps,amounts,percents):
                denominators=assets.get((period,scope),[])
                if not denominators or any(not equal(a,denominators[0]) for a in denominators):continue
                denominator=min(denominators,key=lambda x:x.tolerance)
                numerator=fact(doc,b,m[1],period,raw,unit,scope,m[0]);reported=dec(rate)
                precision=len(rate.rsplit('.',1)[1]) if '.' in rate else 0;tolerance=Decimal(5).scaleb(-precision-1)
                if denominator.value<=denominator.tolerance or numerator.value<0:continue
                lower=(numerator.value-numerator.tolerance)/(denominator.value+denominator.tolerance)*100
                upper=(numerator.value+numerator.tolerance)/(denominator.value-denominator.tolerance)*100
                if lower<=reported+tolerance and upper>=reported-tolerance:continue
                computed=numerator.value/denominator.value*100
                issues.append(finding(doc,b.index,'asset-percentage','资产占比与金额计算不一致：'+m[1],period+'，'+scope+'。'+raw+unit+' ÷ 总资产 '+denominator.display+' ×100，约为 '+f'{computed:.4f}'+'%；该处写 '+rate+'%。已考虑显示位数的舍入区间，请确认分母和口径。',[numerator.evidence(),denominator.evidence()],source=True))
    return issues

def target_checks(doc,sources,source_facts):
    issues=[];by=defaultdict(list)
    for f in source_facts:by[f.key].append(f)
    for f in extract_facts(doc):
        matches=by.get(f.key,[])
        if matches and not any(equal(f,g) for g in matches):
            issues.append(finding(doc,f.block,'target-number','数字与来源不一致：'+f.metric,f.period+'，'+f.scope+'。底稿 '+f.display+'；来源 '+ '、'.join(dict.fromkeys(g.display for g in matches))+'。数字证据不授权程序改写相邻的核查结论。',[f.evidence()]+[g.evidence() for g in matches[:3]]))
    # Explicit contemporary issuer statement using a reporting period absent from
    # the new source glossary. Historical events and issuer histories are excluded.
    current=set(p for s in sources for p in glossary_periods(s))
    for b in doc.blocks:
        if b.is_toc or b.kind!='paragraph':continue
        s=compact(b.text)
        if re.match(r'^(?:截至)?20\d{2}',s) and re.search(r'发行人(?:本部)?(?:营业|净利润|有息|总资产|资产负债|一年以内)',s[:110]) and not re.search(r'历史|沿革|前次|设立|决议', ' '.join(b.path)):
            ps=set(periods(s.split('，')[0]))
            if current and ps and ps.isdisjoint(current):
                issues.append(finding(doc,b.index,'reporting-period','本期分析仍使用旧报告期','此处使用 '+ '、'.join(sorted(ps))+'；本项目来源释义为 '+ '、'.join(sorted(current))+'。这是期间适用性提示，不把历史数据一律认定错误。'))
    # Corporate control names can be compared because these labels have a defined
    # subject. Ordinary counterparties are deliberately not flagged by city/name.
    controls=set();control_aliases=set()
    for sd in sources:
        for b in sd.blocks:
            if b.kind!='table':continue
            for tr in b.el.findall(w('tr')):
                cs=[compact(text(tc)) for tc in tr.findall(w('tc'))]
                if cs and ('控股股东' in cs[0] or '实际控制人' in cs[0]) and len(cs)>1:
                    controls.add(''.join(cs[1:]).removeprefix('指'))
                    control_aliases.update(re.split(r'[/／、]',cs[0]))
    corp=r'[\u4e00-\u9fff（）()A-Za-z]{3,45}(?:有限公司|集团)'
    for b in doc.blocks:
        if b.is_toc or b.kind!='paragraph' or re.search(r'历史|沿革|前次', ' '.join(b.path)):continue
        s=compact(b.text)
        m=re.search(r'(?:控股股东(?:为|系|是)|实际控制人(?:为|系|是))('+corp+')',s)
        if not m:
            shareholder=re.search(r'(?:^|[，；。])('+corp+r')持有发行人(?:的)?(\d+(?:\.\d+)?)%(?:的)?股权',s)
            if shareholder and dec(shareholder[2])>50:m=shareholder
        if m and controls and not any(m[1] in c or c in m[1] for c in controls|control_aliases):
            issues.append(finding(doc,b.index,'control-subject','控制关系主体与来源不同','底稿写明 '+m[1]+'；来源释义为 '+ '、'.join(sorted(controls))+'。不能把主体差异按近似文字替换。'))
    # Local document-type leftovers, but quoted external report titles are exempt.
    for b in doc.blocks:
        if b.is_toc:continue
        if re.search(r'截至本立项申请报告(?:出具|签署)日',b.text) and '立项' not in doc.name:
            issues.append(finding(doc,b.index,'document-context','当前文种仍带有立项报告表述','此处使用“本立项申请报告”作当前文档自称。请确认适用时点与文种，程序不把签署日替换成今天。'))
    # Every live header/footer is checked separately from historical deletions.
    if doc.profile.get('bond') and '可续期' in doc.profile['bond']:
        for part,data in doc.pkg.entries.items():
            if re.fullmatch(r'word/header\d+\.xml',part):
                s=text(xml(data))
                if '非公开发行公司债券之核查意见' in s and '可续期' not in s:
                    i=finding(doc,None,'header-bond-type','页眉债券名称缺少“可续期”',part+'：'+s+'。页眉原样保留并报告，不能只检查正文。')
                    i['locator']=part;issues.append(i)
    # Group repeated same-rule/same-block findings to keep the work queue usable.
    seen={};result=[]
    for i in issues:
        key=(i['rule'],i['block'],i['title'])
        if key in seen:
            existing=seen[key]
            if i['detail'] not in existing['detail']:existing['detail']+='\n'+i['detail']
            for evidence in i.get('evidence',[]):
                if evidence not in existing.setdefault('evidence',[]):existing['evidence'].append(evidence)
            continue
        seen[key]=i;result.append(i)
    return result


def replacement_risks(target,source):
    """Independent vetoes on automated whole-block copying."""
    a=target.text;b=source.text;risks=[]
    embedded=re.findall(r'(?:。|\n)\s*\d+[、.]\s*([^。；\n]{2,25})(?=$|\n)',a)
    if any(compact(h) not in compact(b) for h in embedded):
        risks.append('旧段落含内嵌小标题，来源未覆盖整个子事项；不能自动删掉余下内容。')
    if len(compact(b))<len(compact(a))*.78 and len(a)>120:
        risks.append('来源明显短于目标，可能遗漏原段落中的事实或子事项。')
    # Preserve informative table captions; a table alone doesn't authorize deleting them.
    if target.has_table and source.has_table:
        prefix=[]
        for block in target.blocks:
            if block.kind=='table':break
            if re.search(r'如下表|下表|情况如下|余额情况|^表[:：]',block.text):prefix.append(block.text)
        source_prefix=[]
        for block in source.blocks:
            if block.kind=='table':break
            if block.text.strip() and not re.match(r'^单位[:：]',block.text.strip()):source_prefix.append(block.text)
        old_labels={compact(text(tr.find(w('tc')))) for block in target.blocks if block.kind=='table' for tr in block.el.findall(w('tr'))[2:] if tr.find(w('tc')) is not None}
        new_labels={compact(text(tr.find(w('tc')))) for block in source.blocks if block.kind=='table' for tr in block.el.findall(w('tr'))[2:] if tr.find(w('tc')) is not None}
        lost={label for label in old_labels-new_labels if len(label)>3 and not re.fullmatch(r'[\d.、（）()-]+',label)}
        if lost:risks.append('来源表格缺少原稿部分行项目或行名改变：'+'、'.join(sorted(lost)[:5])+'。需确认是本版调整还是遗漏，未自动删行。')
        if prefix and not source_prefix:
            risks.append('候选表格没有携带原稿的表题或引导句；请核对，避免表格更新时丢失事项说明。')
    # Do not drop explicit measurement units while reusing otherwise similar tables.
    if target.has_table and re.search(r'单位[:：]',a) and not re.search(r'单位[:：]',b):
        risks.append('目标含计量单位，候选未携带对应单位；请核对后再采用，不能静默丢失单位。')
    if re.search(r'发行人本部|母公司本部',a) and not re.search(r'发行人本部|母公司本部',b):
        risks.append('来源未覆盖目标中明确列示的发行人本部，不能自动删除业务主体。')
    units=lambda t:set(re.findall(r'(?:单位[:：])([^\n]{1,24})',t))
    if units(a) and units(b) and units(a)!=units(b):
        risks.append('计量单位发生变化，需要核对换算与原始金额，不能仅按表格形状自动替换。')
    return risks

def precise_bond_updates(doc,sources):
    """Local patches to explicitly defined current-bond amount/type, not sentences.

    Requires one consistent source amount, excludes historical resolutions and
    independent opinion blocks. Targets keep their own runs and paragraph style.
    """
    amounts=[(sd,bond_amount(sd)) for sd in sources]
    amounts=[x for x in amounts if x[1]]
    if not amounts or len({x[1][0] for x in amounts})!=1:return []
    sd,(amount,sb,_,_)=amounts[0];out=[]
    for b in doc.blocks:
        if b.kind!='paragraph' or b.heading or b.protected or b.is_toc or b.images:continue
        s=b.text
        if '本次债券' not in s or re.search(r'董事会|股东决议|审议通过|前次|曾经|已发行|历史',s):continue
        # A displayed amount near the explicit issue cap; not an unrelated balance.
        m=re.search(r'(?:发行总额|发行规模|发行金额|拟申请.{0,36}发行总额|非公开发行总额)[^。；\d]{0,20}('+NUM+r')\s*(亿元|万元)',s)
        replacements=[]
        if m:
            raw,unit=m[1],m[2];expected=amount/SCALE[unit]
            if dec(raw)!=expected:
                dp=len(raw.rsplit('.',1)[1]) if '.' in raw else 0
                newraw=f'{expected:.{max(dp,max(0,-expected.normalize().as_tuple().exponent))}f}'
                # Changing the cap and its parenthetical repetition only.
                parenthetical=re.match(r'\s*[（(]\s*含\s*'+re.escape(raw)+r'\s*'+unit+r'\s*[）)]',s[m.end():])
                segment=s[m.start():m.end()+(parenthetical.end() if parenthetical else 0)]
                revised=re.sub(r'(?<![\d.])'+re.escape(raw)+r'(?=\s*'+unit+r')',newraw,segment)
                replacements.append((segment,revised))
        if doc.profile.get('kind')=='opinion' and '可续期' in sd.profile.get('bond','') and re.search(r'公司债券[（(]以下简称[“"\']本次债券',s) and '可续期公司债券' not in s:
            replacements.append(('的公司债券（以下简称','的可续期公司债券（以下简称'))
        new=s
        for old,value in replacements:new=new.replace(old,value)
        if new!=s:
            out.append({'block':b.index,'text':new,'replacements':replacements,'source':sd,'source_block':sb})
    return out
