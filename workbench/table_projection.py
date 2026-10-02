"""Deterministic whole-table projection from a single source financial statement.

Only exact, complete row/period/unit mappings are allowed. No arithmetic, missing
value imputation, year rollover, fuzzy row lookup or cross-statement cell mixing.
"""
from __future__ import annotations
import copy,re
from .docxio import w,text
from .precision import nearest_scope,table_identity,source_elements,table_cells
from .assurance import periods,compact


def row_key(s):
    s=compact(s)
    s=re.sub(r'[（(](?:损失|亏损|收益|净亏损)[^）)]*[）)]','',s)
    return re.sub(r'^(?:[一二三四五六七八九十]+[、.．]|(?:加|减)[：:])','',s)


def simple_grid(u):
    tables=[b for b in u.blocks if b.kind=='table']
    if len(tables)!=1:return None
    b=tables[0];rows=b.el.findall(w('tr'))
    if len(rows)<3:return None
    if b.el.findall('.//'+w('vMerge')) or any(int(x.get(w('val'),'1'))>1 for x in b.el.findall('.//'+w('gridSpan'))):return None
    cs=rows[0].findall(w('tc'));header=[text(c).strip() for c in cs]
    labels=[periods(s) for s in header[1:]]
    if not labels or not all(len(p)==1 for p in labels) or len(set(p[0] for p in labels))!=len(labels):return None
    if not all(len(row.findall(w('tc')))==len(cs) for row in rows):return None
    return b,rows,[p[0] for p in labels]


def candidates(target,sources,source_flags):
    tg=simple_grid(target)
    if not tg:return []
    tb,tr,period=tg;identity=table_identity(target)
    # Only financial metric extracts; never customer/supplier sets or peer tables.
    wanted=[row_key(text(r.findall(w('tc'))[0])) for r in tr[1:]]
    if len(wanted)!=len(set(wanted)) or any(not re.search(r'收益|利润|收入|费用|成本|资产|负债|现金|权益',k) for k in wanted):return []
    if not identity['units'] or len(wanted)<2:return []
    mapped=[]
    for doc in sources:
        seen=set()
        for su in doc.source_units:
            sg=simple_grid(su)
            if not sg:continue
            sb,sr,sp=sg
            if sb.index in seen:continue
            seen.add(sb.index)
            si=table_identity(su)
            if nearest_scope(target.path)!=nearest_scope(su.path) or set(period)!=set(sp) or identity['units']!=si['units']:continue
            # Projection is explicitly from a larger statement, not a competing
            # table version with missing rows.
            if len(sr)<=len(tr) or not re.search(r'利润表|资产负债表|现金流量表',su.locator+' '+su.text[:180]):continue
            lookup={}
            for n,row in enumerate(sr[1:],1):lookup.setdefault(row_key(text(row.findall(w('tc'))[0])),[]).append(n)
            if any(len(lookup.get(k,[]))!=1 for k in wanted):continue
            if any(source_flags.get((doc.hash,i)) for i in range(su.start,su.end)):continue
            rows=[0]+[lookup[k][0] for k in wanted];cols=[0]+[sp.index(p)+1 for p in period]
            descriptor={'table_block':sb.index,'rows':rows,'columns':cols,'target_prefix':[b.index for b in target.blocks if b.kind!='table'], 'periods':period,'row_labels':wanted}
            candidate={'doc_hash':doc.hash,'source_name':doc.name,'unit_id':'projection'+str(sb.index),'start':sb.index,'end':sb.index+1,'indices':[sb.index],'score':1.0,'content_score':1.0,'heading_score':1.0,'table_score':1.0,'literal_score':0,'locator':su.locator,'text':su.text,'exact':False,'unsupported':'','source_warnings':[],'span':None,'identity':si,'projection':descriptor}
            projected=project_elements(doc,candidate)
            candidate['_matrix']=table_cells(projected)
            mapped.append(candidate)
    # Two source tables disagree? No automatic projection, even if one matches the draft.
    if not mapped or any(c['_matrix']!=mapped[0]['_matrix'] for c in mapped[1:]):return []
    for c in mapped:del c['_matrix']
    return sorted(mapped,key=lambda c:(c['source_name'],c['start']))[:4]


def project_elements(doc,c):
    pr=c['projection'];original=doc.blocks[pr['table_block']].el;table=copy.deepcopy(original)
    source_rows=original.findall(w('tr'))
    for row in list(table.findall(w('tr'))):table.remove(row)
    for ri in pr['rows']:
        if not 0<=ri<len(source_rows):raise ValueError('来源行位置失效。')
        row=copy.deepcopy(source_rows[ri]);cells=row.findall(w('tc'))
        for cell in list(row.findall(w('tc'))):row.remove(cell)
        for ci in pr['columns']:
            if not 0<=ci<len(cells):raise ValueError('来源列位置失效。')
            row.append(copy.deepcopy(cells[ci]))
        table.append(row)
    grid=table.find(w('tblGrid'))
    if grid is not None:
        original_cols=list(grid)
        if len(original_cols)==len(pr['columns']):
            grid[:]=[copy.deepcopy(original_cols[i]) for i in pr['columns']]
    return [table]
