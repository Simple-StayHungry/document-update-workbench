"""Analysis plans, explicit decisions, provenance and export validation."""
from __future__ import annotations
import collections,copy,difflib,html,json,re,zipfile,time
from pathlib import Path
from datetime import date as calendar_date
from .docxio import Package,w,text,dump,xml,sha,transformed,transform_paragraph,resolve_revisions,semantic_blocks,validate_package,structural_digest_content
from .model import Document,Unit
from .matching import Index,normal
from .quality import inspect as quality_inspect
from .assurance import internal_checks,target_checks,precise_bond_updates
from .table_projection import candidates as projection_candidates
from .precision import atomic_units,table_identity,source_elements,fingerprint,assess_complete_section
from .plans import build,freeze,font_normalize,content_digest,retained_plan
from .version import BUILD,SCHEMA
from .ingest import validate_material_hash,validate_source_record
from .source_policy import prepare_target,classify_proposal,paragraph_merges,inventory,code_fingerprint,retain_unmatched

class Engine:
    def __init__(self):self.cache={};self.fact_cache={}
    def load(self,file):
        key=file['id']
        # Hash once when the document first enters the engine cache.  The old
        # implementation re-read and SHA256-hashed multi-megabyte DOCX files on
        # every candidate freeze/build call, causing hundreds of redundant disk
        # reads and apparent UI hangs.  Export still performs a full boundary
        # hash verification for every target/source before writing anything.
        cached=self.cache.get(key)
        if cached is not None:return cached
        path=Path(file['path'])
        if not path.is_file() or sha(path.read_bytes())!=key:raise ValueError('材料内容已经变化，请重新导入后分析。')
        cached=Document(path,file['name']);self.cache[key]=cached
        return cached
    def analyze(self,files,params,progress=lambda *a:None):
        if params.get('author') not in (None,'','柒'):raise ValueError('当前核查同步规则要求修订作者为柒。')
        for f in files:
            if f.get('role') in ('source','target'):validate_material_hash(f)
            if f.get('role')=='source':validate_source_record(f)
        targets=[f for f in files if f.get('role')=='target'];sources=[f for f in files if f.get('role')=='source'];refs=[f for f in files if f.get('role')=='reference']
        if not targets:raise ValueError('请选择至少一份待更新的 Word。')
        if any(f['name'].strip().startswith('老师') for f in sources):raise ValueError('老师稿不能作为事实来源。')
        if not sources:raise ValueError('请选择至少一份最新版来源 Word；示例定稿不能代替独立来源材料。')
        if params.get('date'):
            try:calendar_date.fromisoformat(params['date'])
            except (ValueError,TypeError):raise ValueError('调查日期格式必须为 YYYY-MM-DD，且必须是真实有效的日期。')
        progress(5,'正在读取当前可见稿与材料结构')
        source_by_hash={}
        for f in sorted(sources,key=lambda f:(f['name'],f['id'])):source_by_hash.setdefault(f['id'],f)
        sd=sorted([self.load(f) for f in source_by_hash.values()],key=lambda d:(0 if d.profile['kind']=='prospectus' else 1,d.name,d.hash))
        issuers=sorted({d.profile['issuer'] for d in sd if d.profile['issuer']})
        issuer=(params.get('issuer') or (issuers[0] if len(issuers)==1 else '')).strip()
        if not issuer:raise ValueError('来源涉及不同发行人或未识别发行人。请保留同一项目材料并填写目标发行人。')
        bad=[d.name for d in sd if d.profile['issuer'] and d.profile['issuer']!=issuer]
        if bad:raise ValueError('来源发行人与目标不一致：'+'；'.join(bad)+'。未进行跨公司混用。')
        if any(not d.profile['issuer'] for d in sd) and not params.get('allow_unknown_source'):
            raise ValueError('部分来源无法识别发行人。请核对来源后，在更多设置中确认无发行人标识的来源。')
        progress(8,'正在交叉检查来源数字、报表勾稽与发行条款')
        source_issues=[];source_facts=[]
        for source in sd:
            issues,facts=internal_checks(source);source_issues.extend(issues);source_facts.extend(facts)
        primary=[d for d in sd if d.profile['kind']=='prospectus']
        if not primary:raise ValueError('募集同步需要明确的当前募集说明书；核查意见不自动替代募集来源。')
        if len(primary)>1:raise ValueError('存在多份不同字节的募集说明书；请明确唯一当前版本，其余设为对照后重新分析。')
        if any(d.profile['historical_revisions'] for d in sd):raise ValueError('正式来源含历史修订；请提供明确当前版本或先确认来源基线策略。')
        idx=Index(sd,source_issues);proposals=[];reports=[]
        for n,f in enumerate(targets):
            d=self.load(f);cross=bool(d.profile['issuer'] and d.profile['issuer']!=issuer)
            prepare_target(d,sd)
            progress(10+int(n/max(1,len(targets))*75),'正在匹配 '+d.name)
            localprops=[]
            for group in d.units:
                section_assessment=assess_complete_section(d,group,sd,idx.source_flags)
                if section_assessment:
                    st,cc,why=section_assessment
                    sub=group
                    p={'id':f['id']+':section:'+sub.id,'target_id':f['id'],'target_name':d.name,'start':sub.start,'end':sub.end,'matter':sub.matter or '正文','heading':sub.heading,'locator':sub.locator,'old_text':sub.text,'status':st,'reason':why,'candidates':cc,'selected':0 if st in ('auto','unchanged') else None,'decision':'accept' if st=='auto' else 'same' if st=='unchanged' else 'pending','note':'','kind':'material','has_table':sub.has_table,'table_mode':'block','identity':table_identity(sub) if sub.has_table else None,'sync_mode':'block'}
                    localprops.append(p)
                    continue
                for sub in atomic_units(d,group):
                    st,cc,why=idx.assess(sub)
                    p={'id':f['id']+':'+sub.id,'target_id':f['id'],'target_name':d.name,'start':sub.start,'end':sub.end,'matter':sub.matter or '正文','heading':sub.heading,'locator':sub.locator,'old_text':sub.text,'status':st,'reason':why,'candidates':cc,'selected':0 if st in ('auto','unchanged') else None,'decision':'accept' if st=='auto' else 'same' if st=='unchanged' else 'pending' if st in ('review','missing') else 'protected','note':'','kind':'material','has_table':sub.has_table,'table_mode':'projection' if cc and cc[0].get('projection') else 'copy','identity':table_identity(sub) if sub.has_table else None}
                    localprops.append(p)
            from .scopes import reconcile,source_gaps
            localprops,scope_notes=reconcile(d,idx,localprops)
            localprops=[classify_proposal(p,d) for p in localprops]
            localprops=paragraph_merges(d,sd,localprops,idx.source_flags)
            from .figures import figure_proposals
            figures=figure_proposals(d,sd,idx.source_flags)
            for fp in figures:
                if not fp.get('content_class'):fp['content_class']='prospectus_source' if fp['decision']=='accept' else 'external_protected' if fp['decision']=='protected' else 'source_missing_or_ambiguous'
            for fp in figures:
                overlap=[p for p in localprops if p['start']<fp['end'] and fp['start']<p['end']]
                if any(p['start']<fp['start'] or p['end']>fp['end'] for p in overlap):
                    fp.update(status='review',decision='pending',selected=None,content_class='copy_boundary_pending',reason='图文范围与正文计划交叉，保留待核对。')
                    for p in overlap:p.update(status='review',decision='pending',selected=None,content_class='copy_boundary_pending',reason='图文范围交叉，未自动覆盖。')
                    continue
                localprops=[p for p in localprops if p not in overlap]+[fp]
            for proposal in localprops:retain_unmatched(proposal)
            from .source_coverage import recover_source_coverage,recover_quoted_clause
            localprops=recover_source_coverage(d,sd,localprops,idx.source_flags)
            localprops=recover_quoted_clause(d,sd,localprops,idx.source_flags)
            from .source_additions import recover_source_additions
            localprops,addition_notes=recover_source_additions(d,sd,localprops,idx.source_flags)
            # Cover fields are a separate, source/user-backed update; no blind global replacement.
            changes=[]
            if cross:changes.append((d.profile['issuer'],issuer))
            bond=params.get('bond') or sd[0].profile['bond']
            if d.profile['bond'] and bond and bond!=d.profile['bond']:changes.append((d.profile['bond'],bond))
            period=params.get('period','').strip();date=params.get('date','').strip()
            for b in d.blocks[:min((m['start'] for m in d.matters),default=20)]:
                if b.kind!='paragraph' or not b.text or b.images:continue
                new=b.text
                for old,repl in changes:new=new.replace(old,repl)
                if period and re.search(r'报告期[:：]',new):new=re.sub(r'(?<=报告期)[:：].*', '：'+period+'）',new)
                if date and re.search(r'^调查日期[:：]',new):
                    m=re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})',date)
                    if not m:raise ValueError('调查日期格式应为 YYYY-MM-DD。')
                    new='调查日期：【%s】年【%s】月【%s】日'%(m[1],int(m[2]),int(m[3]))
                if new!=b.text:
                    localprops.append({'id':f['id']+':meta'+str(b.index),'target_id':f['id'],'target_name':d.name,'start':b.index,'end':b.index+1,'matter':'封面信息','heading':'项目名称与日期','locator':'封面 > '+b.text[:60],'old_text':b.text,'new_text':new,'status':'auto','reason':'项目名称来自所选来源；日期和报告期仅按你明确填写的值更新。','candidates':[],'selected':None,'decision':'accept','note':'','kind':'metadata','replacements':changes})
                    edits=[{'field':'issuer' if old==d.profile['issuer'] else 'bond','old':old,'new':repl} for old,repl in changes if old in b.text]
                    if period and re.search(r'报告期[:：]',b.text):edits.append({'field':'period','old':b.text,'new':period})
                    if date and re.search(r'^调查日期[:：]',b.text):edits.append({'field':'date','old':b.text,'new':date})
                    localprops[-1]['parameter_changes']=edits
            for p in localprops:
                retain_unmatched(p)
                freeze(self,p,files)
                p['context_start']=max(0,p['start']-2);p['context_end']=min(len(d.blocks),p['end']+2)
            # Group adjacent source-backed objects under the same workpaper heading into
            # ONE source-range review group. This is only a web review convenience:
            # it does not alter or intercept WPS tracked-change behavior.
            ordered=sorted(localprops,key=lambda x:(x['start'],x['end'],x['id']))
            groups=[];cur=[];curkey=None;curend=None
            for p in ordered:
                sel=p.get('selected')
                cand=(p.get('candidates') or [None])[sel] if isinstance(sel,int) and p.get('candidates') and 0<=sel<len(p['candidates']) else None
                eligible=p.get('kind')=='material' and p.get('action') not in ('source_merge','source_insert') and p.get('decision') in ('accept','same') and cand is not None
                key=(p['target_id'],p.get('matter',''),p.get('heading',''),p.get('locator',''),cand.get('doc_hash')) if eligible else None
                adjacent=bool(eligible and cur and key==curkey and p['start']==curend)
                if not adjacent:
                    if cur:groups.append(cur)
                    cur=[p];curkey=key;curend=p['end']
                else:
                    cur.append(p);curend=p['end']
            if cur:groups.append(cur)
            for g in groups:
                txid=g[0]['target_id']+':txn:'+str(g[0]['start'])+'-'+str(g[-1]['end'])
                title=g[0].get('heading') or g[0].get('matter') or '来源替换'
                for p in g:
                    p['transaction_id']=txid
                    p['transaction_mode']='source-range-group'
                    p['transaction_member_count']=len(g)
                    p['transaction_start']=g[0]['start'];p['transaction_end']=g[-1]['end'];p['transaction_title']=title
            if not d.units:
                reports.append({'id':f['id'],'name':d.name,'profile':d.profile,'matters':d.matters,'cross_project':cross,'warning':'未识别出可自动替换的核查情况区域；这份文件不会被宣称更新完成。','protected':d.protected_summary()})
            else:reports.append({'id':f['id'],'name':d.name,'profile':d.profile,'matters':d.matters,'cross_project':cross,'warning':'跨项目更新。旧稿的结论、截图和未匹配内容需重新核验，不能仅替换公司名称。' if cross else '', 'protected':d.protected_summary()})
            reports[-1]['scope_reconciliation']=scope_notes
            reports[-1]['source_additions']=addition_notes
            reports[-1]['source_gaps']=source_gaps(d,idx,localprops)
            # Read-only side channel: never handed the proposals, so it cannot
            # change the synchronisation plan by construction.
            from .section_items import detect_missing_items
            reports[-1]['missing_candidates']=detect_missing_items(d,sd)
            reports[-1]['sha256']=d.hash
            reports[-1]['quality_issues']=quality_inspect(d,sd)+target_checks(d,sd,source_facts)+d.safety_issues+reports[-1]['source_gaps']
            reports[-1]['coverage']={'eligible_blocks':sum(u.end-u.start for u in d.units),'total_blocks':len(d.blocks),'protected_blocks':sum(b.protected for b in d.blocks)}
            reports[-1]['object_inventory']=inventory(d,localprops)
            reports[-1]['coverage']['dispositions']=dict(collections.Counter(x['status'] for x in reports[-1]['object_inventory']))
            proposals.extend(localprops)
        progress(91,'正在检查匹配边界、重复来源和历史修订')
        # The teacher reference list is metadata only; no teacher fact ever enters Index.
        source_key=tuple(d.hash for d in sd);self.fact_cache[source_key]=source_facts
        result={'issuer':issuer,'params':dict(params,issuer=issuer),'proposals':proposals,'documents':reports,'sources':[{'id':f['id'],'name':f['name'],'sha256':self.load(f).hash,'profile':self.load(f).profile,'source_role':'prospectus' if self.load(f).profile['kind']=='prospectus' else 'authorized_formal_source'} for f in sources],'references':[{'id':f['id'],'name':f['name']} for f in refs],'created':time.strftime('%Y-%m-%d %H:%M:%S'),'schema':SCHEMA,'engine_version':BUILD,'source_issues':source_issues,'fact_count':len(source_facts),'code_sha256':code_fingerprint(),'input_roles':sorted((f['id'],f['role']) for f in files),'frozen_params':dict(params,issuer=issuer),'parameter_authorities':{'issuer':issuer,'bond':params.get('bond') or primary[0].profile['bond'],'date':params.get('date',''),'period':params.get('period','')}}
        result['summary']=summarize(result);progress(100,'匹配完成');return result
    def export(self,result,files,outdir,progress=lambda *a:None):
        if result.get('params')!=result.get('frozen_params'):raise ValueError('项目参数已变更，请重新分析。')
        if result.get('engine_version')!=BUILD or result.get('schema')!=SCHEMA or result.get('code_sha256')!=code_fingerprint():
            raise ValueError('程序或计划版本已变更，请重新分析。')
        if result.get('input_roles')!=sorted((f['id'],f['role']) for f in files) and result.get('input_roles')!=[list(x) for x in sorted((f['id'],f['role']) for f in files)]:
            raise ValueError('材料集合或来源用途已变更，请重新分析。')
        for f in files:
            if f.get('role') in ('source','target'):validate_material_hash(f)
            if f.get('role')=='source':validate_source_record(f)
        from .output_audit import audit_docx,element_paths
        outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True);fmap={f['id']:f for f in files}
        author=result['params'].get('author') or '柒';logs=[];artifacts=[];qa=[]
        for record in result['documents']+result['sources']:
            current=fmap.get(record['id']);expected_hash=record.get('sha256') or record['id']
            if not current or sha(Path(current['path']).read_bytes())!=expected_hash:
                raise ValueError('分析后输入文件已变更，请重新分析；不会按旧计划导出。')
        current_sources=[self.load(f) for f in files if f.get('role')=='source']
        # Reuse in-process source assurance when available; after a server restart
        # the cache is absent and we conservatively recompute from hash-verified inputs.
        current_key=tuple(d.hash for d in sorted(current_sources,key=lambda d:(0 if d.profile['kind']=='prospectus' else 1,d.name,d.hash)))
        current_facts=self.fact_cache.get(current_key)
        if current_facts is None:
            current_facts=[v for d in current_sources for v in internal_checks(d)[1]]
            self.fact_cache[current_key]=current_facts
        for n,record in enumerate(result['documents']):
            target=fmap[record['id']];doc=self.load(target)
            # Reloading once yields exactly the same block indices as the frozen analysis plan.
            fresh=Document(target['path'],target['name']);pkg=fresh.pkg
            if fresh.profile['historical_revisions']:
                raise ValueError('输入含历史修订；请先明确历史审阅基线，当前版本不自动接受历史修订后导出。')
            pristine=copy.deepcopy(pkg.root);base_parts=dict(pkg.entries);expected=semantic_blocks(pkg)
            ps=[p for p in result['proposals'] if p['target_id']==target['id']]
            for proposal in ps:
                if proposal.get('kind')=='material' and proposal.get('decision') in ('accept','same'):
                    if any(fresh.blocks[i].heading for i in range(proposal['start'],proposal['end'])):
                        raise ValueError('底稿标题及编号须保留原样，来源替换范围不得覆盖标题：'+proposal['id'])
                    selected=proposal.get('selected')
                    if isinstance(selected,int) and proposal.get('candidates'):
                        candidate=proposal['candidates'][selected]
                        source_doc=next((s for s in current_sources if s.hash==candidate['doc_hash']),None)
                        if source_doc and any(source_doc.blocks[i].heading for i in candidate.get('indices',range(candidate['start'],candidate['end']))):
                            raise ValueError('仅更新对应内容，来源替换范围不得新增标题：'+proposal['id'])
            prepare_target(fresh,current_sources)
            record['object_inventory']=inventory(fresh,ps)
            record['coverage']['dispositions']=dict(collections.Counter(x['status'] for x in record['object_inventory']))
            for p in ps:
                if p.get('action')=='source_merge' and p['decision'] in ('accept','same'):
                    parent=next((x for x in ps if x['id']==p.get('merge_into')),None)
                    if not parent or parent['decision'] not in ('accept','same'):
                        raise ValueError('合段删除所依赖的完整来源复制尚未采用。')
            for p in ps:
                if p.get('action')=='source_insert' and p['decision'] in ('accept','same'):
                    evidence=p['candidates'][p.get('selected') or 0].get('correspondence',{})
                    for anchor_id in evidence.get('anchors',[]):
                        anchor=next((a for a in ps if a['id']==anchor_id),None)
                        if not anchor or anchor['decision'] not in ('accept','same'):
                            raise ValueError('新增正文所依赖的相邻来源复制尚未采用。')
                        selected=anchor.get('selected')
                        ac=anchor['candidates'][selected] if isinstance(selected,int) else {}
                        expected_anchor=p.get('anchor_sources',{}).get(anchor_id)
                        identity={key:ac.get(key) for key in ('doc_hash','start','end')}
                        identity['indices']=ac.get('indices',list(range(ac.get('start',0),ac.get('end',0))))
                        if not expected_anchor or identity!=expected_anchor:
                            raise ValueError('新增正文依赖的来源范围已经变化，请重新分析。')
            pending=[p for p in ps if p['decision']=='pending'];applied=[];ranges=[];source_inserted=[];audit_records=[];audit_refs=[]
            baseline_paths=element_paths(pkg.root,[b.el for b in fresh.blocks])
            touched={i for p in ps if p['decision'] in ('accept','same') for i in range(p['start'],p['end'])}
            protected_refs=[(baseline_paths[b.index],b.el) for b in fresh.blocks if b.index not in touched]
            for p in ps:
                if p['decision'] not in ('accept','same'):continue
                rng=set(range(p['start'],p['end']))
                if any(rng & r for r in ranges):raise ValueError('检测到重叠替换，已停止导出：'+p['id'])
                ranges.append(rng)
            doc_count=max(1,len(result['documents']))
            doc_base=int(n/doc_count*85);doc_span=max(1,int(85/doc_count))
            progress(doc_base,'正在写入修订 '+doc.name)
            # Export by review transaction, not by tiny atomic proposal.  Adjacent
            # source-backed members that the web UI already treats as one transaction
            # are now written as ONE large delete-then-paste range in Word/WPS.
            txmap={}
            for p in ps:txmap.setdefault(p.get('transaction_id',p['id']),[]).append(p)
            ordered_groups=sorted((sorted(v,key=lambda x:x['start']) for v in txmap.values()),key=lambda g:g[0]['start'],reverse=True)

            def validate_candidate(p):
                if not p['candidates']:return
                c=p['candidates'][p.get('selected') or 0]
                if c.get('blocked_reason'):raise ValueError('不安全的来源替换已阻止：'+c['blocked_reason'])
                if c.get('source_warnings') and not p.get('note','').strip():raise ValueError('来源存在一致性提示，需记录核实说明后才能采用。')

            def record_one(p,plan,inserted,old,table_count=None):
                new=plan['new'];newtext='\n'.join(text(e) for e in new)
                applied.append(p['id'])
                c=p['candidates'][p.get('selected') or 0] if p.get('candidates') else None
                if plan['source'] is None:
                    if p['kind']!='metadata' or not p.get('parameter_changes'):raise ValueError('非来源动作缺少明确项目参数，不能写回。')
                    ar={'id':p['id'],'action':'parameter_edit','target_file':doc.name,'target_sha256':doc.hash,
                        'target_paths':[baseline_paths[i] for i in range(p['start'],p['end'])],
                        'parameter_changes':p['parameter_changes'],'expected_text':p['new_text'],'parameters':result['parameter_authorities']}
                else:
                    sf=fmap[plan['source_hash']];sd=self.load(sf)
                    ar={'id':p['id'],'action':p.get('action','source_copy'),'target_file':doc.name,'target_sha256':doc.hash,
                        'target_paths':[baseline_paths[i] for i in range(p['start'],p['end'])],
                        'source_file':sf['path'],'source_sha256':sd.hash,'source_role':'prospectus' if sd.profile['kind']=='prospectus' else 'authorized_formal_source','source_kind':sd.profile['kind'],'source_name':sd.name,
                        'source_paths':element_paths(sd.pkg.root,[sd.blocks[i].el for i in c.get('indices',range(c['start'],c['end']))]),
                        'source_span':c.get('span'),
                        'source_locator':c['locator'],'source_range':c.get('indices',list(range(c['start'],c['end']))),
                        'source_object_fingerprint':c.get('source_fingerprint'),'old_fingerprint':p.get('target_fingerprint'),
                        'anchor_sources':p.get('anchor_sources',{}),
                        'transforms':['fonts','technical_ids','table_layout'],
                        'wording_policy':plan.get('wording_policy'),'wording_changes':plan.get('wording_changes',[]),
                        'correspondence':c.get('correspondence') or c.get('figure_evidence'), 'merge_into':p.get('merge_into')}
                    if p.get('content_class')=='approved_missing_item':
                        # Lets the independent audit expect exactly one extra,
                        # target-created heading whose text is the approved one,
                        # bound to the persisted human decision.
                        ar['content_class']='approved_missing_item';ar['approved_title']=p.get('approved_title','')
                        ar['decision_fingerprints']=p.get('decision_fingerprints')
                if p.get('action')=='source_insert':
                    ar['insertion_boundary']={side:{'baseline_path':baseline_paths[p['insert_'+side]],'index':p['insert_'+side]} for side in ('before','after') if p.get('insert_'+side) is not None}
                audit_records.append(ar);audit_refs.append((ar,list(inserted),list(old)))
                logs.append({'target':doc.name,'id':p['id'],'matter':p['matter'],'locator':p['locator'],'before':p['old_text'],'after':newtext,'source':plan['origin'],'source_sha256':plan['source_hash'],'replacement_digest':plan['digest'],'preview_equals_export':True,'tables_copied':sum(e.tag==w('tbl') for e in new) if table_count is None else table_count,'tables_exactly_equal':0,'projection':p['candidates'][p.get('selected') or 0].get('projection') if p['candidates'] else None,'decision':'采用来源完整复制','transaction_id':p.get('transaction_id',p['id']),'transaction_mode':'large-range-delete-paste','author':author})

            for gi,members in enumerate(ordered_groups):
                if gi==0 or gi%max(1,len(ordered_groups)//8)==0:
                    progress(min(84,doc_base+int((gi/max(1,len(ordered_groups)))*doc_span*.72)),'正在写入修订 '+doc.name+'（'+str(gi+1)+'/'+str(len(ordered_groups))+'）')
                active=[p for p in members if p['decision'] in ('accept','same')]
                if not active:continue
                # Decisions are transaction-wide in the UI. If an old persisted project
                # somehow contains mixed decisions, fall back to independent ranges
                # rather than silently broadening a replacement.
                contiguous=all(members[i]['end']==members[i+1]['start'] for i in range(len(members)-1))
                material_group=(len(active)==len(members) and len(members)>1 and contiguous and all(p.get('kind')=='material' and p.get('action') not in ('source_merge','source_insert') for p in members))
                plans=[]
                if material_group:
                    for p in members:
                        validate_candidate(p);plans.append(build(self,p,files))
                    hashes={q['source_hash'] for q in plans};sources={q['source'].hash if q['source'] is not None else None for q in plans}
                    material_group=(len(hashes)==1 and None not in sources and len(sources)==1)
                if material_group:
                    old=[b.el for b in fresh.blocks[members[0]['start']:members[-1]['end']]]
                    new=[el for q in plans for el in q['new']];src=plans[0]['source'];txkey=members[0].get('transaction_id',members[0]['id'])
                    pkg.begin_transaction(txkey)
                    try:inserted=pkg.replace(old,new,author=author,source=src,whole=True)
                    finally:pkg.end_transaction()
                    visible=copy.deepcopy(inserted)
                    for el in visible:resolve_revisions(el,True)
                    if content_digest(visible)!=content_digest(new):raise ValueError('整块替换后的实际写入内容与来源不一致：'+txkey)
                    source_inserted.extend((e,plans[0].get('source_kind','unknown')) for e in inserted)
                    cursor=0
                    for p,plan in zip(members,plans):
                        count=len(plan['new']);record_one(p,plan,inserted[cursor:cursor+count],[b.el for b in fresh.blocks[p['start']:p['end']]])
                        cursor+=count
                    continue
                for p in members:
                    if p['decision'] not in ('accept','same'):continue
                    validate_candidate(p)
                    old=[b.el for b in fresh.blocks[p['start']:p['end']]];plan=build(self,p,files)
                    new,src=plan['new'],plan['source'];whole_source=True
                    txkey=p.get('transaction_id',p['id']);pkg.begin_transaction(txkey)
                    try:
                        if p.get('action')=='source_insert':
                            side='before' if p.get('insert_before') is not None else 'after'
                            inserted=pkg.insert_source(fresh.blocks[p['insert_'+side]].el,new,before=side=='before',author=author,source=src)
                        else:inserted=pkg.replace(old,new,author=author,source=src,whole=whole_source)
                    finally:pkg.end_transaction()
                    if src is not None:source_inserted.extend((e,plan.get('source_kind','unknown')) for e in inserted)
                    visible=copy.deepcopy(inserted)
                    for el in visible:resolve_revisions(el,True)
                    if content_digest(visible)!=plan['digest']:raise ValueError('实际写入内容与替换后预览不一致，已阻止导出：'+p['id']+' / '+p.get('heading',''))
                    record_one(p,plan,inserted,old)
            # Approved wording was frozen in build(), then independently checked
            # against the original source after save. Deleted originals stay intact.
            # Count the actual current-view source objects after archive wording has
            # been applied. This avoids treating a deliberately repeated source table
            # in two independent matters as an accidental duplicate.
            intended_source_blocks=collections.Counter()
            for el,_kind in source_inserted:
                tx=text(el).strip()
                if len(tx)>60:intended_source_blocks[tx]+=1
            pkg.track()
            font_changes=font_normalize(pkg,author)
            # Comments are review metadata, not a delivery channel for source tracing.
            # Preserve only comments that already existed in the uploaded draft and
            # hard-stop if replacement/import code introduces any new comment anchor,
            # comment part, or comment relationship.
            def _comment_markers(root):
                tags={w('commentRangeStart'),w('commentRangeEnd'),w('commentReference')}
                return [(e.tag,e.get(w('id'))) for e in root.iter() if e.tag in tags]
            if _comment_markers(pkg.root)!=_comment_markers(pristine):
                raise ValueError('检测到程序新增或改动批注锚点，已停止导出。来源说明不得写入 Word 批注。')
            base_comment_parts={k:v for k,v in base_parts.items() if re.search(r'(^|/)comments[^/]*\.xml$',k,re.I)}
            now_comment_parts={k:v for k,v in pkg.entries.items() if re.search(r'(^|/)comments[^/]*\.xml$',k,re.I)}
            if base_comment_parts!=now_comment_parts:
                raise ValueError('检测到程序新增或改动 Word 批注，已停止导出。仅保留上传稿原有人工批注。')
            def _comment_rels(entries):
                raw=entries.get('word/_rels/document.xml.rels')
                if not raw:return []
                rr=xml(raw)
                return sorted((x.get('Id'),x.get('Type'),x.get('Target'),x.get('TargetMode')) for x in rr if 'comment' in (x.get('Type') or '').lower())
            if _comment_rels(pkg.entries)!=_comment_rels(base_parts):
                raise ValueError('检测到程序新增 Word 批注关系，已停止导出。')
            # Reject THIS round's changes must restore the accepted input baseline.
            reject=copy.deepcopy(pkg.root);resolve_revisions(reject,False)
            accept=copy.deepcopy(pkg.root);resolve_revisions(accept,True)
            def sig(root):
                return [(e.tag,text(e).strip()) for e in root.find(w('body')) if e.tag in (w('p'),w('tbl')) and (text(e).strip() or e.findall('.//'+w('drawing')))]
            roundtrip=structural_digest_content(reject)==structural_digest_content(pristine)
            if not roundtrip:raise ValueError('修订拒绝还原校验未通过，已阻止导出该结果。')
            # Compare introduced duplicates rather than counting legitimate repeats already in the original.
            before=collections.Counter(t for _,t in sig(pristine) if len(t)>60)
            after=collections.Counter(t for _,t in sig(accept) if len(t)>60)
            # A source block may legitimately be reused in multiple independent
            # workpaper matters. Flag only instances beyond both the original
            # baseline count and the number of source copies we intentionally made.
            dupe=sum(max(0,v-max(1,before.get(t,0)+intended_source_blocks.get(t,0))) for t,v in after.items())
            residue=[]
            if record['cross_project']:
                oldissuer=doc.profile['issuer']
                for i,e in enumerate(accept.find(w('body'))):
                    tx=text(e)
                    if oldissuer and oldissuer in tx:residue.append({'locator':'接受稿正文块 '+str(i+1),'text':tx[:220]})
                for k,v in pkg.entries.items():
                    if re.fullmatch(r'word/(header\d+|footer\d+)\.xml',k) and oldissuer in text(xml(v)):residue.append({'locator':k,'text':'页眉或页脚仍含旧发行人全称。'})
            internal_name=safe_name(Path(doc.name).stem)
            unclassified=[x for x in record.get('object_inventory',[]) if x['status']=='pending']
            caution=bool(pending or unclassified or residue or record['cross_project'] or record['warning'] or result.get('source_issues') or any(x['severity']=='warning' for x in record.get('quality_issues',[])))
            filename=internal_name+('_待复核.docx' if caution else '_修订版.docx')
            dst=outdir/filename
            if dst.exists():dst=outdir/(internal_name+'_'+target['id'][:6]+('_待复核.docx' if caution else '_修订版.docx'))
            progress(min(94,doc_base+int(doc_span*.78)),'正在保存并检查 Word 结构 '+doc.name)
            pkg.save(dst);errors=validate_package(dst)
            def revision_ids(elements,kind):
                return [x.get(w('id')) for el in elements for x in el.iter(w(kind))]
            for ar,inserted,old in audit_refs:
                for boundary in ar.get('insertion_boundary',{}).values():
                    boundary['output_path']=element_paths(pkg.root,[fresh.blocks[boundary['index']].el])[0]
                ar.update(output_paths=element_paths(pkg.root,inserted),deletion_paths=element_paths(pkg.root,old),
                          insertion_revision_ids=revision_ids(inserted,'ins'),deletion_revision_ids=revision_ids(old,'del'))
            protected_ranges=[{'baseline_path':bp,'output_path':op} for (bp,el),op in zip(protected_refs,element_paths(pkg.root,[e for _,e in protected_refs]))]
            progress(min(96,doc_base+int(doc_span*.88)),'正在独立核对实际插入、来源对象与保护区 '+doc.name)
            independent=audit_docx(dst,target['path'],audit_records,protected_ranges=protected_ranges,missing_decisions=result.get('missing_decisions'))
            (outdir/(internal_name+'_输出审计.json')).write_text(json.dumps(independent,ensure_ascii=False,indent=2),encoding='utf-8')
            (outdir/(internal_name+'_来源追溯.json')).write_text(json.dumps({'records':audit_records,'protected_ranges':protected_ranges,'object_inventory':record.get('object_inventory',[]),'engine_version':BUILD,'code_sha256':result['code_sha256']},ensure_ascii=False,indent=2),encoding='utf-8')
            if not independent['passed']:raise ValueError('独立输出审计未通过：'+'；'.join(str(x) for x in independent['errors'][:5]))
            exported_doc=Document(dst,doc.name)
            output_issues=quality_inspect(exported_doc,current_sources)+target_checks(exported_doc,current_sources,current_facts)+exported_doc.safety_issues+record.get('source_gaps',[])
            if errors:raise ValueError('Word 结构校验未通过：'+'；'.join(errors[:5]))
            # All untouched package parts (including all existing media) remain byte-identical.
            allowed={'word/document.xml','word/settings.xml','word/_rels/document.xml.rels','[Content_Types].xml','word/styles.xml','word/numbering.xml','word/footnotes.xml','word/endnotes.xml'}
            unchanged_parts=all(pkg.entries.get(k)==v for k,v in base_parts.items() if k not in allowed)
            if not unchanged_parts:raise ValueError('检测到无关部件变化，已停止导出。')
            check={'file':dst.name,'applied':sum(p['decision']=='accept' for p in ps),'source_units_synchronized':len(applied),'tables_copied':sum(v.get('tables_copied',0) for v in logs if v['target']==doc.name),'tables_exactly_equal':sum(v.get('tables_exactly_equal',0) for v in logs if v['target']==doc.name),'font_runs_standardized':font_changes,'preview_equals_export':True,'roundtrip_full_structure':True,'generated_comments':0,'existing_comments_preserved':pkg.entries.get('word/comments.xml')==base_parts.get('word/comments.xml'),'pending':len(pending),'kept_by_user':sum(p['decision']=='keep' and p.get('content_class')!='no_correspondence_retained' for p in ps),'source_matched_unchanged':sum(p['status']=='unchanged' for p in ps),'roundtrip':roundtrip,'unrelated_parts_preserved':unchanged_parts,'structure_errors':errors,'introduced_duplicate_instances':dupe,'old_issuer_residue':residue,'cross_project':record['cross_project'],'images_preserved_unverified':doc.profile['images'],'historical_revisions_normalized':doc.profile['historical_revisions'],'warning':record['warning'],'quality_issues':output_issues,'pre_update_quality_issues':record.get('quality_issues',[]),'preserved_scope':record.get('protected',{}),'caution':caution}
            qa.append(check);artifacts.append(dst)
            check['independent_output_audit']=independent
            check['code_sha256']=result['code_sha256'];check['engine_version']=BUILD
            check['object_dispositions']=record.get('coverage',{}).get('dispositions',{})
            check['figures_source_copied']=sum(bool(p.get('figure')) and p['decision'] in ('accept','same') for p in ps)
            check['source_copy_count']=independent.get('source_copy_count',0)
            check['source_insert_count']=independent.get('source_insert_count',0)
            check['wording_adapted_paragraphs']=sum(len(r.get('wording_changes',[])) for r in audit_records)
            check['retained_without_correspondence']=sum(x.get('content_class')=='no_correspondence_retained' for x in record.get('object_inventory',[]))
            check['source_merge_count']=sum(p.get('action')=='source_merge' and p['decision'] in ('accept','same') for p in ps)
            check['compatibility_fixes']=pkg.compatibility_fixes
            check['layout_validation']='not_verified';check['word_wps_validation']='not_verified'
            # Keep original bytes and a separate current-view baseline for transparent recovery.
            original=outdir/'原始文件'/safe_name(doc.name);original.parent.mkdir(exist_ok=True)
            if original.exists() and sha(original.read_bytes())!=doc.hash:original=original.with_name(target['id'][:6]+'_'+original.name)
            original.write_bytes(doc.pkg.raw);artifacts.append(original)
            if doc.profile['historical_revisions']:
                bp=Package(target['path']).baseline();baseline=outdir/'更新基线'/(internal_name+'_当前可见稿.docx');baseline.parent.mkdir(exist_ok=True)
                if baseline.exists():baseline=baseline.with_name(target['id'][:6]+'_'+baseline.name)
                bp.save(baseline);artifacts.append(baseline)
        open_items=[{k:p.get(k) for k in ('target_name','id','matter','locator','status','decision','content_class','reason','note','old_text')} for p in result['proposals'] if p['decision'] in ('pending','protected','keep')]
        payload={'open_items':open_items,'summary':summarize(result),'checks':qa,'changes':logs,'sources':result['sources'],'teacher_files_not_used_as_sources':result['references'],'source_issues':result.get('source_issues',[]),'boundary':'自动结果是有来源的修订稿，不是独立尽职调查或法律意见。截图内容与核查判断没有被自动重新验证。'}
        payload.update(engine_version=BUILD,code_sha256=result['code_sha256'],parameters=result['frozen_params'],delivery_state='review_draft')
        (outdir/'未完成与保护清单.json').write_text(json.dumps([{'file':d['name'],'sha256':d['sha256'],'coverage':d['coverage'],'objects':[x for x in d['object_inventory'] if x['status']!='copied']} for d in result['documents']],ensure_ascii=False,indent=2),encoding='utf-8')
        (outdir/'修改与来源记录.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
        (outdir/'核查报告.html').write_text(report_html(payload),encoding='utf-8')
        (outdir/'阅读说明.txt').write_text('Word 为真实修订稿，修订作者为 '+author+'。\n募集或核查意见有明确对应（冲突以募集为准）：开启修订，删除底稿旧块，从当前正式来源实际复制粘贴来源块，并按核查分析文件口吻调整称谓；文字、表格、图片完全相同也执行实际复制，不给原稿套修订样式。\n两份正式来源均无明确对应：底稿原样保留，不改写、不新增修订，不需逐项决定。此类保留不计作待核对。\n原始文件文件夹保留上传件，字节未变。原稿存在历史修订时，更新基线是其当前可见稿；拒绝本轮修订恢复该基线，不恢复历史审阅状态。\n带“待复核”的文件仍存在来源冲突、对象迁移、版式等实际待验证项。\n核查报告.html 与修改与来源记录.json 记录来源复制、保留项和校验结果。\n来源使用标题与关键句定位，不把目录页码当成 WPS 物理页码。',encoding='utf-8')
        zip_path=outdir.parent/(outdir.name+'.zip')
        with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
            for p in outdir.rglob('*'):
                if p.is_file():z.write(p,p.relative_to(outdir))
        progress(100,'审阅草稿与独立来源审计已生成；待核对事项及Office实机验证仍须完成');return {'zip':str(zip_path),'checks':qa,'delivery_state':'review_draft','engine_version':BUILD,'code_sha256':result['code_sha256'],'files':[p.name for p in artifacts if p.parent==outdir]}

def safe_name(s):return re.sub(r'[\\/:*?"<>|\x00-\x1f]','_',s)[:180] or '未命名'
def summarize(result):
    ps=result.get('proposals',[])
    # Refresh placeholder warnings against the selected replacement, including
    # undo/redo; resolved old placeholders aren't presented as live failures.
    for d in result.get('documents',[]):
        all_issues=d.setdefault('quality_issues_all',copy.deepcopy(d.get('quality_issues',[])))
        active=[];resolved=[]
        for issue in all_issues:
            covered=next((p for p in ps if p['target_id']==d['id'] and p['start']<=(issue.get('block') if isinstance(issue.get('block'),int) else -1)<p['end'] and p['decision'] in ('accept','same')),None)
            selected=(covered.get('selected') or 0) if covered else 0
            value=(covered.get('new_text') or ((covered.get('candidates') or [{}])[selected].get('planned_text'))) if covered else None
            if issue['rule']=='visible-placeholder' and value is not None and not re.search(r'待更新|待补充|需要VIP|需要\s*VIP|未找到|靠你了老师|【\s*】|\[\s*待填\s*\]',value):resolved.append(issue)
            else:active.append(issue)
        d['quality_issues']=active;d['resolved_quality_issues']=resolved
    tx={}
    for p in ps:tx.setdefault(p.get('transaction_id',p['id']),[]).append(p)
    return {'documents':len(result.get('documents',[])),'matters':sum(len(d['matters']) for d in result.get('documents',[])),'proposals':len(ps),'transactions':len(tx),'auto':sum(p['decision']=='accept' for p in ps),'source_sync':sum(p['decision'] in ('accept','same') for p in ps),'tables_sync':sum((p.get('candidates') or [{}])[p.get('selected') or 0].get('table_count',int(p.get('has_table',False))) for p in ps if p['decision'] in ('accept','same')),'table_total':sum(p.get('table_count',int(p.get('has_table',False))) for p in ps),'unchanged':sum(p['status']=='unchanged' for p in ps),'pending':sum(p['decision']=='pending' for p in ps),'pending_transactions':sum(any(x['decision']=='pending' for x in xs) for xs in tx.values()),'kept':sum(p['decision'] in ('keep','protected') for p in ps),'kept_transactions':sum(all(x['decision'] in ('keep','protected') for x in xs) for xs in tx.values()),'cross_project':sum(d.get('cross_project',False) for d in result.get('documents',[])),'source_warnings':len(result.get('source_issues',[])),'quality_warnings':sum(i.get('severity')=='warning' for d in result.get('documents',[]) for i in d.get('quality_issues',[]))}

def compare_documents(draft,teacher):
    """Evaluation only. The returned differences never feed the matching engine."""
    a=Document(draft['path'],draft['name']);b=Document(teacher['path'],teacher['name'])
    if a.profile['issuer']!=b.profile['issuer']:raise ValueError('老师稿与对照稿不是同一发行人，不能作同项目质量评分。')
    if a.profile['chapter']!=b.profile['chapter']:raise ValueError('老师稿与对照稿章节不同，不能直接比较。')
    aa=[x.text for x in a.blocks if x.text];bb=[x.text for x in b.blocks if x.text]
    sm=difflib.SequenceMatcher(None,[normal(x) for x in aa],[normal(x) for x in bb],autojunk=False)
    changes=[]
    for op,i,j,k,l in sm.get_opcodes():
        if op!='equal':changes.append({'type':op,'before':'\n'.join(aa[i:j]),'after':'\n'.join(bb[k:l]),'before_block':i+1,'after_block':k+1})
    return {'draft':a.name,'teacher':b.name,'issuer':b.profile['issuer'],'chapter':b.profile['chapter'],'equal_block_ratio':round(sm.ratio(),4),'changes':changes,'note':'这是当前可见内容的差异比较；相似度不是事实准确率或自动更新验收率。老师稿只用于对照，不进入来源检索。'}

def report_html(payload):
    esc=html.escape
    cards=''.join(f'<section><h2>{esc(q["file"])}</h2><p>本轮采用 {q["applied"]} 处；来源完整复制 {q["source_units_synchronized"]} 处；待核对 {q["pending"]} 处。</p><p>拒绝修订恢复更新基线：{q["roundtrip"]}；无关部件保留：{q["unrelated_parts_preserved"]}；内部关系与表格结构错误：{len(q["structure_errors"])}。独立输出审计：{q.get("independent_output_audit",{}).get("passed",False)}。Word/WPS实机与最终版式尚未验证。</p><p>从来源复制图组 {q.get("figures_source_copied",0)} 处；未对应来源的对象按规则保留 {q.get("retained_without_correspondence",0)} 个，不计作待核对；其余图示见明细。新增重复实例诊断 {q["introduced_duplicate_instances"]} 处，需结合跨事项复用判断，不直接判错。</p><p>{esc(q["warning"])}</p><p>{esc("存在未完成事项，不能作为直接提交的定稿。" if q["caution"] else "已生成来源支持的修订稿；仍需履行最终业务审阅。")}</p></section>' for q in payload['checks'])
    open_html=''.join('<details><summary>'+esc(p['target_name']+' · '+p['matter']+' · '+('待核对' if p['decision']=='pending' else '保留原文'))+'</summary><p>'+esc(p['locator'])+'</p><p>'+esc(p['reason'])+'</p><p>'+esc(p.get('note') or '')+'</p><pre>'+esc(p['old_text'])+'</pre></details>' for p in payload.get('open_items',[]))
    quality_html=''.join('<section><h3>'+esc(q['file']+' · '+i['title'])+'</h3><p>'+esc(i['locator'])+'</p><p>'+esc(i['detail'])+'</p></section>' for q in payload['checks'] for i in q.get('quality_issues',[]))
    residue_html=''.join('<section><h3>'+esc(q['file']+' · 旧发行人残留')+'</h3><p>'+esc(i['locator'])+'</p><p>'+esc(i['text'])+'</p></section>' for q in payload['checks'] for i in q.get('old_issuer_residue',[]))
    source_html=''.join('<section><h3>'+esc(i['title'])+'</h3><p>'+esc(i['file']+' · '+i['locator'])+'</p><p>'+esc(i['detail'])+'</p></section>' for i in payload.get('source_issues',[]))
    changes=''.join('<details><summary>'+esc(c['matter']+' · '+c['locator'][:100])+'</summary><p class="source">'+esc(c['source'])+'</p><div class="compare"><pre>'+esc(c['before'])+'</pre><pre>'+esc(c['after'])+'</pre></div></details>' for c in payload['changes'])
    return '<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>核查更新报告</title><style>body{background:#f7f5ef;color:#302d29;font:15px/1.8 system-ui;max-width:1120px;margin:56px auto;padding:0 24px}h1,h2{font-family:Georgia,"Songti SC",serif}section,details{background:white;border:1px solid #e1ddd3;border-radius:12px;padding:22px;margin:18px 0}summary{cursor:pointer}.compare{display:grid;grid-template-columns:1fr 1fr;gap:18px}pre{font:14px/1.8 system-ui;white-space:pre-wrap;word-break:break-word;background:#faf9f6;padding:20px;max-height:600px;overflow:auto}.source{color:#8c5b42;word-break:break-all}@media(max-width:700px){.compare{grid-template-columns:1fr}}</style><h1>核查更新报告</h1><p>'+esc(payload['boundary'])+'</p>'+cards+'<h2>独立来源自检</h2>'+source_html+'<h2>结构与口径提示</h2>'+quality_html+residue_html+'<h2>逐处修改与来源</h2>'+changes+'<h2>未完成与人工保留</h2>'+open_html+'</html>'
