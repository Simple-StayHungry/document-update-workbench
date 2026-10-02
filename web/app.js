const token=new URLSearchParams(location.search).get('token')||'';
let result=null,decisions={},activeManualId='';
const selected=new Map();
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=s=>(s??'').toString().replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
async function api(path,opt={}){opt.headers={...(opt.headers||{}),'X-Workbench-Token':token};const r=await fetch(path,opt);const j=await r.json().catch(()=>({}));if(!r.ok)throw new Error(j.error||'请求失败');return j}
function setLoading(on,title='正在分析',text='正在整理材料…'){const el=$('#loading');$('#loadingTitle').textContent=title;$('#loadingText').textContent=text;el.classList.toggle('hidden',!on)}
function setStage(n,doneThrough=0){$$('.step').forEach((s,i)=>{s.classList.toggle('active',i===n-1);s.classList.toggle('done-step',i<doneThrough)})}
function keyFor(f){return [f.webkitRelativePath||f.name,f.size,f.lastModified].join('|')}
function addFiles(files){for(const f of files){if(!/\.(docx|zip)$/i.test(f.name))continue;selected.set(keyFor(f),f)}renderSelection()}
function renderSelection(){const files=[...selected.values()];$('#selection').classList.toggle('hidden',!files.length);$('#selectionSummary').textContent=files.length?`${files.length} 项材料`:'';$('#selectedFiles').innerHTML=files.map(f=>{const k=keyFor(f),label=/\.zip$/i.test(f.name)?'ZIP':/募集说明书/.test(f.name)?'募集说明书':'DOCX';return `<span class="file-chip"><em>${label}</em><b title="${esc(f.webkitRelativePath||f.name)}">${esc(f.webkitRelativePath||f.name)}</b><button data-remove="${esc(k)}" aria-label="移除">×</button></span>`}).join('');$$('[data-remove]').forEach(b=>b.onclick=e=>{e.stopPropagation();selected.delete(b.dataset.remove);renderSelection()})}
$('#pickFiles').onclick=e=>{e.stopPropagation();$('#fileInput').click()};
$('#pickFolder').onclick=e=>{e.stopPropagation();$('#folderInput').click()};
$('#fileInput').onchange=e=>{addFiles(e.target.files);e.target.value=''};
$('#folderInput').onchange=e=>{addFiles(e.target.files);e.target.value=''};
$('#clearFiles').onclick=()=>{selected.clear();renderSelection()};
const zone=$('#dropzone');
zone.onclick=e=>{if(!e.target.closest('button,input'))$('#fileInput').click()};
zone.onkeydown=e=>{if(['Enter',' '].includes(e.key)){e.preventDefault();$('#fileInput').click()}};
['dragenter','dragover'].forEach(n=>zone.addEventListener(n,e=>{e.preventDefault();zone.classList.add('drag')}));
['dragleave','drop'].forEach(n=>zone.addEventListener(n,e=>{e.preventDefault();zone.classList.remove('drag');if(n==='drop')addFiles(e.dataTransfer.files)}));

$('#analyzeBtn').onclick=async()=>{
 const files=[...selected.values()];if(!files.length)return;
 $('#status').textContent='';setLoading(true);
 const fd=new FormData();for(const f of files)fd.append('files',f,f.webkitRelativePath||f.name);
 try{
  const j=await api('/api/analyze',{method:'POST',body:fd});result=j.result;decisions={};activeManualId='';
  $('#uploadPanel').classList.add('hidden');$('#reviewPanel').classList.remove('hidden');
  renderReviewState();
 }catch(err){$('#status').textContent=err.message}
 finally{setLoading(false)}
};

function manualItems(){return (result?.proposals||[]).filter(p=>p.status==='review'&&p.review_required!==false)}
function pendingManual(){return manualItems().filter(p=>!decisions[p.id])}
function chosenSource(p){return decisions[p.id]?.source_index??null}
function sourceByIndex(p,index){return (p.candidates||[]).find(c=>c.source_index===index)||null}
function targetKind(p){return p.figure?'图形':p.kind==='table'?'整表':'整段'}
function locationText(p){const t=(p.target_path_text||'').trim();return t||'正文'}
function safeReason(p){return p.reason||'存在多个可用来源，程序无法安全替你选择。'}
function renderReviewState(){
 const pending=pendingManual();
 if(!manualItems().length){renderNoManual();return}
 if(!pending.length){renderManualDone();return}
 renderManualReview(pending)
}
function summaryNumbers(){const s=result.summary||{},auto=(s.source_updates||0)+(s.wording_updates||0);return {files:s.documents||0,auto,retained:s.retained||0,review:s.review||0}}
function renderNoManual(){
 $('#manualReviewView').classList.add('hidden');$('#reviewDoneView').classList.remove('hidden');
 const n=summaryNumbers(),i=result.intake||{};
 $('#doneSummaryLine').textContent=`${i.targets||n.files} 份说明性文件已分析完成，可直接导出。`;
 $('#doneSummary').innerHTML=`<span><strong>${n.auto}</strong> 自动处理</span><span><strong>${n.retained}</strong> 自动保留</span>`;
 setStage(3,2)
}
function renderManualDone(){
 $('#manualReviewView').classList.add('hidden');$('#reviewDoneView').classList.remove('hidden');
 const n=summaryNumbers();$('#doneSummaryLine').textContent='所有需要人工确认的事项均已处理，可以导出。';
 $('#doneSummary').innerHTML=`<span><strong>${n.auto}</strong> 自动处理</span><span><strong>${manualItems().length}</strong> 已确认</span><span><strong>${n.retained}</strong> 自动保留</span>`;
 setStage(3,2)
}
function renderManualReview(pending){
 $('#reviewDoneView').classList.add('hidden');$('#manualReviewView').classList.remove('hidden');
 if(!activeManualId||!pending.some(p=>p.id===activeManualId))activeManualId=pending[0].id;
 $('#manualTitle').textContent=`需要确认 ${pending.length} 项`;
 $('#manualSubtitle').textContent='程序只保留无法安全自动决定的事项。';
 $('#manualCounter').textContent=`剩余 ${pending.length} 项`;
 $('#manualExportBtn').disabled=true;$('#manualExportBtn').title='处理完待确认事项后即可导出';
 $('#manualList').innerHTML=pending.map(p=>`<button class="manual-item ${p.id===activeManualId?'active':''}" data-manual-id="${esc(p.id)}"><b>${esc(locationText(p))}</b><span>${esc(p.target_name)}</span><em>${targetKind(p)}</em></button>`).join('');
 $$('[data-manual-id]').forEach(b=>b.onclick=()=>{activeManualId=b.dataset.manualId;renderManualReview(pendingManual())});
 const p=pending.find(x=>x.id===activeManualId)||pending[0];renderManualDetail(p);setStage(2,1)
}
function renderManualDetail(p){
 $('#manualLocation').textContent=`修改章节 · ${locationText(p)}`;
 const chosen=chosenSource(p),selectedSource=sourceByIndex(p,chosen),previewSource=selectedSource||(p.candidates||[])[0]||null;
 const sourceTitle=chosen==null?'候选来源 · 尚未选择':'选定来源';
 const candidates=(p.candidates||[]).length?`<div class="source-choices"><div class="source-choices-title">选择来源</div>${p.candidates.map((c,i)=>`<label class="source-choice ${chosen===c.source_index?'selected':''}"><input type="radio" name="src_${p.id}" data-source-choice data-id="${p.id}" value="${c.source_index}" ${chosen===c.source_index?'checked':''}><span><b>候选 ${i+1}</b><small>${Math.round((c.score||0)*100)}% · ${esc(c.source_path||'募集说明书')}</small></span></label>`).join('')}</div>`:'<div class="no-source">未找到可供人工选择的安全来源；建议保持原文。</div>';
 $('#manualDetailBody').innerHTML=`
  <div class="manual-fileline"><span>${targetKind(p)}</span><b>${esc(p.target_name)}</b></div>
  <p class="manual-reason">${esc(safeReason(p))}</p>
  <div class="decision-compare">
   <section class="compare-pane"><div class="compare-title">原稿</div><div class="compare-content">${esc(p.old_text||'[图形对象]')}</div></section>
   <section class="compare-pane after"><div class="compare-title">${sourceTitle}</div><div class="compare-content">${esc(previewSource?.source_text||'请选择一个来源后再应用修改。')}</div></section>
  </div>
  ${candidates}
  <div class="manual-actionbar"><div class="action-state">${chosen==null?'先选择来源，或保持原文':'已选择来源，可应用修改'}</div><div class="action-buttons"><button class="secondary" data-decision="keep" data-id="${p.id}">保持原文</button><button class="primary" data-decision="update" data-id="${p.id}" ${chosen==null?'disabled':''}>应用修改</button></div></div>`;
 $$('input[data-source-choice]').forEach(r=>r.onchange=()=>{decisions[p.id]={...(decisions[p.id]||{}),source_index:Number(r.value)};renderManualDetail(p)});
 $$('[data-decision]').forEach(b=>b.onclick=()=>setDecision(p,b.dataset.decision));
}
async function setDecision(p,decision){
 const chosen=chosenSource(p);if(decision==='update'&&chosen==null)return;
 const body={id:p.id,decision};if(chosen!=null)body.source_index=chosen;
 setLoading(true,'正在保存','正在记录本次决定…');
 try{await api('/api/decision',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});decisions[p.id]=body;activeManualId='';renderReviewState()}
 catch(e){alert(e.message)}finally{setLoading(false)}
}

async function exportResult(button){
 button.disabled=true;setLoading(true,'正在生成','正在写入修订…');
 try{
  const j=await api('/api/export',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  $('#manualReviewView').classList.add('hidden');$('#reviewDoneView').classList.remove('hidden');
  $('.done-kicker').textContent='导出完成';$('.done-card h1').textContent='修订包已生成';
  $('#doneSummaryLine').textContent='文件已准备好，可以直接下载。';
  const n=summaryNumbers();$('#doneSummary').innerHTML=`<span><strong>${n.files}</strong> 份文件</span><span><strong>${n.auto}</strong> 自动处理</span>`;
  $('#doneExportBtn').classList.add('hidden');$('#doneDownloadBtn').classList.remove('hidden');$('#doneDownloadBtn').href=j.download;
  setStage(3,3)
 }
 catch(e){alert(e.message)}finally{setLoading(false);button.disabled=false}
}
$('#doneExportBtn').onclick=()=>exportResult($('#doneExportBtn'));
$('#manualExportBtn').onclick=()=>exportResult($('#manualExportBtn'));
