"""OOXML reader and reversible block patches using only Python's standard library.

Original uploads are never overwritten. A working baseline is the accepted/current
view of an input with historical revisions. No Office rendering, OCR, network, or
third-party XML package is required by the production document engine.
"""
from __future__ import annotations
import copy, difflib, hashlib, io, posixpath, re, zipfile, threading, weakref
from pathlib import Path
from datetime import datetime, timezone, timedelta
from xml.etree import ElementTree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
W14='http://schemas.microsoft.com/office/word/2010/wordml'
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
NS={'w':W,'r':R,'a':A,'pr':PR,'m':M}
for prefix,uri in [('w',W),('r',R),('a',A),('wp',WP),('w14',W14),('pr',PR),('ct',CT),('m',M)]:
    try:E.register_namespace(prefix,uri)
    except ValueError:pass

def w(t): return '{'+W+'}'+t
def m(t): return '{'+M+'}'+t

_NS_LOCK=threading.RLock()
_ROOT_NS=weakref.WeakKeyDictionary()
_KNOWN_NS={**NS,'wp':WP,'w14':W14,'ct':CT,'xml':'http://www.w3.org/XML/1998/namespace'}
MC='http://schemas.openxmlformats.org/markup-compatibility/2006'

def _register_namespaces(raw:bytes):
    found={}
    for _event,(prefix,uri) in E.iterparse(io.BytesIO(raw),events=('start-ns',)):
        prefix=prefix or '';found[prefix]=uri;_KNOWN_NS[prefix]=uri
        try:E.register_namespace(prefix,uri)
        except ValueError:pass
    return found

def xml(b):
    raw=b.encode() if isinstance(b,str) else b
    with _NS_LOCK:
        ns=_register_namespaces(raw)
        parser=E.XMLParser(target=E.TreeBuilder(insert_comments=True,insert_pis=True))
        root=E.fromstring(raw,parser=parser);_ROOT_NS[root]=ns
        return root

def dump(e):
    # ElementTree otherwise discards declarations used ONLY in mc:Ignorable or
    # mc:Choice/@Requires values. Those are significant to Office, not decoration.
    from xml.sax.saxutils import quoteattr
    with _NS_LOCK:
        ns=dict(_KNOWN_NS);ns.update(_ROOT_NS.get(e,{}))
        for prefix,uri in ns.items():
            try:E.register_namespace(prefix,uri)
            except ValueError:pass
        needed=set()
        for node in e.iter():
            for key,value in node.attrib.items():
                if key in ('{'+MC+'}Ignorable','{'+MC+'}PreserveElements','{'+MC+'}PreserveAttributes','{'+MC+'}ProcessContent') or (node.tag=='{'+MC+'}Choice' and key=='Requires'):
                    needed.update(t.split(':',1)[0] for t in value.split())
        raw=E.tostring(e,encoding='UTF-8',xml_declaration=True,short_empty_elements=True)
        begin=raw.find(b'?>')+2;begin=raw.find(b'<',begin);end=raw.find(b'>',begin)
        tag=raw[begin:end];declared={x.decode() for x in re.findall(rb'xmlns:([\w.-]+)\s*=',tag)}
        additions=[]
        for prefix in sorted(needed-declared):
            uri=ns.get(prefix)
            if not uri:raise ValueError('文档命名空间依赖无法解析，已停止写入：'+prefix)
            if not re.fullmatch(r'[A-Za-z_][\w.-]*',prefix):raise ValueError('无效命名空间前缀。')
            additions.append((' xmlns:'+prefix+'='+quoteattr(uri)).encode('utf-8'))
        if additions:
            insert=end-1 if raw[end-1:end]==b'/' else end
            raw=raw[:insert]+b''.join(additions)+raw[insert:]
        if raw.startswith(b'<?xml') and b'standalone=' not in raw.split(b'?>',1)[0]:raw=raw.replace(b'?>',b' standalone="yes"?>',1)
        return raw

def sha(b): return hashlib.sha256(b).hexdigest()
def local(e):
    tag=e.tag
    if not isinstance(tag,str):return ''
    return tag.rsplit('}',1)[-1] if '}' in tag else tag

def _parents(root):return {c:p for p in root.iter() for c in p}
def _ancestor_tags(node,parents):
    cur=parents.get(node)
    while cur is not None:
        yield cur.tag
        cur=parents.get(cur)
def _remove_with(parents,e):
    parent=parents.get(e)
    if parent is not None:
        try:parent.remove(e)
        except ValueError:pass

def text(el):
    """Current visible text, excluding property-only OOXML such as tab stops.

    A `<w:tab>` under paragraph properties means a tab-stop definition, not a
    visible tab character.  Only run-level tabs/breaks are rendered as content.
    """
    out=[]
    def walk(node,hidden=False,deleted_row=False,in_run=False):
        tag=node.tag
        hidden=hidden or tag in (w('del'),w('moveFrom'))
        if tag==w('tr') and node.find('./'+w('trPr')+'/'+w('del')) is not None:deleted_row=True
        now_in_run=in_run or tag==w('r')
        if not hidden and not deleted_row:
            if tag in (w('t'),m('t')):out.append(node.text or '')
            elif tag==w('tab') and in_run:out.append('\t')
            elif tag==w('br') and in_run:out.append('\n')
        for child in list(node):walk(child,hidden,deleted_row,now_in_run)
    walk(el)
    return ''.join(out)

def native_math_structure_reason(elements):
    """Allow complete self-contained OMML objects, never a sliced math subtree.

    ISO 29500 w:ins/w:del (RunTrackChangeType) allow m:oMath and m:oMathPara
    children. Wrapping that complete object preserves its native editability;
    wrapping only its w:r descendants silently leaves the actual m:r untracked.
    See Microsoft OpenXML SDK InsertedRun/DeletedRun child-element definitions.
    """
    for el in elements:
        parents=_parents(el)
        maths=[e for e in el.iter() if str(e.tag).startswith('{'+M+'}')]
        if not maths:continue
        roots=[e for e in maths if not any(str(t).startswith('{'+M+'}') for t in _ancestor_tags(e,parents))]
        for root in roots:
            parent=parents.get(root)
            if root.tag not in (m('oMath'),m('oMathPara')) or parent is None or parent.tag!=w('p'):
                return '原生公式不是段落中的完整独立对象，保留该范围待核对。'
            # Cross-range annotations and external content require an independent
            # import protocol. Do not strip them from a supposedly exact math copy.
            forbidden={'bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd','commentReference',
                       'permStart','permEnd','footnoteReference','endnoteReference','fldSimple','fldChar',
                       'instrText','drawing','pict','object','altChunk','sdt','customXml','hyperlink',
                       'ins','del','moveFrom','moveTo','proofErr'}
            if any(local(e) in forbidden or any(k.startswith('{'+R+'}') for k in e.attrib) for e in root.iter()):
                return '原生公式包含批注、修订、外部引用或复合对象，保留完整范围待核对。'
    return ''

def native_math_copy_reason(target,source,elements):
    """Check package-level math dependencies before any target mutation.

    m:mathPr is document-wide; replacing it would change untouched formulas.
    Therefore differing math defaults or run defaults remain unsupported. Source
    paragraph/character styles are imported normally, including implicit Normal.
    A theme-dependent source is admitted only with the identical target theme.
    """
    elements=list(elements)
    if not any(str(e.tag).startswith('{'+M+'}') for el in elements for e in el.iter()):return ''
    reason=native_math_structure_reason(elements)
    if reason:return reason
    if source is None:return '原生公式缺少可核验的来源文档，不能复制。'
    def setting(pkg):
        root=xml(pkg.entries.get('word/settings.xml',f'<w:settings xmlns:w="{W}"/>'.encode()))
        return root.find(m('mathPr'))
    def signature(el):return structural_digest(el) if el is not None else None
    if signature(setting(target))!=signature(setting(source)):
        return '来源与目标的原生公式全局设置不同，不能改动整份底稿的公式默认值；保留待核对。'
    def styles(pkg):return xml(pkg.entries.get('word/styles.xml',f'<w:styles xmlns:w="{W}"/>'.encode()))
    ts,ss=styles(target),styles(source)
    if signature(ts.find(w('docDefaults')))!=signature(ss.find(w('docDefaults'))):
        return '来源与目标的默认字体或段落设置不同，原生公式的继承格式尚不能完整移植；保留待核对。'
    by_id={s.get(w('styleId')):s for s in ss.findall(w('style'))}
    selected=[e for el in elements for e in el.iter()]
    defaults=[s for s in ss.findall(w('style')) if s.get(w('default')) in ('1','true','on') and s.get(w('type')) in ('paragraph','character')]
    style_ids=[s.get(w('styleId')) for s in defaults]
    style_ids += [e.get(w('val')) for e in selected if e.tag in (w('pStyle'),w('rStyle'),w('tblStyle'))]
    seen=set()
    while style_ids:
        sid=style_ids.pop()
        if sid in seen:continue
        seen.add(sid);style=by_id.get(sid)
        if style is None:return '原生公式引用的来源样式缺失，保留待核对。'
        selected.extend(style.iter())
        for tag in ('basedOn','link'):
            ref=style.find(w(tag))
            if ref is not None:style_ids.append(ref.get(w('val')))
    dd=ss.find(w('docDefaults'))
    if dd is not None:selected.extend(dd.iter())
    if any('theme' in key.lower() for node in selected for key in node.attrib):
        def themes(pkg):
            out=[]
            for rel in pkg.rels():
                if not (rel.get('Type') or '').endswith('/theme'):continue
                if rel.get('TargetMode')=='External':return None
                part=rel.get('Target','');part=part.lstrip('/') if part.startswith('/') else posixpath.normpath(posixpath.join('word',part))
                if part not in pkg.entries:return None
                out.append(sha(pkg.entries[part]))
            return sorted(out) or None
        if themes(source) is None or themes(source)!=themes(target):
            return '原生公式依赖的来源主题与目标不同，尚不能完整移植；保留待核对。'
    return ''

def complex_reference_reason(elements):
    """Refuse cross-document note/field imports whose target IDs are not remapped."""
    elements=list(elements)
    tags={w('altChunk'),w('subDoc')}
    reason=native_math_structure_reason(elements)
    if reason:return reason
    for el in elements:
        if el.find('.//'+w('fldSimple')) is not None or el.tag==w('fldSimple'):
            return '含简单域对象，接受或拒绝修订时可能被阅读器重新生成；本范围保留，需在 Word 中处理。'
        if any(e.tag in tags for e in el.iter()):
            return '含脚注、尾注或外部文档块，当前版本不自动移植其引用关系。'
        fields=[e.text or '' for e in el.iter(w('instrText'))]
        fields += [e.get(w('instr'),'') for e in el.iter(w('fldSimple'))]
        if any(re.search(r'\b(?:REF|PAGEREF|NOTEREF|INCLUDETEXT|INCLUDEPICTURE|DDEAUTO|DDE)\b',v,re.I) for v in fields):
            return '含跨文档引用域或外部内容域，需要在 Word 中人工处理。'
    return ''

def equivalent_table(a,b):
    """Exact logical values + grid merges, excluding tables with rich objects."""
    forbidden={'drawing','pict','object','fldChar','fldSimple','sdt','hyperlink','sym','altChunk','hMerge'}
    def sig(t):
        if any(local(e) in forbidden or str(e.tag).startswith('{http://schemas.openxmlformats.org/officeDocument/2006/math}') for e in t.iter()):return None
        rows=[]
        for tr in t.findall(w('tr')):
            cells=[]
            for tc in tr.findall(w('tc')):
                if tc.find('.//'+w('tbl')) is not None:return None
                gs=tc.find('./'+w('tcPr')+'/'+w('gridSpan'));vm=tc.find('./'+w('tcPr')+'/'+w('vMerge'))
                cells.append((tuple(text(p) for p in tc.findall(w('p'))),gs.get(w('val'),'1') if gs is not None else '1',vm.get(w('val'),'continue') if vm is not None else ''))
            before=tr.find('./'+w('trPr')+'/'+w('gridBefore'));after=tr.find('./'+w('trPr')+'/'+w('gridAfter'))
            rows.append((before.get(w('val'),'0') if before is not None else '0',after.get(w('val'),'0') if after is not None else '0',tuple(cells)))
        return tuple(rows) if rows else None
    x=sig(a);return x is not None and x==sig(b)

def equivalent_plain_block(a,b):
    if a.tag!=b.tag:return False
    if a.tag==w('tbl'):return equivalent_table(a,b)
    if a.tag!=w('p') or text(a)!=text(b):return False
    forbidden={'drawing','pict','object','fldChar','fldSimple','sdt','hyperlink','sym','numPr','altChunk'}
    return not any(local(e) in forbidden or str(e.tag).startswith('{'+M+'}') for el in (a,b) for e in el.iter())

def resolve_revisions(root,accept=True):
    """Resolve textual/row/paragraph/property changes in place, retaining current fields."""
    discard='del' if accept else 'ins'
    parents=_parents(root)
    for tr in list(root.iter(w('tr'))):
        if tr.find('./'+w('trPr')+'/'+w(discard)) is not None:_remove_with(parents,tr)

    parents=_parents(root)
    for e in list(root.iter()):
        if local(e).endswith('PrChange') or e.tag==w('tblGridChange'):
            parent=parents.get(e)
            if not accept and parent is not None and len(e):
                saved=copy.deepcopy(e[0]);grand=parents.get(parent)
                if grand is not None:
                    try:grand[ list(grand).index(parent) ]=saved
                    except ValueError:pass
            else:_remove_with(parents,e)

    mark_paras=[]
    for p in root.iter(w('p')):
        if p.find('./'+w('pPr')+'/'+w('rPr')+'/'+w(discard)) is not None:mark_paras.append(p)

    parents=_parents(root)
    for e in list(root.iter()):
        tag=local(e)
        if tag in ('ins','del','moveFrom','moveTo'):
            parent=parents.get(e)
            if parent is None:continue
            if local(parent) in ('rPr','trPr','numPr','tcPr'):
                _remove_with(parents,e);continue
            delete=tag in (discard,'moveFrom' if accept else 'moveTo')
            if delete:_remove_with(parents,e)
            else:
                try:idx=list(parent).index(e)
                except ValueError:continue
                children=list(e)
                for child in children:
                    try:e.remove(child)
                    except ValueError:pass
                    parent.insert(idx,child);idx+=1
                _remove_with(parents,e)
        elif tag.startswith(('moveFromRange','moveToRange','customXmlInsRange','customXmlDelRange')):
            _remove_with(parents,e)
    for t in root.iter(w('delText')):t.tag=w('t')
    for t in root.iter(w('delInstrText')):t.tag=w('instrText')

    parents=_parents(root)
    for p in mark_paras:
        parent=parents.get(p)
        if parent is None:continue
        siblings=list(parent)
        try:i=siblings.index(p)
        except ValueError:continue
        nxt=siblings[i+1] if i+1<len(siblings) else None
        if nxt is not None and nxt.tag==w('p'):
            payload=[c for c in list(p) if c.tag!=w('pPr')]
            insert=1 if len(nxt) and nxt[0].tag==w('pPr') else 0
            for c in payload:
                try:p.remove(c)
                except ValueError:pass
                nxt.insert(insert,c);insert+=1
            try:parent.remove(p)
            except ValueError:pass
        elif not text(p).strip() and not p.findall('.//'+w('drawing')) and not p.findall('.//'+w('pict')):
            try:parent.remove(p)
            except ValueError:pass

    parents=_parents(root)
    for tbl in list(root.iter(w('tbl'))):
        if not tbl.findall(w('tr')):_remove_with(parents,tbl)
    for tc in root.iter(w('tc')):
        if not any(c.tag in (w('p'),w('tbl')) for c in tc):E.SubElement(tc,w('p'))
        elif len(tc) and tc[-1].tag==w('tbl'):E.SubElement(tc,w('p'))
    for ppr in root.iter(w('pPr')):
        rpr=ppr.find(w('rPr'))
        if rpr is not None and len(rpr)==0 and not rpr.attrib:ppr.remove(rpr)
    return root

class Package:
    def __init__(self,path):
        self.path=Path(path);self.raw=self.path.read_bytes();self.hash=sha(self.raw)
        with zipfile.ZipFile(self.path) as z:
            if sum(i.file_size for i in z.infolist())>350*1024**2:raise ValueError('单个 Word 解压体积超过安全上限。')
            self.entries={i.filename:z.read(i) for i in z.infolist() if not i.is_dir()}
        if 'word/document.xml' not in self.entries:raise ValueError('不是有效的 DOCX 文件。')
        self.root=xml(self.entries['word/document.xml']);self.body=self.root.find(w('body'))
        if self.body is None:raise ValueError('Word 缺少正文。')
        revisions={w('ins'),w('del'),w('moveFrom'),w('moveTo')}
        self.revision_count=sum(1 for e in self.root.iter() if e.tag in revisions)
        self._styles=None;self._numbering=None;self._relroot=None;self._import_cache={};self._style_map={};self._num_map={}
        self._parents_cache=None;self._comment_covered_cache=None;self._track_configured=False
        ids=[];drawids=[]
        for e in self.root.iter():
            v=e.get(w('id'))
            if v and str(v).isdigit():ids.append(int(v))
            if local(e)=='docPr':
                v=e.get('id')
                if v and str(v).isdigit():drawids.append(int(v))
        self.rid=max(ids+[0])+1;self.drawid=max(drawids+[0])+1
        # Paragraph/row identities are navigation anchors in WPS.  A pasted source
        # block must never reuse the target's w14:paraId; WPS can render duplicates
        # but Reject/Accept navigation may then lose the current viewport and jump
        # to the beginning of the document.  Keep a package-wide identity pool and
        # mint deterministic fresh 8-hex ids for every pasted paragraph/table row.
        self._used_para_ids={e.get('{'+W14+'}paraId') for e in self.root.iter() if e.get('{'+W14+'}paraId')}
        self._used_para_ids.discard(None);self._para_id_seq=0
        # Keep the deletion/insertion nodes produced by one source replacement on
        # the same normal revision timestamp, as WPS commonly does for a manual
        # replacement.  No plug-in or custom document metadata consumes this value.
        self._revision_epoch=datetime.now(timezone.utc).replace(microsecond=0)
        self._transaction_dates={}
        self._transaction_seq=0
        self._active_transaction=None
        self.date=self._revision_epoch.strftime('%Y-%m-%dT%H:%M:%SZ')
    def baseline(self):
        resolve_revisions(self.root,True)
        revisions={w('ins'),w('del'),w('moveFrom'),w('moveTo')}
        for k,v in list(self.entries.items()):
            if re.fullmatch(r'word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml',k):
                r=xml(v)
                if any(e.tag in revisions for e in r.iter()):
                    resolve_revisions(r,True);self.entries[k]=dump(r)
        # Revision IDs form their own change stream.  After historical changes are
        # accepted into the working baseline, start the new review pass from the
        # next remaining revision ID instead of borrowing bookmark/comment IDs.
        # WPS itself writes a compact sequential revision stream; matching that
        # shape avoids thousands-high synthetic IDs that have no review meaning.
        rev_tags={w('ins'),w('del'),w('moveFrom'),w('moveTo'),w('pPrChange'),w('rPrChange'),w('tblPrChange'),w('trPrChange'),w('tcPrChange')}
        rev_ids=[]
        for e in self.root.iter():
            if e.tag in rev_tags:
                v=e.get(w('id'))
                if v and v.isdigit():rev_ids.append(int(v))
        self.rid=max(rev_ids+[-1])+1
        self._parents_cache=None;self._comment_covered_cache=None
        return self
    def _finalize_wps_review_shape(self):
        """Canonicalize the package to the teacher-approved WPS review shape.

        The golden WPS file has three important properties: every paragraph/table row
        has a unique ``w14:paraId``; ``document.xml`` contains no Word ``rsid*`` or
        ``w14:textId`` editing-session metadata; and every tracked-change marker has a
        unique sequential revision id.  Duplicate paragraph identities are especially
        dangerous: WPS still renders the page, but rejecting a table row or large block
        can make its review cursor lose the anchor and jump to page 1.

        The metadata normalized here is non-visible bookkeeping.  Text, styles, tables,
        relationships and Accept/Reject semantics are unchanged.
        """
        revtags={w('ins'),w('del'),w('moveFrom'),w('moveTo'),w('pPrChange'),w('rPrChange'),w('tblPrChange'),w('trPrChange'),w('tcPrChange'),w('tblGridChange')}

        # WPS' persisted document.xml does not keep Microsoft's editing-session rsids
        # or w14:textId.  Remove them globally, not only from directly revised runs,
        # because a table row can contain untouched-looking descendants that still
        # participate in the same review/navigation context.
        for node in self.root.iter():
            for key in list(node.attrib):
                if key=='{'+W14+'}textId' or (key.startswith('{'+W+'}') and key.rsplit('}',1)[-1].startswith('rsid')):
                    node.attrib.pop(key,None)

        # Enforce one identity per paragraph/row.  Preserve the first occurrence
        # (normally the original/deleted block) and remint later duplicates (normally
        # the pasted/inserted copy).  Also add identities when a foreign DOCX omitted
        # them, matching the WPS golden file where every w:p and w:tr has one.
        seen=set()
        for node in self.root.iter():
            if node.tag not in (w('p'),w('tr')):continue
            key='{'+W14+'}paraId';pid=node.get(key)
            valid_pid=bool(pid and re.fullmatch(r'[0-9A-Fa-f]{8}',pid))
            if valid_pid:
                value=int(pid,16)
                valid_pid=(0 < value < 0x80000000)
            if not valid_pid or pid.upper() in seen:
                pid=self._fresh_para_id();node.set(key,pid)
            else:
                pid=pid.upper();node.set(key,pid);seen.add(pid);self._used_para_ids.add(pid)
            seen.add(pid)

        # Match WPS' persisted shape for every paragraph/row/table that actually
        # participates in this review pass. Untouched paragraphs remain byte-identical.
        for p in self.root.iter(w('p')):
            if any(e.tag in revtags for e in p.iter()):
                self._wps_revision_context(p)
        for tr in self.root.iter(w('tr')):
            pr=tr.find(w('trPr'))
            if pr is not None and any(e.tag in revtags for e in pr.iter()):
                self._wps_revision_context(tr)
        for tbl in self.root.iter(w('tbl')):
            pr=tbl.find(w('tblPr'))
            grid=tbl.find(w('tblGrid'))
            if ((pr is not None and any(e.tag in revtags for e in pr.iter())) or
                (grid is not None and any(e.tag in revtags for e in grid.iter()))):
                self._wps_revision_context(tbl)

        # Rebase any legacy duplicate footnote/endnote ids before final save.
        self._deduplicate_note_references()

        # The validated WPS teacher sample uses one unique revision id for every
        # persisted revision marker, in strict document order.  In particular:
        # paragraph-mark insertion/deletion, paragraph content insertion/deletion,
        # and every inserted/deleted table row all have distinct ids.  Reusing one
        # id across a large block renders, but WPS review navigation/reject behavior
        # is unstable.  Canonicalize to the native WPS shape exactly.
        next_id=0
        for e in self.root.iter():
            if e.tag not in revtags:continue
            e.set(w('id'),str(next_id));next_id+=1
        self.rid=next_id

    def save(self,path):
        # Finalize only review bookkeeping. Visible content and untouched document
        # parts remain unchanged, while WPS receives the same paragraph/revision shape
        # it writes itself after a native tracked edit.
        self._finalize_wps_review_shape()
        self.entries['word/document.xml']=dump(self.root)
        if self._relroot is not None:self.entries['word/_rels/document.xml.rels']=dump(self._relroot)
        if self._styles is not None:self.entries['word/styles.xml']=dump(self._styles)
        if self._numbering is not None:self.entries['word/numbering.xml']=dump(self._numbering)
        # WordprocessingML property lists are schema sequences, even when an
        # office suite appears to tolerate arbitrary order. Font overrides and
        # target table geometry can introduce order violations. Normalize only
        # those property children, preserving every value and content object.
        from .compatibility import normalize_property_order
        self.compatibility_fixes=[]
        live={'word/document.xml':self.root,
              'word/styles.xml':self._styles,'word/numbering.xml':self._numbering}
        for part,data in list(self.entries.items()):
            if not part.startswith('word/') or not part.endswith('.xml'):continue
            partroot=live.get(part)
            if partroot is None:partroot=xml(data)
            result=normalize_property_order(partroot)
            if result['issues']:
                raise ValueError('Word 兼容结构检查未通过：'+part+'；'+result['issues'][0])
            if result['changed']:
                self.entries[part]=dump(partroot)
                self.compatibility_fixes.append({'part':part,**result})
        # Word has a finite style table. Legacy workpapers may already exceed
        # it; remove only provably unreachable definitions, never used formats.
        from .compatibility import enforce_style_capacity
        style_parts={name:xml(data) for name,data in self.entries.items()
                     if name.endswith('.xml')}
        capacity=enforce_style_capacity(style_parts)
        for part in capacity['parts_changed']:
            self.entries[part]=dump(style_parts[part])
            if part=='word/styles.xml':self._styles=style_parts[part]
        if capacity['parts_changed']:
            self.compatibility_fixes.append({'kind':'style_capacity',**capacity})
        path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
        with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
            for k,v in self.entries.items():z.writestr(k,v)
    def _parents_cached(self):
        # Replacement proposals always point at original baseline blocks. Their
        # container parent never changes while we process non-overlapping ranges in
        # reverse order, so rebuilding a full XML parent map for every proposal is
        # pure overhead on large workpapers.
        if self._parents_cache is None:self._parents_cache=_parents(self.root)
        return self._parents_cache
    def _comment_covered_cached(self):
        # Generated source synchronization never creates comments. The set of
        # original human-comment anchors is therefore immutable during one export.
        if self._comment_covered_cache is None:self._comment_covered_cache=comment_covered_elements(self.root)
        return self._comment_covered_cache
    def rels(self):
        if self._relroot is None:self._relroot=xml(self.entries.get('word/_rels/document.xml.rels',f'<Relationships xmlns="{PR}"/>'.encode()))
        return self._relroot
    def add_rel(self,typ,target,mode=None):
        root=self.rels()
        for r in root:
            if r.get('Type')==typ and r.get('Target')==target and r.get('TargetMode')==mode:return r.get('Id')
        used={r.get('Id') for r in root};n=1
        while 'rId'+str(n) in used:n+=1
        r=E.SubElement(root,'{'+PR+'}Relationship',{'Id':'rId'+str(n),'Type':typ,'Target':target})
        if mode:r.set('TargetMode',mode)
        return r.get('Id')
    def content_type(self,part,kind):
        root=xml(self.entries['[Content_Types].xml']);name='/'+part
        if not any(c.get('PartName')==name for c in root):E.SubElement(root,'{'+CT+'}Override',{'PartName':name,'ContentType':kind})
        self.entries['[Content_Types].xml']=dump(root)
    def track(self):
        if self._track_configured:return
        r=xml(self.entries.get('word/settings.xml',f'<w:settings xmlns:w="{W}"/>'.encode()))
        t=r.find(w('trackRevisions'))
        if t is None:
            t=E.Element(w('trackRevisions'));t.set(w('val'),'1')
            later={'doNotTrackMoves','doNotTrackFormatting','documentProtection','autoFormatOverride','styleLockTheme','styleLockQFSet','defaultTabStop','autoHyphenation','consecutiveHyphenLimit','hyphenationZone','doNotHyphenateCaps','showEnvelope','summaryLength','clickAndTypeStyle','defaultTableStyle','evenAndOddHeaders','bookFoldRevPrinting','bookFoldPrinting','bookFoldPrintingSheets','drawingGridHorizontalSpacing','compat','rsids','mathPr','themeFontLang','clrSchemeMapping','decimalSymbol','listSeparator'}
            ix=next((i for i,c in enumerate(r) if local(c) in later),len(r));r.insert(ix,t)
        else:t.set(w('val'),'1')
        self.entries['word/settings.xml']=dump(r)
        self.add_rel(R+'/settings','settings.xml');self.content_type('word/settings.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
        self._track_configured=True
    def begin_transaction(self,key):
        """Keep the whole export in one native WPS editing-session timestamp.

        The teacher-approved WPS file persists essentially the entire review pass with
        one author and one timestamp while every revision marker still has its own id.
        Splitting one export into dozens of synthetic second-by-second timestamps makes
        WPS treat the document like many unrelated editing sessions and destabilizes
        review navigation after Accept/Reject.  Transactions remain logical grouping
        metadata for our own export log only; they no longer alter w:date.
        """
        key=str(key or ('tx-'+str(self._transaction_seq+1)))
        if key not in self._transaction_dates:
            self._transaction_seq+=1
            self._transaction_dates[key]=self._revision_epoch.strftime('%Y-%m-%dT%H:%M:%SZ')
        self._active_transaction=key
        self.date=self._transaction_dates[key]
        return self.date
    def end_transaction(self):
        self._active_transaction=None
    def transaction_date(self,key):
        return self._transaction_dates.get(str(key))
    def change(self,tag,revision_id=None):
        e=E.Element(w(tag))
        if revision_id is None:
            revision_id=self.rid;self.rid+=1
        e.set(w('id'),str(revision_id));e.set(w('author'),self.author);e.set(w('date'),self.date);return e
    def reserve_revision_id(self):
        rid=self.rid;self.rid+=1;return rid
    def _fresh_para_id(self):
        """Return a deterministic package-unique Word/WPS paragraph identity.

        Microsoft Word requires w14:paraId to be a positive 31-bit value
        (00000001..7FFFFFFF).  WPS is more permissive and will open values with the
        high bit set, but Word may treat such DOCX files as corrupt and offer to
        repair them.  Keep generated identities inside Word's valid range.
        """
        while True:
            self._para_id_seq+=1
            digest=hashlib.sha256((self.hash+':para:'+str(self._para_id_seq)).encode('utf-8')).digest()
            value=int.from_bytes(digest[:4],'big') & 0x7FFFFFFF
            if value==0:
                continue
            pid=f'{value:08X}'
            if pid not in self._used_para_ids:
                self._used_para_ids.add(pid);return pid
    def _refresh_block_identity(self,el):
        """Give a pasted block new navigation identities and clear session metadata."""
        for node in el.iter():
            if node.tag in (w('p'),w('tr')):node.set('{'+W14+'}paraId',self._fresh_para_id())
            for key in list(node.attrib):
                if key=='{'+W14+'}textId' or (key.startswith('{'+W+'}') and key.rsplit('}',1)[-1].startswith('rsid')):
                    node.attrib.pop(key,None)
        return el
    def _wps_revision_context(self,el):
        """Normalize only the touched revision context to WPS' native save shape.

        A real WPS-saved teacher file in the regression corpus keeps w14:paraId but
        strips Microsoft's rsid/textId bookkeeping from edited paragraphs/runs.
        Those attributes are not content and do not participate in Accept/Reject.
        Removing them only from generated revision contexts makes our XML match the
        native WPS shape more closely without rewriting untouched document parts.
        """
        for node in el.iter():
            for key in list(node.attrib):
                if key=='{'+W14+'}textId' or (key.startswith('{'+W+'}') and key.rsplit('}',1)[-1].startswith('rsid')):
                    node.attrib.pop(key,None)
        return el
    def _mark(self,el,kind,revision_id=None):
        """Mark one block using the persisted shape produced by native WPS.

        The teacher-approved WPS sample is the golden shape: each revision marker has
        its own id.  A deleted/inserted paragraph therefore has one marker for the
        paragraph mark and another for its content; every table row has its own row
        marker.  `revision_id` is retained only for API compatibility and is ignored.
        """
        # Do NOT special-case a table as row-only revision. Native WPS marks the
        # complete table replacement at three levels: every row, every paragraph
        # mark inside the cells, and the actual cell runs. Row-only <w:trPr><w:del/>
        # or <w:ins/> produces a blank “删除:”/“插入:” balloon because the revision
        # has no textual payload, and Reject/Accept navigation may jump away from the
        # current page. Let the generic block walker below mark the full table content
        # exactly like the teacher-approved WPS sample.
        if el.tag==w('p') and self._paragraph_patchable(el):
            runs=[c for c in list(el) if c.tag==w('r')]
            self._wrap_runs_revision(el,runs,kind)
            pr=el.find(w('pPr'))
            if pr is None:pr=E.Element(w('pPr'));el.insert(0,pr)
            if pr.find(w('sectPr')) is None:
                rpr=pr.find(w('rPr'))
                if rpr is None:rpr=E.SubElement(pr,w('rPr'))
                rpr.insert(0,self.change(kind))
            return
        parents=_parents(el)
        # Wrap the outermost native equation as one revision unit. Its m:t and
        # m:r nodes stay unchanged; delText is a Word text-run concept, not OMML.
        # The supported boundary is a direct w:p child, checked before mutation.
        for formula in list(el.iter()):
            if formula.tag not in (m('oMath'),m('oMathPara')):continue
            ancestors=list(_ancestor_tags(formula,parents))
            if any(str(t).startswith('{'+M+'}') for t in ancestors) or any(t in (w('ins'),w('del')) for t in ancestors):continue
            parent=parents.get(formula)
            if parent is None or parent.tag!=w('p'):raise ValueError('原生公式修订边界不受支持。')
            idx=list(parent).index(formula);rev=self.change(kind)
            parent[idx]=rev;rev.append(formula)
        for run in list(el.iter(w('r'))):
            ancestors=list(_ancestor_tags(run,parents))
            if w('r') in ancestors or any(a in (w('ins'),w('del')) or str(a).startswith('{'+M+'}') for a in ancestors):continue
            if kind=='del':
                for t in run.iter(w('t')):t.tag=w('delText')
                for t in run.iter(w('instrText')):t.tag=w('delInstrText')
            parent=parents.get(run)
            if parent is None:continue
            try:idx=list(parent).index(run)
            except ValueError:continue
            rev=self.change(kind);parent[idx]=rev;rev.append(run)
        for p in el.iter(w('p')):
            pr=p.find(w('pPr'))
            if pr is None:pr=E.Element(w('pPr'));p.insert(0,pr)
            if pr.find(w('sectPr')) is not None:continue
            rpr=pr.find(w('rPr'))
            if rpr is None:rpr=E.SubElement(pr,w('rPr'))
            rpr.insert(0,self.change(kind))
        for tr in el.iter(w('tr')):
            pr=tr.find(w('trPr'))
            if pr is None:pr=E.Element(w('trPr'));tr.insert(0,pr)
            pr.append(self.change(kind))
    def _paragraph_patchable(self,p):
        """Content-only tracked replacement that preserves the target paragraph shell.

        Bookmarks/proofing markers may remain in the target. Rich objects and nested
        containers are deliberately excluded because their relationship semantics are
        not a text replacement.
        """
        forbidden={'drawing','pict','object','fldChar','fldSimple','sdt','hyperlink','sym','altChunk','footnoteReference','endnoteReference'}
        if p.tag!=w('p'):return False
        if any(local(e) in forbidden or str(e.tag).startswith('{http://schemas.openxmlformats.org/officeDocument/2006/math}') for e in p.iter()):return False
        # Runs nested under smart tags/custom XML are not direct, independently
        # reversible revision units in WPS. Plain runs plus invisible anchors are safe.
        direct_runs={id(c) for c in p if c.tag==w('r')}
        if any(id(r) not in direct_runs for r in p.iter(w('r'))):return False
        return True

    def _prepared_source(self,el,source):
        clone=copy.deepcopy(el)
        if source is not None:self.import_block(clone,source)
        else:
            parents_new=_parents(clone)
            for node in list(clone.iter()):
                if local(node) in ('bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd','commentReference'):
                    _remove_with(parents_new,node)
        # A pasted block is a new document object, not a second copy of the target's
        # navigation identity.  Refresh all paragraph/row ids before revision markup
        # is added; otherwise WPS may associate Accept/Reject with the wrong twin.
        self._refresh_block_identity(clone)
        return clone

    def _wrap_runs_revision(self,p,runs,kind,revision_id=None):
        """Wrap contiguous direct runs using as few revision nodes as possible."""
        if not runs:return
        wanted=set(map(id,runs));children=list(p);i=0
        while i<len(children):
            if children[i].tag!=w('r') or id(children[i]) not in wanted:
                i+=1;continue
            group=[];j=i
            while j<len(children) and children[j].tag==w('r') and id(children[j]) in wanted:
                group.append(children[j]);j+=1
            rev=self.change(kind,revision_id)
            for r in group:
                if kind=='del':
                    for t in r.iter(w('t')):t.tag=w('delText')
                    for t in r.iter(w('instrText')):t.tag=w('delInstrText')
                p.remove(r);rev.append(r)
            p.insert(i,rev)
            children=list(p);i+=1

    def _insert_revision_runs(self,p,at,runs):
        if not runs:return
        rev=self.change('ins')
        for r in runs:rev.append(copy.deepcopy(r))
        p.insert(max(0,min(at,len(p))),rev)

    def _replace_paragraph_whole_inplace(self,old,new,source=None):
        """Replace ONE paragraph as a full tracked delete + full tracked insert.

        This deliberately does not compute a sentence/run diff.  The archive workflow
        first locates a source paragraph, then copies that paragraph literally.  The
        target paragraph shell stays in place so WPS shows one paragraph position,
        while Reject restores the old text and Accept yields the copied source text.
        """
        if not self._paragraph_patchable(old) or not self._paragraph_patchable(new):return False
        prepared=self._prepared_source(new,source)
        old_runs=[c for c in list(old) if c.tag==w('r')]
        new_runs=[c for c in list(prepared) if c.tag==w('r')]
        # A truly empty paragraph has nothing textual to redline.  Keep its shell.
        if not old_runs and not new_runs:return True
        if old_runs:
            self._wrap_runs_revision(old,old_runs,'del')
            children=list(old);del_positions=[i for i,c in enumerate(children) if c.tag==w('del')]
            at=(max(del_positions)+1) if del_positions else (1 if old.find(w('pPr')) is not None else 0)
        else:
            at=1 if old.find(w('pPr')) is not None else 0
        self._insert_revision_runs(old,at,new_runs)
        return True

    def _patch_paragraph(self,old,new,source=None):
        """Synchronize one paragraph in-place as one reviewer-friendly replacement.

        Equal prefix/suffix runs stay untouched. Everything between the first and
        last difference is represented as ONE inserted revision plus the minimum
        number of deletion wrappers needed to preserve invisible bookmark anchors.
        Thus a changed paragraph no longer explodes into dozens of red revision
        records, while rejecting all changes still restores byte-equivalent content.
        """
        if text(old)==text(new):return True
        if not self._paragraph_patchable(old) or not self._paragraph_patchable(new):return False
        prepared=self._prepared_source(new,source)
        old_runs=[c for c in list(old) if c.tag==w('r')]
        new_runs=[c for c in list(prepared) if c.tag==w('r')]
        if not old_runs and not new_runs:return True
        a=[text(r) for r in old_runs];b=[text(r) for r in new_runs]
        sm=difflib.SequenceMatcher(None,a,b,autojunk=False)
        changes=[x for x in sm.get_opcodes() if x[0]!='equal']
        if not changes:return True
        first,last=changes[0],changes[-1]
        i1=first[1];i2=last[2];j1=first[3];j2=last[4]
        affected=old_runs[i1:i2]
        children=list(old)
        if affected:
            try:at=children.index(affected[0])
            except ValueError:return False
        elif i1<len(old_runs):
            try:at=children.index(old_runs[i1])
            except ValueError:return False
        elif old_runs:
            try:at=children.index(old_runs[-1])+1
            except ValueError:return False
        else:at=1 if old.find(w('pPr')) is not None else 0
        if affected:
            self._wrap_runs_revision(old,affected,'del')
            children=list(old)
            # Insert after the last deletion wrapper covering the changed range.
            dels=[]
            affected_ids={id(x) for x in affected}
            for k,c in enumerate(children):
                if c.tag==w('del') and any(id(r) in affected_ids for r in list(c)):dels.append(k)
            if dels:at=max(dels)+1
        self._insert_revision_runs(old,at,new_runs[j1:j2])
        return True

    def _table_structure(self,tbl):
        forbidden={'drawing','pict','object','fldChar','fldSimple','sdt','hyperlink','sym','altChunk','hMerge'}
        if any(local(e) in forbidden or str(e.tag).startswith('{http://schemas.openxmlformats.org/officeDocument/2006/math}') for e in tbl.iter()):return None
        rows=[]
        for tr in tbl.findall(w('tr')):
            cells=[]
            for tc in tr.findall(w('tc')):
                if tc.find('.//'+w('tbl')) is not None:return None
                gs=tc.find('./'+w('tcPr')+'/'+w('gridSpan'));vm=tc.find('./'+w('tcPr')+'/'+w('vMerge'))
                ps=tc.findall(w('p'))
                if not all(self._paragraph_patchable(x) for x in ps):return None
                cells.append((gs.get(w('val'),'1') if gs is not None else '1',vm.get(w('val'),'continue') if vm is not None else '',len(ps)))
            before=tr.find('./'+w('trPr')+'/'+w('gridBefore'));after=tr.find('./'+w('trPr')+'/'+w('gridAfter'))
            rows.append((before.get(w('val'),'0') if before is not None else '0',after.get(w('val'),'0') if after is not None else '0',tuple(cells)))
        return tuple(rows) if rows else None

    def _replace_table_whole_inplace(self,old,new,source=None):
        """Copy a whole source table into ONE existing table with tracked cell text.

        When the row/column/merge grid is identical, creating a second table is both
        unnecessary and confusing in WPS.  Keep the target table shell and replace
        every cell paragraph as a full delete+insert, including text-identical cells.
        Reject restores the original table; Accept yields the literal source values.
        """
        shape=self._table_structure(old)
        if shape is None or shape!=self._table_structure(new):return False
        prepared=self._prepared_source(new,source)
        old_rows=old.findall(w('tr'));new_rows=prepared.findall(w('tr'))
        for otr,ntr in zip(old_rows,new_rows):
            for otc,ntc in zip(otr.findall(w('tc')),ntr.findall(w('tc'))):
                ops=otc.findall(w('p'));nps=ntc.findall(w('p'))
                if len(ops)!=len(nps):return False
        for otr,ntr in zip(old_rows,new_rows):
            for otc,ntc in zip(otr.findall(w('tc')),ntr.findall(w('tc'))):
                for op,np in zip(otc.findall(w('p')),ntc.findall(w('p'))):
                    if not self._replace_paragraph_whole_inplace(op,np,None):return False
        return True

    def _patch_table(self,old,new,source=None):
        """For the same table grid, update source-backed cell contents in place.

        This keeps ONE table in WPS instead of an inserted red table plus a deleted
        original table. Every changed cell remains a real reversible revision.
        """
        shape=self._table_structure(old)
        if shape is None or shape!=self._table_structure(new):return False
        prepared=self._prepared_source(new,source)
        old_rows=old.findall(w('tr'));new_rows=prepared.findall(w('tr'))
        for otr,ntr in zip(old_rows,new_rows):
            for otc,ntc in zip(otr.findall(w('tc')),ntr.findall(w('tc'))):
                ops=otc.findall(w('p'));nps=ntc.findall(w('p'))
                if len(ops)!=len(nps):return False
        # preflight complete; actual edits cannot partially fall back now.
        for otr,ntr in zip(old_rows,new_rows):
            for otc,ntc in zip(otr.findall(w('tc')),ntr.findall(w('tc'))):
                for op,np in zip(otc.findall(w('p')),ntc.findall(w('p'))):
                    if text(op)!=text(np) and not self._patch_paragraph(op,np,None):return False
        return True

    def _replace_block_fallback(self,old,new,source=None):
        """Fallback for structural block changes; kept local to ONE block.

        For tables, mirror the manual WPS workflow literally: delete the original
        table first, then paste the complete source table immediately after it.
        That keeps the reviewer-visible sequence as OLD(deleted) -> NEW(inserted),
        instead of making a same-table cell diff or placing the new table before
        the deleted original.  Paragraph fallback keeps the historical ordering.
        """
        parents=self._parents_cached();parent=parents.get(old)
        if parent is None:return []
        clone=self._prepared_source(new,source)
        if old.tag==w('tbl') and new.tag==w('tbl'):
            self._match_table_geometry(old,clone)
        self._mark(clone,'ins')
        try:idx=list(parent).index(old)
        except ValueError:raise ValueError('替换锚点已失效。')
        if old.tag==w('tbl'):
            self._mark(old,'del')
            parent.insert(idx+1,clone)
        else:
            parent.insert(idx,clone);self._mark(old,'del')
        self._wps_revision_context(clone)
        return [clone]

    def _stabilize_bookmarks_for_whole_replace(self,blocks,parent):
        """Lift bookmark endpoints out of blocks that will be deleted as a whole.

        A native WPS delete/paste can leave bookmark anchors as persistent siblings of
        the revised block range.  If an endpoint stays nested inside a block carrying
        deletion revisions, Accept All can remove only that endpoint while its mate
        survives elsewhere, producing a dangling bookmark and unstable review
        navigation. For top-level body replacements, lift leading endpoints before
        their block and the remaining endpoints after it, in original encounter
        order. Endpoint *type* does not determine its side: several empty bookmarks
        can precede one heading as Start/End/Start/End. Partitioning starts from ends
        would turn those into crossing non-empty ranges and corrupt their order.
        This changes only navigation anchors; visible content is untouched.
        """
        if parent is None or parent.tag!=w('body'):
            return 0
        moved=0
        for block in list(blocks):
            try:base_index=list(parent).index(block)
            except ValueError:continue
            before=[];after=[];seen_payload=False
            for node in list(block.iter()):
                if node.tag in (w('r'),m('oMath'),m('oMathPara')):
                    seen_payload=True
                if node.tag not in (w('bookmarkStart'),w('bookmarkEnd')):continue
                # Endpoints already at body level are not descendants of a block and
                # therefore never reach this loop.
                (after if seen_payload else before).append(node)
            if not before and not after:continue
            parents=_parents(block)
            for node in before+after:
                par=parents.get(node)
                if par is not None:
                    try:par.remove(node);moved+=1
                    except ValueError:pass
            # Preserve the original endpoint sequence, including zero-width pairs.
            for off,node in enumerate(before):
                parent.insert(base_index+off,node)
            try:after_index=list(parent).index(block)+1
            except ValueError:continue
            for off,node in enumerate(after):
                parent.insert(after_index+off,node)
        if moved:
            self._parents_cache=None
            self._comment_covered_cache=None
        return moved

    def replace(self,old,new,author='柒',source=None,comment='',whole=False):
        if not old:raise ValueError('替换区间为空。')
        # A paragraph carrying sectPr also owns the preceding section's page
        # geometry. Deleting only its runs and pasting after it moves new tables
        # into the next section; deleting its paragraph mark loses the boundary.
        # Neither is a supported whole-source copy. Stop before any mutation.
        if any(e.tag==w('sectPr') for block in old for e in block.iter()):
            raise ValueError('目标范围包含原生节分界；保留该段及节布局并列为待核对，不能自动删除后粘贴。')
        parents=self._parents_cached();parent=parents.get(old[0])
        if parent is None or any(parents.get(e) is not parent for e in old):raise ValueError('替换区间跨越了不兼容的表格或正文容器。')
        reason=complex_reference_reason(old) or complex_reference_reason(new)
        reason=reason or native_math_copy_reason(self,source,new)
        if reason:raise ValueError(reason)
        anchored=self._comment_covered_cached()
        if any(id(e) in anchored or any(id(ch) in anchored for ch in e.iter()) for e in old):
            raise ValueError('此范围含原有人工批注；保留批注及其定位，不能直接覆盖。')
        self.author=author
        if (not whole) and len(old)==len(new) and all(equivalent_plain_block(a,b) for a,b in zip(old,new)):
            return list(old)

        # Source-backed synchronization is deliberately coarse-grained.  It must
        # look like the manual archive workflow in WPS: select the complete old
        # paragraph/table range, delete it with Track Changes on, then paste the
        # complete source range immediately after the deleted range.  Do NOT fold
        # the new text into the old paragraph shell and do NOT patch table cells.
        # Keeping old and new as separate block ranges makes the review surface one
        # large replacement instead of hundreds of sentence/cell edits.
        if whole:
            # The same anchor guard used by preflight must also protect direct
            # writer callers, before any bookmarks or revisions are mutated.
            from .layout_policy import floating_table_anchor_reason
            old_tables=[e for e in old if e.tag==w('tbl')]
            new_tables=[e for e in new if e.tag==w('tbl')]
            table_map={id(n):o for o,n in zip(old_tables,new_tables)} if len(old_tables)==len(new_tables) else {}
            for table in new_tables:
                anchor_reason=floating_table_anchor_reason(old,new,table_map.get(id(table)),table)
                if anchor_reason:raise ValueError('整表布局待核对：'+anchor_reason)
            # Keep bookmark endpoints outside blocks that are about to become fully
            # deleted.  Otherwise Accept All can remove one endpoint and leave its
            # mate dangling, which is a known trigger for WPS jumping to page 1.
            self._stabilize_bookmarks_for_whole_replace(old,parent)
            # Mirror the teacher-approved native WPS workflow literally: old blocks
            # remain as deleted blocks, followed by separately inserted source blocks.
            # Do not reuse revision ids across paragraphs, paragraph marks or rows.
            for e in old:self._mark(e,'del')
            try:insert_at=list(parent).index(old[-1])+1
            except ValueError:raise ValueError('替换锚点已失效。')
            inserted=[]
            target_width=self._target_section_width(old[0])
            for srcel in new:
                clone=self._prepared_source(srcel,source)
                fit_width=target_width
                if srcel.tag==w('tbl') and id(srcel) in table_map:
                    # A matching source/old fixed table may deliberately extend
                    # into the margins. Preserve it only after proving identical
                    # native geometry and safe actual paper bounds.
                    from .layout_policy import preserved_native_table_layout,table_has_preceding_paragraph
                    separate_anchors=(table_has_preceding_paragraph(old,table_map[id(srcel)]) and
                                      table_has_preceding_paragraph(new,srcel))
                    native=preserved_native_table_layout(self.root,table_map[id(srcel)],clone,allow_floating=separate_anchors)
                    if native:fit_width=max(fit_width,native['fit_width_twips'])
                    if not native or not native.get('copy_source_grid'):
                        self._match_table_geometry(table_map[id(srcel)],clone)
                if clone.tag==w('tbl') or clone.find('.//'+w('tbl')) is not None:
                    self.fit_new_tables(clone,width=fit_width)
                self._mark(clone,'ins')
                parent.insert(insert_at,clone);insert_at+=1;inserted.append(clone)
            self.track();return inserted

        # Most source refreshes preserve block count and order. Reconcile each
        # paragraph/table independently so an unchanged paragraph inside a larger
        # section never becomes a fake full-paragraph redline.
        if len(old)==len(new) and all(a.tag==b.tag for a,b in zip(old,new)):
            out=[None]*len(old)
            for ix in range(len(old)-1,-1,-1):
                a,b=old[ix],new[ix]
                if equivalent_plain_block(a,b):out[ix]=a;continue
                if a.tag==w('p') and self._patch_paragraph(a,b,source):out[ix]=a;continue
                if a.tag==w('tbl') and self._patch_table(a,b,source):out[ix]=a;continue
                repl=self._replace_block_fallback(a,b,source)
                out[ix]=repl[0] if repl else a
            self.track();return out

        # Structural additions/deletions still need block-level revisions. Equal
        # edge blocks are trimmed first so they are not unnecessarily redlined.
        left=0
        while left<min(len(old),len(new)) and equivalent_plain_block(old[left],new[left]):left+=1
        right=0
        while right<min(len(old)-left,len(new)-left) and equivalent_plain_block(old[-1-right],new[-1-right]):right+=1
        if left or right:
            middle_old=old[left:len(old)-right if right else len(old)]
            middle_new=new[left:len(new)-right if right else len(new)]
            middle=[]
            if middle_old:
                if len(middle_old)==len(middle_new) and all(a.tag==b.tag for a,b in zip(middle_old,middle_new)):
                    middle=self.replace(middle_old,middle_new,author=author,source=source)
                else:
                    anchor=middle_old[0];inserted=[]
                    for srcel in middle_new:
                        clone=self._prepared_source(srcel,source);self._mark(clone,'ins')
                        parent.insert(list(parent).index(anchor),clone);inserted.append(clone)
                    for e in middle_old:self._mark(e,'del')
                    middle=inserted
            elif middle_new:
                anchor=old[len(old)-right] if right else None;inserted=[]
                for srcel in middle_new:
                    clone=self._prepared_source(srcel,source);self._mark(clone,'ins')
                    if anchor is None:parent.append(clone)
                    else:parent.insert(list(parent).index(anchor),clone)
                    inserted.append(clone)
                middle=inserted
            self.track();return list(old[:left])+middle+(list(old[len(old)-right:]) if right else [])

        # Last resort for a genuinely restructured range.
        anchor=old[0];inserted=[]
        for srcel in new:
            clone=self._prepared_source(srcel,source);self._mark(clone,'ins')
            try:idx=list(parent).index(anchor)
            except ValueError:raise ValueError('替换锚点已失效。')
            parent.insert(idx,clone);inserted.append(clone)
        for e in old:self._mark(e,'del')
        if comment and inserted:self.comment(inserted[0],comment,author)
        self.track();return inserted

    def insert_source(self,anchor,new,before=False,author='柒',source=None):
        """Paste new source blocks with insertion revisions at an existing anchor.

        This adds corresponding material missing from the workpaper. It never
        deletes or re-marks the anchor and therefore Reject All restores the
        original document. Semantic topic/heading authorization belongs to the
        planner; this writer enforces native block, object and layout safety.
        """
        new=list(new)
        if not new:raise ValueError('新增来源区间为空。')
        if source is None:raise ValueError('新增内容缺少正式来源文档。')
        parents=self._parents_cached();parent=parents.get(anchor)
        if parent is not self.body or anchor.tag not in (w('p'),w('tbl')):
            raise ValueError('新增内容必须定位到原始正文段落或表格，不能插入未知容器。')
        if not before and any(e.tag==w('sectPr') for e in anchor.iter()):
            raise ValueError('新增锚点包含原生节分界，不能在其后跨节粘贴。')
        if any(e.tag not in (w('p'),w('tbl')) for e in new):
            raise ValueError('新增来源只能包含完整段落或表格。')
        if any(e.tag==w('sectPr') for block in new for e in block.iter()):
            raise ValueError('新增来源范围包含原生节分界，不能改变底稿节布局。')
        reason=complex_reference_reason(new) or native_math_copy_reason(self,source,new)
        if reason:raise ValueError(reason)
        # A floating table without the proven original/source anchor pair is
        # outside the supported layout contract. Do not silently convert it.
        from .layout_policy import floating_table_anchor_reason
        for block in new:
            for table in block.iter(w('tbl')):
                reason=floating_table_anchor_reason([],new,None,table)
                if reason:raise ValueError('新增表格布局待核对：'+reason)
        anchored=self._comment_covered_cached()
        if id(anchor) in anchored or any(id(ch) in anchored for ch in anchor.iter()):
            raise ValueError('新增锚点含原有人工批注，不能改变批注范围。')
        try:insert_at=list(parent).index(anchor)+(0 if before else 1)
        except ValueError:raise ValueError('新增锚点已失效。')
        target_width=self._target_section_width(anchor)
        self.author=author
        # Prepare every source block before changing the body's block sequence.
        # This also ensures unsupported later objects cannot leave a partial paste.
        inserted=[]
        for srcel in new:
            clone=self._prepared_source(srcel,source)
            if clone.tag==w('tbl') or clone.find('.//'+w('tbl')) is not None:
                self.fit_new_tables(clone,width=target_width)
            self._mark(clone,'ins');inserted.append(clone)
        for offset,clone in enumerate(inserted):parent.insert(insert_at+offset,clone)
        self._parents_cache=None;self._comment_covered_cache=None
        self.track();return inserted

    def comment(self,el,message,author='柒'):
        p=el if el.tag==w('p') else el.find('.//'+w('p'))
        if p is None:return
        cr=xml(self.entries.get('word/comments.xml',f'<w:comments xmlns:w="{W}"/>'.encode()))
        cid=max([int(c.get(w('id'),'0')) for c in cr if c.get(w('id'),'').isdigit()]+[-1])+1
        c=E.SubElement(cr,w('comment'));c.set(w('id'),str(cid));c.set(w('author'),author);c.set(w('date'),self.date)
        cp=E.SubElement(c,w('p'));rr=E.SubElement(cp,w('r'));tt=E.SubElement(rr,w('t'));tt.text=message
        start=E.Element(w('commentRangeStart'));start.set(w('id'),str(cid));p.insert(1 if p.find(w('pPr')) is not None else 0,start)
        end=E.SubElement(p,w('commentRangeEnd'));end.set(w('id'),str(cid));run=E.SubElement(p,w('r'));ref=E.SubElement(run,w('commentReference'));ref.set(w('id'),str(cid))
        self.entries['word/comments.xml']=dump(cr)
        self.add_rel(R+'/comments','comments.xml');self.content_type('word/comments.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
    def _import_part(self,src,part):
        key=(src.hash,part)
        if key in self._import_cache:return self._import_cache[key]
        if part not in src.entries:raise ValueError('来源附件缺失：'+part)
        directory=posixpath.dirname(part);base=posixpath.basename(part);dst=directory+'/wb_'+src.hash[:10]+'_'+base
        self._import_cache[key]=dst;self.entries[dst]=src.entries[part]
        sct=xml(src.entries['[Content_Types].xml']);typ=next((e.get('ContentType') for e in sct if e.get('PartName')=='/'+part),None)
        if not typ:
            ext=base.rsplit('.',1)[-1];typ=next((e.get('ContentType') for e in sct if e.get('Extension')==ext),None)
        if typ:self.content_type(dst,typ)
        relpart=directory+'/_rels/'+base+'.rels'
        if relpart in src.entries:
            root=xml(src.entries[relpart])
            for r in root:
                if r.get('TargetMode')=='External':continue
                target=r.get('Target','');child=posixpath.normpath(posixpath.join(directory,target)).lstrip('/') if not target.startswith('/') else target.lstrip('/')
                if child not in src.entries:raise ValueError('来源附件关系缺失：'+child)
                r.set('Target',posixpath.relpath(self._import_part(src,child),posixpath.dirname(dst)))
            self.entries[posixpath.dirname(dst)+'/_rels/'+posixpath.basename(dst)+'.rels']=dump(root)
        return dst
    def _import_style(self,src,sid):
        key=(src.hash,sid)
        if key in self._style_map:return self._style_map[key]
        if 'word/styles.xml' not in src.entries:return sid
        sr=xml(src.entries['word/styles.xml']);style=next((s for s in sr if s.get(w('styleId'))==sid),None)
        if style is None:return sid
        if self._styles is None:self._styles=xml(self.entries.get('word/styles.xml',f'<w:styles xmlns:w="{W}"/>'.encode()))
        old=next((s for s in self._styles if s.get(w('styleId'))==sid),None)
        if old is not None and E.tostring(old)==E.tostring(style):self._style_map[key]=sid;return sid
        new_id='wb'+src.hash[:8]+'_'+sid;self._style_map[key]=new_id
        clone=copy.deepcopy(style);clone.set(w('styleId'),new_id);clone.attrib.pop(w('default'),None)
        for t in ['basedOn','next','link']:
            e=clone.find(w(t))
            if e is not None:e.set(w('val'),self._import_style(src,e.get(w('val'))))
        for num in clone.findall('.//'+w('numId')):num.set(w('val'),self._import_num(src,num.get(w('val'))))
        self._styles.append(clone);return new_id
    def _import_num(self,src,nid):
        if nid=='0':return '0'
        key=(src.hash,nid)
        if key in self._num_map:return self._num_map[key]
        if 'word/numbering.xml' not in src.entries:return nid
        sr=xml(src.entries['word/numbering.xml']);num=next((n for n in sr if n.tag==w('num') and n.get(w('numId'))==nid),None)
        if num is None:return nid
        if self._numbering is None:self._numbering=xml(self.entries.get('word/numbering.xml',f'<w:numbering xmlns:w="{W}"/>'.encode()))
        new_id=str(max([int(n.get(w('numId'),'0')) for n in self._numbering.findall(w('num'))]+[0])+1)
        self._num_map[key]=new_id;clone=copy.deepcopy(num);clone.set(w('numId'),new_id)
        aid=clone.find(w('abstractNumId'));abstract=next((n for n in sr if n.tag==w('abstractNum') and aid is not None and n.get(w('abstractNumId'))==aid.get(w('val'))),None)
        if abstract is not None:
            ac=copy.deepcopy(abstract);new_a=str(max([int(n.get(w('abstractNumId'),'0')) for n in self._numbering.findall(w('abstractNum'))]+[-1])+1)
            ac.set(w('abstractNumId'),new_a);aid.set(w('val'),new_a)
            for sty in ac.findall('.//'+w('pStyle')):sty.set(w('val'),self._import_style(src,sty.get(w('val'))))
            ix=next((i for i,n in enumerate(self._numbering) if n.tag==w('num')),len(self._numbering));self._numbering.insert(ix,ac)
        self._numbering.append(clone);self.add_rel(R+'/numbering','numbering.xml');self.content_type('word/numbering.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
        return new_id
    def _note_specs(self):
        return {
            'footnote':('word/footnotes.xml','footnoteReference','footnote',R+'/footnotes','application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'),
            'endnote':('word/endnotes.xml','endnoteReference','endnote',R+'/endnotes','application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'),
        }
    def _ensure_note_root(self,kind,src=None):
        part,refname,notename,reltype,ctype=self._note_specs()[kind]
        if part in self.entries:
            return xml(self.entries[part])
        # Build the target note part from the source's required special separator
        # notes when available.  This is safer than inventing Word-specific separator
        # markup and keeps note rendering identical to the source application.
        root=E.Element(w(kind+'s'))
        if src is not None and part in src.entries:
            sr=xml(src.entries[part])
            for note in sr.findall(w(notename)):
                nid=note.get(w('id'))
                if nid in ('-1','0'):
                    root.append(copy.deepcopy(note))
        self.entries[part]=dump(root)
        self.add_rel(reltype,part.split('/',1)[1])
        self.content_type(part,ctype)
        return root
    def _write_note_root(self,kind,root):
        part=self._note_specs()[kind][0]
        self.entries[part]=dump(root)
    def _fresh_note_id(self,root,notename):
        used=set()
        for note in root.findall(w(notename)):
            raw=note.get(w('id'))
            if raw and re.fullmatch(r'-?\d+',raw):
                val=int(raw)
                if val>0:used.add(val)
        n=max(used or {0})+1
        while n in used:n+=1
        return n
    def _import_note_references(self,el,src):
        """Import referenced note bodies with fresh ids for a pasted block.

        A tracked delete+paste keeps OLD and NEW content in the package at the same
        time.  Reusing one positive footnote/endnote id on both sides creates two
        references to one note body.  WPS may render it, but Microsoft Word can reject
        the package or show a broken note.  Every pasted reference therefore receives
        a fresh id and its own cloned note body.
        """
        for kind,(part,refname,notename,reltype,ctype) in self._note_specs().items():
            refs=list(el.iter(w(refname)))
            if not refs:continue
            if src is None or part not in src.entries:
                raise ValueError(('脚注' if kind=='footnote' else '尾注')+'引用缺少来源注释正文，已阻止复制。')
            srcroot=xml(src.entries[part])
            srcnotes={n.get(w('id')):n for n in srcroot.findall(w(notename))}
            target=self._ensure_note_root(kind,src)
            # Notes containing images/hyperlinks have their own relationships part.
            # Do not silently copy them without remapping; keep this hard and explicit.
            for ref in refs:
                oldid=ref.get(w('id'))
                note=srcnotes.get(oldid)
                if note is None:
                    raise ValueError(('脚注' if kind=='footnote' else '尾注')+'编号 '+str(oldid)+' 在来源中不存在，已阻止复制。')
                if any(k.startswith('{'+R+'}') for n in note.iter() for k in n.attrib):
                    raise ValueError(('脚注' if kind=='footnote' else '尾注')+'包含图片、超链接或外部关系，当前版本保留原文并要求人工处理。')
                clone=copy.deepcopy(note)
                # Import style/numbering references used by the note body itself.
                for styletag in ('pStyle','rStyle','tblStyle'):
                    for e in clone.findall('.//'+w(styletag)):
                        e.set(w('val'),self._import_style(src,e.get(w('val'))))
                for e in clone.findall('.//'+w('numId')):
                    e.set(w('val'),self._import_num(src,e.get(w('val'))))
                new_id=self._fresh_note_id(target,notename)
                clone.set(w('id'),str(new_id))
                target.append(clone)
                ref.set(w('id'),str(new_id))
            self._write_note_root(kind,target)
    def _deduplicate_note_references(self):
        """Repair legacy duplicate positive note ids already present in a redline."""
        for kind,(part,refname,notename,reltype,ctype) in self._note_specs().items():
            refs=list(self.root.iter(w(refname)))
            if not refs:continue
            if part not in self.entries:
                continue
            root=xml(self.entries[part])
            notes={n.get(w('id')):n for n in root.findall(w(notename))}
            seen=set();changed=False
            for ref in refs:
                rid=ref.get(w('id'))
                if rid not in seen:
                    seen.add(rid);continue
                note=notes.get(rid)
                if note is None:continue
                new_id=self._fresh_note_id(root,notename)
                clone=copy.deepcopy(note);clone.set(w('id'),str(new_id));root.append(clone)
                notes[str(new_id)]=clone;ref.set(w('id'),str(new_id));seen.add(str(new_id));changed=True
            if changed:self._write_note_root(kind,root)
    def import_block(self,el,src):
        reason=native_math_copy_reason(self,src,[el])
        if reason:raise ValueError(reason)
        self._import_note_references(el,src)
        reason=complex_reference_reason([el])
        if reason:raise ValueError(reason)
        # A source paragraph without pStyle inherits the source's Normal style.
        # Explicitly bind that imported style for math paragraphs so the target's
        # different Normal/theme cannot alter an otherwise verbatim equation.
        if any(str(e.tag).startswith('{'+M+'}') for e in el.iter()):
            sr=xml(src.entries.get('word/styles.xml',f'<w:styles xmlns:w="{W}"/>'.encode()))
            normal=next((s for s in sr.findall(w('style')) if s.get(w('type'))=='paragraph' and s.get(w('default')) in ('1','true','on')),None)
            for para in el.iter(w('p')):
                if not any(str(e.tag).startswith('{'+M+'}') for e in para.iter()):continue
                pr=para.find(w('pPr'))
                if normal is not None and (pr is None or pr.find(w('pStyle')) is None):
                    if pr is None:pr=E.Element(w('pPr'));para.insert(0,pr)
                    ref=E.Element(w('pStyle'));ref.set(w('val'),normal.get(w('styleId')));pr.insert(0,ref)
        parents=_parents(el)
        for e in list(el.iter()):
            if local(e) in ('bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd','commentReference','sectPr','proofErr'):_remove_with(parents,e)
        styles=[]
        for tag in ('pStyle','rStyle','tblStyle'):styles.extend(el.findall('.//'+w(tag)))
        for e in styles:e.set(w('val'),self._import_style(src,e.get(w('val'))))
        for e in el.findall('.//'+w('numId')):e.set(w('val'),self._import_num(src,e.get(w('val'))))
        rels={r.get('Id'):r for r in src.rels()}
        for e in el.iter():
            if e.tag=='{'+WP+'}docPr':e.set('id',str(self.drawid));self.drawid+=1
            for key,value in list(e.attrib.items()):
                if key.startswith('{'+R+'}'):
                    rel=rels.get(value)
                    if rel is None:raise ValueError('来源关系无法解析：'+value)
                    target=rel.get('Target');mode=rel.get('TargetMode')
                    if mode=='External':new_target=target
                    else:
                        part=posixpath.normpath(posixpath.join('word',target)) if not target.startswith('/') else target.lstrip('/')
                        new_target=posixpath.relpath(self._import_part(src,part),'word')
                    e.set(key,self.add_rel(rel.get('Type'),new_target,mode))
        # Table geometry is matched to the actual target during replacement.
    def _match_table_geometry(self,old_tbl,new_tbl):
        """Project target placement/column geometry onto a pasted replacement table.

        This prevents source tables from inheriting a different section width and
        becoming visibly narrower or shifted in WPS. Only geometry is projected;
        source text and visual cell styling remain unchanged.
        """
        if old_tbl is None or new_tbl is None or old_tbl.tag!=w('tbl') or new_tbl.tag!=w('tbl'):
            return False
        old_grid=old_tbl.find(w('tblGrid'));new_grid=new_tbl.find(w('tblGrid'))
        if old_grid is None or new_grid is None:return False
        old_cols=old_grid.findall(w('gridCol'));new_cols=new_grid.findall(w('gridCol'))
        if not old_cols or len(old_cols)!=len(new_cols):return False
        try:widths=[max(1,int(c.get(w('w'),'0'))) for c in old_cols]
        except Exception:return False
        if not all(widths):return False

        old_pr=old_tbl.find(w('tblPr'));new_pr=new_tbl.find(w('tblPr'))
        if new_pr is None:
            new_pr=E.Element(w('tblPr'));new_tbl.insert(0,new_pr)
        geometry=('tblW','tblInd','jc','tblLayout','tblCellSpacing','tblpPr')
        for name in geometry:
            cur=new_pr.find(w(name))
            if cur is not None:new_pr.remove(cur)
            src=old_pr.find(w(name)) if old_pr is not None else None
            if src is not None:new_pr.append(copy.deepcopy(src))

        pos=list(new_tbl).index(new_grid)
        new_tbl.remove(new_grid);new_tbl.insert(pos,copy.deepcopy(old_grid))

        for tr in new_tbl.findall(w('tr')):
            trpr=tr.find(w('trPr'))
            before=trpr.find(w('gridBefore')) if trpr is not None else None
            try:col=int(before.get(w('val'),'0')) if before is not None else 0
            except Exception:col=0
            for tc in tr.findall(w('tc')):
                tcpr=tc.find(w('tcPr'))
                if tcpr is None:
                    tcpr=E.Element(w('tcPr'));tc.insert(0,tcpr)
                gs=tcpr.find(w('gridSpan'))
                try:span=max(1,int(gs.get(w('val'),'1'))) if gs is not None else 1
                except Exception:span=1
                end=min(len(widths),col+span)
                if col<len(widths) and end>col:
                    tcw=tcpr.find(w('tcW'))
                    if tcw is None:
                        tcw=E.Element(w('tcW'));tcpr.insert(0,tcw)
                    tcw.set(w('w'),str(sum(widths[col:end])));tcw.set(w('type'),'dxa')
                col=end
        return True

    def _target_section_width(self,anchor):
        return target_section_width(self.root,anchor)

    def fit_new_tables(self,el,width=None,target=None):
        """Shrink pasted real tables to the target section without changing grids."""
        if width is None:width=self._target_section_width(target if target is not None else el)
        for tbl in el.iter(w('tbl')):
            # Keep each pasted row readable across page boundaries. This is a
            # recorded table-layout adjustment; old/protected rows stay untouched.
            # Very tall rows and multi-row headings still require visual review.
            for row in tbl.findall(w('tr')):
                rp=row.find(w('trPr'))
                if rp is None:rp=E.Element(w('trPr'));row.insert(0,rp)
                marker=rp.find(w('cantSplit'))
                if marker is None:
                    marker=E.Element(w('cantSplit'))
                    later={w(n) for n in ('trHeight','tblHeader','tblCellSpacing','jc','hidden','ins','del','trPrChange')}
                    rp.insert(next((i for i,n in enumerate(rp) if n.tag in later),len(rp)),marker)
                marker.set(w('val'),'1')
            grid=tbl.find(w('tblGrid'))
            if grid is None:continue
            cols=grid.findall(w('gridCol'))
            values=[int(c.get(w('w'),'0')) for c in cols];total=sum(values)
            if not cols or not total:continue
            if any(v<=0 for v in values):raise ValueError('来源表格网格宽度无效，保留待核对。')
            pr=tbl.find(w('tblPr'));indent=pr.find(w('tblInd')) if pr is not None else None
            allowance=width-max(0,int(indent.get(w('w'),'0'))) if indent is not None and indent.get(w('type'))=='dxa' else width
            if allowance<len(cols):raise ValueError('目标表格可用宽度不足，保留待核对。')
            if total<=allowance:continue
            factor=allowance/total
            scaled=[max(1,int(v*factor)) for v in values]
            scaled[-1]+=allowance-sum(scaled)
            for col,value in zip(cols,scaled):col.set(w('w'),str(value))
            if pr is not None:
                tablew=pr.find(w('tblW'))
                if tablew is not None and tablew.get(w('type'))=='dxa':tablew.set(w('w'),str(allowance))
            for row in tbl.findall(w('tr')):
                before=row.find('./'+w('trPr')+'/'+w('gridBefore'))
                index=int(before.get(w('val'),'0')) if before is not None else 0
                for cell in row.findall(w('tc')):
                    cp=cell.find(w('tcPr'))
                    if cp is None:cp=E.Element(w('tcPr'));cell.insert(0,cp)
                    span=cp.find(w('gridSpan'));count=int(span.get(w('val'),'1')) if span is not None else 1
                    if count<1 or index+count>len(scaled):raise ValueError('来源表格跨列结构超出网格，保留待核对。')
                    cw=cp.find(w('tcW'))
                    if cw is None:cw=E.SubElement(cp,w('tcW'))
                    cw.set(w('type'),'dxa');cw.set(w('w'),str(sum(scaled[index:index+count])));index+=count


def transform_paragraph(p,mapping):
    """Replace across fragmented runs without flattening unrelated formatting."""
    for old,new in mapping:
        nodes=[t for t in p.iter(w('t'))];combined=''.join(t.text or '' for t in nodes)
        positions=[];cur=0
        for t in nodes:positions.append((cur,cur+len(t.text or ''),t));cur+=len(t.text or '')
        for m in reversed(list(re.finditer(re.escape(old),combined))):
            first=True
            for start,end,t in positions:
                if end<=m.start() or start>=m.end():continue
                value=t.text or '';a=max(0,m.start()-start);b=min(end-start,m.end()-start)
                t.text=value[:a]+(new if first else '')+value[b:];first=False
                t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    return p

def transformed(el,document_kind='workpaper'):
    el=copy.deepcopy(el)
    term='本核查意见出具日' if document_kind=='opinion' else '本核查分析文件出具日'
    for p in el.iter(w('p')):transform_paragraph(p,[('本募集说明书签署之日',term),('本募集说明书签署日',term),('本公司','发行人')])
    return el

def semantic_blocks(pkg):
    return [(local(e),text(e).strip()) for e in pkg.body if e.tag in (w('p'),w('tbl')) and (text(e).strip() or e.findall('.//'+w('drawing')))]

def target_section_width(root,anchor):
    """Use the section ending at/after this target, never the final section."""
    body_root=root.find(w('body'))
    parents=_parents(root);top=anchor
    while top in parents and parents[top] is not body_root:top=parents[top]
    if top not in list(body_root):raise ValueError('无法定位目标表格所属节，已停止宽度调整。')
    body=list(body_root);start=body.index(top);sect=None
    for block in body[start:]:
        if block.tag==w('p'):
            sect=block.find('./'+w('pPr')+'/'+w('sectPr'))
        elif block.tag==w('sectPr'):sect=block
        if sect is not None:break
    width=9360
    if sect is not None:
        size=sect.find(w('pgSz'));margin=sect.find(w('pgMar'))
        page=int(size.get(w('w'),'12240')) if size is not None else 12240
        left=int(margin.get(w('left'),'1440')) if margin is not None else 1440
        right=int(margin.get(w('right'),'1440')) if margin is not None else 1440
        gutter=int(margin.get(w('gutter'),'0')) if margin is not None else 0
        width=page-left-right-gutter
        cols=sect.find(w('cols'))
        if cols is not None:
            explicit=cols.findall(w('col'))
            if explicit:width=min(int(c.get(w('w'),str(width))) for c in explicit)
            else:
                count=max(1,int(cols.get(w('num'),'1')))
                width=(width-(count-1)*int(cols.get(w('space'),'720')))//count
    # For layout tables, the containing cell is a tighter bound than the page.
    cur=anchor
    while cur in parents:
        cur=parents[cur]
        if cur.tag==w('tc'):
            cellw=cur.find('./'+w('tcPr')+'/'+w('tcW'))
            if cellw is not None and cellw.get(w('type'))=='dxa':
                width=min(width,int(cellw.get(w('w'),str(width))))
    if width<=0:raise ValueError('目标节可用宽度无效，已停止表格调整。')
    return width


def revision_shape_issues(root):
    """Static checks for WPS-native persisted tracked-change bookkeeping."""
    issues=[];seen=set();W14='http://schemas.microsoft.com/office/word/2010/wordml'
    revtags={w('ins'),w('del'),w('moveFrom'),w('moveTo'),w('pPrChange'),w('rPrChange'),w('tblPrChange'),w('trPrChange'),w('tcPrChange'),w('tblGridChange')}
    noisy={w('rsidR'),w('rsidRPr'),w('rsidRDefault'),w('rsidP'),w('rsidDel'),w('rsidTr'),'{'+W14+'}textId'}
    # WPS navigation requires unique paragraph/table-row identities.  The source
    # golden file has no duplicates and no rsid/textId metadata in document.xml.
    para_ids=[]
    for node in root.iter():
        if node.tag in (w('p'),w('tr')):
            pid=node.get('{'+W14+'}paraId')
            valid_pid=bool(pid and re.fullmatch(r'[0-9A-Fa-f]{8}',pid or ''))
            if valid_pid:
                value=int(pid,16)
                valid_pid=(0 < value < 0x80000000)
            if not valid_pid:
                issues.append('段落或表格行的 w14:paraId 超出 Microsoft Word 兼容范围（00000001-7FFFFFFF）。')
            else:
                para_ids.append(pid.upper())
        for key in node.attrib:
            if key=='{'+W14+'}textId' or (key.startswith('{'+W+'}') and key.rsplit('}',1)[-1].startswith('rsid')):
                issues.append('document.xml 残留 rsid/textId 导航会话元数据。')
                break
    dup_para=[pid for pid,c in __import__('collections').Counter(para_ids).items() if c>1]
    if dup_para:issues.append('w14:paraId 重复，WPS 拒绝修订可能丢失当前位置：'+dup_para[0])
    ids=[]
    for e in root.iter():
        if e.tag in revtags:
            rid=e.get(w('id'))
            if rid is None or not rid.isdigit():issues.append('修订缺少数字 w:id。')
            else:
                if rid in seen:issues.append('同一修订 w:id 被重复使用：'+rid)
                else:seen.add(rid);ids.append(int(rid))
    if ids and ids!=list(range(len(ids))):
        issues.append('逻辑修订 w:id 未按正文顺序连续编号。')
    for p in root.iter(w('p')):
        if any(e.tag in revtags for e in p.iter()) and any(k in noisy for node in p.iter() for k in node.attrib):
            issues.append('修订段落残留 rsid/textId 元数据。')
    for p in root.iter(w('p')):
        ch=list(p)
        for i in range(len(ch)-1):
            a,b=ch[i],ch[i+1]
            if a.tag==w('del') and b.tag==w('ins'):
                issues.append('同一段落内同时存在删除和插入；整段替换应拆成两个独立段落。')
                if a.get(w('author'))!=b.get(w('author')) or a.get(w('date'))!=b.get(w('date')):
                    issues.append('相邻替换的删除/插入作者或时间不一致。')

    # A native WPS whole-table replacement is not row-marker-only.  Each revised
    # row also carries revised paragraph marks and revised cell text.  Without that
    # payload WPS shows an empty review balloon (“删除:” with nothing after it).
    for tr in root.iter(w('tr')):
        trpr=tr.find(w('trPr'))
        if trpr is None:continue
        for kind in ('del','ins'):
            marker=trpr.find(w(kind))
            if marker is None:continue
            paras=tr.findall('.//'+w('p'))
            for p in paras:
                ppr=p.find(w('pPr'));rpr=ppr.find(w('rPr')) if ppr is not None else None
                if rpr is None or rpr.find(w(kind)) is None:
                    issues.append('整表修订存在仅行级标记、缺少单元格段落修订：'+kind)
                    break
            has_visible=any((t.text or '') for t in tr.iter(w('t'))) or any((t.text or '') for t in tr.iter(w('delText')))
            if has_visible:
                if kind=='del' and not any((t.text or '') for t in tr.iter(w('delText'))):
                    issues.append('整表删除修订没有实际删除文字，WPS 审阅栏会显示空内容。')
                if kind=='ins':
                    payload=False
                    for ins in tr.iter(w('ins')):
                        if ins is marker:continue
                        if any((t.text or '') for t in ins.iter(w('t'))):payload=True;break
                    if not payload:issues.append('整表插入修订没有实际插入文字，WPS 审阅栏会显示空内容。')

    # One export should look like one WPS editing session: one author/date stream
    # with unique ids.  Historical mixed-author files are allowed by the generic
    # validator, but a single-author generated review must not have dozens of dates.
    authors={e.get(w('author')) for e in root.iter() if e.tag in revtags and e.get(w('author'))}
    dates={e.get(w('date')) for e in root.iter() if e.tag in revtags and e.get(w('date'))}
    if len(authors)<=1 and len(dates)>1:
        issues.append('单一修订作者对应多个修订时间批次；应合并为一次 WPS 审阅会话。')

    body=root.find(w('body'))
    if body is not None:
        blocks=list(body)
        def _rev_kind(tbl,kind):
            for tr in tbl.findall(w('tr')):
                pr=tr.find(w('trPr'))
                if pr is not None and pr.find(w(kind)) is not None:return True
            return False
        def _grid_widths(tbl):
            g=tbl.find(w('tblGrid'))
            if g is None:return []
            try:return [int(c.get(w('w'),'0')) for c in g.findall(w('gridCol'))]
            except Exception:return []
        def _table_width(tbl):
            e=tbl.find('./'+w('tblPr')+'/'+w('tblW'))
            if e is not None and e.get(w('type'))=='dxa':
                try:return int(e.get(w('w'),'0'))
                except Exception:pass
            return sum(_grid_widths(tbl))
        for i in range(len(blocks)-1):
            a,b=blocks[i],blocks[i+1]
            if a.tag!=w('tbl') or b.tag!=w('tbl'):continue
            if not _rev_kind(a,'del') or not _rev_kind(b,'ins'):continue
            ga,gb=_grid_widths(a),_grid_widths(b)
            if ga and len(ga)==len(gb):
                wa,wb=_table_width(a),_table_width(b)
                # A pre-existing oversized table may legitimately shrink to the
                # target section. Still reject arbitrary shrink/growth, including
                # a width borrowed from the document's final section.
                indent=b.find('./'+w('tblPr')+'/'+w('tblInd'))
                inset=max(0,int(indent.get(w('w'),'0'))) if indent is not None and indent.get(w('type'))=='dxa' else 0
                from .layout_policy import preserved_native_table_layout
                native=preserved_native_table_layout(root,a,b)
                expected=wa if native else min(wa,target_section_width(root,b)-inset)
                if expected and wb and abs(wb-expected)/expected>0.01:
                    issues.append('整表替换宽度不符合原表宽度及目标节可用宽度，可能出现缩窄或越界。')
                ja=a.find('./'+w('tblPr')+'/'+w('jc'));jb=b.find('./'+w('tblPr')+'/'+w('jc'))
                va=ja.get(w('val'),'') if ja is not None else '';vb=jb.get(w('val'),'') if jb is not None else ''
                if va!=vb:issues.append('整表替换前后表格对齐方式不一致。')
    return sorted(set(issues))

def validate_package(path):
    issues=[]
    try:
        p=Package(path)
        for name,data in p.entries.items():
            if name.endswith(('.xml','.rels')):
                root=xml(data)
                if name.startswith('word/') and name.endswith('.xml'):
                    from .compatibility import property_order_issues
                    issues.extend(name+'：'+issue for issue in property_order_issues(root))
                    if root.tag==w('styles') and len(root.findall(w('style')))>4079:
                        issues.append(name+'：样式定义超过 Word 4079 上限。')
                if name.endswith('.rels'):
                    base=posixpath.dirname(posixpath.dirname(name))
                    for r in root:
                        if r.get('TargetMode')=='External':continue
                        target=r.get('Target','');dest=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base,target))
                        if dest not in p.entries:issues.append('缺少内部关系目标：'+dest)
        relids={r.get('Id') for r in p.rels()}
        for e in p.root.iter():
            for k,v in e.attrib.items():
                if k.startswith('{'+R+'}') and v not in relids:issues.append('正文关系缺失：'+v)
        for tr in p.root.iter(w('tr')):
            if not tr.findall(w('tc')):issues.append('表格行没有单元格。')
        for tc in p.root.iter(w('tc')):
            if not any(c.tag in (w('p'),w('tbl')) for c in tc):issues.append('表格单元格没有内容容器。')
        # Broken/duplicated bookmark ids can also destabilize WPS review navigation.
        # Names may legitimately repeat, ids may not; every start must have one end.
        import collections as _collections
        bs=[e.get(w('id')) for e in p.root.iter(w('bookmarkStart')) if e.get(w('id')) is not None]
        be=[e.get(w('id')) for e in p.root.iter(w('bookmarkEnd')) if e.get(w('id')) is not None]
        if any(c>1 for c in _collections.Counter(bs).values()):issues.append('书签起点 w:id 重复。')
        if _collections.Counter(bs)!=_collections.Counter(be):issues.append('书签起止锚点不成对。')
        # Both reviewer outcomes must also preserve a valid bookmark graph.  A
        # bookmark endpoint hidden inside only the deleted side of a replacement can
        # look fine in the redline yet become dangling after Accept, which is another
        # source of navigation jumps.
        for _accept,_label in ((False,'拒绝'),(True,'接受')):
            rr=copy.deepcopy(p.root);resolve_revisions(rr,_accept)
            aa=[e.get(w('id')) for e in rr.iter(w('bookmarkStart')) if e.get(w('id')) is not None]
            bb=[e.get(w('id')) for e in rr.iter(w('bookmarkEnd')) if e.get(w('id')) is not None]
            if any(c>1 for c in _collections.Counter(aa).values()) or _collections.Counter(aa)!=_collections.Counter(bb):
                issues.append(_label+'全部修订后书签锚点不完整。')
        # Word requires a one-to-one mapping between visible positive note
        # references and note bodies.  Duplicate reference ids are a known source of
        # Word repair/open failures even when WPS/LibreOffice continue rendering.
        for _kind,_part,_ref,_note in (
            ('脚注','word/footnotes.xml','footnoteReference','footnote'),
            ('尾注','word/endnotes.xml','endnoteReference','endnote')):
            refs=[e.get(w('id')) for e in p.root.iter(w(_ref)) if e.get(w('id')) not in (None,'-1','0')]
            if any(c>1 for c in _collections.Counter(refs).values()):
                issues.append(_kind+'引用 w:id 重复，Microsoft Word 可能拒绝打开。')
            if refs:
                if _part not in p.entries:
                    issues.append(_kind+'引用存在但 '+_part+' 缺失。')
                else:
                    nr=xml(p.entries[_part]);defined=[e.get(w('id')) for e in nr.iter(w(_note))]
                    if any(c>1 for c in _collections.Counter(defined).values()):issues.append(_kind+'正文 w:id 重复。')
                    missing=[x for x in refs if x not in set(defined)]
                    if missing:issues.append(_kind+'引用缺少正文编号：'+missing[0])
        issues.extend(revision_shape_issues(p.root))
    except Exception as exc:issues.append(str(exc))
    return sorted(set(issues))


def comment_covered_elements(root):
    """All elements intersecting an existing comment, including range middles."""
    covered=set();active=set()
    def walk(node):
        if node.tag==w('commentRangeStart'):active.add(node.get(w('id')))
        hit=bool(active) or node.tag in (w('commentRangeStart'),w('commentRangeEnd'),w('commentReference'))
        for child in node:
            if walk(child):hit=True
        if node.tag==w('commentRangeEnd'):active.discard(node.get(w('id')))
        if hit:covered.add(id(node))
        return hit
    walk(root);return covered


def structural_digest_content(root):
    """Structural digest for Reject-All content fidelity, ignoring bookmark position.

    Whole-block tracked replacement may lift bookmark endpoints to persistent body
    siblings so Accept All never leaves a dangling navigation anchor.  That movement
    is intentionally invisible and does not alter document content.  Compare all
    other structure byte-semantically while validating bookmark pairing separately.
    """
    r=copy.deepcopy(root)
    parents=_parents(r)
    for node in list(r.iter()):
        if node.tag in (w('bookmarkStart'),w('bookmarkEnd')):
            par=parents.get(node)
            if par is not None:
                try:par.remove(node)
                except ValueError:pass
    return structural_digest(r)

def structural_digest(root):
    """Prefix-independent full-tree digest; ignore only empty property containers.

    No text, tables, merging, bookmarks, formula, field, relationship IDs, or
    drawing geometry is excluded. Empty pPr/rPr/trPr are semantically neutral
    artifacts of resolving tracked paragraph/row/font properties. Recognized
    property children are compared in schema order; their values, multiplicity
    and full subtrees still participate in the comparison.
    """
    empty={w('pPr'),w('rPr'),w('trPr')}
    from .compatibility import ordered_property_children
    def tree(e):
        children=tuple(x for c in ordered_property_children(e) if (x:=tree(c)) is not None)
        if e.tag in empty and not children and not e.attrib and not (e.text or '').strip():return None
        return (str(e.tag),tuple(sorted(e.attrib.items())),e.text or '',children)
    import json
    return sha(json.dumps(tree(root),ensure_ascii=False,separators=(',',':')).encode())
