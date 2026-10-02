"""Read-back audit of persisted DOCX files, independent of the writer.

This module deliberately imports neither docxio nor matching/model/precision. XML
child-index paths are frozen against the input and collected again after save.
A source hash alone never proves a copy: actual revisions, logical cell structure,
relationships, media bytes and note bodies are read from the saved ZIP.
"""
from __future__ import annotations
import copy
import hashlib
import json
import posixpath
import re
import zipfile
import unicodedata
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET
from .compatibility import ordered_property_children,property_order_issues
from .wording import POLICY as WORDING_POLICY,mappings as wording_mappings

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
REV_NAMES = {'ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange',
             'tblPrChange', 'trPrChange', 'tcPrChange', 'tblGridChange', 'numPrChange'}
REV_TAGS = {'{' + W + '}' + n for n in REV_NAMES}
NS = {'w': W}


def q(name):
    return '{' + W + '}' + name


def name(node):
    return node.tag.rsplit('}', 1)[-1] if isinstance(node.tag, str) else ''


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(',', ':')).encode()).hexdigest()


def element_paths(root, elements):
    """Register positions only; never use the writer's content interpretation."""
    wanted = {id(e) for e in elements}
    found = {}
    def visit(node, path):
        if id(node) in wanted:
            found[id(node)] = path
        for i, child in enumerate(node):
            visit(child, path + [i])
    visit(root, [])
    if any(id(e) not in found for e in elements):
        raise ValueError('审计对象不在实际文档树中。')
    return [found[id(e)] for e in elements]


def at_path(root, path):
    node = root
    for index in path:
        if type(index) is not int or index < 0:
            raise ValueError('无效的对象路径。')
        node = node[index]
    return node


def revision_ids(elements, kind=None):
    return [e.get(q('id'), '') for node in elements for e in node.iter()
            if e.tag in REV_TAGS and (kind is None or name(e) == kind)]


class Reader:
    def __init__(self, path):
        self.path = Path(path)
        raw = self.path.read_bytes()
        self.sha256 = hashlib.sha256(raw).hexdigest()
        with zipfile.ZipFile(self.path) as z:
            if len(z.namelist()) != len(set(z.namelist())):
                raise ValueError('DOCX 含重复 ZIP 部件名。')
            self.parts = {i.filename: z.read(i) for i in z.infolist() if not i.is_dir()}
        self.root = ET.fromstring(self.parts['word/document.xml'])
        self.relations = {}
        self.dependencies = []
        self._style_cache = {}
        self._style_stack = set()
        self._styles = None

    def rel(self, part, rid):
        rp = posixpath.dirname(part) + '/_rels/' + posixpath.basename(part) + '.rels'
        if rp not in self.relations:
            root = ET.fromstring(self.parts[rp]) if rp in self.parts else []
            rels = {}
            for e in root:
                if e.get('Id') in rels:
                    raise ValueError('重复的关系 ID：' + rp + '/' + e.get('Id', ''))
                rels[e.get('Id')] = e.attrib
            self.relations[rp] = rels
        if rid not in self.relations[rp]:
            raise ValueError('悬空关系：' + part + '/' + str(rid))
        rel = self.relations[rp][rid]
        if rel.get('TargetMode') == 'External':
            return ('external', rel.get('Type'), rel.get('Target'))
        target = rel.get('Target', '')
        absolute = (target.lstrip('/') if target.startswith('/') else
                    posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
        if absolute.startswith('../') or absolute not in self.parts:
            raise ValueError('依赖部件不存在：' + absolute)
        return ('internal', rel.get('Type'), absolute)

    def resource(self, part, seen=()):
        if part in seen:
            raise ValueError('循环对象依赖暂不支持：' + part)
        raw = self.parts[part]
        if part.lower().endswith('.xml'):
            root = ET.fromstring(raw)
            result = self.tree(root, part, seen=seen + (part,))
        else:
            result = ('bytes', hashlib.sha256(raw).hexdigest())
        self.dependencies.append({'part': part, 'sha256': hashlib.sha256(raw).hexdigest(),
                                  'semantic_digest': digest(result)})
        return result

    def tree(self, node, part='word/document.xml', seen=(), ignore_bookmarks=False):
        """Structure including formatting; ignore only explicitly technical IDs."""
        tag = name(node)
        if tag in {'proofErr'} or (ignore_bookmarks and tag in {'bookmarkStart', 'bookmarkEnd'}):
            return None
        attrs = []
        for key, value in sorted(node.attrib.items()):
            local = key.rsplit('}', 1)[-1]
            if key in {'{' + W14 + '}paraId', '{' + W14 + '}textId'}:
                continue
            if key.startswith('{' + W + '}') and local.startswith('rsid'):
                continue
            if tag in {'docPr', 'cNvPr'} and key == 'id':
                continue
            if key.startswith('{' + R + '}'):
                rel = self.rel(part, value)
                value = rel if rel[0] == 'external' else ('internal', rel[1], self.resource(rel[2], seen))
            if tag in {'pStyle', 'rStyle', 'tblStyle', 'basedOn', 'next', 'link'} and key == q('val'):
                value = self.style(value)
            if tag in {'footnoteReference', 'endnoteReference'} and key == q('id'):
                kind = 'footnote' if tag == 'footnoteReference' else 'endnote'
                np = 'word/' + kind + 's.xml'
                if np not in self.parts:
                    raise ValueError('注释部件缺失：' + np)
                notes = ET.fromstring(self.parts[np])
                matches = [x for x in notes if x.get(q('id')) == value]
                if len(matches) != 1:
                    raise ValueError('注释正文缺失或重复：' + np + '/' + value)
                clone = copy.deepcopy(matches[0]); clone.attrib.pop(q('id'), None)
                value = self.tree(clone, np, seen=seen)
                self.dependencies.append({'part': np, 'note_id': node.get(q('id')),
                                          'semantic_digest': digest(value)})
            attrs.append((key, value))
        children = [self.tree(c, part, seen, ignore_bookmarks) for c in ordered_property_children(node)]
        children = [c for c in children if c is not None]
        content = node.text or ''
        # Generated paragraph-mark revisions may leave an empty properties shell.
        if tag in {'pPr', 'rPr', 'trPr', 'tcPr'} and not attrs and not children and not content.strip():
            return None
        return (node.tag, attrs, content if tag in {'t', 'delText', 'instrText', 'delInstrText'} else content.strip(), children)

    def style(self, style_id):
        # Imported styles are renamed. Compare the actual definition and its
        # dependencies, never make a wb... prefix alone evidence of equivalence.
        if style_id in self._style_cache:
            return self._style_cache[style_id]
        if style_id in self._style_stack:
            return ('style-cycle', re.sub(r'^wb[0-9a-f]{8}_', '', style_id))
        if self._styles is None:
            raw = self.parts.get('word/styles.xml')
            self._styles = {} if raw is None else {x.get(q('styleId')): x for x in ET.fromstring(raw)}
        if style_id not in self._styles:
            return ('missing-style', style_id)
        self._style_stack.add(style_id)
        try:
            clone = copy.deepcopy(self._styles[style_id])
            clone.attrib.pop(q('styleId'), None); clone.attrib.pop(q('default'), None)
            value = self.tree(clone, 'word/styles.xml')
            self._style_cache[style_id] = value
            return value
        finally:
            self._style_stack.remove(style_id)

    def content(self, node):
        """Ordered semantic payload, including full tables and object dependencies."""
        tag = name(node)
        if tag == 'tbl':
            grid = node.find(q('tblGrid'))
            rows = []
            for row in node.findall(q('tr')):
                pr = row.find(q('trPr'))
                def rv(key, default='0'):
                    e = None if pr is None else pr.find(q(key))
                    return default if e is None else e.get(q('val'), '1')
                cells = []
                for cell in row.findall(q('tc')):
                    cp = cell.find(q('tcPr'))
                    merges = []
                    for key, default in [('gridSpan', '1'), ('vMerge', ''), ('hMerge', '')]:
                        e = None if cp is None else cp.find(q(key))
                        merges.append(default if e is None else e.get(q('val'), 'continue'))
                    blocks = [self.content(c) for c in cell if name(c) not in {'tcPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr'}]
                    cells.append((merges, blocks))
                rows.append((rv('gridBefore'), rv('gridAfter'), rv('tblHeader', ''), cells))
            return ('table', len(grid) if grid is not None else None, rows)
        if tag == 'p':
            return ('paragraph', self._inline(node))
        return (tag, self.tree(node))

    def _inline(self, node):
        tokens = []
        def push(kind, value):
            if kind == 'text' and tokens and tokens[-1][0] == 'text':
                tokens[-1] = ('text', tokens[-1][1] + value)
            else:
                tokens.append((kind, value))
        def walk(e):
            tag = name(e)
            if tag in {'pPr', 'rPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr'}:
                return
            if tag in {'t', 'delText'}:
                push('text', e.text or '')
            elif tag in {'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen'}:
                push('control', (tag, sorted(e.attrib.items())))
            elif tag in {'drawing', 'pict', 'object', 'footnoteReference', 'endnoteReference',
                         'hyperlink', 'fldSimple', 'sym'} or str(e.tag).startswith('{' + M + '}'):
                push('object', self.tree(e))
            elif tag in {'instrText', 'delInstrText', 'fldChar'}:
                push('field', self.tree(e))
            elif tag in {'commentReference', 'commentRangeStart', 'commentRangeEnd'}:
                push('comment', self.tree(e))
            else:
                for child in e:
                    walk(child)
        for child in node:
            walk(child)
        return tokens


def _resolve(root, accept):
    """Independent resolver for supported block/run revisions.

    Whole inserted/deleted paragraph marks remove that paragraph. Move/cell-range
    revisions are refused by the audit, rather than guessed from visible strings.
    """
    root = copy.deepcopy(root)
    remove = 'del' if accept else 'ins'
    for parent in list(root.iter()):
        for child in list(parent):
            tag = name(child)
            if tag == 'tr' and child.find('./w:trPr/w:' + remove, NS) is not None:
                parent.remove(child)
            elif tag == 'p' and child.find('./w:pPr/w:rPr/w:' + remove, NS) is not None:
                parent.remove(child)
    def clean(parent):
        for child in list(parent):
            tag = name(child)
            if tag in {'ins', 'del'}:
                index = list(parent).index(child)
                parent.remove(child)
                if tag != remove and name(parent) not in {'rPr', 'trPr', 'tcPr', 'numPr'}:
                    for offset, grandchild in enumerate(list(child)):
                        parent.insert(index + offset, grandchild)
                        clean(grandchild)
            elif tag.endswith('PrChange') or tag == 'tblGridChange':
                parent.remove(child)
                if not accept and len(child):
                    old = child[0]
                    parent.attrib.clear(); parent.attrib.update(old.attrib)
                    for sibling in list(parent):
                        parent.remove(sibling)
                    for grandchild in old:
                        parent.append(copy.deepcopy(grandchild))
            else:
                clean(child)
                if child.tag == q('delText'):
                    child.tag = q('t')
                elif child.tag == q('delInstrText'):
                    child.tag = q('instrText')
        for child in list(parent):
            if child.tag == q('tbl') and not child.findall(q('tr')):
                parent.remove(child)
    clean(root)
    return root


def _unsupported(nodes):
    for node in nodes:
        parents={child:parent for parent in node.iter() for child in parent}
        for e in node.iter():
            n = name(e)
            if isinstance(e.tag,str) and e.tag.startswith('{'+M+'}'):
                parent=parents.get(e)
                if parent is None or not (isinstance(parent.tag,str) and parent.tag.startswith('{'+M+'}')):
                    if e.tag not in {'{'+M+'}oMath','{'+M+'}oMathPara'} or parent is None or parent.tag!=q('p'):
                        return '来源公式不是段落中的完整原生对象。'
            if n in {'chart', 'object', 'OLEObject', 'altChunk',
                'subDoc', 'fldChar', 'fldSimple', 'instrText', 'group', 'grpSp', 'smartTag', 'customXml'}:
                return '源对象含尚未支持的原生图表、公式、域或复合容器：' + n
    return ''


def _fully_marked(node, kind):
    # A paragraph-mark revision alone is insufficient for math: its complete OMML
    # object must be a payload child of the corresponding run-level revision.
    parents={child:parent for parent in node.iter() for child in parent}
    for formula in node.iter():
        if formula.tag not in {'{'+M+'}oMath', '{'+M+'}oMathPara'}:
            continue
        parent=parents.get(formula)
        if parent is not None and parent.tag.startswith('{'+M+'}'):
            continue
        if parent is None or parent.tag!=q(kind):
            return False
    if node.tag == q('p'):
        return node.find('./w:pPr/w:rPr/w:' + kind, NS) is not None
    if node.tag == q('tbl'):
        rows = node.findall(q('tr'))
        return bool(rows) and all(row.find('./w:trPr/w:' + kind, NS) is not None for row in rows)
    return False


def _plain_text(nodes):
    return '\n'.join(''.join(e.text or '' for e in node.iter() if e.tag in {q('t'), q('delText'), '{'+M+'}t'}) for node in nodes)


def _audit_math_dependencies(source,output,sources,inserted):
    """Independently read math defaults and effective paragraph styles from ZIPs.

    The body comparison already fingerprints every OMML node, including m:t and
    math formatting. This extra check catches an otherwise identical formula whose
    appearance changes by inheriting the target document's Normal style/settings.
    It does not call the writer or treat a generated style-id prefix as evidence.
    """
    def paragraphs(nodes):
        return [p for node in nodes for p in node.iter(q('p'))
                if any(e.tag.startswith('{'+M+'}') for e in p.iter() if isinstance(e.tag,str))]
    wanted=paragraphs(sources)
    if not wanted:return
    actual=paragraphs([_resolve(e,True) for e in inserted])
    if len(wanted)!=len(actual):raise ValueError('实际插入的完整公式段落数量与来源不一致。')
    def part_element(reader,part,child):
        raw=reader.parts.get(part)
        return None if raw is None else ET.fromstring(raw).find(child)
    def semantic(reader,node,part):return None if node is None else reader.tree(node,part)
    for part,child in [('word/settings.xml','{'+M+'}mathPr'),('word/styles.xml',q('docDefaults'))]:
        a=part_element(source,part,child);b=part_element(output,part,child)
        if semantic(source,a,part)!=semantic(output,b,part):
            raise ValueError('实际公式继承的全局设置或默认字体与来源不一致。')
    def bound_style(reader,para):
        ref=para.find('./w:pPr/w:pStyle',NS)
        if ref is not None:
            sid=ref.get(q('val'))
        else:
            raw=reader.parts.get('word/styles.xml')
            styles=[] if raw is None else ET.fromstring(raw).findall(q('style'))
            matches=[s for s in styles if s.get(q('type'))=='paragraph' and s.get(q('default')) in ('1','true','on')]
            if len(matches)>1:raise ValueError('文档含多个默认段落样式，公式继承关系不明确。')
            sid=matches[0].get(q('styleId')) if matches else None
        if sid is None:return None
        result=reader.style(sid)
        if result and result[0]=='missing-style':raise ValueError('实际公式引用的段落样式缺失。')
        return result
    for a,b in zip(wanted,actual):
        if bound_style(source,a)!=bound_style(output,b):
            raise ValueError('实际公式段落未继承完整的来源段落样式。')


def _bookmark_manifest(root):
    return [(name(e), tuple(sorted(e.attrib.items()))) for e in root.iter()
            if name(e) in {'bookmarkStart', 'bookmarkEnd'}]


def _pure_paragraph(node):
    if node.tag != q('p'):
        return False
    for child in node:
        if name(child) in {'pPr', 'proofErr'}:
            if any(name(e) in {'bookmarkStart', 'bookmarkEnd'} for e in child.iter()):
                return False
            continue
        if child.tag != q('r'):
            return False
        if any(name(e) not in {'rPr', 't'} for e in child):
            return False
    return True


def _excerpt(sources, old_nodes, span):
    if len(sources) != 1 or len(old_nodes) != 1 or not all(_pure_paragraph(p) for p in sources + old_nodes):
        raise ValueError('来源摘录只支持无书签、图示、注释、超链接或域的完整纯文字目标段。')
    whole = _plain_text(sources)
    a, b = span.get('start'), span.get('end')
    if type(a) is not int or type(b) is not int or not 0 <= a < b <= len(whole):
        raise ValueError('来源摘录字符范围无效。')
    ending = '。！？!?；;'
    if (whole[:a].strip() and whole[:a].rstrip()[-1] not in ending) or (b != len(whole) and whole[a:b].rstrip()[-1] not in ending):
        raise ValueError('来源摘录边界不是完整句界。')
    value = whole[a:b]
    def normalized(text):
        return re.sub(r'[\s，,。.;；：:（）()“”"、]', '', unicodedata.normalize('NFKC', text))
    if normalized(_plain_text(old_nodes)) != normalized(value):
        raise ValueError('来源摘录与完整旧目标段并非相同事实，禁止模糊局部替换。')
    return [('paragraph', [('text', value)])]


def _wording_expected(sources,kind):
    """Independent readback: derive only authorized wording from pristine input."""
    nodes=copy.deepcopy(sources);changes=[]
    for i,el in enumerate(nodes):
        for j,p in enumerate(el.iter(q('p'))):
            parents={c:a for a in p.iter() for c in a};ts=[]
            for t in p.iter(q('t')):
                ancestor=parents.get(t)
                while ancestor is not None and ancestor.tag!=q('p'):ancestor=parents.get(ancestor)
                if ancestor is p:ts.append(t)
            before=''.join(t.text or '' for t in ts);after=before;rules=wording_mappings(before,kind)
            for old,new in rules:after=after.replace(old,new)
            if before!=after:
                changes.append({'element':i,'paragraph':j,'before':before,'after':after,'rules':[list(x) for x in rules]})
                # Audit content compares logical text, not run fragmentation.
                for k,t in enumerate(ts):t.text=after if k==0 else ''
    return nodes,changes


def _audit_insertion_boundary(rec,base_root,output,baseline,inserted):
    if rec['target_paths'] or rec['deletion_paths'] or rec.get('deletion_revision_ids'):
        raise ValueError('纯新增不得删除原稿对象。')
    boundary=rec.get('insertion_boundary',{})
    paths=rec['output_paths']
    if not boundary or not paths or len({tuple(p[:-1]) for p in paths})!=1:
        raise ValueError('新增正文缺少同父节点的原始插入边界。')
    parent_path=paths[0][:-1];positions=[p[-1] for p in paths]
    if positions!=list(range(min(positions),max(positions)+1)):
        raise ValueError('新增来源对象未连续写入。')
    for side,b in boundary.items():
        original=at_path(base_root,b['baseline_path']);current=at_path(output.root,b['output_path'])
        if baseline.tree(original,ignore_bookmarks=True)!=output.tree(_resolve(current,False),ignore_bookmarks=True):
            raise ValueError('新增正文的原始边界内容被改变。')
        path=b['output_path']
        if path[:-1]!=parent_path or (side=='before' and path[-1]<=max(positions)) or (side=='after' and path[-1]>=min(positions)):
            raise ValueError('新增正文不在已冻结的前后边界之间。')


def _audit_parameter(rec, result, baseline, output, base_root, claimed_ids):
    old = [at_path(base_root, p) for p in rec['target_paths']]
    deleted = [at_path(output.root, p) for p in rec['deletion_paths']]
    inserted = [at_path(output.root, p) for p in rec['output_paths']]
    ins = revision_ids(inserted, 'ins'); dels = revision_ids(deleted, 'del')
    claimed_ids.extend(ins + dels)
    if (len(old) != 1 or len(deleted) != 1 or len(inserted) != 1 or
        not ins or not dels or ins != list(map(str, rec.get('insertion_revision_ids', []))) or
        dels != list(map(str, rec.get('deletion_revision_ids', [])))):
        raise ValueError('参数编辑必须有一对完整段落及真实插入/删除修订 ID。')
    if not _fully_marked(deleted[0], 'del') or not _fully_marked(inserted[0], 'ins'):
        raise ValueError('参数编辑未标记完整旧新段落。')
    if not _pure_paragraph(old[0]) or not _pure_paragraph(_resolve(inserted[0], True)):
        raise ValueError('参数编辑仅支持纯文字段落。')
    if baseline.content(old[0]) != output.content(_resolve(deleted[0], False)):
        raise ValueError('参数编辑删除的段落与原始输入不符。')
    expected = _plain_text(old)
    changes = rec.get('parameter_changes')
    if not changes:
        raise ValueError('参数编辑缺少明确冻结的参数变化。')
    for change in changes:
        field, previous, value = change['field'], change['old'], change['new']
        if rec.get('parameters') is not None and rec['parameters'].get(field) != value:
            raise ValueError('参数改动与明确输入或当前来源确定的字段不一致。')
        if not isinstance(value, str) or not value or not isinstance(previous, str) or previous not in expected:
            raise ValueError('参数旧值不在原段或新值为空。')
        if field in {'issuer', 'bond'}:
            if not previous:
                raise ValueError('公司或债券旧值为空。')
            expected = expected.replace(previous, value)
        elif field == 'date':
            if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', value) or not re.match(r'^调查日期[:：]', expected):
                raise ValueError('调查日期参数或原始字段格式无效。')
            dt = date.fromisoformat(value)
            expected = '调查日期：【%s】年【%s】月【%s】日' % (dt.year, dt.month, dt.day)
        elif field == 'period':
            if not re.search(r'报告期[:：]', expected):
                raise ValueError('报告期参数没有对应的原始字段。')
            expected = re.sub(r'(?<=报告期)[:：].*', '：' + value + '）', expected)
        else:
            raise ValueError('未知或未经授权的参数字段：' + str(field))
    if rec.get('expected_text') != expected:
        raise ValueError('参数预期正文不能由冻结参数独立推导。')
    actual = output.content(_resolve(inserted[0], True))
    want = ('paragraph', [('text', expected)])
    if actual != want:
        raise ValueError('实际参数插入正文与明确冻结的参数不一致。')
    result['inserted_object_digests'] = [digest(actual)]
    result['expected_parameter_digest'] = digest(want)
    result['passed'] = True


def audit_docx(output_path, baseline_path, records, protected_ranges=(), baseline_policy='original', missing_decisions=None):
    """Audit final saved bytes. Every failure blocks verified/final status.

    Records: id/action, source_file, source_sha256, source_role, source_paths,
    target_paths, output_paths, deletion_paths, insertion_revision_ids,
    deletion_revision_ids. Paths are root-based lists of zero-based child indices.
    source_merge additionally requires merge_into; it may delete plain paragraph
    fragments only when another audited full source paragraph already includes them.
    protected_ranges: baseline_path/output_path pairs for every retained input block.
    """
    report = {'passed': False, 'errors': [], 'records': [], 'reject_restores_baseline': False,
              'protected_unchanged': False, 'parser': 'independent-stdlib-xml-v1',
              'baseline_policy': baseline_policy, 'baseline_sha256': '', 'output_sha256': '',
              'dependencies': [], 'technical_changes': []}
    errors = report['errors']
    try:
        baseline = Reader(baseline_path); output = Reader(output_path)
        report.update(baseline_sha256=baseline.sha256, output_sha256=output.sha256)
        historical = [e for e in baseline.root.iter() if e.tag in REV_TAGS]
        if historical and baseline_policy != 'accepted-historical':
            raise ValueError('原始输入含历史修订；未明确授权接受历史修订基线，禁止导出。')
        if baseline_policy not in {'original', 'accepted-historical'}:
            raise ValueError('未知原始基线策略。')
        base_root = _resolve(baseline.root, True) if historical else baseline.root
        all_revisions = [e for e in output.root.iter() if e.tag in REV_TAGS]
        if any(name(e) not in {'ins', 'del'} for e in all_revisions):
            raise ValueError('输出含未登记支持的移动或格式修订。')
        ids = [e.get(q('id'), '') for e in all_revisions]
        if '' in ids or len(ids) != len(set(ids)):
            errors.append('输出修订 ID 缺失或重复。')
        if any(e.get(q('author')) != '柒' for e in all_revisions):
            errors.append('修订作者不是柒。')
        rejected = _resolve(output.root, False)
        # Bookmark endpoints may be relocated out of fully deleted blocks. Their
        # identities and order are checked separately; content/format is exact.
        same = (baseline.tree(base_root, ignore_bookmarks=True) ==
                output.tree(rejected, ignore_bookmarks=True))
        if not same:
            errors.append('拒绝本轮修订后正文结构、格式或对象依赖未恢复原始基线。')
        if _bookmark_manifest(base_root) != _bookmark_manifest(rejected):
            errors.append('拒绝修订后书签身份或顺序未恢复。')
            same = False
        report['reject_restores_baseline'] = same
        source_readers = {}
        claimed_ids = []
        passed_by_id = {}
        protected_ok = True
        for protected in protected_ranges:
            try:
                old = at_path(base_root, protected['baseline_path'])
                now = at_path(output.root, protected['output_path'])
                if revision_ids([now]):
                    raise ValueError('保护对象被伪修订。')
                if baseline.tree(old, ignore_bookmarks=True) != output.tree(now, ignore_bookmarks=True):
                    raise ValueError('保护对象内容、格式或关系被改写。')
            except (ValueError, IndexError, KeyError) as exc:
                protected_ok = False
                errors.append('保护区 ' + str(protected.get('baseline_path')) + '：' + str(exc))
        report['protected_unchanged'] = protected_ok
        deferred = []
        for rec in records:
            result = {'id': rec.get('id'), 'action': rec.get('action', 'source_copy'), 'passed': False,
                      'errors': [], 'source_object_digests': [], 'inserted_object_digests': []}
            report['records'].append(result)
            try:
                action = result['action']
                if rec.get('target_sha256') and rec['target_sha256'] != baseline.sha256:
                    raise ValueError('目标原始输入哈希与冻结计划不一致。')
                if action == 'parameter_edit':
                    _audit_parameter(rec, result, baseline, output, base_root, claimed_ids)
                    continue
                if action not in {'source_copy', 'source_merge', 'source_insert'}:
                    raise ValueError('未授权的非来源写入动作：' + action)
                if rec.get('source_role') not in {'source', 'prospectus', 'authorized_formal_source'}:
                    raise ValueError('插入来源角色未经授权。')
                if rec.get('transforms') not in (None, [], ['fonts', 'technical_ids', 'table_layout']):
                    raise ValueError('含未支持或未授权的事实文字变换。')
                source_file = rec['source_file']
                if str(Path(source_file).resolve()) == str(Path(baseline_path).resolve()):
                    raise ValueError('旧底稿不可作为本轮新内容来源。')
                if source_file not in source_readers:
                    source_readers[source_file] = Reader(source_file)
                source = source_readers[source_file]
                if not rec.get('source_sha256') or rec['source_sha256'] != source.sha256:
                    raise ValueError('来源输入哈希已改变或未冻结。')
                if any(e.tag in REV_TAGS for e in source.root.iter()):
                    raise ValueError('正式来源含历史修订；未声明来源接受策略。')
                sources = [at_path(source.root, p) for p in rec['source_paths']]
                if not sources:
                    raise ValueError('缺少真实来源对象范围。')
                for source_node in sources:
                    for node in source_node.iter():
                        if name(node) in {'blip', 'imagedata'}:
                            for key, value in node.attrib.items():
                                if key.startswith('{' + R + '}') and source.rel('word/document.xml', value)[0] == 'external':
                                    raise ValueError('来源为外链图片，无法核验当前实际媒体字节，保留待核对。')
                unsupported = _unsupported(sources)
                if unsupported:
                    raise ValueError(unsupported)
                old_nodes = [at_path(base_root, p) for p in rec['target_paths']]
                deleted = [at_path(output.root, p) for p in rec['deletion_paths']]
                inserted = [at_path(output.root, p) for p in rec['output_paths']]
                ins_ids = revision_ids(inserted, 'ins'); del_ids = revision_ids(deleted, 'del')
                claimed_ids.extend(ins_ids + del_ids)
                if ins_ids != [str(x) for x in rec.get('insertion_revision_ids', [])]:
                    raise ValueError('实际插入修订 ID 与登记不一致。')
                if del_ids != [str(x) for x in rec.get('deletion_revision_ids', [])] or (not del_ids and action!='source_insert'):
                    raise ValueError('实际删除修订 ID 缺失或与登记不一致。')
                if action=='source_insert':_audit_insertion_boundary(rec,base_root,output,baseline,inserted)
                if not all(_fully_marked(e, 'del') for e in deleted):
                    raise ValueError('旧对象未完整标记为删除。')
                restored_deleted = [_resolve(e, False) for e in deleted]
                if [baseline.content(e) for e in old_nodes] != [output.content(e) for e in restored_deleted]:
                    raise ValueError('删除的实际旧对象与冻结原始范围不一致。')
                expected = (_excerpt(sources, old_nodes, rec['source_span']) if rec.get('source_span')
                            else [source.content(e) for e in sources])
                result['source_object_digests'] = [digest(e) for e in expected]
                if rec.get('wording_policy'):
                    if rec['wording_policy']!=WORDING_POLICY or rec.get('source_kind') not in ('prospectus','opinion'):
                        raise ValueError('称谓转换规则或来源文种未经授权。')
                    wording_sources=sources
                    if rec.get('source_span'):
                        p=ET.Element(q('p'));ET.SubElement(ET.SubElement(p,q('r')),q('t')).text=expected[0][1][0][1]
                        wording_sources=[p]
                    adapted,changes=_wording_expected(wording_sources,rec['source_kind'])
                    if changes!=rec.get('wording_changes',[]):raise ValueError('称谓转换记录与原始来源不一致。')
                    expected=[source.content(e) for e in adapted]
                    result['wording_changes']=changes
                elif rec.get('wording_changes'):
                    raise ValueError('称谓转换缺少授权规则。')

                result['source_file'] = str(source_file)
                result['source_paths'] = rec['source_paths']
                if action == 'source_merge':
                    if inserted or ins_ids:
                        raise ValueError('source_merge 不应产生第二份来源插入。')
                    if any(e.tag != q('p') or _unsupported([e]) or any(name(x) in {'drawing', 'pict', 'hyperlink', 'footnoteReference', 'endnoteReference'} or str(x.tag).startswith('{'+M+'}') for x in e.iter()) for e in old_nodes):
                        raise ValueError('合并删除暂仅支持纯文字段落。')
                    deferred.append((rec, result, source, sources, old_nodes))
                    continue
                pre_title = rec.get('approved_title') if rec.get('content_class') == 'approved_missing_item' else None
                if pre_title is not None:
                    # One target-created heading precedes the copied body. The
                    # extra heading is only accepted when a persisted human
                    # decision exists for exactly this candidate and every
                    # frozen fingerprint still matches that decision. A plain
                    # field on some other proposal cannot fake this.
                    fingerprints = rec.get('decision_fingerprints') or {}
                    decision = (missing_decisions or {}).get(fingerprints.get('candidate_id'))
                    if not decision or decision.get('decision') != 'add':
                        raise ValueError('新增事项缺少持久化的人工确认，已拒绝写入。')
                    for key in ('source_hash', 'source_body_hash', 'parent_fingerprint',
                                'previous_sibling_fingerprint', 'next_boundary_fingerprint'):
                        if fingerprints.get(key) != decision.get(key):
                            raise ValueError('新增事项的冻结证据与人工确认记录不一致，已拒绝写入。')
                    if pre_title != decision.get('approved_title'):
                        raise ValueError('新增事项标题与人工确认的标题不一致，已拒绝写入。')
                    if len(inserted) != len(sources) + 1 or not all(_fully_marked(e, 'ins') for e in inserted):
                        raise ValueError('实际插入不是完整来源对象的一一复制。')
                    if _plain_text([inserted[0]]).strip() != pre_title:
                        raise ValueError('新增事项标题与人工确认的标题不一致。')
                    rest = inserted[1:]
                else:
                    rest = inserted
                if len(rest) != len(sources) or not ins_ids or not all(_fully_marked(e, 'ins') for e in rest):
                    raise ValueError('实际插入不是完整来源对象的一一复制。')
                actual = [output.content(_resolve(e, True)) for e in rest]
                result['inserted_object_digests'] = [digest(e) for e in actual]
                if actual != expected:
                    raise ValueError('实际插入正文、整表结构或媒体依赖与已登记当前来源不一致。')
                _audit_math_dependencies(source,output,sources,rest)
                result['passed'] = True
                passed_by_id[rec['id']] = (rec, expected, _plain_text(sources))
            except (ValueError, KeyError, IndexError, ET.ParseError) as exc:
                result['errors'].append(str(exc)); errors.append(str(rec.get('id')) + '：' + str(exc))
        for rec, result, source, sources, old_nodes in deferred:
            dest = passed_by_id.get(rec.get('merge_into'))
            if dest is None or dest[0]['source_sha256'] != rec['source_sha256'] or dest[0]['source_paths'] != rec['source_paths']:
                result['errors'].append('合并目标未通过同一当前来源实际插入审计。')
            elif any(not _plain_text([e]).strip() or _plain_text([e]).strip() not in dest[2] for e in old_nodes):
                result['errors'].append('合并删除的旧段并未被实际复制的完整来源覆盖。')
            else:
                result['passed'] = True
            errors.extend(str(rec.get('id')) + '：' + msg for msg in result['errors'])
        if len(claimed_ids) != len(set(claimed_ids)) or set(claimed_ids) != set(ids):
            errors.append('存在未登记或被多个来源计划重复认领的实际修订。')
        _audit_parts(baseline, output, report)
        report['dependencies'] = output.dependencies
        report['protected_count'] = len(protected_ranges)
        report['revision_count'] = len(ids)
        for part,raw in output.parts.items():
            if part.startswith('word/') and part.endswith('.xml'):
                errors.extend(part+'：'+issue for issue in property_order_issues(ET.fromstring(raw)))
        report['source_insert_count'] = sum(r['passed'] and r['action'] == 'source_insert' for r in report['records'])
        report['source_copy_count'] = sum(r['passed'] and r['action'] == 'source_copy' for r in report['records'])
    except (ValueError, KeyError, IndexError, ET.ParseError, zipfile.BadZipFile, OSError) as exc:
        errors.append(str(exc))
    report['passed'] = not errors
    return report


def allowed_style_pruning(old_parts, new_parts):
    """Independently prove a minimal, unreachable style removal for Word's cap.

    No writer trace is trusted. All body, header, note, numbering and revision
    XML remains part of the reference scan, including field style names.
    """
    part = 'word/styles.xml'
    if part not in old_parts or part not in new_parts:
        return set(), None
    old_root = ET.fromstring(old_parts[part]); new_root = ET.fromstring(new_parts[part])
    old = {n.get(q('styleId')): n for n in old_root if n.tag == q('style')}
    new = {n.get(q('styleId')): n for n in new_root if n.tag == q('style')}
    removed = set(old) - set(new)
    if not removed:
        return set(), None
    union_count = len(set(old) | set(new))
    evidence = {'part': part, 'change': 'minimal-unreferenced-style-pruning',
                'before_including_imports': union_count, 'after': len(new),
                'removed_style_ids': sorted(removed), 'limit': 4079}
    if union_count <= 4079 or len(new) != 4079 or len(removed) != union_count - 4079:
        return set(), {**evidence, 'allowed': False, 'reason': 'not-minimal-limit-repair'}
    protected = set(); fields = []
    style_refs = {q(n) for n in ('pStyle', 'rStyle', 'tblStyle', 'basedOn',
                                'next', 'link', 'styleLink', 'numStyleLink')}
    for entries in (old_parts, new_parts):
        for filename, raw in entries.items():
            if not filename.endswith('.xml'):
                continue
            root = ET.fromstring(raw)
            if filename == part:
                # Deleted internal nodes cannot create a live dependency.
                for child in list(root):
                    if child.tag == q('style') and child.get(q('styleId')) in removed:
                        root.remove(child)
            for node in root.iter():
                for attr, val in node.attrib.items():
                    # Revision IDs, font sizes and numbering IDs can equal a
                    # numeric style ID without referring to that style.
                    style_ref = node.tag in style_refs and attr == q('val')
                    extension_ref = not node.tag.startswith('{' + W + '}')
                    if (style_ref or extension_ref) and val in removed:
                        protected.add(val)
                if node.tag == q('instrText') and node.text:
                    fields.append(node.text)
                if node.tag == q('fldSimple'):
                    fields.append(node.get(q('instr'), ''))
    field_text = ' '.join(fields).casefold()
    compact_field_text = ''.join(fields).casefold()
    for sid in removed:
        style = old[sid]
        if style.get(q('default')) in ('1', 'true', 'on'):
            protected.add(sid)
        labels = [sid]
        for tag in ('name', 'aliases'):
            node = style.find(q(tag))
            if node is not None:
                labels.extend(node.get(q('val'), '').split(','))
        if any(label.strip() and (label.strip().casefold() in field_text or
               label.strip().casefold() in compact_field_text) for label in labels):
            protected.add(sid)
    evidence.update({'allowed': not protected, 'protected_removed_ids': sorted(protected)})
    return (removed if not protected else set()), evidence


def _audit_parts(baseline, output, report):
    """Old non-body parts stay intact; note additions require actual copy evidence."""
    allowed = {'word/document.xml', 'word/settings.xml', 'word/_rels/document.xml.rels',
               '[Content_Types].xml', 'word/styles.xml', 'word/numbering.xml'}
    # Imported definitions may be appended, but existing definitions determine
    # the appearance of protected content and must never be silently overwritten.
    allowed_style_ids, pruning = allowed_style_pruning(baseline.parts, output.parts)
    if pruning is not None:
        report['technical_changes'].append(pruning)
    for part, keynames in [('word/styles.xml', ('styleId',)),
                           ('word/numbering.xml', ('numId', 'abstractNumId'))]:
        if part not in baseline.parts:
            continue
        if part not in output.parts:
            report['errors'].append('原始样式或编号部件丢失：' + part)
            continue
        a = ET.fromstring(baseline.parts[part]); b = ET.fromstring(output.parts[part])
        def identity(node):
            return (node.tag, tuple((key, node.get(q(key))) for key in keynames))
        current = {identity(e): e for e in b}
        for old in a:
            new = current.get(identity(old))
            if (new is None and part == 'word/styles.xml' and old.tag == q('style')
                    and old.get(q('styleId')) in allowed_style_ids):
                continue
            if new is None or baseline.tree(old, part) != output.tree(new, part):
                report['errors'].append('原始样式或编号定义被覆盖：' + part + '/' + str(identity(old)))
    if 'word/settings.xml' in baseline.parts:
        if 'word/settings.xml' not in output.parts:
            report['errors'].append('原始设置部件丢失。')
        else:
            a = ET.fromstring(baseline.parts['word/settings.xml']); b = ET.fromstring(output.parts['word/settings.xml'])
            for root in (a, b):
                for child in list(root):
                    if child.tag == q('trackRevisions'):
                        root.remove(child)
            if baseline.tree(a, 'word/settings.xml') != output.tree(b, 'word/settings.xml'):
                report['errors'].append('除修订开关外的原始文档设置被改写。')
    note_parts = {'word/footnotes.xml', 'word/endnotes.xml'}
    referenced_notes = {(d['part'], d.get('note_id')) for d in output.dependencies if 'note_id' in d}
    for part, raw in baseline.parts.items():
        if part in allowed or output.parts.get(part) == raw:
            continue
        if part not in note_parts or part not in output.parts:
            report['errors'].append('原始无关部件被改动或移除：' + part)
            continue
        oldroot = ET.fromstring(raw); newroot = ET.fromstring(output.parts[part])
        oldnotes = {e.get(q('id')): e for e in oldroot}; newnotes = {e.get(q('id')): e for e in newroot}
        if len(newnotes) != len(newroot):
            report['errors'].append('注释部件含重复 ID：' + part)
        for note_id, oldnote in oldnotes.items():
            if note_id not in newnotes or baseline.tree(oldnote, part) != output.tree(newnotes[note_id], part):
                report['errors'].append('原始脚注或尾注被覆盖：' + part + '/' + str(note_id))
        additions = set(newnotes) - set(oldnotes)
        for note_id in additions:
            if (part, note_id) not in referenced_notes:
                report['errors'].append('新增注释无实际已登记引用：' + part + '/' + str(note_id))
        report['technical_changes'].append({'part': part, 'change': 'preserved-existing-notes-added-referenced-notes',
                                            'added_note_ids': sorted(additions)})
    # Added media/dependencies are legal only when actually traversed by the audit.
    touched = {d['part'] for d in output.dependencies}
    for part in set(output.parts) - set(baseline.parts):
        if (part.startswith('word/media/') or part.startswith('word/charts/') or part.startswith('word/embeddings/')) and part not in touched:
            report['errors'].append('新增对象依赖未被已登记来源插入实际引用：' + part)
