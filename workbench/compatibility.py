"""Conservative WordprocessingML property-order compatibility checks.

The generated documents must satisfy the property's schema sequence. Word for
Windows can reject malformed sequences that another office application silently
repairs. These helpers only reorder *property elements*: no text, attributes,
relationships, revisions, drawing objects or document blocks are discarded.

Orders are taken from the Microsoft Open XML SDK's schema definitions:
https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/
schemas_openxmlformats_org_wordprocessingml_2006_main.json

This is a targeted check, not a complete OOXML schema validator or evidence of a
successful Microsoft Word open/save/accept/reject test.
"""
from __future__ import annotations

from collections import Counter
import re
import shlex

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
NS = {'w': W, 'w14': W14, 'w15': W15}


def _q(value):
    prefix, local = value.split(':', 1)
    return '{' + NS[prefix] + '}' + local


def _w_words(value):
    return [_q('w:' + word) for word in value.split()]


_RUN = _w_words('''rStyle rFonts b bCs i iCs caps smallCaps strike dstrike
outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing
w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em
lang eastAsianLayout specVanish''') + [
    _q('w14:' + word) for word in '''glow shadow reflection textOutline textFill
    scene3d props3d ligatures numForm numSpacing stylisticSets cntxtAlts'''.split()
]

_ORDERS = {
    _q('w:rPr'): _RUN + [_q('w:rPrChange')],
    _q('w:pPr'): _w_words('''pStyle keepNext keepLines pageBreakBefore framePr
        widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens
        kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi
        adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents
        suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl
        divId cnfStyle rPr sectPr pPrChange'''),
    _q('w:tblPr'): _w_words('''tblStyle tblpPr tblOverlap bidiVisual tblW jc
        tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption
        tblDescription tblPrChange'''),
    _q('w:tcPr'): _w_words('''cnfStyle tcW gridSpan hMerge vMerge tcBorders shd
        noWrap tcMar textDirection tcFitText vAlign hideMark cellIns cellDel
        cellMerge tcPrChange'''),
    _q('w:sectPr'): _w_words('''headerReference footerReference footnotePr
        endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols
        formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid
        printerSettings''') + [_q('w15:footnoteColumns'), _q('w:sectPrChange')],
}
_PARA_RUN = (_w_words('ins del moveFrom moveTo') +
             [_q('w14:conflictIns'), _q('w14:conflictDel')] +
             _RUN + _w_words('oMath rPrChange'))
_ROW_PROPERTIES = _w_words('''cnfStyle divId gridBefore gridAfter wBefore wAfter
    trHeight hidden cantSplit tblHeader tblCellSpacing jc''')
_ROW_REVISIONS = _w_words('ins del trPrChange') + [
    _q('w14:conflictIns'), _q('w14:conflictDel')]


def _ranks(node, parent_tag=None):
    """Return sequence ranks; an equal rank represents a schema choice group."""
    tag = node.tag
    if tag == _q('w:trPr'):
        # The pre-revision row properties form a repeated choice, not a sequence.
        return {**{name: 0 for name in _ROW_PROPERTIES},
                **{name: i + 1 for i, name in enumerate(_ROW_REVISIONS)}}
    if tag == _q('w:rPr'):
        # A caller comparing a subtree may not have its parent available. The
        # extended rank set has exactly the same order for shared run properties;
        # validation still uses the actual parent and rejects illegal children.
        order = _PARA_RUN if parent_tag in (None, _q('w:pPr')) else _ORDERS[tag]
    else:
        order = _ORDERS.get(tag)
    if order is None:
        return None
    ranks = {name: i for i, name in enumerate(order)}
    if tag == _q('w:sectPr'):
        # Header/footer references may alternate and repeat by reference type.
        ranks[_q('w:footerReference')] = ranks[_q('w:headerReference')]
    return ranks


def is_property_container(node):
    return node.tag in _ORDERS or node.tag == _q('w:trPr')


def ordered_property_children(node, parent_tag=None):
    """Return canonical children without mutation, preserving duplicates/content.

    Only closed, recognized property sequences are sorted. Unknown extensions
    leave the entire sequence unchanged; they must never be silently relocated.
    """
    children = list(node)
    ranks = _ranks(node, parent_tag)
    if ranks is None or any(c.tag not in ranks for c in children):
        return children
    return sorted(children, key=lambda child: ranks[child.tag])


def property_order_issues(root):
    """Report actual schema-order/duplicate violations in supported properties."""
    issues = []
    parents = {child: parent for parent in root.iter() for child in parent}
    for index, node in enumerate(root.iter()):
        parent = parents.get(node)
        ranks = _ranks(node, parent.tag if parent is not None else None)
        if ranks is None:
            continue
        tag = str(node.tag).rsplit('}', 1)[-1]
        children = list(node)
        values = [ranks[c.tag] for c in children if c.tag in ranks]
        if values != sorted(values):
            issues.append(f'{tag}[{index}] 的属性子元素顺序不符合 WordprocessingML schema。')
        # Repeated choice groups have distinct cardinality rules. Do not invent
        # a uniqueness rule for those groups when checking sequence properties.
        repeatable = ({_q('w:headerReference'), _q('w:footerReference')}
                      if node.tag == _q('w:sectPr') else set())
        if node.tag == _q('w:trPr'):
            repeatable.update(_ROW_PROPERTIES)
        counts = Counter(c.tag for c in children if c.tag in ranks)
        for child_tag, count in counts.items():
            if count > 1 and child_tag not in repeatable:
                child_name = str(child_tag).rsplit('}', 1)[-1]
                issues.append(f'{tag}[{index}] 的单值属性 {child_name} 重复 {count} 次。')
    return issues


def normalize_property_order(root):
    """Normalize only recognized property sequences in place; return diagnostics.

    Duplicate children are preserved and reported, never deduplicated. Unknown
    extensions are preserved in position. The caller must reject unresolved
    issues after this function; a changed count is not a general validity claim.
    """
    changed = Counter()
    parents = {child: parent for parent in root.iter() for child in parent}
    for node in root.iter():
        parent = parents.get(node)
        before = list(node)
        after = ordered_property_children(node, parent.tag if parent is not None else None)
        if before != after:
            node[:] = after
            changed[str(node.tag).rsplit('}', 1)[-1]] += 1
    return {'changed': sum(changed.values()), 'by_property': dict(changed),
            'issues': property_order_issues(root)}


WORD_STYLE_LIMIT = 4079


def _unreferenced_style_groups(definitions, dependencies, reachable):
    """Return unreachable SCCs with no incoming dependency from another SCC.

    A paragraph/character link pair is one indivisible group. Iterative graph
    traversal avoids recursion limits in documents containing thousands of
    legacy style definitions.
    """
    graph = {sid: set(dependencies[sid]) - {sid} for sid in definitions}
    reverse = {sid: set() for sid in graph}
    for sid, deps in graph.items():
        for dep in deps:
            reverse[dep].add(sid)
    seen, order = set(), []
    for sid in graph:
        if sid in seen:
            continue
        todo = [(sid, False)]
        while todo:
            node, finished = todo.pop()
            if finished:
                order.append(node)
            elif node not in seen:
                seen.add(node)
                todo.append((node, True))
                todo.extend((dep, False) for dep in sorted(graph[node]) if dep not in seen)
    groups, assigned = [], set()
    for sid in reversed(order):
        if sid in assigned:
            continue
        group, todo = set(), [sid]
        while todo:
            node = todo.pop()
            if node in assigned:
                continue
            assigned.add(node)
            group.add(node)
            todo.extend(reverse[node] - assigned)
        if group.isdisjoint(reachable) and not any(reverse[node] - group for node in group):
            groups.append(group)
    index = {sid: i for i, sid in enumerate(definitions)}
    return sorted(groups, key=lambda group: min(index[sid] for sid in group))


def _style_fields(root):
    """Read simple/complex field instructions, including deleted review history."""
    instructions, stack, orphan = [], [], []
    for node in root.iter():
        if node.tag == _q('w:fldSimple'):
            instructions.append(node.get(_q('w:instr'), ''))
        elif node.tag == _q('w:fldChar'):
            kind = node.get(_q('w:fldCharType'))
            if kind == 'begin':
                if stack and not stack[-1]['result']:
                    stack[-1]['nested'] = True
                stack.append({'text': [], 'result': False, 'nested': False})
            elif kind == 'separate' and stack:
                stack[-1]['result'] = True
            elif kind == 'end' and stack:
                field = stack.pop()
                text = ''.join(field['text'])
                if field['nested'] and re.search(r'\b(STYLEREF|TOC|AUTOTEXTLIST)\b', text, re.I):
                    raise ValueError('Word 样式容量检查：样式字段包含动态嵌套参数，不能安全裁除样式。')
                instructions.append(text)
        elif node.tag in (_q('w:instrText'), _q('w:delInstrText')):
            if stack and not stack[-1]['result']:
                stack[-1]['text'].append(node.text or '')
            elif not stack:
                orphan.append(node.text or '')
    if stack or orphan:
        # Broken field boundaries cannot safely delimit a style name.
        text = ''.join(orphan + [p for f in stack for p in f['text']])
        if re.search(r'\b(STYLEREF|TOC|AUTOTEXTLIST)\b', text, re.I):
            raise ValueError('Word 样式容量检查：样式字段边界不完整，不能安全裁除样式。')
    return instructions


def enforce_style_capacity(parts, limit=WORD_STYLE_LIMIT, _precise_refs=False):
    """Remove only enough unreachable style definitions to meet Word's limit.

    ``parts`` contains every final package XML part as an Element root. Only
    styles.xml/stylesWithEffects.xml nodes are mutated, after the whole plan is
    validated. Actual content, formatting, IDs, lists and revisions are untouched.
    Microsoft documents 4,079 style definitions as Word's operating limit:
    https://learn.microsoft.com/en-us/previous-versions/troubleshoot/
    microsoft-365/microsoft-365-apps/word/operating-parameter-limitation
    """
    style_parts = {name: root for name, root in parts.items()
                   if name in ('word/styles.xml', 'word/stylesWithEffects.xml')}
    counts = {name: len(root.findall(_q('w:style'))) for name, root in style_parts.items()}
    report = {'limit': limit, 'before': counts, 'after': dict(counts),
              'removed': [], 'parts_changed': []}
    if not any(n > limit for n in counts.values()):
        return report
    definitions, name_to_ids = {}, {}
    nodes_by_part = {}
    for name, root in style_parts.items():
        nodes = root.findall(_q('w:style'))
        ids = [n.get(_q('w:styleId')) for n in nodes]
        if None in ids or len(set(ids)) != len(ids):
            raise ValueError('Word 样式容量检查：样式 ID 缺失或重复，不能安全裁除。')
        nodes_by_part[name] = dict(zip(ids, nodes))
        for sid, node in zip(ids, nodes):
            definitions.setdefault(sid, []).append(node)
            labels = [sid]
            for tag in ('name', 'aliases'):
                item = node.find(_q('w:' + tag))
                if item is not None:
                    labels.extend((item.get(_q('w:val')) or '').split(','))
            for label in labels:
                name_to_ids.setdefault(label.strip().casefold(), set()).add(sid)
    all_ids = set(definitions)
    roots, dependencies = set(), {sid: set() for sid in all_ids}

    def matching(value):
        return name_to_ids.get((value or '').strip().casefold(), set())

    def scan_attrs(node, found):
        # Conservative: any full attribute matching an ID/name/alias is a
        # reference, including unfamiliar extensions and custom XML parts.
        # If numeric IDs collide with revision IDs/measurements and prevent
        # safe trimming, retry with the known Word style-reference elements.
        # This never changes handling of unknown extension/custom XML values.
        references = {'pStyle', 'rStyle', 'tblStyle', 'basedOn', 'next', 'link',
                      'numStyleLink', 'styleLink', 'clickAndTypeStyle',
                      'defaultTableStyle', 'lsdException'}
        for child in node.iter():
            if _precise_refs and str(child.tag).startswith('{' + W + '}'):
                local = str(child.tag).split('}', 1)[1]
                if local not in references and 'style' not in local.lower():
                    continue
                if local in ('style', 'styles', 'latentStyles'):
                    continue
            for value in child.attrib.values():
                found.update(matching(value))

    for sid, nodes in definitions.items():
        for node in nodes:
            if node.get(_q('w:default')) in ('1', 'true', 'on'):
                roots.add(sid)
            scan_attrs(node, dependencies[sid])
            # Preserve implicit built-in defaults, even in older source files
            # that omit the explicit default flag.
            label = node.find(_q('w:name'))
            labels = {sid.casefold(), (label.get(_q('w:val'), '') if label is not None else '').casefold()}
            if labels & {'normal', 'defaultparagraphfont', 'default paragraph font',
                         'tablenormal', 'normal table', 'table normal', 'nolist', 'no list'}:
                roots.add(sid)
    for part, root in parts.items():
        if part in style_parts:
            for child in root:
                if child.tag != _q('w:style'):
                    scan_attrs(child, roots)
        else:
            scan_attrs(root, roots)
        for instruction in _style_fields(root):
            lexer = shlex.shlex(instruction, posix=True)
            lexer.whitespace_split = True
            lexer.escape = ''
            try:
                words = list(lexer)
            except ValueError as exc:
                raise ValueError('Word 样式容量检查：字段引号不完整，不能安全裁除。') from exc
            if not words:
                continue
            command = words[0].upper()
            if command == 'STYLEREF':
                if len(words) < 2 or words[1].startswith('\\'):
                    raise ValueError('Word 样式容量检查：STYLEREF 缺少明确样式名。')
                refs = matching(words[1])
                if not refs:
                    # Numeric built-in style identifiers/localized names need
                    # Word's own resolution; do not guess and remove its target.
                    raise ValueError('Word 样式容量检查：无法解析 STYLEREF 样式 ' + words[1])
                roots.update(refs)
            elif command == 'TOC':
                for i, word in enumerate(words[1:], 1):
                    if word.lower() == '\\t':
                        if i + 1 >= len(words):
                            raise ValueError('Word 样式容量检查：TOC 样式参数不完整。')
                        tokens = words[i + 1].split(',')
                        if len(tokens) % 2:
                            raise ValueError('Word 样式容量检查：TOC 样式参数不完整。')
                        for label in tokens[::2]:
                            refs = matching(label)
                            if not refs:
                                raise ValueError('Word 样式容量检查：无法解析 TOC 样式 ' + label)
                            roots.update(refs)
                    elif word.lower() in ('\\o', '\\u'):
                        roots.update(sid for sid, nodes in definitions.items()
                                     if any(n.find('.//' + _q('w:outlineLvl')) is not None for n in nodes))
            elif command == 'AUTOTEXTLIST':
                # This field can resolve a template style at runtime.
                raise ValueError('Word 样式容量检查：AUTOTEXTLIST 样式不能安全静态解析。')
            else:
                # If another field quotes an exact style name, retain it too.
                for word in words[1:]:
                    roots.update(matching(word))
    reachable, stack = set(roots), list(roots)
    while stack:
        sid = stack.pop()
        for dep in dependencies.get(sid, ()) - reachable:
            reachable.add(dep)
            stack.append(dep)
    plan = {}
    # Unused definitions that remain are still real package objects. Do not
    # leave their dependencies dangling merely because the body cannot reach
    # them. Unreferenced linked-style cycles are removed as indivisible groups.
    groups = _unreferenced_style_groups(definitions, dependencies, reachable)
    incoming = {dep for sid, deps in dependencies.items() for dep in deps if dep != sid}
    for part, count in counts.items():
        need = max(0, count - limit)
        leaves = [sid for sid in nodes_by_part[part]
                  if sid not in reachable and sid not in incoming]
        if len(leaves) >= need:
            # Keep the already validated single-definition plan stable.
            plan[part] = set(leaves[:need])
            continue
        # Exact bounded subset selection keeps the edit to the minimum required
        # definitions; it never splits a link cycle merely to hit the limit.
        subsets = {0: set()}
        for group in groups:
            cost = len(group.intersection(nodes_by_part[part]))
            if not cost or cost > need:
                continue
            for size, selected in list(subsets.items())[::-1]:
                if size + cost <= need and size + cost not in subsets:
                    subsets[size + cost] = selected | group
            if need in subsets:
                break
        if need not in subsets:
            if not _precise_refs:
                return enforce_style_capacity(parts, limit, _precise_refs=True)
            raise ValueError('Word 样式容量超过 ' + str(limit) + '，无法在保留全部引用和依赖的前提下安全降低：' + part)
        plan[part] = subsets[need]
    # Keep a style present in both style stores consistent. Removing extra
    # unreachable duplicates does not alter any referenced formatting.
    removed_ids = {sid for ids in plan.values() for sid in ids}
    for part, nodes in nodes_by_part.items():
        removed_here = [sid for sid in nodes if sid in removed_ids]
        if not removed_here:
            continue
        for sid in removed_here:
            style_parts[part].remove(nodes[sid])
            report['removed'].append({'part': part, 'style_id': sid,
                'reason': 'No package reference or style-field name reference; unreachable from default/used styles through all style dependencies.'})
        report['parts_changed'].append(part)
        report['after'][part] -= len(removed_here)
    report['reachable_count'] = len(reachable)
    report['reference_mode'] = 'semantic_word_refs_and_conservative_extensions' if _precise_refs else 'conservative_all_attributes'
    return report
