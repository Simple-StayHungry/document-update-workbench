"""Turn one human-approved missing item into parameters for the existing insert.

This is deliberately the narrowest possible bridge.  It invents no Word writing
logic: it validates the frozen fingerprints, derives a title in the target's own
numbering style, and returns a proposal shaped exactly like the ones
``recover_source_additions`` already produces, so the ordinary ``source_insert``
channel does the actual writing.

``derive_new_item_title`` may only be reached from here.  No existing heading is
ever re-numbered, re-derived or normalised: this module never looks at items
that already exist in the workpaper.
"""
from __future__ import annotations
import copy
import re
from .docxio import w
from .matching import title_key
from .model import strip_lead
from .precision import fingerprint

LEAD = re.compile(r'^\s*(?:[（(]([^）)]*)[）)]|(\d+(?:[.．]\d+)*)\s*[、.．)）])')
DIGITS = re.compile(r'\d+')


def derive_new_item_title(previous_sibling_heading, source_item_heading):
    """Number a NEW item the way the target already numbers its siblings.

    Returns None when the sibling pattern cannot be read, so the caller can fall
    back to the source wording instead of guessing a format.
    """
    if not previous_sibling_heading or not source_item_heading:
        return None
    match = LEAD.match(previous_sibling_heading)
    if not match:
        return None
    current = match.group(1) or match.group(2)
    if not current or not DIGITS.fullmatch(current.strip()):
        return None
    body = strip_lead(source_item_heading)
    if not body:
        return None
    number = str(int(current) + 1)
    if match.group(1) is not None:
        return '（%s）%s' % (number, body)
    return '%s、%s' % (number, body)


def decision_fields(candidate):
    """The fingerprints a stored human decision is bound to."""
    return {'candidate_id': candidate.get('candidate_id', ''),
            'source_hash': candidate.get('source_hash', ''),
            'source_body_hash': candidate.get('source_body_hash', ''),
            'parent_fingerprint': candidate.get('parent_fingerprint', ''),
            'previous_sibling_fingerprint': candidate.get('previous_sibling_fingerprint', ''),
            'next_boundary_fingerprint': candidate.get('next_boundary_fingerprint', '')}


def validate(candidate, decision, doc=None):
    """Refuse to reuse a decision once the source or the anchors moved.

    The candidate is re-detected on every run, so it carries the fingerprints of
    the workpaper as it stands now.  A stored decision is compared against that,
    not merely searched for somewhere in the file: an anchor that drifted to a
    different position is a changed anchor.  Empty means the frozen evidence
    still holds; otherwise the item goes back to human review.
    """
    why = []
    if decision.get('candidate_id') != candidate.get('candidate_id'):
        why.append('候选标识与已确认记录不一致')
    for key, label in (('source_hash', '来源文件'),
                       ('source_body_hash', '来源正文'),
                       ('parent_fingerprint', '目标父事项'),
                       ('previous_sibling_fingerprint', '前一相邻事项'),
                       ('next_boundary_fingerprint', '后一边界')):
        frozen = decision.get(key)
        if frozen and frozen != candidate.get(key):
            why.append('%s已变化，需重新确认' % label)
    return why


def create_missing_item_heading(target_doc, shell_index, approved_title):
    """Create the target-side heading paragraph for one approved missing item.

    The style shell is copied from the target's own sibling heading — paragraph
    properties, run properties, spacing, indentation — never from the source.
    Only the visible text becomes the approved title. Automatic numbering is
    stripped so the number cannot appear twice; the approved title carries it.
    """
    if not approved_title or shell_index is None:
        return None
    if not 0 <= shell_index < len(target_doc.blocks):
        return None
    shell = copy.deepcopy(target_doc.blocks[shell_index].el)
    props = shell.find(w('pPr'))
    if props is not None:
        numbering = props.find(w('numPr'))
        if numbering is not None:
            props.remove(numbering)
    texts = [t for t in shell.iter(w('t'))]
    if not texts:
        return None
    texts[0].text = approved_title
    for t in texts[1:]:
        t.text = ''
    return shell


def build_proposal(candidate, decision, doc, source):
    """One approved missing item, expressed as an ordinary source_insert plan."""
    # Only the body is copied. The export guard forbids a source range that
    # contains a heading, so the derived title stays a suggestion until a
    # separate, explicit decision allows writing one.
    start, end = candidate['source_body_range']
    body = [b for b in source.blocks[start:end] if not b.heading and not b.is_toc]
    if not body:
        return None
    title = decision.get('approved_title') or derive_new_item_title(
        candidate.get('previous_sibling_heading'), candidate.get('source_item_heading'))
    indices = [b.index for b in body]
    position = candidate['suggested_insert_position']
    if not 0 <= position <= len(doc.blocks):
        return None
    cand = {'doc_hash': source.hash, 'source_name': source.name, 'start': body[0].index,
            'end': body[-1].index + 1, 'indices': indices, 'unit_id': 'missing' + str(start),
            'text': '\n'.join(b.text for b in body), 'score': 1.0, 'literal_score': 1.0,
            'content_score': 1.0, 'heading_score': 1.0, 'table_score': 0.0, 'exact': False,
            'locator': candidate.get('source_parent_heading', ''),
            'unsupported': '', 'source_warnings': [], 'scope_verified': True,
            'source_fingerprint': fingerprint([b.el for b in body]),
            'correspondence': {'method': 'approved-missing-item', 'anchors': [],
                               'target_heading': None, 'source_heading': None}}
    return {'id': doc.hash + ':missing:' + candidate['candidate_id'],
            'target_id': doc.hash, 'target_name': doc.name,
            'start': position, 'end': position,
            'insert_before': position if position < len(doc.blocks) else None,
            'insert_after': position - 1 if position else None,
            'insert_before_fingerprint': fingerprint([doc.blocks[position].el]) if position < len(doc.blocks) else None,
            'insert_after_fingerprint': fingerprint([doc.blocks[position - 1].el]) if position else None,
            'matter': '正文', 'heading': candidate.get('target_parent_heading', ''),
            'locator': candidate.get('target_parent_heading', ''), 'old_text': '',
            'status': 'auto', 'decision': 'accept', 'selected': 0, 'candidates': [cand],
            'action': 'source_insert', 'anchor_sources': {}, 'kind': 'material',
            'content_class': 'approved_missing_item',
            # The target-side heading is created here, from the target's own
            # style shell; the copied source range stays heading-free.
            'missing_heading': create_missing_item_heading(
                doc, candidate.get('previous_sibling_index')
                if candidate.get('previous_sibling_index') is not None
                else candidate.get('next_boundary_index'), title),
            'suggested_title': title or '',
            'approved_title': title or '',
            # Frozen evidence snapshot: the export audit re-checks these against
            # the persisted human decision before accepting the extra heading.
            'decision_fingerprints': dict(decision_fields(candidate)),
            'reason': '人工确认的来源新增事项；按目标同级体例生成标题，已有标题与编号不变。',
            'has_table': False, 'table_mode': 'copy', 'sync_mode': 'block', 'identity': None, 'note': ''}


def primary_evidence(rows):
    """The one evidence whose fingerprints a decision is bound to.

    Several formal sources can report the same item. They share one identity, so
    one representative must carry the version state; the prospectus wins, matching
    the existing authority rule.
    """
    return next((r for r in rows if r.get('source_kind') == 'prospectus'), rows[0])


def approved_inserts(result, files, decisions, engine):
    """Proposals for every still-valid approved decision, plus the stale ones.

    Decisions are keyed by the stable candidate id, so approving an item once
    covers every source that reported it. The prospectus copy is the one written,
    and a same-level conflict is refused rather than inserted twice.
    """
    by_name = {f['id']: f for f in files}
    groups, order = {}, []
    for record in result.get('documents', []):
        doc_file = by_name.get(record.get('id'))
        if not doc_file:
            continue
        doc = engine.load(doc_file)
        for candidate in record.get('missing_candidates') or []:
            cid = candidate.get('candidate_id')
            if cid not in groups:
                groups[cid] = {'doc': doc, 'rows': []}
                order.append(cid)
            groups[cid]['rows'].append(candidate)
    plans, stale = [], []
    for cid in order:
        entry = groups[cid]
        decision = decisions.get(cid)
        if not decision or decision.get('decision') != 'add':
            continue
        rows, doc = entry['rows'], entry['doc']
        chosen = primary_evidence(rows)
        why = validate(chosen, decision, doc)
        if why:
            stale.append({'candidate_id': cid, 'target_name': doc.name, 'errors': why})
            continue
        same_level = [r for r in rows if r.get('source_kind') == chosen.get('source_kind')]
        if len(same_level) > 1:
            stale.append({'candidate_id': cid, 'target_name': doc.name,
                          'errors': ['同一新增事项存在多份同级别来源，未唯一定位，需人工核对']})
            continue
        source = next((engine.load(by_name[f['id']]) for f in files
                       if f.get('role') == 'source'
                       and engine.load(by_name[f['id']]).hash == chosen.get('source_hash')), None)
        if source is None:
            stale.append({'candidate_id': cid, 'target_name': doc.name,
                          'errors': ['来源文件已不在项目中']})
            continue
        plan = build_proposal(chosen, decision, doc, source)
        if plan is None:
            stale.append({'candidate_id': cid, 'target_name': doc.name,
                          'errors': ['无法按冻结位置生成插入计划']})
            continue
        plans.append(plan)
    plans.sort(key=lambda p: (p['start'], p['id']))
    return plans, stale
