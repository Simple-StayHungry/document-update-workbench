"""Read-only detection of source items that have no sibling in the workpaper.

Stage one only.  This module never touches the synchronisation path: it takes
documents, returns brand new dictionaries, and writes nothing back.  Whether a
found item is ever inserted is somebody else's decision.

Identity rules, deliberately narrow:

* An item is identified by its *structural position* inside a parent matter,
  plus the normalised key of its own title.  Similar body text elsewhere is
  evidence, never proof of existence, so a repeated amount in another chapter
  cannot hide a missing risk item here.
* ``title_key`` only generates candidate parents.  Chinese outline numbering
  collapses easily ("（一）财务风险" and "1、财务风险" share a key), so pairing is
  decided by how many direct children actually overlap, and the best score wins.
  An apparent parent with no real overlap is not a parent.
"""
from __future__ import annotations
from .matching import title_key, normal
from .precision import fingerprint

MIN_SIBLINGS = 3
RATIO = 0.6
RULE = 'same-level-missing-item'


def _headings(doc):
    return [b for b in doc.blocks if b.heading and not b.is_toc]


def _level(block):
    return block.level if block.level is not None else 99


def _direct_children(headings, index):
    """Headings of the shallowest level below ``headings[index]``."""
    parent = headings[index]
    inside = []
    for block in headings[index + 1:]:
        if _level(block) <= _level(parent):
            break
        inside.append(block)
    if not inside:
        return []
    shallowest = min(_level(b) for b in inside)
    return [b for b in inside if _level(b) == shallowest]


def _subtree_end(headings, index):
    """Index just past every heading nested under ``headings[index]``."""
    parent = headings[index]
    for block in headings[index + 1:]:
        if _level(block) <= _level(parent):
            return block.index
    return None


def _support_needed(count):
    if count < MIN_SIBLINGS:
        return count
    return max(MIN_SIBLINGS, -(-int(RATIO * 100 * count) // 100))


def _keys(blocks):
    return [title_key(b.text) for b in blocks]


def _pair_parents(source_headings, source_childgens, target_headings, target_children):
    """Return index pairs whose direct children really agree, best score wins."""
    pairs = []
    for i, children in enumerate(source_childgens):
        if not children:
            continue
        wanted = set(_keys(children))
        best = None
        for j, others in enumerate(target_children):
            if not others:
                continue
            score = len(wanted & set(_keys(others)))
            if best is None or score > best[0]:
                best = (score, j)
        if not best or best[0] < _support_needed(len(children)):
            continue
        pairs.append((i, best[1], best[0]))
    return pairs


def _signature_in(doc, body):
    """Whether the leading body line already appears anywhere in this workpaper.

    Evidence only.  Another chapter carrying the same wording never proves that
    the item exists under this parent, so this annotates a candidate instead of
    filtering it out.
    """
    if not body:
        return False
    signature = normal(body[0].text)[:40]
    return bool(signature) and any(normal(b.text)[:40] == signature for b in doc.blocks)


def detect_missing_items(doc, sources):
    """List source siblings missing under their paired target parent.

    Each row carries everything a later human decision needs: what is missing,
    where it belongs, which surviving siblings prove that position, and frozen
    fingerprints so a moved outline is refused rather than guessed.
    """
    rows = []
    if not sources:
        return rows
    headings = _headings(doc)
    target_children = [_direct_children(headings, j) for j in range(len(headings))]
    for sd in sources:
        if not sd or sd.profile.get('kind') not in ('prospectus', 'opinion'):
            continue
        if sd.profile.get('issuer') != doc.profile.get('issuer'):
            continue
        source_headings = _headings(sd)
        childgens = [_direct_children(source_headings, i) for i in range(len(source_headings))]
        for si, ti, score in _pair_parents(source_headings, childgens, headings, target_children):
            src_parent = source_headings[si]
            tgt_parent = headings[ti]
            src_items = childgens[si]
            tgt_items = target_children[ti]
            have = set(_keys(tgt_items))
            order = _keys(src_items)
            for n, key in enumerate(order):
                if key in have:
                    continue
                item = src_items[n]
                previous = next((src_items[k] for k in range(n - 1, -1, -1) if order[k] in have), None)
                following = next((src_items[k] for k in range(n + 1, len(src_items)) if order[k] in have), None)
                prev_target = None
                next_target = None
                if previous is not None:
                    prev_target = next((b for b in tgt_items if title_key(b.text) == title_key(previous.text)), None)
                if following is not None:
                    next_target = next((b for b in tgt_items if title_key(b.text) == title_key(following.text)), None)
                # When the source item ends its parent matter, the surviving lower
                # boundary is whatever closes the matching target parent.
                boundary = next_target or next(
                    (b for b in headings[ti + 1:] if _level(b) <= _level(tgt_parent)), None)
                if boundary is not None:
                    position = boundary.index
                elif prev_target is not None:
                    end = _subtree_end(headings, headings.index(prev_target))
                    position = end if end is not None else len(doc.blocks)
                else:
                    position = tgt_parent.index + 1
                parent_end = _subtree_end(source_headings, si) or len(sd.blocks)
                limit = src_items[n + 1].index if n + 1 < len(src_items) else parent_end
                body = [b for b in sd.blocks[item.index + 1:limit] if not b.heading and not b.is_toc]
                # The chapter identity comes from the document's own content, not
                # from the uploaded file name and never from a hash: renaming
                # "第九章.docx" must not orphan the decisions made about it.
                chapter = doc.profile.get('chapter') or doc.name
                # The parent key must be the whole structural path, not the nearest
                # heading text: two different structures in one chapter can both
                # contain a heading called 财务风险, and keying on the short text
                # would make them one identity.
                parent_path = ' > '.join(title_key(x) for x in tgt_parent.path) or title_key(tgt_parent.text)
                rows.append({
                    'rule': RULE,
                    # Three layers, kept strictly apart: this id says what the item
                    # is and where it goes; the evidences say which formal sources
                    # prove it; the fingerprints below say whether the content is
                    # still the version a human approved.
                    'target_doc_key': chapter,
                    'target_parent_key': parent_path,
                    'candidate_id': chapter + '|' + parent_path + '|' + key,
                    'source_name': sd.name,
                    'source_hash': sd.hash,
                    'source_kind': sd.profile.get('kind'),
                    'source_parent_heading': src_parent.text,
                    'source_item_heading': item.text,
                    'source_item_key': key,
                    'source_item_index': item.index,
                    'source_body_range': [body[0].index, body[-1].index + 1] if body else [item.index, item.index + 1],
                    'source_text': '\n'.join(b.text for b in body),
                    'source_body_hash': fingerprint([b.el for b in body]) if body else '',
                    'target_name': doc.name,
                    'target_parent_heading': tgt_parent.text,
                    'previous_sibling_heading': prev_target.text if prev_target else '',
                    'next_boundary_heading': boundary.text if boundary else '',
                    'previous_sibling_index': prev_target.index if prev_target else None,
                    'next_boundary_index': boundary.index if boundary else None,
                    'previous_sibling_fingerprint': fingerprint([prev_target.el]) if prev_target else '',
                    'next_boundary_fingerprint': fingerprint([boundary.el]) if boundary else '',
                    'parent_fingerprint': fingerprint([tgt_parent.el]),
                    'suggested_insert_position': position,
                    'matched_sibling_count': score,
                    'sibling_total': len(src_items),
                    'evidence': {'body_exists_elsewhere_in_target': _signature_in(doc, body)},
                    'reason': '同一父事项下来源存在该同级事项，目标事项组缺少对应项，尚未写入文档。',
                    'status': 'pending',
                })
    rows.sort(key=lambda r: (r['suggested_insert_position'], r['candidate_id']))
    return rows
