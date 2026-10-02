"""User-authorized editorial voice adaptation, never factual rewriting."""
from __future__ import annotations
import re
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def w(n):return '{'+W+'}'+n

POLICY='workpaper-voice-2026-09-23-v1'
DATES=('本募集说明书签署之日','本募集说明书签署日','募集说明书签署之日','募集说明书签署日',
       '本募集说明书出具之日','本募集说明书出具日','募集说明书出具之日','募集说明书出具日',
       '本核查意见出具之日','本核查意见出具日','核查意见出具之日','核查意见出具日',
       '本核查文件出具之日','本核查文件出具日','本核查分析文件出具之日')

def mappings(value,source_kind):
    result=[(old,'本核查分析文件出具日') for old in DATES if old in value]
    # Prospectus is issuer voice. In an opinion the company's own first person
    # remains GF Securities, except an explicitly introduced issuer declaration.
    issuer_quote=bool(re.search(r'发行人[^。]{0,80}(?:声明|承诺)[：:]?[“「]',value))
    if '本公司' in value and (source_kind=='prospectus' or issuer_quote):result.append(('本公司','公司'))
    # Preserve references to source chapters: the workpaper has different numbering.
    if '本募集说明书' in value:result.append(('本募集说明书','募集说明书'))
    if '本核查意见' in value:result.append(('本核查意见','核查意见'))
    return result

def adapt(elements,source_kind):
    """Adapt fresh copied XML in place; return reproducible, per-paragraph evidence."""
    from .docxio import transform_paragraph
    changes=[]
    for i,el in enumerate(elements):
        paras=list(el.iter(w('p')))
        for j,p in enumerate(paras):
            # Nested text-box paragraphs have a different semantic scope. Only
            # select this paragraph's own runs, while preserving all run styles.
            parents={c:a for a in p.iter() for c in a}
            ts=[]
            for t in p.iter(w('t')):
                node=parents.get(t)
                while node is not None and node.tag!=w('p'):node=parents.get(node)
                if node is p:ts.append(t)
            before=''.join(t.text or '' for t in ts);rules=mappings(before,source_kind)
            if not rules:continue
            # A temporary paragraph references the same text nodes only, avoiding
            # concatenation across nested objects or separate table-cell paragraphs.
            from xml.etree.ElementTree import Element
            view=Element(w('p'))
            for t in ts:view.append(t)
            transform_paragraph(view,rules)
            after=''.join(t.text or '' for t in ts)
            if before!=after:changes.append({'element':i,'paragraph':j,'before':before,'after':after,'rules':[list(x) for x in rules]})
    return changes
