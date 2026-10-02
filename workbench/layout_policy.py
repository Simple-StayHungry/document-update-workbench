"""Conservative preflight for whole source tables requiring narrower layout.

No font engine is used: widths below are a documented estimate for inserted
Times New Roman numeric text, NOT a measurement or a rendering certificate.
Unresolved sizes are retained for review. This module never mutates XML.
"""
from __future__ import annotations
import re
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NUMERIC = re.compile(r'[+\-−]?(?:\d{4}[-/.]\d{1,2}(?:[-/.]\d{1,2})?|\d{4}年\d{1,2}月(?:\d{1,2}日)?|\d+(?:,\d{3})*(?:\.\d+)?)[%％]?')
ESTIMATOR = 'Times New Roman em-width estimate with 8% allowance; not measured font metrics'
PAPER_EDGE_CLEARANCE = 360  # 0.25 inch, after any floating-table text clearance.


def q(n):
    return '{' + W + '}' + n


def val(parent, child, default=None):
    e = None if parent is None else parent.find(q(child))
    return default if e is None else e.get(q('val'), default)


def _int(v):
    if v is None:
        raise ValueError('缺少尺寸定义')
    return int(v)


def target_text_width(root, anchor):
    parents = {c: p for p in root.iter() for c in p}
    body = root.find(q('body')); top = anchor
    while top in parents and parents[top] is not body:
        top = parents[top]
    if body is None or top not in list(body):
        raise ValueError('无法定位目标节')
    section = None
    for block in list(body)[list(body).index(top):]:
        section = (block.find('./' + q('pPr') + '/' + q('sectPr')) if block.tag == q('p')
                   else block if block.tag == q('sectPr') else None)
        if section is not None:
            break
    if section is None:
        raise ValueError('目标节缺少页面尺寸')
    size = section.find(q('pgSz')); margin = section.find(q('pgMar'))
    if size is None or margin is None:
        raise ValueError('目标节尺寸或页边距不完整')
    width = _int(size.get(q('w'))) - _int(margin.get(q('left'))) - _int(margin.get(q('right'))) - _int(margin.get(q('gutter'), '0'))
    cols = section.find(q('cols'))
    if cols is not None:
        explicit = cols.findall(q('col'))
        if explicit:
            width = min(_int(c.get(q('w'))) for c in explicit)
        else:
            count = max(1, _int(cols.get(q('num'), '1')))
            width = (width - (count - 1) * _int(cols.get(q('space'), '720'))) // count
    current = anchor
    while current in parents:
        current = parents[current]
        if current.tag == q('tc'):
            cw = current.find('./' + q('tcPr') + '/' + q('tcW'))
            if cw is not None and cw.get(q('type')) == 'dxa':
                width = min(width, _int(cw.get(q('w'))))
    if width <= 0:
        raise ValueError('目标节有效宽度非正数')
    return width


def table_has_preceding_paragraph(elements, table):
    index = next((i for i, el in enumerate(elements) if el is table), -1)
    if index < 1 or elements[index - 1].tag != q('p'):
        return False
    return bool(''.join(n.text or '' for n in elements[index - 1].iter()
                        if n.tag in (q('t'), q('delText'))).strip())


def floating_table_anchor_reason(old_elements, source_elements, old, source):
    """Guard separate old/new anchors before either native fitting or shrinking."""
    floating = [t.find('./' + q('tblPr') + '/' + q('tblpPr'))
                for t in (old, source) if t is not None]
    floating = [p for p in floating if p is not None]
    if not floating:
        return ''
    if any(p.get(q('vertAnchor')) != 'text' for p in floating):
        return '浮动表纵向锚点不是 text；绝对页面或页边距定位可能使删除表与新增表重叠'
    if (old is None or not table_has_preceding_paragraph(old_elements, old) or
            not table_has_preceding_paragraph(source_elements, source)):
        return '浮动表的新旧完整范围均须包含独立前置文字段，避免删除表与新增表共用锚点'
    return ''


def preserved_native_table_layout(root, old, source, allow_floating=False):
    """Prove a narrow exception to text-margin fitting, without changing XML.

    Some originals and their matching source tables deliberately extend into the
    margins. Keep that SAME fixed grid/placement only when it remains inside the
    actual single-column page with a conservative edge clearance. This is not
    permission to place arbitrary source tables in margins or to change sections.
    Floating-table vertical interactions still require review of actual output.
    """
    try:
        body = root.find(q('body'))
        if (body is None or old not in list(body) or old.tag != q('tbl') or
                source.tag != q('tbl') or old.find('.//' + q('tbl')) is not None or
                source.find('.//' + q('tbl')) is not None):
            return None
        section = None
        for block in list(body)[list(body).index(old):]:
            section = (block.find('./' + q('pPr') + '/' + q('sectPr')) if block.tag == q('p')
                       else block if block.tag == q('sectPr') else None)
            if section is not None:
                break
        if section is None:
            return None
        size = section.find(q('pgSz')); margins = section.find(q('pgMar'))
        columns = section.find(q('cols'))
        if (size is None or margins is None or
                (columns is not None and (columns.findall(q('col')) or
                                          int(columns.get(q('num'), '1')) != 1)) or
                int(margins.get(q('gutter'), '0')) != 0 or
                section.find(q('bidi')) is not None or section.find(q('textDirection')) is not None):
            return None
        page = int(size.get(q('w'))); left = int(margins.get(q('left'))); right = int(margins.get(q('right')))
        text_width = page - left - right
        # Symmetric margins make odd/even mirror-margin placement identical.
        if min(page, text_width) <= 0 or min(left, right) < 0 or left != right:
            return None
        og = [int(c.get(q('w'))) for c in old.findall('./' + q('tblGrid') + '/' + q('gridCol'))]
        sg = [int(c.get(q('w'))) for c in source.findall('./' + q('tblGrid') + '/' + q('gridCol'))]
        op = old.find(q('tblPr')); sp = source.find(q('tblPr'))
        if not og or len(og) != len(sg) or min(og + sg) <= 0 or op is None or sp is None:
            return None
        def geometry(node):
            if node is None:
                return None
            return (node.tag, tuple(sorted(node.attrib.items())), tuple(geometry(c) for c in node))
        names = ('tblW', 'tblInd', 'jc', 'tblLayout', 'tblpPr', 'tblCellSpacing', 'bidiVisual')
        if any(geometry(op.find(q(n))) != geometry(sp.find(q(n))) for n in names):
            return None
        layout = op.find(q('tblLayout')); preferred = op.find(q('tblW'))
        if (layout is None or layout.get(q('type')) != 'fixed' or preferred is None or
                op.find(q('bidiVisual')) is not None):
            return None
        width_type = preferred.get(q('type')); preferred_value = int(preferred.get(q('w')))
        if width_type == 'dxa':
            if og != sg or preferred_value != sum(og):
                return None
            occupied_width = sum(sg)
            method = 'identical-native-grid-and-placement-within-paper'
        elif width_type == 'pct':
            # Percentage widths are based on the section text width, not on the
            # stored absolute grid sum. Allow only the same percentage/placement
            # and a tightly bounded serialization rounding difference. Preserve
            # the SOURCE grid verbatim; this is not geometric equality.
            if (preferred_value <= 0 or op.find(q('tblpPr')) is not None or
                    any(abs(a - b) > 2 or abs(a - b) / max(a, b) > .002 for a, b in zip(og, sg))):
                return None
            occupied_width = max(sum(og), sum(sg), text_width * preferred_value / 5000)
            method = 'same-percentage-placement-source-grid-rounding-within-paper'
        else:
            return None
        spacing = op.find(q('tblCellSpacing'))
        if spacing is not None and (spacing.get(q('type')) != 'dxa' or int(spacing.get(q('w'))) != 0):
            return None
        indent = op.find(q('tblInd'))
        if indent is not None and indent.get(q('type')) != 'dxa':
            return None
        inset = int(indent.get(q('w'), '0')) if indent is not None else 0
        floating = op.find(q('tblpPr')); alignment = val(op, 'jc', 'left')
        outer_left = outer_right = 0
        if floating is not None:
            if (not allow_floating or floating.get(q('horzAnchor')) != 'margin' or floating.get(q('tblpXSpec')) != 'center' or
                    floating.get(q('vertAnchor')) != 'text' or floating.get(q('tblpX')) is not None or inset != 0):
                return None
            alignment = 'center'
            outer_left = max(0, int(floating.get(q('leftFromText'), '0')))
            outer_right = max(0, int(floating.get(q('rightFromText'), '0')))
        if alignment == 'center' and inset == 0:
            start = left + (text_width - occupied_width) / 2
        elif alignment in ('left', 'start') and floating is None:
            start = left + inset
        else:
            return None
        end = start + occupied_width
        if start - outer_left < PAPER_EDGE_CLEARANCE or page - end - outer_right < PAPER_EDGE_CLEARANCE:
            return None
        return {'method': method, 'copy_source_grid': width_type == 'pct',
                'grid_width_twips': sum(sg), 'target_text_width_twips': text_width,
                'preferred_width_twips': text_width * preferred_value / 5000 if width_type == 'pct' else preferred_value,
                'source_grid_rounding_differences_twips': [b - a for a, b in zip(og, sg)],
                'table_left_twips': start, 'table_right_twips': end, 'page_width_twips': page,
                'paper_edge_clearance_twips': PAPER_EDGE_CLEARANCE,
                'fit_width_twips': sum(sg) + max(0, inset), 'floating': floating is not None}
    except (ValueError, TypeError, AttributeError):
        return None


class Sizes:
    def __init__(self, entries):
        raw = entries.get('word/styles.xml')
        self.root = ET.fromstring(raw) if raw else None
        self.styles = {} if self.root is None else {s.get(q('styleId')): s for s in self.root.findall(q('style'))}
        self.cache = {}
        self.default_paragraph = next((k for k, e in self.styles.items() if e.get(q('type')) == 'paragraph' and e.get(q('default')) in {'1', 'true', 'on'}), None)
        self.default_table = next((k for k, e in self.styles.items() if e.get(q('type')) == 'table' and e.get(q('default')) in {'1', 'true', 'on'}), None)

    def style_size(self, style_id, seen=()):
        if not style_id or style_id in seen:
            return None
        if style_id in self.cache:
            return self.cache[style_id]
        style = self.styles.get(style_id)
        if style is None:
            return None
        result = val(style.find(q('rPr')), 'sz')
        if result is None:
            result = self.style_size(val(style, 'basedOn'), seen + (style_id,))
        self.cache[style_id] = result
        return result

    def run_size(self, run, paragraph, table):
        rp = run.find(q('rPr')); pp = paragraph.find(q('pPr')); tp = table.find(q('tblPr'))
        sources = [val(rp, 'sz'), self.style_size(val(rp, 'rStyle')),
                   val(None if pp is None else pp.find(q('rPr')), 'sz'),
                   self.style_size(val(pp, 'pStyle') or self.default_paragraph), self.style_size(val(tp, 'tblStyle') or self.default_table)]
        if self.root is not None:
            defaults = self.root.find('./' + q('docDefaults') + '/' + q('rPrDefault') + '/' + q('rPr'))
            sources.append(val(defaults, 'sz'))
        found = next((x for x in sources if x is not None), None)
        if found is None:
            raise ValueError('数字字符字号未能从来源直接格式或样式可靠解析')
        # OOXML size is in half-points, i.e. ten twips per unit.
        size = float(found)
        if not 4 <= size <= 144:
            raise ValueError('来源字号超出可靠估算范围')
        if rp is not None and any(rp.find(q(k)) is not None for k in ('fitText', 'w', 'spacing', 'position')):
            raise ValueError('数字字符含压缩、字距或位置格式，估算可靠性不足')
        return size * 10


def _numeric_runs(paragraph, table, sizes):
    text = ''; fonts = []
    for run in paragraph.findall(q('r')):
        value = ''.join(e.text or '' for e in run.findall(q('t')))
        # Runs unrelated to numeric tokens may have no font size. Defer resolution
        # until a numeric/date token actually overlaps the run.
        start = len(text); text += value
        fonts.append((start, len(text), run))
    for token in NUMERIC.finditer(text):
        widths = []
        for start, end, run in fonts:
            a, b = max(start, token.start()), min(end, token.end())
            if a >= b:
                continue
            em = sizes.run_size(run, paragraph, table)
            for char in text[a:b]:
                factor = .5 if char.isascii() and char.isdigit() else .25 if char in ',.' else .85 if char in '%％' else .36 if char in '-−+/' else 1.0
                widths.append(em * factor)
        yield token.group(), sum(widths) * 1.08


def _cell_margin(cell, table, sizes):
    cp = cell.find(q('tcPr')); tp = table.find(q('tblPr'))
    cm = None if cp is None else cp.find(q('tcMar'))
    tm = None if tp is None else tp.find(q('tblCellMar'))
    style_margins = []
    style_id = val(tp, 'tblStyle') or sizes.default_table
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id); style = sizes.styles.get(style_id)
        if style is None:
            raise ValueError('单元格边距所依赖的表格样式缺失')
        if any(e.find('.//' + q('tcMar')) is not None or e.find('.//' + q('tblCellMar')) is not None for e in style.findall(q('tblStylePr'))):
            raise ValueError('单元格边距受条件表格样式控制，估算可靠性不足')
        sp = style.find(q('tblPr'))
        if sp is not None:
            style_margins.append(sp.find(q('tblCellMar')))
        style_id = val(style, 'basedOn')
    sides = []
    for names in (('start', 'left'), ('end', 'right')):
        entry = None
        for parent in (cm, tm, *style_margins):
            if parent is not None:
                entry = next((parent.find(q(n)) for n in names if parent.find(q(n)) is not None), None)
            if entry is not None:
                break
        if entry is None:
            sides.append(108)  # Word's default table cell side margin, twips.
        elif entry.get(q('type'), 'dxa') == 'dxa':
            sides.append(max(0, _int(entry.get(q('w')))))
        elif entry.get(q('type')) == 'nil':
            sides.append(0)
        else:
            raise ValueError('单元格边距不是可可靠换算的固定尺寸')
    return sum(sides)


def table_layout_risks(target_pkg, old_elements, source_pkg, source_elements, insertion_anchor=None):
    """Return layout_pending reasons; no throwing for an individual bad table.

    Same-count/same-grid-column tables inherit the old grid exactly as the writer
    does. Different grids use the complete source grid. Numeric/date tokens that
    previously fit but fail a narrower estimated cell are kept for review.
    """
    old_tables = [e for e in old_elements if e.tag == q('tbl')]
    new_tables = [e for e in source_elements if e.tag == q('tbl')]
    if not new_tables:
        return []
    mapping = dict(zip(map(id, new_tables), old_tables)) if len(old_tables) == len(new_tables) else {}
    risks = []; sizes = Sizes(source_pkg.entries)
    try:
        width = target_text_width(target_pkg.root, old_elements[0] if old_elements else insertion_anchor)
    except (ValueError, TypeError, KeyError) as exc:
        return [{'status': 'layout_pending', 'reason': '整表布局待核对：' + str(exc), 'estimator': ESTIMATOR}]
    for number, table in enumerate(new_tables, 1):
        detail = {'status': 'layout_pending', 'table_number': number, 'estimator': ESTIMATOR,
                  'target_text_width_twips': width}
        try:
            if table.find('.//' + q('tbl')) is not None:
                raise ValueError('含嵌套表，无法独立确认内部单元格可用宽度')
            sg = table.find(q('tblGrid'))
            source_grid = [] if sg is None else [_int(e.get(q('w'))) for e in sg.findall(q('gridCol'))]
            if not source_grid or min(source_grid) <= 0:
                raise ValueError('来源整表缺少有效完整网格宽度')
            grid = source_grid; old = mapping.get(id(table)); geometry = table
            anchor_reason = floating_table_anchor_reason(old_elements, source_elements, old, table)
            if anchor_reason:
                raise ValueError(anchor_reason)
            separate_anchors = (old is not None and table_has_preceding_paragraph(old_elements, old) and
                                table_has_preceding_paragraph(source_elements, table))
            if old is not None and preserved_native_table_layout(target_pkg.root, old, table, allow_floating=separate_anchors):
                continue
            if old is not None:
                og = old.find(q('tblGrid'))
                old_grid = [] if og is None else [_int(e.get(q('w'))) for e in og.findall(q('gridCol'))]
                if len(old_grid) == len(source_grid) and old_grid and min(old_grid) > 0:
                    grid = old_grid; geometry = old
            pr = geometry.find(q('tblPr')); ind = None if pr is None else pr.find(q('tblInd'))
            indent = _int(ind.get(q('w'), '0')) if ind is not None and ind.get(q('type')) == 'dxa' else 0
            available = width - max(0, indent)
            total = sum(grid)
            if total <= available:
                continue
            if available <= len(grid):
                raise ValueError('目标可用宽度不足以放置整表')
            scaled = [max(1, int(v * available / total)) for v in grid]
            scaled[-1] += available - sum(scaled)
            detail.update(source_grid_twips=source_grid, geometry_grid_twips=grid, scaled_grid_twips=scaled,
                          scale_ratio=round(available / total, 6))
            issues = []
            for row_index, row in enumerate(table.findall(q('tr')), 1):
                before = row.find('./' + q('trPr') + '/' + q('gridBefore'))
                col = _int(before.get(q('val'), '0')) if before is not None else 0
                for cell_index, cell in enumerate(row.findall(q('tc')), 1):
                    span_node = cell.find('./' + q('tcPr') + '/' + q('gridSpan'))
                    span = _int(span_node.get(q('val'), '1')) if span_node is not None else 1
                    if col < 0 or span < 1 or col + span > len(grid):
                        raise ValueError('来源合并单元格超出完整网格')
                    margin = _cell_margin(cell, table, sizes)
                    before_width = sum(grid[col:col+span]) - margin
                    after_width = sum(scaled[col:col+span]) - margin
                    for para in cell.findall(q('p')):
                        pp = para.find(q('pPr')); indentp = None if pp is None else pp.find(q('ind'))
                        para_indent = 0 if indentp is None else sum(max(0, _int(indentp.get(q(n), '0'))) for n in ('left', 'right', 'firstLine'))
                        for token, required in _numeric_runs(para, table, sizes):
                            if required + para_indent <= before_width and required + para_indent > after_width:
                                issues.append({'row': row_index, 'cell': cell_index, 'token': token,
                                               'estimated_required_twips': round(required + para_indent, 1),
                                               'before_inner_width_twips': before_width,
                                               'after_inner_width_twips': after_width})
                    col += span
            if issues:
                detail['numeric_fit_risks'] = issues[:20]
                detail['numeric_fit_risk_count'] = len(issues)
                detail['reason'] = ('整表布局待核对：目标节需缩宽，原可单行的数字或日期预计不再容纳'
                                    '（例如“' + issues[0]['token'] + '”）。保留原表，不缩字或重建；字宽为保守估算，仍需实际渲染。')
                risks.append(detail)
        except (ValueError, TypeError, KeyError, ET.ParseError) as exc:
            detail['reason'] = '整表布局待核对：' + str(exc) + '。保留原表，不以低置信尺寸自动缩宽。'
            risks.append(detail)
    return risks
