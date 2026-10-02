"""General quality diagnostics distilled from reviewer feedback, not answer facts.
No reference/teacher document text or issuer-specific values are consumed here.
Diagnostics do not authorize edits or replace independent investigations.
"""
import re

def inspect(target,sources):
    issues=[];source_text='\n'.join(b.text for s in sources for b in s.blocks);target_text='\n'.join(b.text for b in target.blocks)
    def add(rule,severity,b,title,detail):
        issues.append({'rule':rule,'severity':severity,'file':target.name,'block':b.index if b else None,'matter':b.matter if b else '', 'locator':' > '.join(b.path[-4:]) if b else '全文','title':title,'detail':detail})
    # Visible draft markers only; historical comments are not conflated with live document facts.
    for b in target.blocks:
        if b.is_toc:continue
        if re.search(r'待更新|待补充|需要VIP|需要\s*VIP|未找到|靠你了老师|【\s*】|\[\s*待填\s*\]',b.text):
            add('visible-placeholder','warning',b,'正文仍有未完成标记',b.text[:240]+'。这属于原稿待办，不因为来源匹配成功而消失。')
        scope=' '.join(b.path)
        if '不存在对外担保' in b.text and re.search(r'违规.*担保|担保.*违规|关联方.*担保',scope):
            add('guarantee-scope','warning',b,'担保判断的范围不能互换','本事项关注关联方或违规担保，原文却使用“不存在对外担保”。无对外担保、无违规担保和未向关联方担保不是同一命题；保留原句并要求独立核验，不自动改写结论。')
    # Compare source coverage within the relevant target matter, not against every table in the book.
    related=[m for m in target.matters if '关联交易' in m['title']]
    for m in related:
        blocks=target.blocks[m['start']:m['end']];tx='\n'.join(b.text for b in blocks)
        topics=[('采购商品或接受劳务',r'采购商品.{0,4}接受劳务|购买商品.{0,4}接受劳务|采购商品及接受劳务'),('应收关联方款项',r'应收关联方款项|关联方应收款项'),('应付关联方款项',r'应付关联方款项|关联方应付款项')]
        for topic,pattern in topics:
            present=[b for s in sources for b in s.blocks if re.search(pattern,b.text[:160]) and '关联' in ' '.join(b.path)]
            if present and not re.search(pattern,tx):
                sb=present[0]
                add('related-coverage','warning',blocks[0],'来源有披露，但本事项未见对应标题：'+topic,'来源位置：'+' > '.join(sb.path[-4:])+'。可能是旧稿漏项，也可能采用了其他表述，需要确认；不会仅因缺标题就自动追加整张表。')
    # The reviewer preferred business-oriented organization. This is a suggestion, not a truth rule.
    revenue=[b for b in target.blocks if b.heading and re.search(r'营业收入分析',b.text)]
    costs=[b for b in target.blocks if b.heading and re.search(r'营业成本分析',b.text)]
    seen=set()
    for b in revenue:
        if b.matter in seen:continue
        if any(c.matter==b.matter for c in costs):
            seen.add(b.matter);add('business-organization','suggestion',b,'可按业务板块合并收入、成本和毛利分析','目前按财务指标分段。老师反馈曾将同一业务的收入、成本、毛利放在一起，减少跨段重复。这里只提示组织方式，不擅自重排现有判断或复制老师业务名称。')
    return issues
