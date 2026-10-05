from difflib import SequenceMatcher
from decimal import Decimal
from .models import Match, Issue
from .normalize import normalize


def key(item):
    name, spec = normalize(item.name), normalize(item.spec)
    return name if not spec or spec in name else name + spec


def reconcile(quotes, deliveries):
    matches = []
    for d in deliveries:
        m = Match(d)
        if d.validation_errors or d.quantity is None or not d.name.strip():
            m.status = '导入记录待修正'
            matches.append(m)
            continue
        if d.code:
            candidates = [q for q in quotes if q.code and normalize(q.code) == normalize(d.code)]
        else:
            candidates = []
        by_code = bool(candidates)
        if not candidates:
            candidates = [q for q in quotes if key(q) == key(d)]
        compatible = [q for q in candidates if normalize(q.unit) == normalize(d.unit)]
        if len(compatible) == 1:
            q = compatible[0]
            # 编码冲突时名称/规格差异不能被编码相等掩盖。
            if by_code and key(q) != key(d):
                m.status = '编号一致但名称/规格冲突'
            elif q.validation_errors:
                m.status = '报价记录待修正'
            elif q.price is None:
                m.status = '报价缺价'
            else:
                m.quote, m.status = q, '自动匹配'
        elif len(compatible) > 1:
            m.status = '多个报价候选'
        elif candidates:
            m.status = '单位不一致'
        if not m.ready:
            ranked = sorted(quotes, key=lambda q: SequenceMatcher(None, key(d), key(q)).ratio(), reverse=True)
            m.candidates = candidates + [q for q in ranked if q not in candidates][:12]
        if any(word in d.note for word in ('赠送', '退回', '退货')):
            if m.quote and m.quote not in m.candidates:
                m.candidates.insert(0, m.quote)
            m.status = '结算备注待确认'
        matches.append(m)
    return matches


def confirm(match, quote, factor, override, reason):
    if match.delivery.validation_errors or quote.validation_errors or match.delivery.quantity is None:
        raise ValueError('请先通过“修改已导入记录”修正记录错误，再确认商品匹配。')
    if not reason.strip():
        raise ValueError('请填写确认原因，便于追溯。')
    if not factor.is_finite() or factor <= 0:
        raise ValueError('单位倍率必须大于 0。')
    if override is not None and (not override.is_finite() or override < 0):
        raise ValueError('有效单价必须是非负数。')
    if quote.price is None and override is None:
        raise ValueError('所选商品缺价，请填写有效单价。')
    match.quote, match.factor, match.override_price = quote, factor, override
    match.status, match.reason = '人工确认', reason.strip()


def match_issues(matches):
    issues = []
    for m in matches:
        d, q = m.delivery, m.quote
        if not m.ready:
            issues.append(Issue('商品匹配', d.source, f'{d.serial} · {d.name}：{m.status}', '双击记录选择正确报价；缺少商品时补充报价并重导。'))
        elif q.effective and d.day and d.day < q.effective:
            issues.append(Issue('报价生效日期', d.source, f'送货日期 {d.day} 早于报价日期 {q.effective}', '确认允许按当前导入报价核算，或换用当期报价。'))
        if m.ready and d.price is not None and m.price != d.price:
            issues.append(Issue('单价差异', d.source, f'原送货单价 {d.price}，本次有效单价 {m.price}', '核对后确认采用本次价格。'))
    return issues
