"""可编辑的导入草稿：先校验再提交，强制导入不丢失坏记录。"""
from copy import deepcopy
from datetime import datetime
from decimal import Decimal
from .normalize import text, number

FIELDS = [
    ('serial', '商品序号'), ('name', '商品名称'), ('spec', '规格'), ('unit', '单位'),
    ('quantity', '数量'), ('price', '单价'), ('day', '日期'), ('order', '单号'),
    ('code', '商品编号'), ('original_amount', '原金额'), ('note', '备注'),
    ('effective', '报价日期'), ('tax_included', '报价已含税'),
]
LABELS = dict(FIELDS)


def fields_for(kind):
    omitted = {'quantity', 'day', 'order', 'original_amount'} if kind == 'quote' else {'effective', 'tax_included'}
    return [(key, label) for key, label in FIELDS if key not in omitted]


def values_of(item):
    return {key: ('是' if item.tax_included else '否') if key == 'tax_included' else text(getattr(item, key)) for key, _ in FIELDS}


def validate_item(item, kind):
    from .parser import parse_date
    errors = {}
    values = item.edit_values
    for key in ('serial', 'name', 'spec', 'unit', 'code', 'order', 'note'):
        setattr(item, key, values.get(key, '').strip())
    if not item.serial:
        item.serial = item.source
    if not item.name:
        errors['name'] = '商品名称不能为空'
    for key in ('quantity', 'price', 'original_amount'):
        raw = values.get(key, '')
        try:
            value = number(raw)
            if (key == 'quantity' and kind == 'delivery') or (key == 'price' and kind == 'quote'):
                if value is None:
                    raise ValueError('不能为空')
            if key == 'price' and value is not None and value < 0:
                raise ValueError('不能为负数')
            setattr(item, key, value)
        except ValueError as exc:
            setattr(item, key, None)
            errors[key] = f'{LABELS[key]}：{exc}'
    for key in ('day', 'effective'):
        raw = values.get(key, '').strip()
        try:
            value = parse_date(raw) if raw else None
            if raw and value is None:
                raise ValueError('请填写完整日期 YYYY-MM-DD')
            setattr(item, key, value)
        except ValueError:
            setattr(item, key, None)
            errors[key] = f'{LABELS[key]}：请填写有效日期 YYYY-MM-DD'
    tax = values.get('tax_included', '否').strip()
    item.tax_included = tax in ('是', '1', 'true', 'True')
    if kind == 'quote' and tax not in ('是', '否', '1', '0', 'true', 'false', 'True', 'False'):
        errors['tax_included'] = '报价已含税：请选择是或否'
    if (kind == 'delivery' and item.quantity is not None and item.price is not None
            and item.original_amount is not None
            and abs(item.quantity * item.price - item.original_amount) > Decimal('.005')):
        errors['original_amount'] = f'原金额与数量×原单价不一致，应为 {item.quantity * item.price}'
    item.validation_errors = errors
    return errors


def prepare_review(result, kind):
    result.kind = kind
    for item in result.items:
        if not item.original_values:
            item.original_values = values_of(item)
        if not item.edit_values:
            item.edit_values = deepcopy(item.original_values)
        validate_item(item, kind)
    return result


def edit_item(result, index, changes):
    """无效修改也留在草稿中，不转换为0，也不要求反复弹窗才能继续编辑。"""
    item = result.items[index]
    allowed = dict(fields_for(result.kind))
    for key, raw in changes.items():
        if key not in allowed:
            raise ValueError(f'不能修改字段 {key}')
        raw = text(raw)
        previous = item.edit_values.get(key, '')
        if raw != previous:
            item.edit_history.append({'field': key, 'before': previous, 'after': raw,
                                      'time': datetime.now().isoformat(timespec='seconds')})
            item.edit_values[key] = raw
    validate_item(item, result.kind)
    return item


def commit_review(result, *, force=False):
    result = deepcopy(result)
    prepare_review(result, result.kind)
    if not result.items:
        raise ValueError('没有识别到可导入的记录，请先检查表头。')
    if result.unresolved and not force:
        raise ValueError(f'还有 {result.unresolved} 项问题未修正，请继续修改或选择“强制导入”。')
    result.forced = bool(force and result.unresolved)
    result.reviewed = True
    return result


def force_warning(result):
    bad = sum(bool(i.validation_errors) for i in result.items)
    return (f'即将强制导入 {len(result.items)} 条记录，其中 {bad} 条尚有错误，'
            f'{len(result.blocking_issues)} 项表头/工作表问题未解决。\n\n'
            '所有已识别记录都会保留。错误字段不会按0计算，未识别的工作表不会被补造。'
            '导入后仍需修正，当前只能生成待核对稿，不能生成正式账单。\n\n确认强制导入吗？')
