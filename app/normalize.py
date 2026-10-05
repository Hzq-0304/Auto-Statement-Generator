import re
import unicodedata
from decimal import Decimal, InvalidOperation


def text(value):
    return '' if value is None else str(value).strip()


def normalize(value):
    value = unicodedata.normalize('NFKC', text(value)).lower()
    value = re.sub(r'\s+', '', value).replace('×', '*').replace('＊', '*')
    value = re.sub(r'(?<=\d)[x](?=\d)', '*', value)
    # 2.0 与 2 等价，但保留尺寸顺序、颜色、品牌和开口属性。
    value = re.sub(r'\d+\.\d+', lambda m: format(Decimal(m[0]).normalize(), 'f'), value)
    return value


def number(value):
    if value is None or text(value) == '':
        return None
    if isinstance(value, bool):
        raise ValueError('不能把布尔值当作数字')
    raw = unicodedata.normalize('NFKC', text(value)).replace(',', '').replace('￥', '').replace('¥', '')
    try:
        result = Decimal(raw)
        if not result.is_finite():
            raise ValueError('数字必须为有限值')
        return result
    except InvalidOperation:
        raise ValueError(f'不是有效数字：{value}') from None


def extract_spec(name):
    # 仅用于展示和候选排序；完整名称仍参与自动匹配。
    return ' '.join(re.findall(r'(?:[MmNn])?\d+(?:\.\d+)?(?:\s*[*×xX]\s*\d+(?:\.\d+)?)+(?:[A-Za-z]+)?', name))
