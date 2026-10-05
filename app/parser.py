from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
import re
from .excel_reader import read_workbook
from .models import ImportResult, Item, Issue, ImportFailure
from .normalize import normalize, text, number, extract_spec

ALIASES = {
    'name': ['商品名称', '货品名称', '品名', '品名及规格', '产品名称', '名称'],
    'serial': ['序号', '商品序号', '项次'],
    'code': ['货品编号', '商品编号', '产品编号', '物料编码', '商品编码', '货号'],
    'spec': ['规格', '型号', '规格型号', '规格/型号'],
    'unit': ['单位', '计量单位'],
    'price': ['单价', '销售单价', '报价', '圆养报价', '单价二', '参考报价', '单价一'],
    'quantity': ['数量', '送货数量', '实送数量'],
    'date': ['日期', '送货日期', '订单日期'],
    'order': ['单号', '送货单号', '订单号', '送货单编号'],
    'amount': ['金额', '总金额', '小计'],
    'note': ['备注', '说明'],
}
LOOKUP = {normalize(a): k for k, vals in ALIASES.items() for a in vals}
PRICE_PRIORITY = {'圆养报价': 4, '单价二': 4, '单价': 3, '销售单价': 3, '报价': 3, '参考报价': 1, '单价一': 1}
END = ('合计', '总计', '小计', '说明:', '联系人', '收货人', '单价二为')


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = text(value)
    m = re.search(r'(20\d{2})[-/年.]\s*(\d{1,2})[-/月.]\s*(\d{1,2})', raw)
    if not m:
        m = re.search(r'(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)', raw)
    return date(*map(int, m.groups())) if m else None


def header_at(values):
    fields = defaultdict(list)
    for col, val in values.items():
        key = LOOKUP.get(normalize(val))
        if key:
            fields[key].append(col)
    if 'name' not in fields:
        return None
    return fields


def import_excel(path, kind):
    result = ImportResult(str(Path(path).resolve()))
    if Path(path).suffix.lower() not in ('.xlsx', '.xlsm'):
        raise ImportFailure([Issue('打开文件', str(path), '仅支持 .xlsx / .xlsm，旧版 .xls 或 CSV 不能直接导入。', '请用 Excel/WPS 另存为 .xlsx。', '错误')])
    try:
        sheets = read_workbook(path)
    except Exception as exc:
        raise ImportFailure([Issue('打开文件', str(path), str(exc), '确认文件未损坏、未加密，且扩展名确为 xlsx；另存后重试。', '错误')]) from exc
    errors = []
    for sheet in sheets:
        rows = sheet.rows
        if not any(text(v) for v in sheet.cells.values()):
            result.recognized.append(f'{sheet.name}：空白工作表，跳过')
            continue
        required = 'price' if kind == 'quote' else 'quantity'
        candidates = [(r, header_at(v)) for r, v in sorted(rows.items())]
        complete = [(r, f) for r, f in candidates if f and required in f]
        if not complete:
            seen = [(r, f) for r, f in candidates if f]
            message = '找不到单价列' if seen and kind == 'quote' else '找不到数量列' if seen else '找不到表头或商品名称列'
            errors.append(Issue('识别表头', sheet.name, message, '添加一行清晰的表头：商品名称、单位、单价（报价）或数量（送货）。', '错误'))
            result.recognized.append(f'{sheet.name}：发现字段 {seen}；前几行内容 {list(rows.items())[:6]}')
            continue
        layouts = {tuple((k, tuple(v)) for k, v in sorted(f.items())) for _, f in complete}
        if len(layouts) > 1:
            errors.append(Issue('识别表头', sheet.name, f'存在多个不同布局的表头，行号 {[r for r, _ in complete]}', '每张工作表只放一种布局，不同表格拆到独立工作表。', '错误'))
            continue
        first, fields = complete[0]
        cols = {}
        for key, options in fields.items():
            if key == 'price':
                priorities = {c: PRICE_PRIORITY.get(normalize(rows[first][c]), 0) for c in options}
                best = max(priorities.values())
                options = [c for c in options if priorities[c] == best]
                if kind == 'quote' and best < 3:
                    errors.append(Issue('选择报价列', sheet.name, '只有参考报价/单价一，无法确定结算价格', '将实际结算列命名为单价、圆养报价或单价二。', '错误'))
            if len(options) != 1:
                errors.append(Issue('识别列', sheet.name, f'字段 {key} 对应多个列 {options}', '删除重复字段或将不采用的列改为其他名称。', '错误'))
            cols[key] = options[0]
        result.recognized.append(f'{sheet.name}：表头第 {first} 行；字段列号 {cols}')
        before = ' '.join(text(v) for (r, _), v in sheet.cells.items() if r < first)
        all_text = ' '.join(text(v) for v in sheet.cells.values())
        if not result.supplier:
            result.supplier = next((text(v) for (r, _), v in sheet.cells.items() if r < first and '有限公司' in text(v) and not re.match('(?i)to', text(v))), '')
        # 客户名称按实际单元格提取，避免把下一格的报价日期拼进客户名。
        for (mr, _), value in sheet.cells.items():
            if mr >= first:
                continue
            customer = re.search(r'TO\s*[:：]\s*([^\n]+?)(?=\s+(?:FM|TEL)|$)', text(value), re.I)
            if customer:
                found = customer[1].strip()
                if result.customer and normalize(result.customer) != normalize(found):
                    errors.append(Issue('识别客户', sheet.name, f'同一文件出现不同客户：{result.customer} / {found}', '请按客户拆分工作簿后分别对账，避免混入他人订单。', '错误'))
                else:
                    result.customer = found
        current_day = None
        current_order = ''
        effective = None
        tax_included = '含税' in sheet.name or ('含税。' in all_text and '+6%' not in all_text)
        try:
            effective = parse_date(before) if kind == 'quote' else None
        except ValueError:
            result.issues.append(Issue('识别报价日期', sheet.name, '报价日期无效', '修正为 YYYY-MM-DD。'))
        count = 0
        active = False
        header_rows = {r for r, _ in complete}
        for r, values in sorted(rows.items()):
            joined = ' '.join(text(v) for v in values.values())
            # 订单元数据可在每个打印块的表头之前重复出现。
            if kind == 'delivery' and (r < first or not active):
                try:
                    m = re.search(r'日期\s*[:：]\s*(20\d{2}[-/年.]\d{1,2}[-/月.]\d{1,2}|20\d{6})', joined)
                    if m:
                        current_day = parse_date(m[1])
                    om = re.search(r'(?:送货单号|订单号|送货单编号)\s*[:：]\s*([A-Za-z0-9_-]+)', joined)
                    if om:
                        current_order = om[1]
                except ValueError:
                    result.issues.append(Issue('识别日期', f'{sheet.name}!{r}', '单据日期不合法', '修改为真实日期后重新导入。'))
            if r in header_rows:
                if active:  # 连续打印表头可以复用同一订单；不把表头当商品。
                    pass
                active = True
                continue
            if not active or r <= first:
                continue
            get = lambda key, merged=False: sheet.get(r, cols[key], merged) if key in cols else None
            name = text(get('name', True))
            if any(normalize(joined).startswith(p) for p in END):
                active = False
                continue
            raw_qty = get('quantity')
            raw_price = get('price')
            unit = text(get('unit', True))
            # 分类行无单位、无数值价格；缺价但有单位的商品不能丢失。
            if kind == 'quote' and not unit:
                try:
                    category = number(raw_price) is None
                except ValueError:
                    category = True
                if category and not get('code') and not get('serial'):
                    continue
            if kind == 'delivery' and not name and raw_qty is None:
                continue  # 空模板中只有序号和预填单位。
            if not name:
                if raw_qty is not None or raw_price is not None:
                    errors.append(Issue('读取明细', f'{sheet.name}!{r}', '商品名称为空', '补充商品名称，不能只写数量或价格。', '错误'))
                continue
            if normalize(raw_price) in ('圆养报价', '参考报价'):
                continue
            location = f'{sheet.name}!{r}'
            parsed = {}
            for key in ('price', 'quantity', 'amount'):
                try:
                    parsed[key] = number(get(key))
                    if key in cols and (r, cols[key]) in sheet.formulas and get(key) is None:
                        raise ValueError('公式没有已保存的计算结果')
                except ValueError as exc:
                    issue = Issue('读取数值', location, f'{key}：{exc}', '请在 Excel/WPS 中重新计算并保存，或填入纯数值。', '错误' if key == required else '警告')
                    (errors if key == required else result.issues).append(issue)
                    parsed[key] = None
            if kind == 'delivery' and parsed['quantity'] is None:
                errors.append(Issue('读取明细', location, f'{name} 的数量为空', '补充数量；无实际明细的占位行请清空名称。', '错误'))
                continue
            if parsed['price'] is not None and parsed['price'] < 0:
                errors.append(Issue('读取单价', location, '单价为负数', '单价应为非负数，退货请以负数量表达。', '错误'))
            if kind == 'quote' and parsed['price'] is None:
                result.issues.append(Issue('读取报价', location, f'{name} 缺少有效报价', '补充报价后重导，或匹配时人工确认有效单价。'))
            day = current_day
            if kind == 'delivery':
                try:
                    if get('date', True) is not None:
                        day = parse_date(get('date', True))
                    if day is None:
                        result.issues.append(Issue('识别日期', location, '找不到完整的送货日期', '在日期列或单据抬头填入 YYYY-MM-DD。'))
                except ValueError:
                    day = None
                    result.issues.append(Issue('识别日期', location, '明细日期不合法', '修改为真实日期。'))
            explicit_spec = text(get('spec', True))
            item = Item(result.path, sheet.name, r, text(get('serial')) or location, name,
                        explicit_spec or extract_spec(name), unit, text(get('code')), parsed['price'],
                        parsed['quantity'], parsed['amount'], day, text(get('order', True)) or current_order,
                        text(get('note')), effective, tax_included)
            result.items.append(item)
            count += 1
            if kind == 'delivery':
                if item.price is not None and item.original_amount is not None and abs(item.quantity * item.price - item.original_amount) > number('0.005'):
                    result.issues.append(Issue('核验原金额', location, f'原金额 {item.original_amount} 与数量×原单价 {item.quantity * item.price} 不一致', '核对原单；新账单按匹配后的报价计算。'))
                if any(word in item.note for word in ('赠送', '退回', '退货')):
                    result.issues.append(Issue('结算备注', location, item.note, '双击该记录，确认商品、有效单价或赠送零价；退货需补充实际负数量记录。'))
        result.recognized.append(f'{sheet.name}：读取 {count} 条明细')
        if kind == 'delivery' and current_day and re.match(r'^20\d{6}', sheet.name):
            if sheet.name[:8] != current_day.strftime('%Y%m%d'):
                result.issues.append(Issue('核验日期', sheet.name, f'工作表名称与正文日期 {current_day} 不同，采用正文日期', '确认正文日期及单号，必要时修正原表副本后重新导入。'))
        if count == 0:
            result.issues.append(Issue('读取明细', sheet.name, '数据区域为空，已跳过', '若此表是空模板可忽略；否则检查名称和数量。', '提示'))
    if kind == 'delivery':
        orders = defaultdict(set)
        for item in result.items:
            if item.order:
                orders[item.order].add(item.sheet)
        for order, tabs in orders.items():
            if len(tabs) > 1:
                result.issues.append(Issue('核验单号', order, f'单号出现在多张表：{sorted(tabs)}', '检查重复单号是否为重复送货；程序未自动去重。'))
    if not result.items:
        errors.append(Issue('读取明细', str(path), '数据区域为空，没有可用商品', '保留清晰表头并至少填写一条商品。', '错误'))
    if errors:
        raise ImportFailure(errors + result.issues, '\n'.join(result.recognized) + f'\n已读到 {len(result.items)} 条，但为避免漏单本次导入整体未生效。')
    return result
