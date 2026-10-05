"""套用样例版式，并写入可重算公式、缓存与逐条来源。"""
from copy import copy
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import os
import sys
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.workbook.properties import CalcProperties
from .models import ImportFailure


def resource(relative):
    return Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parents[1])) / relative


def literal(cell, value):
    """外部文本必须保持文本，不能变成 Excel 公式。"""
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'


def chinese_money(value):
    digits = '零壹贰叁肆伍陆柒捌玖'
    value = Decimal(value).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
    sign = '负' if value < 0 else ''
    cents = int(abs(value) * 100)
    integer, fraction = divmod(cents, 100)
    if integer >= 10**16:
        return f'{value:,.2f} 元'
    def group(n):
        output, zero = '', False
        for power, unit in [(3, '仟'), (2, '佰'), (1, '拾'), (0, '')]:
            d = n // 10**power % 10
            if d:
                if zero and output:
                    output += '零'
                output += digits[d] + unit
                zero = False
            elif output:
                zero = True
        return output
    parts, n = [], integer
    while n:
        parts.append(n % 10000)
        n //= 10000
    result, gap = '', False
    for i in range(len(parts)-1, -1, -1):
        if not parts[i]:
            gap = True
            continue
        if result and (gap or parts[i] < 1000):
            result += '零'
        result += group(parts[i]) + ['', '万', '亿', '万亿'][i]
        gap = False
    result = sign + (result or '零') + '元'
    jiao, fen = divmod(fraction, 10)
    return result + (digits[jiao] + '角' if jiao else ('零' if fen else '')) + (digits[fen] + '分' if fen else '') + ('整' if not fraction else '')


def _cache_formulas(path, cache):
    """写入自身已核验的公式结果，使未打开 Excel 的预览也能显示金额。"""
    ns = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    ET.register_namespace('', ns['m'])
    temp = path.with_suffix('.cached.xlsx')
    try:
        with ZipFile(path) as src, ZipFile(temp, 'w', ZIP_DEFLATED) as dst:
            for member in src.infolist():
                data = src.read(member.filename)
                if member.filename == 'xl/worksheets/sheet1.xml':
                    root = ET.fromstring(data)
                    for c in root.findall('.//m:sheetData/m:row/m:c', ns):
                        if c.get('r') in cache and c.find('m:f', ns) is not None:
                            v = c.find('m:v', ns)
                            if v is None:
                                v = ET.SubElement(c, '{'+ns['m']+'}v')
                            v.text = str(cache[c.get('r')])
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                dst.writestr(member, data)
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def export_statement(path, quotes, deliveries, matches, *, draft=True, acknowledged=False, tax_rate=Decimal('0')):
    if not matches or len(matches) != len(deliveries.items):
        raise ValueError('送货记录为空或匹配结果不完整，请重新导入。')
    if not draft and (not all(m.ready for m in matches) or not acknowledged):
        raise ValueError('正式对账单要求全部记录匹配完成，并确认所有诊断提示。')
    if not tax_rate.is_finite() or not Decimal('0') <= tax_rate <= Decimal('1'):
        raise ValueError('税率应在 0%～100% 之间。')
    path = Path(path).resolve()
    if path in (Path(quotes.path).resolve(), Path(deliveries.path).resolve(), resource('assets/statement-template.xlsx').resolve()):
        raise ValueError('不能覆盖原始报价单、送货单或程序模板，请选择新的文件名。')
    if path.suffix.lower() != '.xlsx':
        raise ValueError('导出文件扩展名必须是 .xlsx。')
    wb = openpyxl.load_workbook(resource('assets/statement-template.xlsx'))
    ws = wb.active
    # 从模板提取表头/明细样式；动态建立明细与尾部，避免旧月份行数残留。
    detail_styles = [copy(ws.cell(7,c)._style) for c in range(1,12)]
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= 7:
            ws.unmerge_cells(str(merged))
    ws.delete_rows(7, max(1, ws.max_row-6))
    days = sorted({m.delivery.day for m in matches if m.delivery.day})
    period = f'{days[0]:%Y-%m}' if days and days[0].strftime('%Y%m') == days[-1].strftime('%Y%m') else (f'{days[0]} 至 {days[-1]}' if days else '日期待核对')
    literal(ws['A1'], deliveries.supplier or quotes.supplier or '供应方')
    literal(ws['A2'], f'对账单（{period}）' + (' — 待核对稿' if draft else ''))
    literal(ws['A3'], 'TO：' + (deliveries.customer or quotes.customer))
    literal(ws['G3'], 'FM：' + (deliveries.supplier or quotes.supplier))
    ws['G3'].font = Font(name='宋体', size=10)
    ws['G3'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[3].height = 32
    literal(ws['B5'], f'核对日期：{date.today():%Y-%m-%d}')
    cache, total, taxable = {}, Decimal('0'), Decimal('0')
    audit = wb.create_sheet('来源明细')
    headers = ['对账序号','报价文件','报价工作表','报价原始行','报价商品序号','报价商品名称','报价规格','报价单位','原报价','单位倍率','有效单价','送货文件','送货工作表','送货原始行','送货商品序号','原送货单价','原送货金额','匹配状态','确认原因','报价日期','报价已含税']
    audit.append(headers)
    for i, m in enumerate(matches, 1):
        r, d, q = i+6, m.delivery, m.quote
        for c, style in enumerate(detail_styles, 1):
            ws.cell(r,c)._style = copy(style)
        note = d.note
        if not m.ready:
            note = (note + '；' if note else '') + '待确认：' + m.status
        elif m.reason:
            note = (note + '；' if note else '') + m.reason
        values = [i, d.day, d.order, d.name + ((' ' + d.spec) if d.spec and d.spec not in d.name else ''), d.unit, d.quantity,
                  m.price if m.ready else None, None, note, q.serial if q else '待确认', d.serial]
        for c, value in enumerate(values, 1):
            if isinstance(value, Decimal):
                value = float(value)
            literal(ws.cell(r,c), value)
        ws.cell(r,2).number_format = 'yyyy-mm-dd'
        ws.cell(r,6).number_format = '#,##0.#####'
        ws.cell(r,7).number_format = '#,##0.00###'
        ws.cell(r,8).number_format = '#,##0.00'
        ws.row_dimensions[r].height = max(32, 16 * max((len(str(values[c-1] or ''))//(19 if c == 4 else 15)+1) for c in (4,9,10)))
        if m.ready:
            ws.cell(r,8, f'=F{r}*G{r}')
            cache[f'H{r}'] = m.amount
            total += m.amount
            if not q.tax_included:
                taxable += m.amount
        else:
            for c in range(1,12):
                ws.cell(r,c).fill = PatternFill('solid', fgColor='FFF2CC')
        row = [i, Path(q.file).name if q else '', q.sheet if q else '', q.row if q else '', q.serial if q else '', q.name if q else '', q.spec if q else '', q.unit if q else '',
               q.price if q else None, m.factor, m.price if m.ready else None, Path(d.file).name, d.sheet, d.row, d.serial,
               d.price, d.original_amount, m.status, m.reason, q.effective if q else None, '是' if q and q.tax_included else '否']
        for c,v in enumerate(row,1):
            literal(audit.cell(i+1,c), float(v) if isinstance(v,Decimal) else v)
    end = 6 + len(matches)
    total_row = end + 2
    label = '已确认部分小计' if draft else '合计'
    tax = (taxable * tax_rate).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
    for r, caption in [(total_row, f'{label}（大写）：{chinese_money(total)}'),
                       (total_row+1, f'税额（未含税报价加税 {tax_rate*100:g}%）：'),
                       (total_row+2, f'{label}含税金额：')]:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7)
        literal(ws.cell(r,1),caption)
        ws.cell(r,1).font=Font(name='宋体',size=11,bold=True)
        ws.cell(r,8).number_format='#,##0.00'
        ws.row_dimensions[r].height=28
    ws.cell(total_row,8,f'=SUM(H7:H{end})')
    cache[f'H{total_row}']=total
    # 税基直接引用已确认且未含税行；行多时分组求和避免公式长度限制。
    refs = [f'H{i+7}' for i,m in enumerate(matches) if m.ready and not m.quote.tax_included]
    if len(refs) > 200:
        # 明细税务归属在来源表保留，使用 SUMIF 避免超长公式。
        tax_formula = f'=ROUND(SUMIF(\'来源明细\'!U2:U{len(matches)+1},"否",H7:H{end})*{tax_rate},2)'
    else:
        tax_formula = f'=ROUND(SUM({",".join(refs) or "0"})*{tax_rate},2)'
    ws.cell(total_row+1,8,tax_formula); cache[f'H{total_row+1}']=tax
    ws.cell(total_row+2,8,f'=H{total_row}+H{total_row+1}'); cache[f'H{total_row+2}']=total+tax
    warnings = sum(not m.ready for m in matches)
    footer = f'待核对稿：共 {len(matches)} 条，{warnings} 条待确认；金额为空的记录未纳入小计，不可作为最终结算。' if draft else '按已确认报价和人工调整生成；金额显示两位小数，合计按完整精度计算。'
    for offset,line in [(4,footer),(5,'报价商品序号为空时使用“工作表!原始行”；详细来源及人工确认原因见“来源明细”。'),(7,'制表人：                         核对人：                         审核：                         审批：'),(8,'账户名称：'),(9,'账户号码：')]:
        r=total_row+offset
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11)
        literal(ws.cell(r,1),line)
        ws.cell(r,1).font=Font(name='宋体',size=10)
        ws.cell(r,1).alignment=Alignment(wrap_text=True,vertical='center')
        ws.row_dimensions[r].height=32
    ws.freeze_panes='D7'
    ws.print_title_rows='1:6'
    ws.print_area=f'A1:K{total_row+9}'
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    audit.freeze_panes='F2'; audit.auto_filter.ref=audit.dimensions
    for c in audit[1]:
        c.font=Font(name='宋体',bold=True,color='FFFFFF')
        c.fill=PatternFill('solid',fgColor='365F91')
        audit.column_dimensions[c.column_letter].width=22 if c.column not in (6,19) else 40
    for row in audit.iter_rows(min_row=2):
        for c in row:
            c.alignment=Alignment(vertical='center',wrap_text=True)
            if c.column==20:
                c.number_format='yyyy-mm-dd'
    wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
    path.parent.mkdir(parents=True,exist_ok=True)
    fd,tmp=tempfile.mkstemp(prefix='statement-',suffix='.xlsx',dir=path.parent)
    os.close(fd)
    temp=Path(tmp)
    try:
        wb.save(temp)
        _cache_formulas(temp,cache)
        os.replace(temp,path)
    finally:
        temp.unlink(missing_ok=True)
    return total
