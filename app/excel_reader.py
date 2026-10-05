"""稀疏读取 xlsx：不遍历被空白格式撑大的矩形区域，不执行任何公式。"""
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import PurePosixPath
from zipfile import ZipFile, BadZipFile
import posixpath
import xml.etree.ElementTree as ET
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.utils.datetime import from_excel, WINDOWS_EPOCH, MAC_EPOCH

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


@dataclass
class Sheet:
    name: str
    cells: dict = field(default_factory=dict)
    formulas: set = field(default_factory=set)
    merges: list = field(default_factory=list)

    def get(self, row, col, merged=False):
        if (row, col) in self.cells:
            return self.cells[row, col]
        if merged:
            for c1, r1, c2, r2 in self.merges:
                if r1 <= row <= r2 and c1 <= col <= c2:
                    return self.cells.get((r1, c1))
        return None

    @property
    def rows(self):
        rows = {}
        for (r, c), v in self.cells.items():
            rows.setdefault(r, {})[c] = v
        return rows


def read_workbook(path):
    with ZipFile(path) as z:
        # 避免异常压缩包无限占用内存，普通业务表远小于此限制。
        if sum(i.file_size for i in z.infolist()) > 250_000_000:
            raise ValueError('解压后超过 250 MB，请拆分文件后导入')
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            strings = [''.join(t.text or '' for t in si.iterfind('.//m:t', NS)) for si in root]
        styles = ET.fromstring(z.read('xl/styles.xml')) if 'xl/styles.xml' in z.namelist() else None
        formats = dict(BUILTIN_FORMATS)
        if styles is not None:
            for fmt in styles.findall('m:numFmts/m:numFmt', NS):
                formats[int(fmt.get('numFmtId'))] = fmt.get('formatCode')
        date_styles = set()
        if styles is not None:
            for index, xf in enumerate(styles.findall('m:cellXfs/m:xf', NS)):
                if is_date_format(formats.get(int(xf.get('numFmtId', '0')), '')):
                    date_styles.add(index)
        book = ET.fromstring(z.read('xl/workbook.xml'))
        pr = book.find('m:workbookPr', NS)
        epoch = MAC_EPOCH if pr is not None and pr.get('date1904') in ('1', 'true') else WINDOWS_EPOCH
        rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = []
        for info in book.findall('m:sheets/m:sheet', NS):
            target = rels[info.get(REL)]
            member = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            root = ET.fromstring(z.read(member))
            sheet = Sheet(info.get('name'))
            for cell in root.findall('m:sheetData/m:row/m:c', NS):
                key = coordinate_to_tuple(cell.get('r'))
                typ = cell.get('t', '')
                v = cell.find('m:v', NS)
                value = v.text if v is not None else None
                if cell.find('m:f', NS) is not None:
                    sheet.formulas.add(key)
                if typ == 's' and value is not None:
                    value = strings[int(value)]
                elif typ == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.findall('.//m:t', NS))
                elif typ == 'b' and value is not None:
                    value = value == '1'
                elif typ == 'd' and value:
                    value = datetime.fromisoformat(value)
                elif value is not None and typ not in ('str', 'e'):
                    if int(cell.get('s', '0')) in date_styles:
                        value = from_excel(float(value), epoch)
                    # 普通数值保持十进制文本，避免二进制浮点损失。
                if value is not None or key in sheet.formulas:
                    sheet.cells[key] = value
            sheet.merges = [range_boundaries(m.get('ref')) for m in root.findall('m:mergeCells/m:mergeCell', NS)]
            sheets.append(sheet)
        return sheets
