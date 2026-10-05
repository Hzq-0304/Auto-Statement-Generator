"""只读检查原始样例，完整记录非空单元格、公式、样式和打印设置。"""
import json
from pathlib import Path
from collections import Counter
import openpyxl

root = Path(__file__).resolve().parents[1]
out = root / 'analysis'
out.mkdir(exist_ok=True)
for path in root.glob('*.xlsx'):
    wb = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    result = {}
    for ws in wb:
        rows = []
        occupied = {}
        for cell in list(ws._cells.values()):
            if cell.value is not None:
                occupied.setdefault(cell.row, []).append(cell)
        for _, row in sorted(occupied.items()):
            row.sort(key=lambda c: c.column)
            cells = []
            for c in row:
                if c.value is not None:
                    cells.append({'cell': c.coordinate, 'value': str(c.value), 'type': c.data_type,
                                  'cached': str(cached[ws.title][c.coordinate].value),
                                  'style': c.style_id, 'format': c.number_format})
            if cells:
                rows.append(cells)
        result[ws.title] = {'dimensions': ws.calculate_dimension(), 'merged': [str(m) for m in ws.merged_cells.ranges],
                            'widths': {k: v.width for k, v in ws.column_dimensions.items()},
                            'heights': {k: v.height for k, v in ws.row_dimensions.items() if v.height},
                            'print_area': str(ws.print_area), 'print_titles': ws.print_title_rows,
                            'rows': rows, 'images': len(ws._images)}
        print(path.name, ws.title, ws.calculate_dimension(), '非空行', len(rows), '合并', len(ws.merged_cells.ranges), '图片', len(ws._images))
        for row in rows[:12]:
            print(' | '.join(f"{c['cell']}={c['value']}" for c in row))
        print('尾部:')
        for row in rows[-6:]:
            print(' | '.join(f"{c['cell']}={c['value']}" for c in row))
    (out / (path.stem + '.json')).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
