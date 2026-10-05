"""工作进度只保存到用户选择的本地文件；来源变化后拒绝复用旧确认。"""
from dataclasses import asdict
from datetime import date
from decimal import Decimal
from hashlib import sha256
from pathlib import Path
import json
from .parser import import_excel
from .matching import reconcile, confirm
from .normalize import number


def fingerprint(path):
    return sha256(Path(path).read_bytes()).hexdigest()


def save_session(path, quotes, deliveries, matches):
    path=Path(path).resolve()
    if path.suffix.lower() != '.json' or path in (Path(quotes.path).resolve(),Path(deliveries.path).resolve()):
        raise ValueError('工作进度请保存为新的 .json 文件，不能覆盖原始 Excel。')
    source = lambda result: {'path':result.path,'sha256':fingerprint(result.path)}
    data = {'version':1,'quote':source(quotes),'delivery':source(deliveries),
            'quote_snapshot':[asdict(i) for i in quotes.items],
            'confirmations':[{'delivery':m.delivery.source,'quote':m.quote.source,'factor':str(m.factor),
                              'price':str(m.override_price) if m.override_price is not None else None,'reason':m.reason}
                             for m in matches if m.status=='人工确认']}
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2,default=str),encoding='utf-8')


def load_session(path):
    data=json.loads(Path(path).read_text(encoding='utf-8'))
    if data.get('version') != 1:
        raise ValueError('不支持此工作进度版本。')
    for key in ('quote','delivery'):
        item=data[key]
        if fingerprint(item['path']) != item['sha256']:
            raise ValueError(f'来源文件已变化：{item["path"]}。请重新导入，避免复用过期的商品确认。')
    q=import_excel(data['quote']['path'],'quote'); d=import_excel(data['delivery']['path'],'delivery')
    matches=reconcile(q.items,d.items)
    qmap={i.source:i for i in q.items}; mmap={m.delivery.source:m for m in matches}
    for entry in data['confirmations']:
        confirm(mmap[entry['delivery']],qmap[entry['quote']],number(entry['factor']),number(entry['price']),entry['reason'])
    return q,d,matches
