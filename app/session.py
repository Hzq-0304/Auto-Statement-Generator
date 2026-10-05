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
from .models import Item, ImportResult, Issue
from .review import prepare_review


def fingerprint(path):
    return sha256(Path(path).read_bytes()).hexdigest()


def save_session(path, quotes, deliveries, matches):
    path=Path(path).resolve()
    if path.suffix.lower() != '.json' or path in (Path(quotes.path).resolve(),Path(deliveries.path).resolve()):
        raise ValueError('工作进度请保存为新的 .json 文件，不能覆盖原始 Excel。')
    source = lambda result: {'path':result.path,'sha256':fingerprint(result.path)}
    data = {'version':2,'quote':source(quotes),'delivery':source(deliveries),
            'quote_snapshot':asdict(quotes),'delivery_snapshot':asdict(deliveries),
            'confirmations':[{'delivery':m.delivery.source,'quote':m.quote.source,'factor':str(m.factor),
                              'price':str(m.override_price) if m.override_price is not None else None,'reason':m.reason}
                             for m in matches if m.status=='人工确认']}
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2,default=str),encoding='utf-8')


def load_session(path):
    data=json.loads(Path(path).read_text(encoding='utf-8'))
    if data.get('version') not in (1,2):
        raise ValueError('不支持此工作进度版本。')
    for key in ('quote','delivery'):
        item=data[key]
        if fingerprint(item['path']) != item['sha256']:
            raise ValueError(f'来源文件已变化：{item["path"]}。请重新导入，避免复用过期的商品确认。')
    if data['version']==1:
        q=import_excel(data['quote']['path'],'quote'); d=import_excel(data['delivery']['path'],'delivery')
    else:
        q=restore_result(data['quote_snapshot'],'quote');d=restore_result(data['delivery_snapshot'],'delivery')
        if q.path!=data['quote']['path'] or d.path!=data['delivery']['path']:
            raise ValueError('工作进度的来源路径不一致，请重新导入。')
    matches=reconcile(q.items,d.items)
    qmap={i.source:i for i in q.items}; mmap={m.delivery.source:m for m in matches}
    for entry in data['confirmations']:
        confirm(mmap[entry['delivery']],qmap[entry['quote']],number(entry['factor']),number(entry['price']),entry['reason'])
    return q,d,matches


def restore_result(snapshot,kind):
    from .parser import parse_date
    snapshot=dict(snapshot)
    items=[]
    for raw in snapshot.pop('items'):
        raw=dict(raw)
        for key in ('price','quantity','original_amount'):
            raw[key]=number(raw.get(key))
        for key in ('day','effective'):
            raw[key]=parse_date(raw.get(key))
        items.append(Item(**raw))
    for key in ('issues','blocking_issues'):
        snapshot[key]=[Issue(**v) for v in snapshot.get(key,[])]
    result=ImportResult(items=items,**snapshot)
    result.kind=kind
    if result.reviewed or any(i.edit_values for i in result.items):
        prepare_review(result,kind)
    return result
