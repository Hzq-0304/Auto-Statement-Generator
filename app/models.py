from dataclasses import dataclass, field
from decimal import Decimal
from datetime import date


@dataclass
class Issue:
    step: str
    location: str
    message: str
    advice: str
    severity: str = '警告'

    def __str__(self):
        return f'[{self.severity}] {self.step} · {self.location}\n{self.message}\n建议：{self.advice}'


class ImportFailure(Exception):
    def __init__(self, issues, recognized=''):
        self.issues = issues
        self.recognized = recognized
        super().__init__('\n\n'.join(map(str, issues)) + '\n\n已识别：\n' + recognized)


@dataclass
class Item:
    file: str
    sheet: str
    row: int
    serial: str
    name: str
    spec: str = ''
    unit: str = ''
    code: str = ''
    price: Decimal | None = None
    quantity: Decimal | None = None
    original_amount: Decimal | None = None
    day: date | None = None
    order: str = ''
    note: str = ''
    effective: date | None = None
    tax_included: bool = False
    # 原始文本和当前编辑文本分开保存，坏数字也能展示、修正和追溯。
    original_values: dict = field(default_factory=dict)
    edit_values: dict = field(default_factory=dict)
    validation_errors: dict = field(default_factory=dict)
    edit_history: list = field(default_factory=list)

    @property
    def source(self):
        return f'{self.sheet}!{self.row}'


@dataclass
class ImportResult:
    path: str
    items: list[Item] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    recognized: list[str] = field(default_factory=list)
    supplier: str = ''
    customer: str = ''
    blocking_issues: list[Issue] = field(default_factory=list)
    forced: bool = False
    reviewed: bool = False
    kind: str = ''

    @property
    def unresolved(self):
        return len(self.blocking_issues) + sum(bool(i.validation_errors) for i in self.items)

    def diagnostic(self):
        header = f'文件：{self.path}\n识别条数：{len(self.items)}\n'
        if self.forced:
            header += f'强制导入：仍有 {self.unresolved} 项问题待修正\n'
        row_errors = [f'[待修正] {i.source}：' + '；'.join(i.validation_errors.values()) for i in self.items if i.validation_errors]
        return header + '\n'.join(self.recognized) + '\n\n' + '\n\n'.join(map(str, self.issues + self.blocking_issues)) + '\n' + '\n'.join(row_errors)


@dataclass
class Match:
    delivery: Item
    quote: Item | None = None
    status: str = '未匹配'
    candidates: list[Item] = field(default_factory=list)
    factor: Decimal = Decimal('1')
    override_price: Decimal | None = None
    reason: str = ''

    @property
    def price(self):
        if self.override_price is not None:
            return self.override_price
        if self.quote and self.quote.price is not None:
            return self.quote.price * self.factor
        return None

    @property
    def ready(self):
        return (self.quote is not None and self.price is not None and self.price >= 0
                and self.delivery.quantity is not None and not self.delivery.validation_errors
                and not self.quote.validation_errors and bool(self.delivery.name.strip())
                and self.status in ('自动匹配', '人工确认'))

    @property
    def amount(self):
        return self.delivery.quantity * self.price if self.ready else None
