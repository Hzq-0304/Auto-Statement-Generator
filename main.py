"""双击打包程序启动 GUI；命令行支持自动化验收和批量生成核对稿。"""
import argparse
from decimal import Decimal
from pathlib import Path
from app.parser import import_excel
from app.matching import reconcile, match_issues
from app.exporter import export_statement


def main():
    parser=argparse.ArgumentParser(description='本地自动对账工具')
    parser.add_argument('--quote',help='报价单路径')
    parser.add_argument('--delivery',help='送货单路径')
    parser.add_argument('--output',help='输出 xlsx 路径，默认生成待核对稿')
    parser.add_argument('--formal',action='store_true',help='生成正式账单，要求全部匹配')
    parser.add_argument('--acknowledge',action='store_true',help='确认已经核对诊断信息')
    parser.add_argument('--tax-rate',default='0',help='未含税商品附加税率百分数，如 6')
    args=parser.parse_args()
    if not any([args.quote,args.delivery,args.output]):
        from app.gui import main as gui_main
        gui_main();return
    if not all([args.quote,args.delivery,args.output]):parser.error('必须同时提供 --quote、--delivery、--output')
    try:
        q=import_excel(args.quote,'quote');d=import_excel(args.delivery,'delivery');matches=reconcile(q.items,d.items)
        print(q.diagnostic());print(d.diagnostic())
        for issue in match_issues(matches):print(issue)
        total=export_statement(args.output,q,d,matches,draft=not args.formal,acknowledged=args.acknowledge,tax_rate=Decimal(args.tax_rate)/100)
        print(f'报价商品 {len(q.items)}；送货记录 {len(d.items)}；匹配 {sum(m.ready for m in matches)}；待确认 {sum(not m.ready for m in matches)}；已确认部分金额 {total}；输出 {args.output}')
    except Exception as exc:
        parser.exit(1,f'操作未完成：{exc}\n请查看 docs/excel-format-guide.md。\n')


if __name__=='__main__':main()
