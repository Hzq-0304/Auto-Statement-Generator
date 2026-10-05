"""本机启动实际 Tk 窗口，操作预览、帮助、诊断、人工确认窗口和导出按钮。"""
from pathlib import Path
import sys
import time
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app.gui import App
from app.parser import import_excel
from app.models import Item, ImportResult
from app.matching import reconcile, confirm
from decimal import Decimal as D

root=Path(__file__).resolve().parents[1]
app=App();app.update()
qfile=next(root.glob('*报价*.xlsx'));dfile=next(root.glob('送货*.xlsx'))
app.accept_import('quote',import_excel(qfile,'quote'))
app.accept_import('delivery',import_excel(dfile,'delivery'))
app.update()
assert len(app.qtree.get_children())==327
assert len(app.dtree.get_children())==99
assert str(app.formal_button['state'])=='disabled'
app.pending_only.set(True);app.refresh();assert len(app.dtree.get_children())==88
first=app.dtree.get_children()[0];app.dtree.selection_set(first);app.resolve_selected();app.update()
for w in app.winfo_children():
    if w.winfo_class()=='Toplevel':w.destroy()
app.help();app.update()
for w in app.winfo_children():
    if w.winfo_class()=='Toplevel':w.destroy()
app.show_details();app.update()
for w in app.winfo_children():
    if w.winfo_class()=='Toplevel':w.destroy()
with patch('app.gui.messagebox.showinfo'):
    app.copy_diagnostic();assert '327' in app.clipboard_get()
# 通过真实导出按钮方法、文件选择返回值和后台队列完成核对稿导出。
output=root/'outputs/sample/GUI验收_待核对稿.xlsx'
with patch('app.gui.filedialog.asksaveasfilename',return_value=str(output)),patch('app.gui.messagebox.showinfo'):
    app.export(True)
    deadline=time.monotonic()+30
    while app.busy and time.monotonic()<deadline:
        app.update();time.sleep(.05)
assert output.exists() and not app.busy
# 用合成数据验证正式按钮在人工确认及勾选后才解锁。
q=Item('quote.xlsx','商品',2,'Q1','螺丝',unit='盒',price=D('8.2'))
d=Item('delivery.xlsx','单据',2,'D1','螺丝',unit='箱',quantity=D('2'))
app.quotes=ImportResult('quote.xlsx',[q]);app.deliveries=ImportResult('delivery.xlsx',[d]);app.matches=reconcile([q],[d])
app.ack.set(True);app.refresh();assert str(app.formal_button['state'])=='disabled'
confirm(app.matches[0],q,D(20),None,'包装规格确认');app.refresh();assert str(app.formal_button['state'])=='normal'
app.destroy()
print('GUI SMOKE PASS: 实际窗口、99条预览、88条待确认、人工确认弹窗、帮助、剪贴板、后台导出、正式生成限制。')
