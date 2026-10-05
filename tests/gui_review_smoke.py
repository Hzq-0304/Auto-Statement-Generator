"""运行真实Tk控件，验证预览事务、编辑、提醒与DPI（不改变系统显示设置）。"""
from copy import deepcopy
from pathlib import Path
import ctypes
import sys
import tempfile
import time
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
from unittest.mock import patch
import openpyxl
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app.gui import App
from app.matching import confirm
from app.review import edit_item, commit_review
from app.display import apply_fonts
from decimal import Decimal as D


def pump(app):
    deadline=time.monotonic()+20
    while app.busy and time.monotonic()<deadline:
        app.update();time.sleep(.03)
    app.update();assert not app.busy


def children(widget):
    for child in widget.winfo_children():
        yield child
        yield from children(child)


with tempfile.TemporaryDirectory() as directory:
    root=Path(directory)
    for name,headers,rows in [('q',['品名','单位','单价'],[['零件','个',3]]),
                              ('d',['品名','单位','数量'],[['零件','个','错误'],['零件','个',2]])]:
        wb=openpyxl.Workbook();ws=wb.active;ws.append(headers)
        for row in rows:ws.append(row)
        wb.save(root/f'{name}.xlsx');wb.close()
    app=App();errors=[]
    app.report_callback_exception=lambda *args: errors.append(args)
    app.update()
    if sys.platform=='win32':
        user32=ctypes.windll.user32
        user32.GetThreadDpiAwarenessContext.restype=ctypes.c_void_p
        user32.GetAwarenessFromDpiAwarenessContext.argtypes=[ctypes.c_void_p]
        assert user32.GetAwarenessFromDpiAwarenessContext(user32.GetThreadDpiAwarenessContext())==2
    # 走选择文件、后台读取、预览窗口和提交的完整入口。
    for kind,name in [('quote','q'),('delivery','d')]:
        with patch('app.gui.filedialog.askopenfilename',return_value=str(root/f'{name}.xlsx')):
            app.choose(kind);pump(app)
        dialog=next(w for w in app.winfo_children() if isinstance(w,tk.Toplevel))
        if kind=='quote':
            dialog.confirm_button.invoke();app.update();assert app.quotes.reviewed
        else:
            assert app.deliveries is None
            assert str(dialog.confirm_button['state'])=='disabled'
            with patch('app.review_ui.messagebox.askyesno',return_value=False) as warning:
                dialog.force_button.invoke();assert warning.called;assert app.deliveries is None
            with patch('app.review_ui.messagebox.askyesno',return_value=True):
                dialog.force_button.invoke()
            app.update();assert app.deliveries.forced
    app.ack.set(True);app.refresh_gate();assert str(app.formal_button['state'])=='disabled'
    dialog=app.edit_import('delivery');app.update()
    dialog.only_errors.set(True);dialog.refresh();assert len(dialog.tree.get_children())==1
    # 横向滚动后真实Entry修改数量；筛选自动移除已修复行。
    dialog.tree.xview_moveto(.3);app.update()
    column=2+[key for key,_ in dialog.fields].index('quantity')
    dialog.open_cell('0',column);assert dialog.editor
    dialog.editor[0].delete(0,'end');dialog.editor[0].insert(0,'4');dialog.save_cell()
    assert not dialog.tree.get_children()
    assert str(dialog.confirm_button['state'])=='normal'
    assert app.deliveries.items[0].quantity is None  # 未提交前不污染主界面
    dialog.confirm_button.invoke();app.update()
    assert app.deliveries.items[0].quantity==D(4) and not app.deliveries.forced
    # 不改数据直接关闭，不因旧修改历史而重复询问。
    dialog=app.edit_import('delivery');app.update()
    with patch('app.review_ui.messagebox.askyesno') as ask:
        dialog.cancel();assert not ask.called
    # 整条修改窗口保存，然后取消整个草稿，主界面仍保持原值。
    dialog=app.edit_import('delivery');app.update();dialog.tree.selection_set('0');dialog.edit_row();app.update()
    rowwin=next(w for w in dialog.winfo_children() if isinstance(w,tk.Toplevel))
    entries=[w for w in children(rowwin) if isinstance(w,ttk.Entry)]
    quantity_index=[key for key,_ in dialog.fields].index('quantity')
    entries[quantity_index].delete(0,'end');entries[quantity_index].insert(0,'9')
    next(w for w in children(rowwin) if isinstance(w,ttk.Button) and w['text']=='保存修改').invoke()
    with patch('app.review_ui.messagebox.askyesno',return_value=True):dialog.cancel()
    assert app.deliveries.items[0].quantity==D(4)
    # 未变化行的人工确认保留，修改数量的行撤销旧人工价格。
    confirm(app.matches[0],app.quotes.items[0],D(1),D(0),'赠送确认')
    app.accept_import('delivery',commit_review(deepcopy(app.deliveries)))
    assert app.matches[0].status=='人工确认'
    changed=deepcopy(app.deliveries);edit_item(changed,0,{'quantity':'5'})
    app.accept_import('delivery',commit_review(changed))
    assert app.matches[0].status=='自动匹配' and app.matches[0].override_price is None
    # 缩放测试仅改变本进程字体，检查行高没有裁字。
    for dpi in (96,120,144,192):
        apply_fonts(app,dpi);app.update_idletasks()
        assert int(ttk.Style(app).lookup('Treeview','rowheight'))>tkfont.nametofont('TkDefaultFont').metrics('linespace')
    apply_fonts(app,app._dpi);app.update()
    app.ack.set(True);app.refresh_gate();assert str(app.formal_button['state'])=='normal'
    assert not errors,errors
    app.destroy()
print('GUI REVIEW PASS：选择文件、预览、改单元格/整行、确认/取消/强制提醒、匹配失效、PerMonitor DPI及100%-200%字体行高。')
