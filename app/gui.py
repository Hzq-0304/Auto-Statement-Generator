"""Tkinter 主界面；耗时导入/导出在后台执行，界面更新仅在主线程执行。"""
from datetime import datetime
from decimal import Decimal
from pathlib import Path
import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from . import __version__
from .models import Match
from .display import enable_dpi_awareness, apply_fonts, fit_window, window_dpi
from .review_ui import ReviewDialog
from .parser import import_excel
from .matching import reconcile, confirm, match_issues
from .normalize import number
from .exporter import export_statement, resource
from .session import save_session, load_session


class App(tk.Tk):
    def __init__(self):
        enable_dpi_awareness()
        super().__init__()
        self.title('自动对账工具')
        self._dpi=apply_fonts(self)
        fit_window(self,1250,860)
        self.quotes=None; self.deliveries=None; self.matches=[]
        self.messages=queue.Queue(); self.busy=False
        self.quote_status=tk.StringVar(value='尚未导入')
        self.delivery_status=tk.StringVar(value='尚未导入')
        self.summary=tk.StringVar(value='先导入报价单和送货单，再预览匹配结果。')
        self.pending_only=tk.BooleanVar(); self.ack=tk.BooleanVar(); self.tax=tk.StringVar(value='0')
        self.last_error=''; self.buttons=[]
        style=ttk.Style(self)
        if 'vista' in style.theme_names(): style.theme_use('vista')
        apply_fonts(self,self._dpi)
        main=ttk.Frame(self,padding=16); main.pack(fill='both',expand=True)
        # 仅预览表格伸缩，避免高DPI时把下方导出和诊断按钮挤出窗口。
        main.columnconfigure(0,weight=1);main.rowconfigure(5,weight=1)
        ttk.Label(main,text='自动对账工具',font=('Microsoft YaHei UI',19,'bold')).grid(row=0,column=0,sticky='w')
        ttk.Label(main,text='选择文件 → 直接修改错误 → 确认或强制导入 → 匹配并生成对账单').grid(row=1,column=0,sticky='w',pady=(4,12))
        for index,(title,var,kind) in enumerate([('报价单',self.quote_status,'quote'),('送货单',self.delivery_status,'delivery')],2):
            frame=ttk.LabelFrame(main,text=title,padding=10); frame.grid(row=index,column=0,sticky='ew',pady=3)
            self.button(frame,'选择文件…',lambda k=kind:self.choose(k)).pack(side='left')
            self.button(frame,'修改已导入记录',lambda k=kind:self.edit_import(k)).pack(side='left',padx=8)
            ttk.Label(frame,textvariable=var,wraplength=700).pack(side='left',padx=6)
        bar=ttk.Frame(main); bar.grid(row=4,column=0,sticky='ew',pady=10)
        ttk.Label(bar,textvariable=self.summary,font=('Microsoft YaHei UI',10,'bold')).pack(side='left')
        self.button(bar,'保存进度',self.save_progress).pack(side='right',padx=4)
        self.button(bar,'打开进度',self.open_progress).pack(side='right',padx=4)
        tabs=ttk.Notebook(main); tabs.grid(row=5,column=0,sticky='nsew')
        dtab=ttk.Frame(tabs,padding=6); qtab=ttk.Frame(tabs,padding=6)
        tabs.add(dtab,text='送货与匹配预览'); tabs.add(qtab,text='报价商品预览')
        controls=ttk.Frame(dtab); controls.pack(fill='x',pady=4)
        ttk.Checkbutton(controls,text='只看待确认',variable=self.pending_only,command=self.refresh).pack(side='left')
        self.button(controls,'确认选中商品…',self.resolve_selected).pack(side='left',padx=10)
        ttk.Label(controls,text='双击明细修改原记录；“确认选中商品”处理匹配与价格。').pack(side='left')
        self.dtree=self.tree(dtab,['送货序号','商品名称','规格','单位','数量','日期','单号','匹配单价','金额','匹配状态','来源'],[90,250,120,60,90,110,130,90,100,150,170])
        self.dtree.bind('<Double-1>',lambda e:self.edit_import('delivery'))
        self.qtree=self.tree(qtab,['报价序号','商品名称','规格','单位','单价','报价日期','备注'],[190,300,160,70,100,120,250])
        self.qtree.bind('<Double-1>',lambda e:self.edit_import('quote'))
        ops=ttk.LabelFrame(main,text='生成对账单',padding=10); ops.grid(row=6,column=0,sticky='ew',pady=(10,5))
        ttk.Label(ops,text='未含税报价加税（%）：').pack(side='left')
        ttk.Entry(ops,textvariable=self.tax,width=6).pack(side='left')
        ttk.Label(ops,text='已含税商品不重复加税').pack(side='left',padx=8)
        self.button(ops,'生成待核对稿',lambda:self.export(True)).pack(side='right',padx=4)
        self.formal_button=self.button(ops,'生成正式对账单',lambda:self.export(False)); self.formal_button.pack(side='right',padx=4)
        ttk.Checkbutton(main,text='我已检查下方诊断，确认所选报价适用于本次送货，并确认日期、单号、税率和所有人工调整。',variable=self.ack,command=self.refresh_gate).grid(row=7,column=0,sticky='w')
        frame=ttk.LabelFrame(main,text='异常信息与识别诊断',padding=8); frame.grid(row=8,column=0,sticky='ew',pady=(8,0))
        self.diagnostic=ScrolledText(frame,height=4,font='TkTextFont',wrap='word'); self.diagnostic.pack(fill='x')
        row=ttk.Frame(frame); row.pack(fill='x',pady=(6,0))
        ttk.Button(row,text='查看错误详情',command=self.show_details).pack(side='left')
        ttk.Button(row,text='复制诊断信息',command=self.copy_diagnostic).pack(side='left',padx=8)
        ttk.Button(row,text='查看格式帮助',command=self.help).pack(side='left')
        ttk.Label(row,text='原始文件只读 · 不依赖 Excel · 数据仅保存在本地').pack(side='right')
        self.refresh(); self.after(100,self.poll)
        self.report_callback_exception=self.callback_error
        self.bind('<Configure>',self.monitor_dpi,add='+')

    def monitor_dpi(self,event):
        if event.widget is self:
            dpi=window_dpi(self)
            if dpi!=self._dpi:
                self._dpi=apply_fonts(self,dpi)

    def button(self,parent,text,command):
        b=ttk.Button(parent,text=text,command=command); self.buttons.append(b); return b

    @staticmethod
    def tree(parent,columns,widths):
        box=ttk.Frame(parent); box.pack(fill='both',expand=True)
        tree=ttk.Treeview(box,columns=columns,show='headings',selectmode='browse')
        for name,width in zip(columns,widths):
            tree.heading(name,text=name); tree.column(name,width=round(width*window_dpi(parent)/96),minwidth=55,stretch=False)
        y=ttk.Scrollbar(box,orient='vertical',command=tree.yview); x=ttk.Scrollbar(box,orient='horizontal',command=tree.xview)
        tree.configure(yscrollcommand=y.set,xscrollcommand=x.set)
        tree.grid(row=0,column=0,sticky='nsew'); y.grid(row=0,column=1,sticky='ns'); x.grid(row=1,column=0,sticky='ew')
        box.rowconfigure(0,weight=1); box.columnconfigure(0,weight=1)
        tree.tag_configure('pending',background='#fff2cc')
        return tree

    def run(self,worker,success):
        if self.busy:return
        self.busy=True
        for b in self.buttons:b.configure(state='disabled')
        def task():
            try:self.messages.put((True,worker(),success))
            except Exception as exc:self.messages.put((False,(str(exc),traceback.format_exc()),None))
        threading.Thread(target=task,daemon=True).start()

    def poll(self):
        try:
            ok,value,callback=self.messages.get_nowait()
        except queue.Empty:pass
        else:
            self.busy=False
            for b in self.buttons:b.configure(state='normal')
            if ok:callback(value)
            else:
                if self.quotes is None:self.quote_status.set('未导入成功，请查看下方诊断')
                if self.deliveries is None:self.delivery_status.set('未导入成功，请查看下方诊断')
                self.last_error=value[0]+'\n\n技术详情：\n'+value[1]
                self.refresh()
                messagebox.showerror('操作未完成','详细诊断已显示在下方，可复制给 AI 排查。\n\n'+value[0][:650])
            self.refresh_gate()
        self.after(100,self.poll)

    def choose(self,kind):
        path=filedialog.askopenfilename(title='选择报价单' if kind=='quote' else '选择送货单',filetypes=[('Excel 工作簿','*.xlsx *.xlsm'),('所有文件','*.*')])
        if not path:return
        # 先读取到隔离草稿。取消、读取失败或未确认，都不替换已导入的批次。
        self.run(lambda:import_excel(path,kind,preview=True),lambda result:self.open_review(kind,result))

    def open_review(self,kind,result):
        return ReviewDialog(self,result,kind,lambda reviewed:self.accept_import(kind,reviewed))

    def edit_import(self,kind):
        if self.busy:return
        result=self.quotes if kind=='quote' else self.deliveries
        if result is None:
            messagebox.showinfo('请先选择文件','选择Excel后可在预览中直接修改记录。');return
        dialog=self.open_review(kind,result)
        tree=self.qtree if kind=='quote' else self.dtree
        selected=tree.selection()
        if selected and dialog.tree.exists(selected[0]):
            dialog.tree.selection_set(selected[0]);dialog.tree.see(selected[0])
        return dialog

    def accept_import(self,kind,result):
        # 未改变的记录保留既有人工匹配；修改了送货或报价依据的记录必须重算。
        old_matches=self.matches
        if kind=='quote':self.quotes=result
        else:self.deliveries=result
        self.last_error=''; self.ack.set(False)
        if self.quotes and self.deliveries:self.matches=reconcile(self.quotes.items,self.deliveries.items)
        elif self.deliveries:self.matches=[Match(d,status='尚未导入报价单') for d in self.deliveries.items]
        qmap={(q.file,q.source):q for q in self.quotes.items} if self.quotes else {}
        oldmap={(m.delivery.file,m.delivery.source):m for m in old_matches if m.status=='人工确认'}
        for current in self.matches:
            previous=oldmap.get((current.delivery.file,current.delivery.source))
            if previous and previous.quote:
                q=qmap.get((previous.quote.file,previous.quote.source))
                if current.delivery==previous.delivery and q==previous.quote:
                    current.quote=q;current.status=previous.status;current.factor=previous.factor
                    current.override_price=previous.override_price;current.reason=previous.reason
        self.refresh()

    def diagnostic_text(self):
        chunks=[f'自动对账工具 {__version__} · {datetime.now():%Y-%m-%d %H:%M:%S}',self.last_error]
        for result in (self.quotes,self.deliveries):
            if result:chunks.append(result.diagnostic())
        if self.matches:chunks.append('\n\n'.join(map(str,match_issues(self.matches))))
        return '\n\n'.join(c for c in chunks if c) or '暂无诊断。'

    def refresh(self):
        self.qtree.delete(*self.qtree.get_children()); self.dtree.delete(*self.dtree.get_children())
        if self.quotes:
            for i,q in enumerate(self.quotes.items):
                self.qtree.insert('', 'end',iid=str(i),values=[q.serial,q.name,q.spec,q.unit,q.edit_values.get('price','') if q.validation_errors else q.price if q.price is not None else '缺价',q.effective or '',q.note],tags=('pending',) if q.validation_errors else ())
        for i,m in enumerate(self.matches):
            if self.pending_only.get() and m.ready:continue
            d=m.delivery
            self.dtree.insert('','end',iid=str(i),values=[d.serial,d.name,d.spec,d.unit,d.edit_values.get('quantity','') if 'quantity' in d.validation_errors else d.quantity,d.day or '',d.order,
                                                       m.price if m.ready else '',f'{m.amount:,.2f}' if m.ready else '',m.status,d.source],tags=() if m.ready else ('pending',))
        qn=len(self.quotes.items) if self.quotes else 0; dn=len(self.deliveries.items) if self.deliveries else 0
        ready=sum(m.ready for m in self.matches)
        self.summary.set(f'报价商品 {qn}   送货记录 {dn}   已匹配 {ready}   待确认 {dn-ready}')
        for result,var in ((self.quotes,self.quote_status),(self.deliveries,self.delivery_status)):
            if result:
                state=f'强制导入 · {result.unresolved}项待修正' if result.forced else '已确认导入' if result.reviewed else '已导入'
                var.set(f'{Path(result.path).name} · {len(result.items)}条 · {state}')
        self.diagnostic.configure(state='normal'); self.diagnostic.delete('1.0','end'); self.diagnostic.insert('1.0',self.diagnostic_text()); self.diagnostic.configure(state='disabled')
        self.refresh_gate()

    def refresh_gate(self):
        unresolved=any(r and r.unresolved for r in (self.quotes,self.deliveries))
        enabled=not self.busy and not unresolved and bool(self.matches) and all(m.ready for m in self.matches) and self.ack.get()
        self.formal_button.configure(state='normal' if enabled else 'disabled')

    def resolve_selected(self):
        if self.busy or not self.quotes:return
        selected=self.dtree.selection()
        if not selected:return
        m=self.matches[int(selected[0])]
        win=tk.Toplevel(self); win.title('确认商品与结算单价'); fit_window(win,1050,650); win.transient(self); win.grab_set()
        body=ttk.Frame(win,padding=14); body.pack(fill='both',expand=True)
        d=m.delivery
        ttk.Label(body,text=f'送货序号 {d.serial}：{d.name}   单位：{d.unit}   数量：{d.quantity}',wraplength=980,font=('Microsoft YaHei UI',11,'bold')).pack(anchor='w')
        ttk.Label(body,text=f'来源 {d.source}；备注：{d.note or "无"}；候选项只作提示，请核对颜色、品牌、尺寸顺序、开口状态及等级。',wraplength=980).pack(anchor='w',pady=8)
        query=tk.StringVar(); ttk.Label(body,text='搜索全部报价商品（留空显示候选）：').pack(anchor='w')
        ttk.Entry(body,textvariable=query).pack(fill='x')
        tree=self.tree(body,['序号','商品名称','规格','单位','单价','备注'],[180,300,150,65,90,260])
        choices=m.candidates or ([m.quote] if m.quote else [])
        def search(*args):
            tree.delete(*tree.get_children())
            needle=query.get().strip().lower()
            values=[q for q in self.quotes.items if needle in (q.name+' '+q.spec+' '+q.source).lower()] if needle else choices
            for q in values:
                index=self.quotes.items.index(q)
                tree.insert('','end',iid=str(index),values=[q.serial,q.name,q.spec,q.unit,q.price if q.price is not None else '缺价',q.note])
        query.trace_add('write',search); search()
        form=ttk.Frame(body); form.pack(fill='x',pady=8)
        factor=tk.StringVar(value=str(m.factor)); override=tk.StringVar(value='' if m.override_price is None else str(m.override_price)); reason=tk.StringVar(value=m.reason)
        ttk.Label(form,text='每个送货单位包含几个报价单位：').grid(row=0,column=0,sticky='w')
        ttk.Entry(form,textvariable=factor,width=12).grid(row=0,column=1)
        ttk.Label(form,text='例如 1箱=20盒 → 填20；默认1，单位不同必须确认。').grid(row=0,column=2,padx=8,sticky='w')
        ttk.Label(form,text='直接指定有效单价（可选）：').grid(row=1,column=0,sticky='w',pady=8)
        ttk.Entry(form,textvariable=override,width=12).grid(row=1,column=1)
        ttk.Label(form,text='留空采用报价×倍率；赠送可填0。').grid(row=1,column=2,padx=8,sticky='w')
        ttk.Label(body,text='确认原因（必填，写明别名、单位换算依据或价格调整原因）：').pack(anchor='w')
        ttk.Entry(body,textvariable=reason).pack(fill='x',pady=6)
        def apply():
            try:
                selection=tree.selection()
                if not selection:raise ValueError('请先选择一条报价商品。')
                q=self.quotes.items[int(selection[0])]
                f=number(factor.get())
                if f is None:raise ValueError('请输入单位倍率。')
                confirm(m,q,f,number(override.get()),reason.get())
            except ValueError as exc:messagebox.showerror('请检查输入',str(exc),parent=win); return
            self.ack.set(False); self.refresh(); win.destroy()
        ttk.Button(body,text='确认此记录',command=apply).pack(anchor='e',pady=8)

    def export(self,draft):
        if not self.quotes or not self.deliveries or not self.matches:
            messagebox.showinfo('请先导入','请先导入报价单和送货单。');return
        try:
            rate=number(self.tax.get())
            if rate is None or not 0<=rate<=100:raise ValueError('税率请输入 0～100 的数字。')
            rate/=100
        except ValueError as exc:messagebox.showerror('税率错误',str(exc));return
        path=filedialog.asksaveasfilename(title='保存对账单',defaultextension='.xlsx',initialfile='对账单_待核对稿.xlsx' if draft else '对账单.xlsx',filetypes=[('Excel 工作簿','*.xlsx')])
        if not path:return
        acknowledged=self.ack.get()
        self.run(lambda:export_statement(path,self.quotes,self.deliveries,self.matches,draft=draft,acknowledged=acknowledged,tax_rate=rate),
                 lambda total:messagebox.showinfo('生成完成',f'已保存：{path}\n{"已确认部分小计" if draft else "合计"}：{total:,.2f}\n'+('待确认明细已标黄，不能作为最终结算。' if draft else '可使用 Excel/WPS 打开。')))

    def save_progress(self):
        if not self.quotes or not self.deliveries:
            messagebox.showinfo('请先导入','保存进度前需要导入两份文件。');return
        path=filedialog.asksaveasfilename(defaultextension='.json',initialfile='对账工作进度.json',filetypes=[('工作进度','*.json')])
        if path:self.run(lambda:save_session(path,self.quotes,self.deliveries,self.matches),lambda _:messagebox.showinfo('已保存','报价数据及人工确认已保存在本地工作进度文件。'))

    def open_progress(self):
        path=filedialog.askopenfilename(filetypes=[('工作进度','*.json')])
        if path:
            self.quotes=None;self.deliveries=None;self.matches=[];self.ack.set(False)
            self.quote_status.set('正在恢复进度');self.delivery_status.set('正在恢复进度');self.refresh()
            self.run(lambda:load_session(path),self.accept_session)

    def accept_session(self,value):
        self.quotes,self.deliveries,self.matches=value; self.last_error=''; self.ack.set(False)
        self.quote_status.set(Path(self.quotes.path).name); self.delivery_status.set(Path(self.deliveries.path).name); self.refresh()

    def copy_diagnostic(self):
        self.clipboard_clear(); self.clipboard_append(self.diagnostic_text()); self.update()
        messagebox.showinfo('已复制','已复制识别诊断，可粘贴给 AI 排查。内容包含商品和文件信息，请按需脱敏。')

    def show_details(self):
        win=tk.Toplevel(self); win.title('完整识别诊断'); fit_window(win,950,650)
        t=ScrolledText(win,wrap='word',font='TkTextFont'); t.pack(fill='both',expand=True);t.insert('1.0',self.diagnostic_text());t.configure(state='disabled')

    def help(self):
        path=resource('docs/excel-format-guide.md')
        win=tk.Toplevel(self);win.title('Excel 格式帮助');fit_window(win,950,650)
        t=ScrolledText(win,wrap='word',font='TkTextFont');t.pack(fill='both',expand=True);t.insert('1.0',path.read_text(encoding='utf-8'));t.configure(state='disabled')

    def callback_error(self,kind,value,tb):
        self.last_error='界面操作异常：'+str(value)+'\n'+''.join(traceback.format_exception(kind,value,tb))
        self.refresh(); messagebox.showerror('操作未完成','诊断已显示，请复制错误详情以便排查。')


def main():
    App().mainloop()
