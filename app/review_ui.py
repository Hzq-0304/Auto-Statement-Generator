"""导入预览：双击单元格直接编辑，所有修改完成后一次提交。"""
from copy import deepcopy
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from .display import fit_window, window_dpi
from .review import fields_for, prepare_review, edit_item, commit_review, force_warning


class ReviewDialog(tk.Toplevel):
    def __init__(self, parent, result, kind, on_commit):
        super().__init__(parent)
        self.result = prepare_review(deepcopy(result), kind)
        self.initial_values = [deepcopy(i.edit_values) for i in self.result.items]
        self.on_commit = on_commit
        self.fields = fields_for(kind)
        self.editor = None
        self.title('检查并修改报价记录' if kind=='quote' else '检查并修改采购／送货记录')
        self.transient(parent)
        fit_window(self, 1180, 760)
        self.grab_set()
        self.protocol('WM_DELETE_WINDOW', self.cancel)
        body = ttk.Frame(self, padding=14); body.pack(fill='both', expand=True)
        ttk.Label(body, text='先检查和修改，再确认导入', font=('Microsoft YaHei UI', 16, 'bold')).pack(anchor='w')
        ttk.Label(body, text='双击单元格修改；Enter保存、Tab保存并编辑下一格、Esc取消本格。原始Excel不会被改动。').pack(anchor='w', pady=(4,8))
        self.summary = tk.StringVar()
        ttk.Label(body, textvariable=self.summary).pack(anchor='w', pady=(0,8))
        controls=ttk.Frame(body); controls.pack(fill='x', pady=(0,8))
        self.only_errors=tk.BooleanVar(); self.search=tk.StringVar()
        ttk.Checkbutton(controls,text='只看错误记录',variable=self.only_errors,command=self.refresh).pack(side='left')
        ttk.Label(controls,text='搜索：').pack(side='left',padx=(20,4))
        ttk.Entry(controls,textvariable=self.search,width=28).pack(side='left')
        self.search.trace_add('write',lambda *_:self.refresh())
        ttk.Button(controls,text='修改整条记录',command=self.edit_row).pack(side='left',padx=12)
        box=ttk.Frame(body); box.pack(fill='both',expand=True)
        columns=['source','status']+[f for f,_ in self.fields]
        self.tree=ttk.Treeview(box,columns=columns,show='headings',selectmode='browse')
        scale=window_dpi(self)/96
        for key,label in [('source','来源（只读）'),('status','检查结果')]+self.fields:
            width=240 if key in ('status','name','note') else 170 if key=='source' else 120
            self.tree.heading(key,text=label);self.tree.column(key,width=round(width*scale),stretch=False)
        sy=ttk.Scrollbar(box,orient='vertical',command=self.tree.yview)
        sx=ttk.Scrollbar(box,orient='horizontal',command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
        self.tree.grid(row=0,column=0,sticky='nsew');sy.grid(row=0,column=1,sticky='ns');sx.grid(row=1,column=0,sticky='ew')
        box.rowconfigure(0,weight=1);box.columnconfigure(0,weight=1)
        self.tree.tag_configure('error',background='#fff0d6')
        self.tree.bind('<Double-1>',self.begin_cell)
        self.tree.bind('<Return>',lambda _:self.edit_row())
        self.tree.bind('<<TreeviewSelect>>',lambda _:self.show_selected())
        self.details=ScrolledText(body,height=4,wrap='word',font='TkTextFont')
        self.details.pack(fill='x',pady=8)
        bottom=ttk.Frame(body);bottom.pack(fill='x')
        ttk.Button(bottom,text='取消',command=self.cancel).pack(side='left')
        self.force_button=ttk.Button(bottom,text='强制导入…',command=lambda:self.submit(True));self.force_button.pack(side='right',padx=(8,0))
        self.confirm_button=ttk.Button(bottom,text='确认导入',command=lambda:self.submit(False));self.confirm_button.pack(side='right')
        self.refresh()

    def refresh(self):
        # 筛选前保存当前正在编辑的单元格，防止刚输入的值随刷新丢失。
        if self.editor:
            self.save_cell(refresh=False)
        selected=self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        needle=self.search.get().strip().lower()
        for index,item in enumerate(self.result.items):
            if self.only_errors.get() and not item.validation_errors:continue
            if needle and needle not in (' '.join(map(str,item.edit_values.values()))+' '+item.source).lower():continue
            status='；'.join(item.validation_errors.values()) if item.validation_errors else ('已修改 · 正常' if item.edit_history else '正常')
            self.tree.insert('','end',iid=str(index),values=[item.source,status]+[item.edit_values.get(f,'') for f,_ in self.fields],tags=('error',) if item.validation_errors else ())
        if selected and self.tree.exists(selected[0]):self.tree.selection_set(selected[0])
        bad=sum(bool(i.validation_errors) for i in self.result.items)
        changed=sum(bool(i.edit_history) for i in self.result.items)
        self.summary.set(f'共 {len(self.result.items)} 条，已修改 {changed} 条，错误记录 {bad} 条，结构问题 {len(self.result.blocking_issues)} 项')
        self.confirm_button.configure(state='normal' if not self.result.unresolved else 'disabled')
        self.force_button.configure(state='normal' if self.result.unresolved else 'disabled')
        self.show_selected()

    def show_selected(self):
        selected=self.tree.selection()
        lines=[]
        if selected:
            item=self.result.items[int(selected[0])]
            lines.append(f'{item.source}：'+('；'.join(item.validation_errors.values()) or '字段检查通过'))
            changes=[f'{key}：{item.original_values.get(key, "")} → {value}' for key,value in item.edit_values.items() if value!=item.original_values.get(key,'')]
            if changes:lines.append('本次修改：'+'；'.join(changes))
        if self.result.blocking_issues:
            lines.append('结构问题（需修正Excel后重新读取，或明确选择强制导入）：\n'+'\n'.join(map(str,self.result.blocking_issues)))
        notices='\n'.join(map(str,self.result.issues))
        if notices:lines.append('来源提醒：\n'+notices)
        self.details.configure(state='normal');self.details.delete('1.0','end');self.details.insert('1.0','\n\n'.join(lines) or '选择记录可查看详细问题。');self.details.configure(state='disabled')

    def begin_cell(self,event):
        row=self.tree.identify_row(event.y);column=self.tree.identify_column(event.x)
        if not row or not column:return
        index=int(column[1:])-1
        if index<2:return
        self.open_cell(row,index)

    def open_cell(self,row,index):
        if self.editor:self.save_cell()
        if not self.tree.exists(row):return
        self.tree.selection_set(row)
        bounds=self.tree.bbox(row,f'#{index+1}')
        if not bounds:return
        key=self.fields[index-2][0]
        entry=ttk.Entry(self.tree)
        entry.insert(0,self.result.items[int(row)].edit_values.get(key,''))
        entry.place(x=bounds[0],y=bounds[1],width=bounds[2],height=bounds[3])
        entry.focus_set();entry.selection_range(0,'end')
        self.editor=(entry,int(row),key,index)
        entry.bind('<Return>',lambda _:self.save_cell())
        entry.bind('<Escape>',lambda _:self.cancel_cell())
        entry.bind('<Tab>',lambda _:self.next_cell())
        entry.bind('<FocusOut>',lambda _:self.save_cell())

    def save_cell(self,refresh=True):
        if not self.editor:return 'break'
        entry,index,key,_=self.editor
        self.editor=None
        value=entry.get();entry.destroy()
        edit_item(self.result,index,{key:value})
        if refresh:self.refresh()
        return 'break'

    def cancel_cell(self):
        if self.editor:
            entry=self.editor[0];self.editor=None;entry.destroy()
        return 'break'

    def next_cell(self):
        if not self.editor:return 'break'
        _,row,_,col=self.editor
        self.save_cell()
        if col+1<len(self.fields)+2 and self.tree.exists(str(row)):
            self.open_cell(str(row),col+1)
        return 'break'

    def edit_row(self):
        self.save_cell()
        selected=self.tree.selection()
        if not selected:return
        index=int(selected[0]);item=self.result.items[index]
        dialog=tk.Toplevel(self);dialog.title('修改记录');dialog.transient(self);dialog.grab_set()
        frame=ttk.Frame(dialog,padding=16);frame.pack(fill='both',expand=True)
        ttk.Label(frame,text=f'来源：{item.source}（原始文件不变）').grid(row=0,column=0,columnspan=4,sticky='w',pady=(0,12))
        variables={}
        for n,(key,label) in enumerate(self.fields):
            r,c=1+n//2,(n%2)*2
            ttk.Label(frame,text=label).grid(row=r,column=c,sticky='w',padx=(0,8),pady=6)
            variable=tk.StringVar(value=item.edit_values.get(key,''));variables[key]=variable
            field=ttk.Combobox(frame,textvariable=variable,values=['是','否'],state='readonly',width=27) if key=='tax_included' else ttk.Entry(frame,textvariable=variable,width=29)
            field.grid(row=r,column=c+1,padx=(0,16),sticky='ew')
        def close():dialog.destroy();self.grab_set()
        def save():
            edit_item(self.result,index,{key:var.get() for key,var in variables.items()})
            close();self.refresh()
        bottom=1+(len(self.fields)+1)//2
        ttk.Button(frame,text='保存修改',command=save).grid(row=bottom,column=3,pady=(12,0),sticky='e')
        ttk.Button(frame,text='取消',command=close).grid(row=bottom,column=2,pady=(12,0))
        dialog.protocol('WM_DELETE_WINDOW',close)

    def submit(self,force):
        self.save_cell()
        if force and self.result.unresolved:
            if not messagebox.askyesno('强制导入提醒',force_warning(self.result),icon='warning',default='no',parent=self):return
        try:result=commit_review(self.result,force=force)
        except ValueError as exc:messagebox.showerror('尚未完成修改',str(exc),parent=self);return
        self.on_commit(result)
        self.destroy()

    def cancel(self):
        self.save_cell()
        if [i.edit_values for i in self.result.items] != self.initial_values:
            if not messagebox.askyesno('放弃本次修改', '关闭后，本次尚未确认的修改不会应用。确定关闭吗？',default='no',parent=self):return
        self.destroy()
