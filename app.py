"""tkinter application for coal_sampling."""
from __future__ import annotations
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from decimal import Decimal
from pathlib import Path
from coal_sampling import SamplingEngine, Parameters

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title("煤岩地质编录自动取样 V1.0"); self.geometry("1280x760"); self.configure(bg="white")
        self.engine=SamplingEngine(); self.input=tk.StringVar(); self.output=tk.StringVar(); self.status=tk.StringVar(value="就绪")
        self._build()
    def _build(self):
        top=ttk.Frame(self); top.pack(fill="x", padx=12,pady=10)
        for row,var,label,cmd in [(0,self.input,"原始Excel",self.pick_input),(1,self.output,"导出成果",self.pick_output)]:
            ttk.Label(top,text=label).grid(row=row,column=0,sticky="w",padx=4,pady=4); ttk.Entry(top,textvariable=var).grid(row=row,column=1,sticky="ew",padx=4); ttk.Button(top,text="选择文件",command=cmd).grid(row=row,column=2,padx=4)
        top.columnconfigure(1,weight=1)
        actions=ttk.Frame(top); actions.grid(row=0,column=3,rowspan=2,padx=12); ttk.Button(actions,text="开始分析",command=self.analyze).pack(side="left",padx=3); ttk.Button(actions,text="按选择重新计算",command=self.analyze).pack(side="left",padx=3); ttk.Button(actions,text="参数设置",command=self.parameters).pack(side="left",padx=3); ttk.Button(actions,text="导出成果",command=self.export).pack(side="left",padx=3)
        self.nb=ttk.Notebook(self); self.nb.pack(fill="both",expand=True,padx=12); self.tables={}
        for name in ("取样成果","未取样煤层复核","特殊岩性未取样复核","原始编录","运行记录"): self.add_page(name)
        ttk.Label(self,textvariable=self.status,anchor="w",background="white").pack(fill="x",padx=12,pady=5)
    def add_page(self,name):
        frame=ttk.Frame(self.nb); self.nb.add(frame,text=name); tree=ttk.Treeview(frame,show="headings"); y=ttk.Scrollbar(frame,orient="vertical",command=tree.yview); x=ttk.Scrollbar(frame,orient="horizontal",command=tree.xview); tree.configure(yscrollcommand=y.set,xscrollcommand=x.set); tree.grid(row=0,column=0,sticky="nsew"); y.grid(row=0,column=1,sticky="ns"); x.grid(row=1,column=0,sticky="ew"); frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1); self.tables[name]=tree
    def pick_input(self): self.input.set(filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")]))
    def pick_output(self): self.output.set(filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")]))
    def analyze(self):
        if not self.input.get(): return messagebox.showwarning("提示","请选择原始Excel")
        try: self.engine.load_xlsx(self.input.get()); self.refresh(); self.status.set(f"分析完成：{len(self.engine.samples)} 个样品")
        except Exception as e: messagebox.showerror("分析失败",str(e)); self.status.set("分析失败")
    def refresh(self):
        data={"取样成果":self.engine.result_rows(),"未取样煤层复核":[x for x in self.engine.review_rows() if x.get("Reason")=="THIN_COAL_NOT_SAMPLED"],"特殊岩性未取样复核":[x for x in self.engine.review_rows() if x.get("Reason")=="SPECIAL_NOT_SAMPLED"],"原始编录":self.engine.original_rows(),"运行记录":self.engine.review_rows()}
        for name, rows in data.items():
            tree=self.tables[name]; tree.delete(*tree.get_children()); keys=list(rows[0]) if rows else ["信息"]; tree["columns"]=keys
            for k in keys: tree.heading(k,text=k); tree.column(k,width=130,anchor="center")
            for i,row in enumerate(rows): tree.insert("","end",values=[str(row.get(k,"")) for k in keys],tags=("even" if i%2 else "odd",))
            tree.tag_configure("odd",background="white"); tree.tag_configure("even",background="#eef7ff")
    def parameters(self):
        win=tk.Toplevel(self); win.title("参数设置"); win.transient(self); entries={}
        for i,(k,v) in enumerate(self.engine.params.jsonable().items()): ttk.Label(win,text=k).grid(row=i,column=0,padx=8,pady=3,sticky="w"); e=ttk.Entry(win); e.insert(0, json.dumps(v,ensure_ascii=False) if isinstance(v,list) else str(v)); e.grid(row=i,column=1,padx=8); entries[k]=e
        def save():
            try:
                for k,e in entries.items():
                    value=e.get(); old=getattr(self.engine.params,k); setattr(self.engine.params,k,tuple(json.loads(value)) if isinstance(old,tuple) else Decimal(value))
                win.destroy(); self.status.set("参数已保存，将在下一次完整计算时生效")
            except Exception as ex: messagebox.showerror("参数错误",str(ex))
        ttk.Button(win,text="保存",command=save).grid(row=len(entries),column=0,columnspan=2,pady=8)
    def export(self):
        if not self.engine.samples: return messagebox.showwarning("提示","请先分析")
        if not self.output.get(): self.pick_output()
        if not self.output.get(): return
        import openpyxl
        wb=openpyxl.Workbook(); wb.remove(wb.active)
        for name,rows in (("自动取样结果",self.engine.result_rows()),("未取样煤层复核",[x for x in self.engine.review_rows() if x.get("Reason")=="THIN_COAL_NOT_SAMPLED"]),("特殊岩性未取样复核",[x for x in self.engine.review_rows() if x.get("Reason")=="SPECIAL_NOT_SAMPLED"]),("原始编录",self.engine.original_rows()),("运行记录",self.engine.review_rows())):
            ws=wb.create_sheet(name); keys=list(rows[0]) if rows else ["信息"]; ws.append(keys)
            for row in rows: ws.append([row.get(k,"") for k in keys])
            ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
        wb.save(self.output.get()); self.status.set("成果已导出："+self.output.get()); messagebox.showinfo("完成","成果导出成功")
if __name__ == "__main__": App().mainloop()
