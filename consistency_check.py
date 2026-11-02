"""Read-only Excel comparison; independent of the analysis database and calculations."""
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import hashlib
import os
from pathlib import Path
import queue
import re
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def text(v):
    return '' if v is None else str(v).strip()


def name_key(v):
    return re.sub(r'\s+', '', text(v)).casefold()


def file_hash(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    return h.hexdigest()


def bas_key(value):
    # Preserve five-digit suffix: numeric dotted IDs may already have lost zeroes.
    s=text(value)
    if isinstance(value,(int,float)):
        if isinstance(value,float) and not value.is_integer():return None
        s=str(int(value))
    m=re.fullmatch(r'(\d{6})\.?([0-9]{5})',s)
    return m[1]+'.'+m[2] if m else None


def sheet_names(path):
    w=load_workbook(path,read_only=True)
    try:return w.sheetnames
    finally:w.close()


def read_sheet(path,sheet,header_row=1):
    path=Path(path).resolve(); before=file_hash(path)
    if header_row<1:raise ValueError('表头行号必须大于0')
    wf=load_workbook(path,read_only=True,data_only=False)
    wv=None
    try:
        wv=load_workbook(path,read_only=True,data_only=True)
        sf,sv=wf[sheet],wv[sheet]
        heads=list(next(sf.iter_rows(min_row=header_row,max_row=header_row),[]))
        headers=[text(c.value) for c in heads]
        while headers and not headers[-1]:headers.pop()
        if not headers:raise ValueError(f'{sheet} 第{header_row}行没有表头')
        rows=[]
        for ri,(fr,vr) in enumerate(zip(sf.iter_rows(min_row=header_row+1,max_col=len(headers)),sv.iter_rows(min_row=header_row+1,max_col=len(headers))),header_row+1):
            if all(c.value is None for c in fr):continue
            cells=[]
            for ci,(f,v) in enumerate(zip(fr,vr),1):
                error=''
                if f.data_type=='f' and v.value is None:error='公式无缓存结果；用Excel打开、计算并保存后重试'
                elif v.data_type=='e':error='Excel错误值：'+text(v.value)
                cells.append({'value':v.value,'formula':f.value if f.data_type=='f' else '',
                              'error':error,'format':f.number_format,'coord':f'{get_column_letter(ci)}{ri}'})
            rows.append((ri,cells))
        if file_hash(path)!=before:raise ValueError('读取过程中源文件发生变化，请重新选择文件')
        return {'path':str(path),'sheet':sheet,'header_row':header_row,'headers':headers,'rows':rows,'sha256':before}
    finally:
        wf.close()
        if wv:wv.close()


def number(cell):
    value=cell['value']
    if value is None or isinstance(value,(bool,datetime)):return None
    s=text(value);explicit=s.endswith('%')
    if explicit:s=s[:-1].strip()
    try:
        n=Decimal(s)
        if not n.is_finite():return None
    except InvalidOperation:return None
    # Excel 0.25 formatted as 25% represents 25 percentage points.
    fmt=re.sub(r'"[^"]*"|\\.', '', cell.get('format',''))
    if '%' in fmt and not explicit:n*=100
    return n


def display(cell):
    if cell is None:return ''
    return cell['formula'] if cell['error'] and cell['formula'] else text(cell['value'])


def compare_cells(a,b,precision):
    if a['error'] or b['error']:return '无法核查','；'.join(x for x in (a['error'],b['error']) if x)
    av,bv=text(a['value']),text(b['value'])
    if not av and not bv:return '双方空白','未核查，空白不代表数值一致'
    if not av or not bv:return '缺失值','A空白' if not av else 'B空白'
    an,bn=number(a),number(b)
    if an is not None and bn is not None:
        with localcontext() as ctx:
            ctx.prec=400
            aa,bb=an,bn
            if precision is not None:
                q=Decimal(1).scaleb(-precision)
                aa,bb=an.quantize(q,rounding=ROUND_HALF_UP),bn.quantize(q,rounding=ROUND_HALF_UP)
            same=aa==bb
            return ('一致' if same else '不一致'),f'数值差 A−B = {an-bn}'
    if (an is None)!=(bn is None):return '不一致','数值与非数值类型不同'
    return ('一致' if av==bv else '不一致'),'文本比较（只忽略首尾空格）'


def compare_tables(a,b,key_a,key_b,mappings):
    """Mappings: (A column index, B column index, decimal places or None)."""
    if not mappings:raise ValueError('请至少添加一组比较列')
    if len({m[0] for m in mappings})!=len(mappings) or len({m[1] for m in mappings})!=len(mappings):raise ValueError('同一列不能重复映射')
    for ai,bi,precision in mappings:
        if ai==key_a or bi==key_b:raise ValueError('BAS ID已用作匹配键，不需要加入比较列')
        if precision is not None and not 0<=precision<=10:raise ValueError('小数位需为0至10')
    results=[]
    def add(status,bas='',item='',ac=None,bc=None,reason='',aloc='',bloc=''):
        results.append([status,bas,item,display(ac),display(bc),aloc or (ac['coord'] if ac else ''),bloc or (bc['coord'] if bc else ''),reason])
    groups=[]
    for side,table,key in [('A',a,key_a),('B',b,key_b)]:
        grouped=defaultdict(list)
        for ri,cells in table['rows']:
            c=cells[key];bas=bas_key(c['value']) if not c['error'] else None
            if bas is None:
                add('无法匹配',item='BAS ID',ac=c if side=='A' else None,bc=c if side=='B' else None,
                    reason=f'{side}表BAS ID为空、格式无效或公式未计算；该行未参与比较。'+c['error'])
            else:grouped[bas].append((ri,cells))
        groups.append(grouped)
    ga,gb=groups
    for bas in sorted(set(ga)|set(gb)):
        ar,br=ga.get(bas,[]),gb.get(bas,[])
        if len(ar)>1 or len(br)>1:
            add('重复BAS',bas,reason='无法唯一匹配；请选择每个BAS一行的Sheet，不自动取首行或平均',
                aloc=', '.join(f'{get_column_letter(key_a+1)}{r[0]}' for r in ar),bloc=', '.join(f'{get_column_letter(key_b+1)}{r[0]}' for r in br));continue
        if not ar or not br:
            add('缺少样本',bas,reason='A表缺少' if not ar else 'B表缺少',
                aloc=f'{get_column_letter(key_a+1)}{ar[0][0]}' if ar else '',bloc=f'{get_column_letter(key_b+1)}{br[0][0]}' if br else '');continue
        for ai,bi,p in mappings:
            ac,bc=ar[0][1][ai],br[0][1][bi];status,reason=compare_cells(ac,bc,p)
            rule='精确数值' if p is None else f'四舍五入至{p}位小数'
            add(status,bas,a['headers'][ai]+' ↔ '+b['headers'][bi],ac,bc,rule+'；'+reason)
    counts=Counter(r[0] for r in results)
    bad=sum(v for k,v in counts.items() if k!='一致')
    if not counts.get('一致') and not bad:conclusion='无可核查数据'
    elif bad:conclusion='存在差异或未完成核查项，详见差异明细'
    else:conclusion='所选范围内一致（未选列不在结论范围内）'
    return {'a':a,'b':b,'keys':(key_a,key_b),'mappings':mappings,'results':results,'counts':counts,'conclusion':conclusion}


def add_paired_sheet(w,report):
    """Show both source rows together without inventing duplicate-BAS matches."""
    from openpyxl.comments import Comment
    ws=w.create_sheet('逐样本核查',0)
    mappings=report['mappings']
    headers=['BAS ID','数据来源','源Sheet / 行','核查状态']+[
        report['a']['headers'][ai] if report['a']['headers'][ai]==report['b']['headers'][bi]
        else report['a']['headers'][ai]+' ↔ '+report['b']['headers'][bi]
        for ai,bi,_ in mappings]
    ws.append(headers)
    groups=[];invalid=[]
    for side in range(2):
        table=report[('a','b')[side]];key=report['keys'][side]
        grouped=defaultdict(list)
        for ri,cells in table['rows']:
            cell=cells[key];bas=bas_key(cell['value']) if not cell['error'] else None
            if bas:grouped[bas].append((ri,cells))
            else:invalid.append((side,ri,cells))
        groups.append(grouped)
    red=PatternFill('solid',fgColor='FCE8E6')
    def mark(cell,reason):
        cell.fill=red;cell.font=Font(name='Arial',size=10,color='B3261E')
        cell.comment=Comment(reason,'一致性核查')
    def append(bas,side,record,status,issues=None,reason=''):
        table=report[('a','b')[side]]
        ri,cells=record if record else (None,None)
        values=[bas,('A · Summary','B · 人工结果')[side],
                f"{table['sheet']} / {ri}" if ri is not None else '',status]
        source_cells=[cells[m[side]] if cells else None for m in mappings]
        values += [(c['formula'] if c['error'] and c['formula'] else c['value']) if c else None for c in source_cells]
        ws.append(values);r=ws.max_row
        for cell in ws[r]:
            if isinstance(cell.value,str):cell.data_type='s'
            cell.font=Font(name='Arial',size=10)
            cell.alignment=Alignment(vertical='center',wrap_text=True)
            if side==1:cell.fill=PatternFill('solid',fgColor='F5F8FC')
        for j,c in enumerate(source_cells,5):
            if c:
                ws.cell(r,j).number_format=c.get('format') or 'General'
                ws.cell(r,j).comment=Comment(f"{table['sheet']}!{c['coord']}"+(('；'+c['error']) if c['error'] else ''),'一致性核查')
        if issues is not None:
            for j,(state,why) in enumerate(issues,5):
                if state!='一致':mark(ws.cell(r,j),state+'；'+why)
            if any(state!='一致' for state,_ in issues):mark(ws.cell(r,4),'存在差异或未核查项；详见标红单元格')
        elif reason:
            mark(ws.cell(r,1),reason);mark(ws.cell(r,4),reason)
    for bas in sorted(set(groups[0])|set(groups[1])):
        ar,br=groups[0].get(bas,[]),groups[1].get(bas,[])
        if len(ar)==len(br)==1:
            issues=[compare_cells(ar[0][1][ai],br[0][1][bi],p) for ai,bi,p in mappings]
            status='一致' if all(st=='一致' for st,_ in issues) else '存在差异 / 未核查项'
            append(bas,0,ar[0],status,issues);append(bas,1,br[0],status,issues)
        else:
            status='重复BAS · 未匹配' if len(ar)>1 or len(br)>1 else '缺少样本'
            reason='同一BAS有多行，未自动配对；请按同一管或同一统计口径整理后核查。' if '重复' in status else '一侧缺少该样本，未参与数值比较。'
            for side,records in enumerate((ar,br)):
                for record in records or [None]:append(bas,side,record,status,reason=reason)
    for side,ri,cells in invalid:
        key=report['keys'][side]
        append(display(cells[key]),side,(ri,cells),'BAS无效 · 未匹配',reason='BAS ID为空、格式无效或公式未计算。')
    for c in ws[1]:
        c.data_type='s';c.font=Font(name='Arial',size=10,bold=True,color='1F3B60');c.fill=PatternFill('solid',fgColor='E8EEF7');c.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height=36
    ws.freeze_panes='E2';ws.auto_filter.ref=ws.dimensions;ws.sheet_view.showGridLines=False
    for i in range(1,len(headers)+1):ws.column_dimensions[get_column_letter(i)].width=([20,22,30,29][i-1] if i<=4 else 24)


def export_report(report,destination):
    dest=Path(destination).resolve()
    if dest.suffix.lower()!='.xlsx':raise ValueError('报告请保存为.xlsx')
    if dest in (Path(report['a']['path']),Path(report['b']['path'])):raise ValueError('不能覆盖输入文件，请选择新的报告文件名')
    for side in ('a','b'):
        if file_hash(report[side]['path'])!=report[side]['sha256']:raise ValueError('核查后源文件发生变化，请重新核查后导出')
    w=Workbook();overview=w.active;overview.title='核查概览'
    def row(ws,values):
        ws.append(values)
        for c in ws[ws.max_row]:
            if isinstance(c.value,str):c.data_type='s'
    row(overview,['项目','内容']);row(overview,['结论',report['conclusion']]);row(overview,['核查时间',datetime.now().astimezone().isoformat(timespec='seconds')])
    if report['a']['sha256']==report['b']['sha256']:row(overview,['提示','两份文件的SHA256相同，文件内容完全相同；本次不构成两份独立结果的核查。'])
    row(overview,['规则','按BAS ID匹配；重复BAS不自动合并；空白不等于0；不将#转0；差异不判定哪一方正确。'])
    row(overview,['百分比规则','带Excel百分比格式的数值统一为百分数值（0.25→25）；25%文本也转换为25。'])
    for side in ('a','b'):
        t=report[side];key=report['keys'][0 if side=='a' else 1]
        for label,val in [('文件',t['path']),('Sheet',t['sheet']),('表头行',t['header_row']),('SHA256',t['sha256']),('BAS列',f'{get_column_letter(key+1)}: {t["headers"][key]}'),('数据行数',len(t['rows']))]:row(overview,[side.upper()+label,val])
    for k,v in sorted(report['counts'].items()):row(overview,[k,v])
    for ai,bi,p in report['mappings']:row(overview,['比较列',f'A {get_column_letter(ai+1)} {report["a"]["headers"][ai]} ↔ B {get_column_letter(bi+1)} {report["b"]["headers"][bi]}；'+('精确数值' if p is None else f'{p}位小数')])
    for side,index in [('a',0),('b',1)]:
        used={m[index] for m in report['mappings']}|{report['keys'][index]}
        for i,h in enumerate(report[side]['headers']):
            if i not in used:row(overview,[side.upper()+'未核查列',f'{get_column_letter(i+1)}: {h or "（无表头）"}'])
    headers=['状态','BAS ID','指标（A ↔ B）','A原值 / 公式','B原值 / 公式','A单元格','B单元格','说明']
    for title,rows in [('差异明细',[r for r in report['results'] if r[0]!='一致']),('全部比较结果',report['results'])]:
        ws=w.create_sheet(title);row(ws,headers)
        for r in rows:row(ws,r)
    for ws in w:
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions;ws.sheet_view.showGridLines=False
        for cells in ws:
            for c in cells:
                c.font=Font(name='Arial',size=10,bold=c.row==1)
                c.alignment=Alignment(vertical='top',wrap_text=True)
                if c.row==1:c.fill=PatternFill('solid',fgColor='E8EEF7')
                elif ws.title!='核查概览' and ws.cell(c.row,1).value!='一致':c.fill=PatternFill('solid',fgColor='FFF0EE')
        widths=[26,100] if ws.title=='核查概览' else [18,20,46,25,25,20,20,65]
        for i,width in enumerate(widths,1):ws.column_dimensions[get_column_letter(i)].width=width
    add_paired_sheet(w,report)
    fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=dest.parent);os.close(fd)
    try:w.save(tmp);os.replace(tmp,dest)
    finally:
        w.close()
        if os.path.exists(tmp):os.unlink(tmp)


class CheckPage:
    def __init__(self,app,parent):
        self.app=app;self.parent=parent;self.sides=[];self.tables=None;self.mappings=[];self.report=None
        app.add_header_actions('一致性核查',[('读取所选Sheet',self.load,'TButton'),('核查并导出报告',self.run,'Primary.TButton')])
        for i,label in enumerate(('A：自动 Summary','B：人工 Processed Data')):
            frame=ttk.Frame(parent);frame.pack(fill='x',padx=14,pady=10)
            ttk.Label(frame,text=label).grid(row=0,column=0,sticky='w')
            path=tk.StringVar();sheet=tk.StringVar();header=tk.StringVar(value='1')
            ttk.Entry(frame,textvariable=path,state='readonly').grid(row=1,column=0,sticky='ew',padx=(0,12));frame.columnconfigure(0,weight=1)
            button=ttk.Button(frame,text='选择Excel',command=lambda idx=i:app.guard(lambda:self.choose(idx)));button.grid(row=1,column=1,padx=6)
            box=ttk.Combobox(frame,textvariable=sheet,state='readonly',width=24);box.grid(row=1,column=2,padx=6)
            ttk.Label(frame,text='表头行').grid(row=1,column=3,padx=6)
            ttk.Spinbox(frame,from_=1,to=999,textvariable=header,width=5).grid(row=1,column=4)
            self.sides.append({'path':path,'sheet':sheet,'header':header,'box':box})
            for v in (path,sheet,header):v.trace_add('write',self.invalidate_tables)
        f=ttk.Frame(parent);f.pack(fill='x',padx=14,pady=8)
        self.keys=[]
        for i in range(2):
            ttk.Label(f,text=('A','B')[i]+' BAS ID列').grid(row=0,column=i*2,padx=6)
            v=tk.StringVar();box=ttk.Combobox(f,textvariable=v,state='readonly',width=32);box.grid(row=0,column=i*2+1,padx=6);self.keys.append((v,box))
            v.trace_add('write',self.invalidate_result)
        edit=ttk.Frame(parent);edit.pack(fill='x',padx=14,pady=8)
        self.cols=[]
        for i in range(2):
            v=tk.StringVar();box=ttk.Combobox(edit,textvariable=v,state='readonly',width=28);box.pack(side='left',padx=(0,10));self.cols.append((v,box))
        self.precision=tk.StringVar(value='精确数值');ttk.Combobox(edit,textvariable=self.precision,state='readonly',values=['精确数值']+[f'{i}位小数' for i in range(11)],width=12).pack(side='left',padx=6)
        app.button(edit,'添加对应列',self.add)
        app.button(edit,'删除',self.delete)
        self.tree=app.tree(parent,['A列','B列','比较精度'],[330,330,150],height=10)
        self.tree.bind('<Double-1>',lambda e:app.guard(self.edit_precision))
        footer=ttk.Frame(parent);footer.pack(fill='x',padx=14,pady=8)
        app.button(footer,'同名列自动对应',self.auto)
        self.status=tk.StringVar(value='选择两个Excel及Sheet。每个BAS应各有一行；重复BAS会单独报告。')
        ttk.Label(parent,textvariable=self.status,wraplength=950).pack(fill='x',padx=18,pady=10)

    def invalidate_result(self,*_):self.report=None

    def invalidate_tables(self,*_):
        self.tables=None;self.mappings=[];self.report=None
        if hasattr(self,'tree'):self.render()
        if hasattr(self,'status'):self.status.set('文件、Sheet或表头已变更，请读取所选Sheet。')

    def choose(self,i):
        path=filedialog.askopenfilename(parent=self.app,title='选择核查文件',filetypes=[('Excel','*.xlsx *.xlsm')])
        if not path:return
        names=sheet_names(path);s=self.sides[i];s['path'].set(path);s['box']['values']=names;s['sheet'].set(names[0] if names else '')

    def background(self,title,job,done):
        # No Store connection: this utility never mutates the project database.
        win=tk.Toplevel(self.app);win.title(title);win.geometry('420x130');win.transient(self.app);win.grab_set();win.protocol('WM_DELETE_WINDOW',lambda:None)
        ttk.Label(win,text=title).pack(pady=18);bar=ttk.Progressbar(win,mode='indeterminate');bar.pack(fill='x',padx=20);bar.start()
        q=queue.Queue()
        def worker():
            try:q.put((True,job()))
            except Exception as exc:q.put((False,str(exc)))
        threading.Thread(target=worker,daemon=True).start()
        def poll():
            try:ok,value=q.get_nowait()
            except queue.Empty:self.app.after(100,poll);return
            bar.stop();win.destroy()
            if ok:self.app.guard(lambda:done(value))
            else:messagebox.showerror('核查未完成',value,parent=self.app)
        self.app.after(100,poll)

    def load(self):
        specs=[]
        for s in self.sides:
            if not s['path'].get() or not s['sheet'].get():raise ValueError('请先选择两份Excel和各自Sheet')
            specs.append((s['path'].get(),s['sheet'].get(),int(s['header'].get())))
        self.background('读取核查数据',lambda:[read_sheet(*s) for s in specs],self.loaded)

    def loaded(self,tables):
        self.tables=tables;self.mappings=[];self.report=None
        for i,t in enumerate(tables):
            choices=[f'{get_column_letter(n+1)} | {h or "（无表头）"}' for n,h in enumerate(t['headers'])]
            self.keys[i][1]['values']=choices;self.cols[i][1]['values']=choices
            guess=next((n for n,h in enumerate(t['headers']) if name_key(h) in ('basid','bas编号','样品编号','样本编号')),-1)
            self.keys[i][0].set(choices[guess] if guess>=0 else '')
            self.cols[i][0].set('')
        if all(v.get() for v,b in self.keys):self.auto()
        else:self.render();self.status.set('请指定两边的BAS ID列，再点击同名列自动对应。')

    def index(self,box):
        value=box.current()
        if value<0:raise ValueError('请从下拉菜单选择列')
        return value

    def auto(self):
        if not self.tables:raise ValueError('请先读取所选Sheet')
        ka,kb=[self.index(b) for v,b in self.keys];a,b=self.tables
        ac=Counter(name_key(h) for h in a['headers']);bc=Counter(name_key(h) for h in b['headers'])
        other={name_key(h):i for i,h in enumerate(b['headers'])}
        self.mappings=[]
        for ai,h in enumerate(a['headers']):
            k=name_key(h)
            if k and ac[k]==bc[k]==1 and ai!=ka and other[k]!=kb:self.mappings.append((ai,other[k],None))
        self.render();self.status.set(f'已对应{len(self.mappings)}列。请删除不需要核查的列；双击对应行可改比较精度。不同名称可手动添加。')

    def add(self):
        if not self.tables:raise ValueError('请先读取所选Sheet')
        ai,bi=[self.index(b) for v,b in self.cols];ka,kb=[self.index(b) for v,b in self.keys]
        if ai==ka or bi==kb:raise ValueError('BAS ID不需要添加为比较项')
        if any(ai==m[0] or bi==m[1] for m in self.mappings):raise ValueError('该列已存在，请先删除旧对应关系')
        p=None if self.precision.get()=='精确数值' else int(self.precision.get().split('位')[0])
        self.mappings.append((ai,bi,p));self.render()

    def edit_precision(self):
        selected=self.tree.selection()
        if len(selected)!=1:return
        i=int(selected[0]);ai,bi,p=self.mappings[i]
        win=tk.Toplevel(self.app);win.title('比较精度');win.transient(self.app);win.grab_set()
        value=tk.StringVar(value='精确数值' if p is None else f'{p}位小数')
        ttk.Label(win,text='Events通常选精确数值；百分比、ABS按汇报精度选择。').pack(padx=18,pady=14)
        ttk.Combobox(win,textvariable=value,state='readonly',values=['精确数值']+[f'{n}位小数' for n in range(11)]).pack(padx=18,pady=10)
        def save():
            precision=None if value.get()=='精确数值' else int(value.get().split('位')[0])
            self.mappings[i]=(ai,bi,precision);self.render();win.destroy()
        ttk.Button(win,text='保存',command=save).pack(pady=14)

    def delete(self):
        ids={int(i) for i in self.tree.selection()};self.mappings=[m for i,m in enumerate(self.mappings) if i not in ids];self.render()

    def render(self):
        self.report=None;self.tree.delete(*self.tree.get_children())
        if not self.tables:return
        for i,(ai,bi,p) in enumerate(self.mappings):self.tree.insert('','end',iid=str(i),values=(f'{get_column_letter(ai+1)} | {self.tables[0]["headers"][ai]}',f'{get_column_letter(bi+1)} | {self.tables[1]["headers"][bi]}','精确数值' if p is None else f'{p}位小数'))

    def run(self):
        if not self.tables or not self.mappings:raise ValueError('请先读取Sheet并添加要比较的列')
        ka,kb=[self.index(b) for v,b in self.keys]
        destination=filedialog.asksaveasfilename(parent=self.app,title='保存一致性核查报告',initialfile='一致性核查报告.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
        if not destination:return
        tables=self.tables;maps=list(self.mappings)
        def job():
            report=compare_tables(*tables,ka,kb,maps);export_report(report,destination);return report
        def done(report):
            self.report=report;self.status.set(report['conclusion']+'；'+ '，'.join(f'{k} {v}' for k,v in report['counts'].items()))
            messagebox.showinfo('核查报告已导出',self.status.get()+'\n\n'+destination,parent=self.app)
        self.background('核查并导出报告',job,done)
