#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Flow cytometry summary V3.4.4 — SQLite + Tkinter/CustomTkinter + openpyxl.
Run: py express7_v2.py
Business rules and installation are documented in 使用说明.txt.
"""
from __future__ import annotations
from decimal import Decimal, ROUND_HALF_UP, localcontext
import hashlib
import json
import math
import os
import queue
import re
import sqlite3
import statistics
import tempfile
import threading
import traceback
from consistency_check import CheckPage
from datetime import date, datetime, time, timezone
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont
import customtkinter as ctk
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VERSION = '3.4.6'
RAW_META = ['Experiment', 'Specimen', 'Analyst', 'Date', 'Time', 'Tube']
INPUT_FIELDS = RAW_META + ['BAS ID','Subject ID','Run number','Cycle','Visit','Timepoint']
META = ['Experiment', 'Subject ID', 'Cycle', 'Visit', 'Timepoint', 'Run number', 'BAS ID', 'P/F', 'Analyst', 'Date', 'Time', 'Tube']
PCT = ['CD3+%', 'CD4+%', 'CD8+%', 'NK%', 'CD56brightCD16-%', 'CD56lowCD16+%', 'CD4+ CD45RO%', 'CD8+CD45RO%', 'CD56brightCD16-NKG2D%', 'CD56lowCD16+NKG2D%', 'NK+NKG2D+%']
MFI = ['CD3+MFI', 'CD4+MFI', 'CD8+MFI', 'NK MFI', 'CD56brightCD16-MFI', 'CD56lowCD16+MFI', 'CD4+CD45RO MFI', 'CD8+CD45RO MFI', 'CD56brightCD16-NKG2D MFI', 'CD56lowCD16+NKG2D MFI', 'NK+NKG2D MFI']
EVENTS = ['CD3+Events', 'CD4+Events', 'CD8+Events', 'NK Events', 'CD56brightCD16-Events', 'CD56lowCD16+Events', 'CD4+CD45RO Events', 'CD8+CD45RO Events', 'CD56brightCD16-NKG2D Events', 'CD56lowCD16+NKG2D Events', 'NK+NKG2D Events']
ABS = ['ABS CD3+', 'ABS CD4+', 'ABS CD8+', 'ABS NK', 'ABS CD56brightCD16-', 'ABS CD56lowCD16+', 'ABS CD4+CD45RO', 'ABS CD8+CD45RO', 'ABS CD56brightCD16-NKG2D', 'ABS CD56lowCD16+NKG2D', 'ABS NK+NKG2D']
METRICS = ['CD45 Events'] + PCT + ['Beads Events'] + MFI + EVENTS
LEGACY_SUBJECT_FIELDS = ['样品编号', 'Barcode', 'Patient ID', 'Visit-Cycle', 'Time Point', 'Collection Date', 'Collection Time']
SUBJECT_FIELDS = LEGACY_SUBJECT_FIELDS + ['Cycle']
SUBJECT_VIEW = ['样品编号', 'Barcode', 'Patient ID', 'Cycle', 'Visit-Cycle', 'Time Point', 'Collection Date', 'Collection Time']
SUBJECT_LABELS = {'Cycle':'Cycle（给药周期）', 'Visit-Cycle':'Visit（访视）', 'Time Point':'Timepoint（采血时间）'}
SUBJECT_SPLIT_VIEW = ['样品编号', 'Barcode', 'Patient ID', 'Cycle', 'Visit-Cycle', '采血阶段', '给药前后', '间隔时长', 'Collection Date', 'Collection Time']
SUBJECT_SPLIT_LABELS = {**SUBJECT_LABELS, '采血阶段':'采血阶段（W/D）', '给药前后':'Pre / Post', '间隔时长':'时长（h/D）'}
STATE_LABEL = {'current': '当前有效', 'pending': '待审核', 'superseded': '已被替代', 'rejected': '未采用'}
KIND_LABEL = {'number': '数值', 'events': 'Events整数', 'pct100': '百分比0–100', 'pct1': '比例0–1', 'text': '文本'}
CV_HEADERS = ['Experiment', 'Panel', 'Subject ID', 'Visit', 'Timepoint', 'Run number', 'BAS ID', '管组', '汇报项', '复管数', '复管编号', '记录ID', '均值', '样本SD', 'CV%', '阈值%', '判定', 'P/F提示', 'Cycle']
CV_HEADERS += ['验证样品名称']
def strip_retired_config(value):
    if isinstance(value,dict):
        return {k:strip_retired_config(v) for k,v in value.items() if k.casefold()!='dts'}
    if isinstance(value,list):return [strip_retired_config(v) for v in value]
    return value


def collection_value(value, is_time=False):
    """Reference collection fields only; never substitute analysis timestamps."""
    value = txt(value)
    if not value: return None
    if is_time:
        try: return time.fromisoformat(value)
        except ValueError: raise ValueError('采血时间格式不支持：'+value+'；请使用HH:MM或HH:MM:SS') from None
    try: return datetime.fromisoformat(value).date()
    except ValueError: pass
    for fmt in ('%Y/%m/%d','%m/%d/%Y'):
        try: return datetime.strptime(value,fmt).date()
        except ValueError: pass
    raise ValueError('采血日期格式不支持：'+value+'；请在受试者库核对为YYYY-MM-DD')


def now():
    return datetime.now(timezone.utc).isoformat(timespec='microseconds')


def norm(value):
    return re.sub(r'\s+', '', str(value if value is not None else '')).casefold()


def txt(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_date(value):
    if isinstance(value,datetime): return value.date().isoformat()
    if isinstance(value,date): return value.isoformat()
    text=txt(value)
    if not text: return ''
    try: return datetime.fromisoformat(text.replace('Z','+00:00')).date().isoformat()
    except ValueError: pass
    for fmt in ('%Y-%m-%d','%Y/%m/%d','%m/%d/%Y','%d/%m/%Y','%Y.%m.%d','%Y%m%d','%Y年%m月%d日'):
        try: return datetime.strptime(text,fmt).date().isoformat()
        except ValueError: pass
    return text  # Invalid/ambiguous source text is preserved, never guessed into a date.


def json_value(v):
    if isinstance(v, (datetime, date, time)):
        return v.isoformat()
    if isinstance(v, float) and not math.isfinite(v):
        raise ValueError('原始数据含非有限数值')
    return v


def dumps(v):
    return json.dumps(v, ensure_ascii=False, sort_keys=True, allow_nan=False)


def bas_id(v):
    if not isinstance(v, str):
        raise ValueError('样品编号必须为文本；请核对完整编号后填写，避免小数丢位')
    v = v.strip()
    if re.fullmatch(r'\d{11}', v):
        return v[:6] + '.' + v[6:]
    if re.fullmatch(r'\d{6}\.\d{5}', v):
        return v
    raise ValueError(f'样品编号格式不支持：{v}；当前要求6位项目号+5位样本号')


def run_number(v):
    m = re.fullmatch(r'(?:Run\s*)?(\d+)', txt(v), re.I)
    if not m or not 1 <= int(m[1]) <= 100000:
        raise ValueError(f'Run编号不正确：{v}')
    return int(m[1])


def expand_runs(v):
    out = set()
    for part in txt(v).replace('，', ',').split(','):
        part = re.sub('run', '', part, flags=re.I).strip()
        if '-' in part:
            bits = part.split('-')
            if len(bits) != 2:
                raise ValueError('Run范围示例：1-10,12-20 或 Run11')
            a, b = map(run_number, bits)
            if b < a or b - a > 10000:
                raise ValueError('Run范围顺序或长度不正确')
            out.update(range(a, b + 1))
        else:
            out.add(run_number(part))
    if len(out) > 10000:
        raise ValueError('一次最多设置10000个Run')
    return sorted(out)


def numeric(v, name, kind='number'):
    if isinstance(v, bool) or v is None or txt(v) == '':
        raise ValueError(f'{name} 缺失或不是数值')
    try:
        n = float(v)
    except (ValueError, TypeError):
        raise ValueError(f'{name} 不是数值：{v!r}') from None
    if not math.isfinite(n):
        raise ValueError(f'{name} 不是有限数值')
    if kind == 'events' and (n < 0 or not n.is_integer()):
        raise ValueError(f'{name} 必须为非负整数')
    if kind == 'pct100' and not 0 <= n <= 100:
        raise ValueError(f'{name} 必须在0–100之间')
    if kind == 'pct1' and not 0 <= n <= 1:
        raise ValueError(f'{name} 必须在0–1之间')
    return int(n) if kind == 'events' else n


def metric_number(v, name, kind='number'):
    # Project reporting convention: only a literal source metric '#' means zero.
    return numeric(0 if txt(v)=='#' else v,name,kind)


def infer_kind(name):
    if 'events' in norm(name):
        return 'events'
    if '%' in name:
        return 'pct100'
    return 'number'


def default_config():
    return {'columns': [{'source': h, 'name': h, 'mode': 'raw', 'kind': infer_kind(h)} for h in METRICS]
            + [{'source': e, 'name': a, 'mode': 'abs', 'kind': 'events'} for e, a in zip(EVENTS, ABS)],
            'beads_field': 'Beads Events', 'sample_regex': r'(?i)Sample(?:\s+.*)?', 'replicate_cv': True, 'cv_threshold': None}


def empty_config():
    return {'columns': [], 'beads_field': 'Beads Events', 'sample_regex': r'(?i)Sample(?:\s+.*)?', 'replicate_cv': True, 'cv_threshold': None}


def tube_names_regex(value):
    """Literal prefixes, case insensitive; digits and terminal rpt are optional."""
    names = [s.strip() for s in re.split(r'[,，;；\n]+', value) if s.strip()]
    return r'(?i)(?:' + '|'.join(re.escape(s)+r'\s*\d*(?:\s+rpt\s*\d+)?' for s in names) + ')' if names else r'(?!)'


def tube_names_display(cfg):
    if 'abs_tube_names' in cfg:
        return cfg['abs_tube_names']
    pattern = cfg.get('sample_regex', '')
    if re.search(r'\(\?(?:[=!<]|P)',pattern): return ''
    literal_words=re.findall(r'(?<!\\)[A-Za-z][A-Za-z0-9_-]{1,}',pattern)
    if any(w.casefold() not in {'sample','rpt','icfmx','surffmx','fmo','isotype','saturation'} for w in literal_words):
        return ''
    # Read familiar old rules without exposing symbols or rewriting their behavior.
    names = []
    for name in re.findall(r'[A-Za-z][A-Za-z0-9_-]{1,}', pattern):
        if name.casefold() not in {'sample','rpt','icfmx','surffmx','fmo','isotype','saturation'}:
            continue
        if name.casefold() not in [n.casefold() for n in names]: names.append(name)
    return ', '.join(names)


def split_subject(row):
    out = dict(row)
    if any(k in row for k in ('采血阶段','给药前后','间隔时长')):
        return out
    try:
        parts = time_parts(row.get('Time Point',''))
        stage = [parts['anchors'][k] for k in ('W','D') if parts['anchors'].get(k)]
        if parts['eoi']: stage.append('EOI')
        out.update(采血阶段=' '.join(stage), 给药前后=parts['phase'], 间隔时长=parts['duration'])
    except ValueError:
        # Keep invalid imported source text visible for manual repair.
        out.update(采血阶段=txt(row.get('Time Point')), 给药前后='', 间隔时长='')
    return out


def join_subject(row):
    out = dict(row)
    if any(k in row for k in ('采血阶段','给药前后','间隔时长')):
        split = ' '.join(txt(row.get(k)) for k in ('采血阶段','给药前后','间隔时长'))
        old = txt(row.get('Time Point'))
        out['Time Point'] = merge_time(old, split) if old and split else old or split
    return out


def mode_label(column):
    return {'raw':'原始值','abs':'Beads回算','parent_abs':'跨Panel ABS×%'}[column['mode']]


def abs_source_label(column):
    if column['mode']=='parent_abs': return column.get('parent_panel_name','Panel')+' / '+column.get('parent_source','')
    return '本Panel Beads Events' if column['mode']=='abs' else ''


PRECISION_LABELS={'default':'默认','decimals':'小数位','significant':'有效数字'}


def precision_label(column):
    mode=column.get('precision_mode','default')
    return '默认' if mode=='default' else f"{column['precision_digits']}位{PRECISION_LABELS[mode]}"


def precision_options(mode_label,digits):
    mode=next(k for k,v in PRECISION_LABELS.items() if v==mode_label)
    if mode=='default':return {'precision_mode':'default'}
    try:n=int(digits)
    except (TypeError,ValueError):raise ValueError('精度位数请填写整数') from None
    if not (0 if mode=='decimals' else 1)<=n<=10:raise ValueError('小数位支持0–10位，有效数字支持1–10位')
    return {'precision_mode':mode,'precision_digits':n}


def split_summary_visits(ws):
    """Expand only the exported metadata; leave stored identities and formulas intact."""
    from copy import copy
    headers=[c.value for c in ws[1]]
    if 'Timepoint' not in headers:return
    timecol=headers.index('Timepoint')+1
    saved=[[copy(c) for c in row] for row in ws]
    ws.delete_rows(1,ws.max_row)
    for ri,row in enumerate(saved,1):
        ci=1
        for oldcol,src in enumerate(row,1):
            if oldcol==timecol:
                values=['采血阶段（W/D）','Pre / Post','时长（h/D）'] if ri==1 else [split_subject({'Time Point':src.value}).get(k,'') for k in ('采血阶段','给药前后','间隔时长')]
                for value in values:
                    dst=ws.cell(ri,ci,value);dst.data_type='s';dst._style=copy(src._style);ci+=1
            else:
                dst=ws.cell(ri,ci);dst.value=src.value;dst.data_type=src.data_type;dst._style=copy(src._style)
                if src.comment:dst.comment=copy(src.comment)
                ci+=1
    for c in ws[1]:
        if c.value=='Cycle':c.value='Cycle（给药周期）'
        elif c.value=='Visit':c.value='Visit（访视）'
    for i in range(1,ws.max_column+1):ws.column_dimensions[get_column_letter(i)].width=20
    ws.column_dimensions[get_column_letter(ws.max_column)].width=55
    ws.auto_filter.ref=ws.dimensions


def validate_validation_batches(cfg):
    names=set();used=set()
    for batch in cfg.get('batches',[]):
        name=txt(batch.get('name'));runs=batch.get('runs',[])
        if not name or name in names:raise ValueError('验证批次名称不能为空或重复')
        if not runs or len(runs)!=len(set(runs)):raise ValueError('请选择不重复的Run')
        if batch.get('baseline') not in runs:raise ValueError('Baseline必须属于本批次所选Run')
        if used.intersection(runs):raise ValueError('同一个Run不能同时属于两个验证批次')
        names.add(name);used.update(runs)


def validation_metrics(items):
    fields={}
    for x in items:
        p=x['payload']
        for key,label in p['labels'].items():
            value=p['raw'].get(key)
            try:
                if value is not None:metric_number(value,label,infer_kind(label));fields.setdefault(norm(label),label)
            except (ValueError,TypeError):pass
            if infer_kind(label) in ('events','pct100','pct1'):fields.setdefault(norm(label),label)
    return list(fields.values())


def add_validation_summary(wb,items,cv,metrics,threshold,settings):
    ws=wb.create_sheet('验证Summary',0)
    heads=['验证批次','验证材料','分析员','Cycle（给药周期）','Visit（访视）','采血阶段（W/D）','Pre / Post','时长（h/D）','Run','批次用途','Tube','管组','P/F','记录ID']+metrics
    literal_row(ws,heads)
    groups={}
    for item in items:
        m=item['payload']['meta'];key=(item['batch'],item['material'],m['Run number'],txt(m['Analyst']))
        groups.setdefault(key,[]).append(item)
    rows_by_key={(r[0],r[1],r[2],r[3],r[4],r[6]):r for r in cv}
    for (batch,material,run,analyst),rows in sorted(groups.items(),key=lambda kv:natural_sort_key(str(kv[0]))):
        for item in rows:
            m=item['payload']['meta'];split=split_subject({'Time Point':m.get('Timepoint','')})
            prefix=[batch or '未分批',material,analyst,m.get('Cycle',''),m.get('Visit','')]+[split.get(k,'') for k in ('采血阶段','给药前后','间隔时长')]+[run,item['run_role'],m['Tube'],item['tube_group'],item['rec']['pf'],item['rec']['id']]
            values=[item['payload']['raw'].get(norm(metric)) for metric in metrics]
            literal_row(ws,prefix+values)
        samplegroups=sorted({x['tube_group'] for x in rows if x['role']=='样本管'})
        for group in samplegroups:
            for label in ['Average','批内CV%']+(['与Baseline均值CV%'] if batch and rows[0]['run_role']!='Baseline' else []):
                line=[batch or '未分批',material,analyst,'','','','','',run,rows[0]['run_role'],label,group,'','']
                details=[]
                for metric in metrics:
                    kind='与Baseline均值CV' if label=='与Baseline均值CV%' else '批内CV'
                    result=rows_by_key.get((kind,material,group+' / 样本管',run,analyst,metric))
                    value=None;reason='缺少可计算数据'
                    if result:
                        reason=result[10]+'；记录ID：'+result[11]
                        if label=='Average':
                            values=json.loads(result[12])
                            if values and len(values)==result[7]:value=statistics.mean(values)
                        else:value=result[8]
                    line.append(value);details.append(reason)
                literal_row(ws,line);ri=ws.max_row
                for cell in ws[ri]:cell.fill=PatternFill('solid',fgColor='EDF4FC');cell.font=Font(name='Arial',size=10,bold=True)
                for i,reason in enumerate(details,15):
                    cell=ws.cell(ri,i);cell.number_format='0.00';cell.comment=Comment(reason,'流式数据汇总工具')
                    if label!='Average' and (cell.value is None or cell.value>threshold):
                        cell.fill=PatternFill('solid',fgColor='FEE4E2');cell.font=Font(name='Arial',size=10,color='B42318',bold=True)
    ws.freeze_panes='O2';ws.auto_filter.ref=ws.dimensions
    for i in range(1,ws.max_column+1):ws.column_dimensions[get_column_letter(i)].width=22
    ws.column_dimensions['K'].width=28
    note=wb['计算说明']
    literal_row(note,['验证批次',dumps(settings.get('batches',[]))])
    literal_row(note,['未分批','仍输出全部原始管与批内CV；不推测Baseline，不计算与Baseline均值CV。'])
    literal_row(note,['汇总取数','全部当前记录，所有原始数值指标；不同材料/样本管组分别统计。超限标红不拦截、不自动判P/F。'])


def sample_tube_tag(tube):
    """Only primary sample tubes; never Control/FMO/saturation or their rpt suffixes."""
    name=txt(tube)
    match=re.fullmatch(r'(?:sample\s*)?rpt\s*([1-9]\d*|0+[1-9]\d*)',name,re.I)
    if match:return 'rpt',int(match[1])
    match=re.fullmatch(r'sample\s*(\d+)?',name,re.I)
    if match and (match[1] is None or int(match[1])>0):return 'sample',int(match[1] or 1)
    return None


def formatted_report_value(value,column):
    if value is None:return ''
    if column['mode']=='raw' and column['kind']=='text':return txt(value)
    pct=column['mode']=='raw' and column['kind']=='pct1'
    with localcontext() as ctx:
        ctx.prec=400
        number=Decimal(str(value))*(100 if pct else 1)
        mode=column.get('precision_mode','default')
        n=column.get('precision_digits',1)
        if mode=='default':n=0 if column['mode']=='raw' and column['kind']=='events' else 1
        places=n if mode!='significant' else (n-1-number.copy_abs().adjusted() if number else n-1)
        rounded=number.quantize(Decimal(1).scaleb(-places),rounding=ROUND_HALF_UP)
        digits=max(0,places)
        if mode=='significant' and rounded:digits=max(0,n-1-rounded.copy_abs().adjusted())
        return f'{rounded:.{digits}f}'+('%' if pct else '')


def inline_sample_statistics(wb,cfg,prepared,exclude_fail):
    """Group displayed rows, then append sample-only statistics from full-precision inputs."""
    from copy import copy
    old=wb['Summary'];heads=[c.value for c in old[1]];idcol=heads.index('记录ID')+1
    by_id={r['id']:(r,p,vals,calc) for r,p,vals,calc,_ in prepared}
    groups={}
    for row in old.iter_rows(min_row=2):
        ident=row[idcol-1].value
        rec,p,vals,calc=by_id[ident];m=p['meta'];run=m.get('Run number','')
        key=(m.get('BAS ID',''),run_number(run) if run else 'missing:'+str(ident))
        groups.setdefault(key,[]).append((row,ident))
    ws=wb.create_sheet('Summary_new',0)
    for key,dim in old.column_dimensions.items():ws.column_dimensions[key]=copy(dim)
    def copy_row(row):
        n=ws.max_row+1 if ws.max_row>1 or ws.cell(1,1).value is not None else 1
        for col,src in enumerate(row,1):
            dst=ws.cell(n,col);dst.value=src.value;dst.data_type=src.data_type
            dst._style=copy(src._style)
            if src.comment:dst.comment=copy(src.comment)
        return n
    copy_row(old[1]);ws.row_dimensions[1].height=old.row_dimensions[1].height
    calcws=wb['计算明细'];start=21
    for j,h in enumerate(['统计记录ID','统计BAS ID','统计Run','统计Tube','统计汇报项','未舍入统计输入']):calcws.cell(1,start+j,h)
    calc_index={(calcws.cell(r,1).value,calcws.cell(r,2).value):r for r in range(2,calcws.max_row+1)}
    input_row=2
    def group_sort(key):return (key[0],0 if isinstance(key[1],int) else 1,key[1])
    for key,rows in sorted(groups.items(),key=lambda item:group_sort(item[0])):
        for row,ident in rows:copy_row(row)
        families={}
        for row,ident in rows:
            rec,p,vals,calc=by_id[ident];tag=sample_tube_tag(p['meta']['Tube'])
            if tag:families.setdefault(tag[0],[]).append((tag[1],ident))
        for family,members in sorted(families.items()):
            members.sort();ids=[ident for _,ident in members];repnums=[n for n,_ in members]
            first=by_id[ids[0]];meta=first[1]['meta'];missing_run=not meta.get('Run number')
            duplicate=len(repnums)!=len(set(repnums))
            avgrow=ws.max_row+1;cvrow=avgrow+1
            label_suffix=f' ({family})' if len(families)>1 else ''
            for ri,label,color in [(avgrow,'Average','#EDF4FC'),(cvrow,'CV%','#F3F6FA')]:
                for h in ['BAS ID','Subject ID','Cycle','Visit','Timepoint','Run number']:
                    ws.cell(ri,heads.index(h)+1,meta.get(h,''))
                ws.cell(ri,heads.index('Tube')+1,label+label_suffix)
                for ci in range(1,len(heads)+1):
                    ws.cell(ri,ci).fill=PatternFill('solid',fgColor=color[1:])
                    ws.cell(ri,ci).font=Font(name='Arial',size=10,bold=True)
                for h in ('BAS ID','Subject ID'):ws.cell(ri,heads.index(h)+1).number_format='@'
                ws.cell(ri,idcol+1,'统计来源记录ID：'+', '.join(map(str,ids))).data_type='s'
            rule='仅'+family+'样本管；记录ID：'+','.join(map(str,ids))+'。使用未舍入值；Control/FMO/saturation不参与。'
            if exclude_fail:rule+='本表按排除F后显示的样本管计算；独立复管CV页仍保留全部当前复管的核查口径。'
            for i,column in enumerate(cfg['columns']):
                if column['mode']=='raw' and column['kind']=='text':continue
                numbers=[];refs=[]
                for ident in ids:
                    rec,p,values,lines=by_id[ident];v=values[i]
                    if isinstance(v,tuple):v=lines[v[1]][6]
                    if not isinstance(v,(int,float)) or not math.isfinite(v):continue
                    numbers.append(v)
                    for j,x in enumerate([ident,p['meta']['BAS ID'],p['meta']['Run number'],p['meta']['Tube'],column['name'],v]):
                        cell=calcws.cell(input_row,start+j,x)
                        if isinstance(x,str):cell.data_type='s'
                    parent=calc_index.get((ident,column['name']))
                    if column['mode'] in ('abs','parent_abs') and parent:calcws.cell(input_row,26,f'=G{parent}')
                    calcws.cell(input_row,26).number_format='0.###############'
                    refs.append(f"'计算明细'!Z{input_row}");input_row+=1
                ac=ws.cell(avgrow,len(META)+i+1);cc=ws.cell(cvrow,len(META)+i+1)
                reason=''
                if missing_run:reason='缺少Run，无法确认复管批次'
                elif duplicate:reason='样本管复管编号重复，未计算'
                elif len(numbers)!=len(ids):reason='样本管存在缺值，未用部分数据计算'
                if reason:
                    ac.comment=Comment(rule+'\n'+reason,'流式数据汇总工具');cc.comment=Comment(rule+'\n'+reason,'流式数据汇总工具');continue
                mean=statistics.mean(numbers);formula='AVERAGE('+','.join(refs)+')'
                # Statistics always read full precision inputs, never rounded Summary cells.
                precision_col=dict(column)
                if column['kind']=='events' and column['mode']=='raw' and column.get('precision_mode','default')=='default':precision_col['kind']='number'
                rendered=formatted_report_value(mean,precision_col)
                suffix='%' if rendered.endswith('%') else ''
                places=len(rendered.rstrip('%').split('.')[1]) if '.' in rendered else 0
                ac.number_format='0'+('.'+'0'*places if places else '')+suffix
                pm=column.get('precision_mode','default')
                if pm!='default':
                    pct=column['kind']=='pct1' and column['mode']=='raw';scaled=Decimal(str(mean))*(100 if pct else 1);digits=column['precision_digits']
                    rounding=digits if pm=='decimals' else (digits-1-scaled.copy_abs().adjusted() if scaled else digits-1)
                    formula=f'ROUND(({formula})*100,{rounding})/100' if pct else f'ROUND({formula},{rounding})'
                ac.value='='+formula
                wb._summary_results[ac.value]=float(rendered.rstrip('%'))/(100 if suffix else 1) if pm!='default' else mean
                ac.comment=Comment(rule+f'\n未舍入均值：{mean}','流式数据汇总工具')
                cc.number_format='0'+('.'+'0'*cfg.get('cv_digits',2) if cfg.get('cv_digits',2) else '')
                if len(numbers)<2:reason='不足2管，CV留空'
                elif mean<=0:reason='均值≤0，CV未定义，留空'
                else:
                    args=','.join(refs);cc.value=f'=STDEV({args})/AVERAGE({args})*100'
                    wb._summary_results[cc.value]=statistics.stdev(numbers)/mean*100
                cc.comment=Comment(rule+'\n'+(reason or '样本SD(n−1) / 未舍入均值 × 100%；显示精度按独立CV小数位设置。'),'流式数据汇总工具')
    wb.remove(old);ws.title='Summary';ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    for ci in range(21,27):calcws.column_dimensions[get_column_letter(ci)].width=23
    literal_row(wb['_V2说明'],['样本统计','各BAS/Run全部管后追加Average/CV%；仅rpt/Sample样本管；两套命名组分别计算。统计输入见计算明细U:Z；单管或均值≤0的CV留空。'])


def capture_summary_results(wb,cfg,prepared):
    wb._summary_results={}
    ws=wb['Summary'];heads=[c.value for c in ws[1]];idcol=heads.index('记录ID')+1
    by_id={r['id']:(vals,calc) for r,p,vals,calc,ok in prepared}
    for row in ws.iter_rows(min_row=2):
        vals,calc=by_id[row[idcol-1].value]
        for i,col in enumerate(cfg['columns']):
            cell=row[len(META)+i]
            if cell.data_type!='f':continue
            v=vals[i];v=calc[v[1]][6] if isinstance(v,tuple) else v
            if col.get('precision_mode','default')!='default':
                v=float(formatted_report_value(v,col).rstrip('%'))/(100 if col['mode']=='raw' and col['kind']=='pct1' else 1)
            wb._summary_results[cell.value]=v


def materialize_summary(wb):
    ws=wb['Summary'];formulas=wb.copy_worksheet(ws);formulas.title='Summary公式'
    wb.move_sheet(formulas,offset=1-wb.index(formulas))
    for row in ws:
        for cell in row:
            if cell.data_type=='f':
                if cell.value not in wb._summary_results:raise ValueError('Summary公式缺少已计算结果：'+cell.coordinate)
                cell.value=wb._summary_results[cell.value]


def apply_summary_precision(wb,cfg,prepared):
    # Round only final report cells. All calculations and CV use the original data.
    ws=wb['Summary'];heads=[c.value for c in ws[1]];id_col=heads.index('记录ID')+1
    values_by_id={r['id']:[calc[v[1]][6] if isinstance(v,tuple) else v for v in vals] for r,p,vals,calc,ok in prepared}
    for ri in range(2,ws.max_row+1):
        values=values_by_id[ws.cell(ri,id_col).value]
        for i,col in enumerate(cfg['columns']):
            mode=col.get('precision_mode','default');v=values[i]
            if mode=='default' or v is None or (col['mode']=='raw' and col['kind']=='text'):continue
            cell=ws.cell(ri,len(META)+i+1)
            if cell.value is None:continue
            is_pct=col['mode']=='raw' and col['kind']=='pct1'
            n=col['precision_digits']
            with localcontext() as ctx:
                ctx.prec=400
                number=Decimal(str(v))*(100 if is_pct else 1)
                places=n if mode=='decimals' else (n-1-number.copy_abs().adjusted() if number else n-1)
                rounded=number.quantize(Decimal(1).scaleb(-places),rounding=ROUND_HALF_UP)
                # Carries such as 9.999 → 10.0 must still show exactly three significant digits.
                display_places=n if mode=='decimals' else (max(0,n-1-rounded.copy_abs().adjusted()) if rounded else max(0,n-1))
                if cell.data_type=='f':
                    expr=cell.value[1:]
                    cell.value=f'=ROUND(({expr})*100,{places})/100' if is_pct else f'=ROUND({expr},{places})'
                else:cell.value=float(rounded/(100 if is_pct else 1))
            cell.number_format='0'+('.'+'0'*display_places if display_places else '')+('%' if is_pct else '')
            cell.comment=Comment(f'汇报规则：{precision_label(col)}；四舍五入。\n未舍入原值：{v}\n仅最终Summary舍入；ABS中间值和CV使用未舍入值。','流式数据汇总工具')
    note=wb['_V2说明']
    literal_row(note,['Summary精度','只对最终汇报值四舍五入；计算明细、ABS中间值及CV取数不舍入。有效数字大数的尾零不额外使用科学计数法。'])
    for col in cfg['columns']:literal_row(note,['汇报精度：'+col['name'],precision_label(col)])


def validate_config(cfg):
    if type(cfg.get('cv_digits',2)) is not int or not 0<=cfg.get('cv_digits',2)<=10:raise ValueError('CV小数位支持0–10位')
    validate_import_options(cfg.get('import_options', {}))
    if not cfg.get('columns'):
        raise ValueError('请至少添加一个汇报项')
    if cfg.get('cv_threshold') is not None and numeric(cfg['cv_threshold'], 'CV阈值') < 0:
        raise ValueError('CV阈值不能为负数')
    low, limit = cfg.get('cv_low_pct_boundary'), cfg.get('cv_low_pct_threshold')
    if (low is None) != (limit is None):
        raise ValueError('低表达CV规则的两个空格请一起填写，或一起留空')
    if low is not None:
        numeric(low, '低表达%parent界限', 'pct100')
        if numeric(limit, '低表达CV限值') < 0: raise ValueError('低表达CV限值不能为负数')
    try:
        re.compile(cfg['sample_regex'])
    except re.error as exc:
        raise ValueError(f'计数管名规则错误：{exc}') from exc
    names = list(map(norm, META + ['样品类型','验证样品名称','原始Specimen']))
    for c in cfg['columns']:
        pm=c.get('precision_mode','default')
        if pm not in PRECISION_LABELS:raise ValueError('不支持的汇报精度类型')
        if pm!='default':
            if type(c.get('precision_digits')) is not int:raise ValueError('精度位数必须为整数')
            precision_options(PRECISION_LABELS[pm],c['precision_digits'])
        if not txt(c.get('source')) or not txt(c.get('name')):
            raise ValueError('来源表头和输出名称不能为空')
        if norm(c['name']) in names:
            raise ValueError(f'输出列重名：{c["name"]}')
        names.append(norm(c['name']))
        if c.get('mode') not in ('raw', 'abs', 'parent_abs') or c.get('kind') not in KIND_LABEL:
            raise ValueError('不支持的汇报项类型')
        if c['mode'] == 'abs' and '%' in c['source']:
            raise ValueError('ABS来源不能是百分比字段，请选择细胞Events')
        if c['mode']=='parent_abs':
            if c['kind'] not in ('pct100','pct1'): raise ValueError('跨Panel ABS来源必须选择百分比0–100或比例0–1')
            if type(c.get('parent_panel')) is not int or c['parent_panel']<=0 or not txt(c.get('parent_source')):
                raise ValueError('请选择父群ABS的来源Panel和汇报项')
            if c.get('parent_replicates','unique') not in ('unique','mean'): raise ValueError('父群来源复管规则不支持')
            if c.get('parent_confirmed') is not True: raise ValueError('请核对百分比的分母确实是选定父群，再勾选确认')
            if not isinstance(c.get('parent_choices',{}),dict): raise ValueError('父群取管配置格式错误')
            for key,ids in c.get('parent_choices',{}).items():
                if not isinstance(key,str) or not isinstance(ids,list) or not ids or any(type(i) is not int or i<=0 for i in ids):
                    raise ValueError('父群取管须为非空的记录ID列表')
    if any(c['mode'] == 'abs' for c in cfg['columns']) and not txt(cfg.get('beads_field')):
        raise ValueError('ABS需要指定Beads Events来源表头')


def iter_tables(path, anchors, required):
    """Yield header labels + complete raw row, rejecting formula source cells."""
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if '_V2说明' in wb.sheetnames or '_DTS说明' in wb.sheetnames or {'汇总', '来源与检查'}.issubset(wb.sheetnames):
            return
        for ws in wb:
            rows = ws.iter_rows(values_only=True)
            for line, values in enumerate(rows, 1):
                if line > 30:
                    break
                labels = [txt(v) for v in values]
                keys = [norm(v) for v in values]
                if not all(norm(a) in keys for a in anchors):
                    continue
                nonempty = [k for k in keys if k]
                if len(nonempty) != len(set(nonempty)):
                    raise ValueError(f'{path.name}/{ws.title} 表头重名（忽略空格及大小写后）')
                missing = [h for h in required if norm(h) not in keys]
                if missing:
                    raise ValueError(f'{path.name}/{ws.title} 缺少：{missing}')
                for n, values in enumerate(rows, line + 1):
                    if not any(v is not None and v != '' for v in values):
                        continue
                    raw = {k: json_value(v) for k, v in zip(keys, values) if k}
                    if any(isinstance(v, str) and v.startswith('=') for v in raw.values()):
                        raise ValueError(f'{path.name}/{ws.title}/行{n} 含公式，请提供值导出')
                    yield ws.title, n, dict(zip(keys, labels)), raw
                break
    finally:
        wb.close()


META_ALIASES = {
    'Experiment': ['Experiment', 'Experiment Name', 'EXPERIMENT NAME'],
    'Specimen': ['Specimen', '$SRC', 'Sample Name', '样品名'],
    'Tube': ['Tube', 'TUBE NAME'],
    'Date': ['Date', 'FCSE_DATE', '$DATE'],
    'Time': ['Time', '$BTIM'],
    'Analyst': ['Analyst', 'EXPORT USER NAME', '$OP'],
    'BAS ID': ['BAS ID','Sample ID','样品编号'],
    'Subject ID': ['Subject ID','Patient ID','受试者编号'],
    'Run number': ['Run number','Run'],
    'Cycle': ['Cycle'],
    'Visit': ['Visit','Visit-Cycle'],
    'Timepoint': ['Timepoint','Time Point'],
}


def header_token(label):
    m = re.fullmatch(r'Keyword:\s*(.*?)\s+for\s+plot\s*#\d+,\s*overlay\s*#\d+', txt(label), re.I | re.S)
    return norm(m[1] if m else label)


def validate_import_options(options):
    if not isinstance(options, dict) or not isinstance(options.get('mapping', {}), dict):
        raise ValueError('导入字段配置格式不正确')
    for field, source in options.get('mapping', {}).items():
        if field not in INPUT_FIELDS or not isinstance(source, str):
            raise ValueError('不支持的导入字段映射')
    if not isinstance(options.get('tube_default',''),str):
        raise ValueError('缺失Tube时的管名必须为文本')
    runs = options.get('validation_runs', [])
    if not isinstance(runs, list) or any(type(n) is not int or not 1 <= n <= 100000 for n in runs):
        raise ValueError('验证Run配置须为有效Run编号列表')
    panel = options.get('source_panel', '')
    if panel and not re.fullmatch(r'(?i)Panel\s*\d+', panel):
        raise ValueError('来源Panel请填Panel1、Panel2等，或留空')


def unique_headers(values):
    """Never overwrite duplicate metrics; retain their Excel column addresses."""
    labels = [re.sub(r'\s+', ' ', txt(v)) for v in values]
    counts = {norm(v): sum(norm(w) == norm(v) for w in labels) for v in labels if v}
    result = [f'{v} [列{get_column_letter(i+1)}]' if v and counts[norm(v)] > 1 else v for i,v in enumerate(labels)]
    keys = [norm(v) for v in result if v]
    if len(set(keys)) != len(keys):
        raise ValueError('带列号的表头仍重名，请先为来源列设置唯一名称')
    return result


def resolve_metadata(labels, options=None):
    mapping = (options or {}).get('mapping', {})
    result = {}
    for field, aliases in META_ALIASES.items():
        explicit = mapping.get(field)
        if explicit == '':
            result[field] = None; continue
        candidates = [label for label in labels if label and (
            norm(label) == norm(explicit) if explicit is not None else header_token(label) in set(map(norm, aliases)))]
        if len(candidates) > 1:
            raise ValueError(f'{field}对应多个表头，请在“导入字段配置”明确选择：{candidates}')
        result[field] = norm(candidates[0]) if candidates else None
    return result


def iter_raw_tables(path, options=None):
    options = options or {}; validate_import_options(options)
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if '_V2说明' in wb.sheetnames or '_DTS说明' in wb.sheetnames or {'汇总','来源与检查'}.issubset(wb.sheetnames): return
        for ws in wb:
            rows = ws.iter_rows(values_only=True)
            for line, cells in enumerate(rows, 1):
                if line > 30: break
                labels = unique_headers(cells)
                # Only resolve complete metadata rows, so title lines are ignored.
                tokens = {header_token(v) for v in labels if v}
                mapping = options.get('mapping', {})
                if not any((norm(mapping[f]) in set(map(norm, labels))) if mapping.get(f) else
                           (mapping.get(f) != '' and bool(tokens.intersection(map(norm,META_ALIASES[f])))) for f in ['BAS ID','Specimen']):
                    continue
                fields = resolve_metadata(labels, options)
                keys = [norm(v) for v in labels]
                for n, values in enumerate(rows, line+1):
                    if not any(v is not None and v != '' for v in values): continue
                    if any(isinstance(v,str) and v.startswith('=') for v in values):
                        raise ValueError(f'{path.name}/{ws.title}/行{n} 含公式，请提供值导出')
                    if any(v is not None and v != '' and (i >= len(keys) or not keys[i]) for i,v in enumerate(values)):
                        raise ValueError(f'{path.name}/{ws.title}/行{n} 存在无表头的数据列，请补齐表头')
                    raw = {k: json_value(v) for k,v in zip(keys,values) if k}
                    yield ws.title,n,{k:v for k,v in zip(keys,labels) if k},raw,fields
                break
    finally:
        wb.close()


def config_from_excel(path, options=None):
    iterator = iter_raw_tables(Path(path), options)
    try: first = next(iterator, None)
    finally: iterator.close()
    if first is None: raise ValueError('未找到数据表；请配置样品编号（BAS ID）或样品名（Specimen）来源列')
    _,_,labels,raw,fields = first
    cfg = empty_config(); cfg['import_options'] = options or {}
    for key,label in labels.items():
        if key in fields.values(): continue
        kind = infer_kind(label)
        if kind == 'number' and isinstance(raw.get(key),str):
            try: metric_number(raw[key],label)
            except ValueError: kind = 'text'
        cfg['columns'].append({'source':label,'name':label,'mode':'raw','kind':kind})
    validate_config(cfg)
    return cfg, labels, fields


def scan_paths(root, excluded=(), options=None):
    excluded = {Path(p).resolve() for p in excluded}
    paths = []
    for p in Path(root).rglob('*'):
        if not p.is_file() or p.suffix.lower() != '.xlsx' or p.name.startswith('~$'):
            continue
        if p.resolve() in excluded or any(x.endswith('_backups') for x in p.parts):
            continue
        it = iter_raw_tables(p, options)
        try:
            if next(it, None) is not None:
                paths.append(p.resolve())
        except Exception:
            # Keep unreadable candidates visible; importing produces a per-file
            # error instead of aborting the entire directory scan.
            paths.append(p.resolve())
        finally:
            it.close()
    return sorted(paths)


def filename_panels(name):
    """Use the filename only, never a parent folder or a substring of Panel12."""
    return sorted({int(n) for n in re.findall(r'(?i)(?<![a-z0-9])panel[ _-]*0*(\d+)(?![a-z0-9])', Path(name).stem) if int(n)>0})


def numbered_panel(name):
    match = re.fullmatch(r'(?i)panel[ _-]*0*(\d+)', txt(name))
    return int(match[1]) if match and int(match[1])>0 else None


def route_needs_attention(row):
    if row.get('skip'):
        return row.get('detail') not in ('已登记的汇总文件或受试者资料','程序导出的汇总文件')
    return bool(row.get('scan_error') or row.get('attention') or not (row.get('panel_id') or row.get('create_name')))


def scan_project_folder_from_paths(paths):
    for path in paths:
        nums=filename_panels(Path(path).name)
        yield {'file':str(Path(path).resolve()),'number':nums[0] if len(nums)==1 else None,
               'run':'','skip':False,'attention':True,'detail':'失败文件重试；重新核对分配'}


def scan_project_folder(root, excluded=()):
    """Inventory all Excel candidates, including unrecognized/corrupt inputs."""
    excluded = {str(Path(p).resolve()) for p in excluded}
    result = []
    for path in sorted(Path(root).rglob('*')):
        if not path.is_file() or path.suffix.lower() not in ('.xlsx','.xls','.xlsm'): continue
        if path.name.startswith('~$') or any(p.endswith('_backups') for p in path.parts): continue
        path = path.resolve(); numbers = filename_panels(path.name)
        runs = sorted({int(n) for n in re.findall(r'(?i)(?<![a-z0-9])run[ _-]*(\d+)(?![a-z0-9])',path.stem)})
        row = {'file':str(path),'number':numbers[0] if len(numbers)==1 else None,
               'run': 'Run'+str(runs[0]) if len(runs)==1 else '', 'skip':False,
               'detail': '按文件名识别' if len(numbers)==1 else ('多个Panel标记，需分配' if numbers else '文件名无Panel标记，需分配')}
        if str(path) in excluded:
            row.update(skip=True,detail='已登记的汇总文件或受试者资料')
        elif path.suffix.lower() != '.xlsx':
            row.update(skip=True,detail='仅支持.xlsx；请先另存为.xlsx后重试')
        else:
            book = None
            try:
                book = load_workbook(path,read_only=True)
                if any(n in book.sheetnames for n in ('_V2说明','_DTS说明','_说明')):
                    row.update(skip=True,detail='程序导出的汇总文件')
            except Exception as exc:
                row['scan_error']=str(exc)
                row['detail'] += '；文件读取异常，导入时会列出失败：'+str(exc)
            finally:
                if book: book.close()
        result.append(row)
    return result


NUMERIC_TIME_PATTERN = r'(?:(?:pre|post)\s*)?[+-]?\d+(?:\.\d+)?\s*(?:hours?|hrs?|h|minutes?|mins?|m|days?|d|seconds?|secs?|s)'
TIME_PATTERN = r'(?:EOI(?:\s*[+-]\s*\d+(?:\.\d+)?\s*(?:hours?|hrs?|h|minutes?|mins?|m|days?|d|seconds?|secs?|s))?|' + NUMERIC_TIME_PATTERN + r'|pre|post)'
SPECIMEN_PATTERN = r'(?:(.*?)\s+)?(Run\s*\d+)\s+(\d{11}|\d{6}\.\d{5})(?:\s+(?:beads\s*[:=]?\s*)?(\d+(?:\.0+)?)(?:\s*beads)?)?'


def visit_text(value):
    value = txt(value).translate(str.maketrans('０１２３４５６７８９＋－', '0123456789+-'))
    return re.sub(r'\s+', ' ', value).strip()


def visit_tokens(value, letter):
    # V2D0 and W2Pre4h are common compact labels; ignore unrelated words.
    matches = re.findall(r'(?<![A-Za-z])'+letter+r'\s*(\d+)(?!\d)', visit_text(value), re.I)
    numbers = sorted({int(n) for n in matches})
    if len(numbers)>1:
        raise ValueError(f'{letter}编号存在多个不同值：'+', '.join(letter+str(n) for n in numbers))
    return letter+str(numbers[0]) if numbers else ''


def time_parts(value):
    text = visit_text(value)
    anchors = {letter:visit_tokens(text,letter) for letter in ('W','D')}
    pre = bool(re.search(r'(?<![A-Za-z])pre(?=$|[^A-Za-z]|dose)|给药前|用药前|投药前|药前',text,re.I))
    post = bool(re.search(r'(?<![A-Za-z])post(?=$|[^A-Za-z]|dose)|给药后|用药后|投药后|药后',text,re.I))
    if pre and post: raise ValueError('采血时间同时包含Pre和Post，请核对')
    phase = 'Pre' if pre else ('Post' if post else '')
    unit_pattern = r'(hours?|hrs?|h|小时|minutes?|mins?|min|分钟|m|days?|d|天|seconds?|secs?|s|秒)'
    durations = []
    for match in re.finditer(r'(?<![\d.])([+-]?\d+(?:\.\d+)?)\s*'+unit_pattern+r'(?![A-Za-z\d])',text,re.I):
        # Exclude visit/stage/cycle indices, e.g. D0 Post, V2D0 or W1D1.
        before = text[:match.start()]
        if re.search(r'[CVWD]\s*$',before,re.I): continue
        number,unit = match.groups(); n=float(number)
        if not math.isfinite(n): raise ValueError('采血时间不是有限数值')
        unit=unit.lower()
        unit='h' if unit.startswith('h') or unit=='小时' else ('min' if unit.startswith('m') or unit=='分钟' else ('D' if unit.startswith('d') or unit=='天' else 's'))
        direction = 'Pre' if n<0 else ('Post' if number.startswith('+') else '')
        if phase=='Post' and n<0: raise ValueError('Post后使用负时间，前后关系冲突，请核对')
        attached=bool(re.search(r'(?:pre|post)(?:[\s_-]*dose)?\s*[:=]?\s*$|(?:给药|用药|投药|药)[前后]\s*[:=]?\s*$|EOI\s*$',before,re.I))
        durations.append((abs(n),unit,direction,attached))
    # EOI remains an explicit reference point; only its surrounding prose is ignored.
    eoi=bool(re.search(r'(?<![A-Za-z])EOI(?![A-Za-z])',text,re.I))
    if durations:
        # A timing window or note after the primary Pre/Post duration is commentary.
        if any(d[3] for d in durations): durations=[d for d in durations if d[3]]
        durations=[d[:3] for d in durations]
        unique={n*{'h':3600,'min':60,'D':86400,'s':1}[u] for n,u,d in durations}
        if len(unique)>1: raise ValueError('采血时间包含多个不一致时长，请核对')
        n,unit,direction=durations[0]
        directions={d for _,_,d in durations if d}
        if not phase and len(directions)>1: raise ValueError('采血时间前后关系冲突，请核对')
        direction=next(iter(directions),'')
        phase=phase or direction or 'Post'
        duration=txt(n)+unit
        seconds=n*{'h':3600,'min':60,'D':86400,'s':1}[unit]
    else:
        duration=''; seconds=None
    return {'anchors':anchors,'phase':phase,'duration':duration,'seconds':seconds,'eoi':eoi}


def time_from_parts(parts):
    values=[parts['anchors'][letter] for letter in ('W','D') if parts['anchors'].get(letter)]
    if parts.get('eoi'): values.append('EOI')
    if parts.get('phase'): values.append(parts['phase'])
    if parts.get('duration'): values.append(parts['duration'])
    return ' '.join(values)


def canonical_time(value):
    return time_from_parts(time_parts(value))


def normalized_visit(cycle='',visit='',timepoint=''):
    text=' '.join(visit_text(v) for v in (cycle,visit,timepoint))
    return visit_tokens(text,'C'),visit_tokens(text,'V'),canonical_time(text)


def cycle_visit(cycle,visit):
    c,v,_=normalized_visit(cycle,visit)
    return c,v


def visit_fields(prefix,reference=None):
    return normalized_visit('',prefix,'')


def timepoint_key(value):
    p=time_parts(value)
    return (p['anchors']['W'],p['anchors']['D'],p['phase'],p['seconds'],p['eoi'])


def compatible_time(left,right):
    a,b=time_parts(left),time_parts(right)
    for field in ('W','D'):
        if a['anchors'][field] and b['anchors'][field] and a['anchors'][field]!=b['anchors'][field]: return False
    if a['phase'] and b['phase'] and a['phase']!=b['phase']: return False
    if a['seconds'] is not None and b['seconds'] is not None and a['seconds']!=b['seconds']: return False
    # EOI vs unspecified reference is incomplete information, not an automatic conflict.
    return True


def merge_time(left,right):
    if not compatible_time(left,right): raise ValueError(f'Timepoint不一致：{left!r} / {right!r}')
    a,b=time_parts(left),time_parts(right)
    merged={'anchors':{key:a['anchors'][key] or b['anchors'][key] for key in ('W','D')},
            'phase':a['phase'] or b['phase'],'duration':a['duration'] or b['duration'],'eoi':a['eoi'] or b['eoi']}
    return time_from_parts(merged)


def normalized_meta(meta):
    result=dict(meta)
    result['Cycle'],result['Visit'],result['Timepoint']=normalized_visit(meta.get('Cycle',''),meta.get('Visit',''),meta.get('Timepoint',''))
    if 'Date' in result: result['Date']=clean_date(result['Date'])
    return result


def merge_visit_meta(left,right,context=''):
    a,b=normalized_meta(left),normalized_meta(right)
    check_optional_fields(a,b,context)
    out=dict(a)
    for field in ('Subject ID','Cycle','Visit'):
        if not txt(out.get(field)): out[field]=b.get(field,'')
    out['Timepoint']=merge_time(a['Timepoint'],b['Timepoint'])
    return out


def check_optional_fields(left,right,context):
    a,b=normalized_meta(left),normalized_meta(right)
    for field in ('Subject ID','Cycle','Visit'):
        x,y=txt(a.get(field)),txt(b.get(field))
        if x and y and norm(x)!=norm(y):
            label='受试者Subject ID' if field=='Subject ID' else field
            raise ValueError(f'{context} 的 {label} 不一致：{x!r} / {y!r}')
    if not compatible_time(a['Timepoint'],b['Timepoint']):
        raise ValueError(f'{context} 的 Timepoint 不一致：{a["Timepoint"]!r} / {b["Timepoint"]!r}')


def normalized_reference(row):
    row = join_subject(row)
    out = {h: json_value(row.get(h, '')) for h in SUBJECT_FIELDS}
    out['Cycle'], out['Visit-Cycle'], out['Time Point'] = normalized_visit(out['Cycle'], out['Visit-Cycle'], out['Time Point'])
    out['Collection Date']=clean_date(out['Collection Date'])
    return out


def subject_candidate(row):
    """Validate one row without discarding the source cells."""
    errors={}
    for field,value in row.items():
        if isinstance(value,str) and value.startswith('='): errors[field]='含公式，请先转为值'
    try: item=normalized_reference(row)
    except ValueError as exc:
        item={h:json_value(row.get(h,'')) for h in SUBJECT_FIELDS}; errors['Visit-Cycle']=str(exc)
    try: item['样品编号']=bas_id(item.get('样品编号'))
    except ValueError as exc: errors['样品编号']=str(exc)
    if not txt(item.get('Patient ID')): errors['Patient ID']='Patient ID不能为空'
    for field in ('Collection Date','Collection Time'):
        if txt(item.get(field)) and field not in errors:
            try: collection_value(item[field],field=='Collection Time')
            except ValueError as exc: errors[field]=str(exc)
    return item,errors


def iter_subject_candidates(path):
    aliases={'样品编号':['样品编号','BAS ID','Sample ID'], 'Patient ID':['Patient ID','Subject ID','受试者编号'],
             'Barcode':['Barcode'], 'Cycle':['Cycle','Cycle（给药周期）'], 'Visit-Cycle':['Visit','Visit-Cycle','Visit（访视）'],
             'Time Point':['Timepoint','Time Point'], 'Collection Date':['Collection Date'], 'Collection Time':['Collection Time']}
    aliases.update({'采血阶段':['采血阶段','采血阶段（W/D）'], '给药前后':['给药前后','Pre / Post'], '间隔时长':['间隔时长','时长（h/D）']})
    wb=load_workbook(path,read_only=True,data_only=False); found=False
    try:
        for ws in wb:
            iterator=ws.iter_rows(values_only=True)
            for line,values in enumerate(iterator,1):
                if line>30: break
                keys=[norm(v) for v in values]
                if not any(norm(a) in keys for a in aliases['样品编号']): continue
                found=True; mapping={}; header_errors={}
                for field,names in aliases.items():
                    matches=[i for i,k in enumerate(keys) if k in set(map(norm,names))]
                    if len(matches)>1: header_errors[field]='对应多个表头，请保留唯一来源列'
                    elif matches: mapping[field]=matches[0]
                    elif field in ('样品编号','Patient ID'): header_errors[field]='表头缺少'+field
                if header_errors:
                    yield {'sheet':ws.title,'line':line,'row':{h:'' for h in SUBJECT_FIELDS},'errors':header_errors,'header_error':True}
                    break
                for n,cells in enumerate(iterator,line+1):
                    if not any(v is not None and v!='' for v in cells): continue
                    row={h:json_value(cells[mapping[h]]) if h in mapping and mapping[h]<len(cells) else '' for h in SUBJECT_FIELDS}
                    for h in ('采血阶段','给药前后','间隔时长'):
                        if h in mapping: row[h]=json_value(cells[mapping[h]]) if mapping[h]<len(cells) else ''
                    yield {'sheet':ws.title,'line':n,'row':row,'errors':{}}
                break
        if not found:
            yield {'sheet':'','line':0,'row':{h:'' for h in SUBJECT_FIELDS},'errors':{'样品编号':'未找到样品编号表头；请使用样品编号/BAS ID/Sample ID'},'header_error':True}
    finally: wb.close()


def runs_in_name(value):
    return sorted({int(n) for n in re.findall(r'(?i)(?<![a-z0-9])Run[ _-]*(\d+)(?![a-z0-9])', txt(value))})


def confirmed_validation_run(specimen, explicit_run, options):
    registered=set(options.get('validation_runs', []))
    if not registered: return None
    candidates=set(runs_in_name(specimen)) | set(options.get('_file_runs', []))
    if txt(explicit_run): candidates.add(run_number(explicit_run))
    if not candidates.intersection(registered): return None
    if len(candidates)!=1:
        raise ValueError('已标记验证Run，但文件名、Specimen或Run列的Run不一致/不唯一，请核对')
    return run_number(str(next(iter(candidates))))


def parse_raw(raw, labels, references=None, fields=None, options=None):
    get = lambda k: raw.get(fields.get(k) if fields is not None else norm(k))
    specimen = txt(get('Specimen')); prefix = ''; run = ''; bas = ''; validation_name = ''
    confirmed_run = confirmed_validation_run(specimen, get('Run number'), options or {})
    s = re.fullmatch(SPECIMEN_PATTERN, specimen, re.I)
    if confirmed_run is not None and (s or txt(get('BAS ID')) or re.search(r'(?<!\d)\d{6}\.?\d{5}(?!\d)',specimen)):
        confirmed_run=None
    if confirmed_run is not None:
        if not specimen: raise ValueError('验证Run仍需非空Specimen名称，不能仅凭Run区分不同样品')
        validation_name = re.sub(r'(?i)(?<![a-z0-9])Run[ _-]*\d+(?![a-z0-9])', ' ', specimen)
        validation_name = re.sub(r'\s+', ' ', validation_name).strip() or specimen
        run = 'Run'+str(confirmed_run)
    elif s:
        prefix,run,bas,_ = s.groups(); bas = bas_id(bas)
    elif specimen:
        simple = re.fullmatch(r'(?:(.*?)\s+)?(\d{11}|\d{6}\.\d{5})(?:\s+\d+(?:\.0+)?)?', specimen)
        validation = re.fullmatch(r'(LY\d+)\s+(Run\s*\d+)', specimen, re.I)
        if simple:
            prefix,bas = simple.groups();bas = bas_id(bas)
        elif validation:
            validation_name,run = validation.groups()
        else:
            if txt(get('BAS ID')):
                prefix=specimen
                found_runs=runs_in_name(specimen)
                if len(found_runs)>1: raise ValueError('样品名包含多个Run，请核对')
                if found_runs: run='Run'+str(found_runs[0])
            elif re.search(r'\d{6}\.\d+|(?<!\d)\d{11}(?!\d)',specimen):
                raise ValueError('样品名含样品编号但格式不能解析，请核对完整BAS ID或映射BAS ID列')
            else:
                validation_name=re.sub(r'(?i)(?<![a-z0-9])Run[ _-]*\d+(?![a-z0-9])',' ',specimen)
                validation_name=re.sub(r'\s+',' ',validation_name).strip() or specimen
                found_runs=runs_in_name(specimen) or (options or {}).get('_file_runs',[])
                if len(found_runs)>1: raise ValueError('验证样品包含多个Run，请核对')
                if found_runs: run='Run'+str(found_runs[0])
    if confirmed_run is None and get('BAS ID') is not None and txt(get('BAS ID')):
        if validation_name: raise ValueError('验证样品同时带正式BAS ID，身份有歧义，请核对来源')
        explicit_bas = bas_id(get('BAS ID'))
        if bas and bas != explicit_bas: raise ValueError('BAS ID列与样品名中的样品编号不一致')
        bas = explicit_bas
    if not bas and not validation_name: raise ValueError('无法识别完整样品编号；需6位项目号+5位样品流水号')
    if txt(get('Run number')):
        explicit_run = 'Run'+str(run_number(get('Run number')))
        if run and run_number(run) != run_number(explicit_run): raise ValueError('Run列与样品名中的Run不一致')
        run = explicit_run
    if not run and bas and len((options or {}).get('_file_runs',[]))==1:
        run='Run'+str(options['_file_runs'][0])
    experiment = txt(get('Experiment'));source_panel = '';subject = ''
    panel_suffix = re.search(r'\s+(Panel\s*\d+)$',experiment,re.I)
    if panel_suffix:
        source_panel = panel_suffix[1];experiment = experiment[:panel_suffix.start()].strip()
    e = re.fullmatch(r'(.+?)\s+(\d+)',experiment)
    project,subject = e.groups() if e else (experiment,'')
    if not validation_name and txt(get('Subject ID')):
        if subject and subject != txt(get('Subject ID')): raise ValueError('受试者Subject ID列与Experiment不一致')
        subject = txt(get('Subject ID'))
    reference = references.get(bas) if references else None
    if validation_name: subject = ''
    subject_from_reference = bool(reference is not None and not subject)
    if subject_from_reference: subject = txt(reference.get('Patient ID'))
    cycle, visit, tp = visit_fields(prefix, reference)
    explicit_cycle,explicit_visit,explicit_tp = normalized_visit(get('Cycle'),get('Visit'),get('Timepoint'))
    explicit_meta = {'Cycle':explicit_cycle,'Visit':explicit_visit,'Timepoint':explicit_tp}
    meta = {'Experiment': project or bas[:6], 'Subject ID': subject,
            'Cycle':cycle,'Visit':visit,'Timepoint':tp,
            'Run number': 'Run' + str(run_number(run)) if run else '', 'BAS ID':bas}
    metadata_alternatives=[]
    meta=merge_import_meta(meta,explicit_meta,'原始信息列',metadata_alternatives)
    if reference is not None:
        r=normalized_reference(reference)
        meta=merge_import_meta(meta,{'Cycle':r['Cycle'],'Visit':r['Visit-Cycle'],'Timepoint':r['Time Point']},'受试者库',[])
    for h in ['Analyst', 'Date', 'Time', 'Tube']:
        meta[h] = get(h) if get(h) is not None else ''
    meta['Date']=clean_date(meta['Date'])
    manual_tube = not txt(meta['Tube']) and bool(txt((options or {}).get('tube_default')))
    if manual_tube: meta['Tube'] = txt(options['tube_default'])
    if not txt(meta['Tube']):
        raise ValueError('Tube缺失，无法区分同一样品的不同管；请映射Tube列，或确认整份文件管名相同后填写“缺失Tube时管名”')
    if source_panel: meta['Source Panel'] = 'Panel' + str(int(re.search(r'\d+',source_panel)[0]))
    result = {'meta': meta, 'raw': raw, 'labels': labels, 'column_order':list(labels)}
    if metadata_alternatives: result['metadata_alternatives']=metadata_alternatives
    tail=re.search(r'(?<!\d)(\d{6}\.?\d{5})\s+(\d{4,7})(?:\.0+)?\s*$',specimen)
    if bas and tail and bas_id(tail[1])==bas and int(tail[2])>0:
        result['beads_from_name']=int(tail[2])
        result['specimen_original']=specimen
    if validation_name:
        result['sample_kind'] = 'validation'
        result['sample_name'] = validation_name
        result['specimen_original'] = specimen
    if fields is not None: result['input_fields'] = fields
    result['import_options_original']={k:v for k,v in (options or {}).items() if not k.startswith('_')}
    protected={k for k in (fields or {}).values() if k}
    protected.update(norm(x) for names in META_ALIASES.values() for x in names)
    protected.add(norm((options or {}).get('_beads_field','Beads Events')))
    protected.update((options or {}).get('_text_fields',[]))
    replacements={key:value for key,value in raw.items() if key not in protected and txt(value)=='#'}
    if replacements:
        result['hash_originals']=replacements
        result['raw']={key:0 if key in replacements else value for key,value in raw.items()}
    if manual_tube: result['manual_tube'] = True
    if subject_from_reference: result['subject_from_reference'] = True
    return result


def import_warnings(payload):
    m = payload['meta']; warnings = []
    if not txt(m.get('Run number')): warnings.append('Run缺失：可汇总原始指标，不计算复管CV或按Run回算ABS')
    if payload.get('sample_kind') == 'validation':
        warnings.append('验证样品：'+payload['sample_name']+'；按当前项目、Panel、名称、Run和管名区分，不关联受试者库；单独导出到验证样品表')
        if not project_code(payload): warnings.append('验证样品无可识别项目号，按当前选择的项目归档，请核对')
    elif not txt(m.get('Subject ID')): warnings.append('受试者编号缺失；以BAS ID识别样品，不用样品流水号替代受试者号')
    if payload.get('manual_tube'): warnings.append('来源缺失Tube；使用人工指定管名：'+m['Tube'])
    if payload.get('subject_from_reference'): warnings.append('Subject ID由受试者库按BAS ID关联')
    if not txt(m.get('Analyst')): warnings.append('Analyst缺失，已留空')
    if not txt(m.get('Time')): warnings.append('Time缺失，已留空；不使用文件时间补填，也不据此自动采用重分析版本')
    if m.get('Timepoint') in ('Pre','Post'): warnings.append('时间点无具体时长，保留'+m['Timepoint'])
    date_value = txt(m.get('Date'))
    if not date_value: warnings.append('Date缺失，已留空')
    else:
        valid = False
        for fmt in ['%m/%d/%Y','%d/%m/%Y','%Y-%m-%d','%Y/%m/%d','%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M:%S']:
            try: datetime.strptime(date_value,fmt); valid = True; break
            except ValueError: pass
        if not valid: warnings.append(f'Date格式或日期异常：{date_value}；保留源值，请核对')
    if any(re.search(r' \[列[A-Z]+\]$',v) for v in payload['labels'].values()):
        warnings.append('存在重名指标列，已加Excel列号；请核对各列细胞群及输出名称')
    return warnings


def sample_identity(payload):
    # BAS ID is a unique physical sample; optional visit fields are validation
    # metadata, not part of its stable identity.
    m = payload['meta']
    if payload.get('sample_kind') == 'validation':
        return ('validation', txt(payload['sample_name']).casefold(), run_number(m['Run number']) if m['Run number'] else '')
    return (bas_id(m['BAS ID']),)


def sample_display(payload):
    return payload.get('sample_name') or payload['meta']['BAS ID']


def project_code(payload):
    if payload['meta']['BAS ID']: return bas_id(payload['meta']['BAS ID'])[:6]
    match = re.match(r'^\d{6}-(\d{6})-', txt(payload['meta']['Experiment']))
    return match[1] if match else ''


def logical_key(payload):
    return dumps(sample_identity(payload) + (norm(payload['meta']['Tube']),))


def replicate_tag(tube):
    """A terminal rpt integer identifies a replicate, never a revision number."""
    match = re.fullmatch(r'(?:(.+?)\s+)?rpt\s*(\d+)', txt(tube), re.I)
    if match and int(match[2]) > 0:
        return norm(match[1] or 'Sample'), int(match[2])
    return None


def distinct_replicates(left, right):
    a, b = left['meta'], right['meta']
    ra, rb = replicate_tag(a['Tube']), replicate_tag(b['Tube'])
    return bool(ra and rb and ra[0] == rb[0] and ra[1] != rb[1]
                and (not a['Run number'] or not b['Run number'] or run_number(a['Run number']) == run_number(b['Run number'])))


def cv_metric_threshold(cfg, column, group, default):
    boundary, limit = cfg.get('cv_low_pct_boundary'), cfg.get('cv_low_pct_threshold')
    if boundary is None or limit is None or column['kind'] not in ('pct100','pct1'):
        return default
    if column['mode'] not in ('raw','parent_abs'): return default
    percentages=[]
    for _, rec, payload, values, calc in group:
        try: v=metric_number(payload['raw'].get(norm(column['source'])),column['source'],column['kind'])
        except ValueError: return default
        percentages.append(v * (100 if column['kind']=='pct1' else 1))
    return numeric(limit,'低表达CV限值') if percentages and statistics.mean(percentages)<numeric(boundary,'低表达%parent界限','pct100') else default


def replicate_cv_rows(prepared, cfg, panel_name, _allow_partial=False):
    if not cfg.get('replicate_cv', True):
        return []
    groups = {}
    for rec, payload, values, calc, verified in prepared:
        tag = replicate_tag(payload['meta']['Tube'])
        if tag:
            run = payload['meta']['Run number']
            key = sample_identity(payload) + (run_number(run) if run else 'missing:'+str(rec['id']), tag[0])
            groups.setdefault(key, []).append((tag[1], rec, payload, values, calc))
    output = []
    threshold = cfg.get('cv_threshold')
    threshold = numeric(threshold, 'CV阈值') if threshold is not None else None
    for group in groups.values():
        group.sort(key=lambda g: g[0])
        indices = [g[0] for g in group]
        duplicate_indices=len(indices)!=len(set(indices))
        if duplicate_indices and not _allow_partial:
            raise ValueError('同一个Run/样品存在复管编号重复（如rpt1与rpt01），请先核对管名')
        m = dict(group[0][2]['meta'])
        for field in ['Subject ID','Cycle', 'Visit', 'Timepoint']:
            m[field] = next((g[2]['meta'].get(field) for g in group if g[2]['meta'].get(field)), '')
        prefix = [m['Experiment'], panel_name, m['Subject ID'], m['Visit'], m['Timepoint'], m['Run number'], m['BAS ID'], replicate_tag(m['Tube'])[0]]
        pf_note = '包含F复管，需复核' if any(g[1]['pf'] == 'F' for g in group) else ('存在P/F未判定' if any(not g[1]['pf'] for g in group) else '')
        if duplicate_indices:
            output.append(prefix+['复管组（编号重复）',len(group),','.join('rpt'+str(i) for i in indices),','.join(str(g[1]['id']) for g in group)]+[None]*4+['复管编号重复，未计算CV',pf_note,m.get('Cycle',''),group[0][2].get('sample_name','')])
            continue
        for i, c in enumerate(cfg['columns']):
            if c['mode'] == 'raw' and c['kind'] == 'text':
                continue
            numbers = []
            metric_threshold = cv_metric_threshold(cfg,c,group,threshold)
            for _, rec, p, values, calc in group:
                v = values[i]
                if isinstance(v, tuple):
                    v = calc[v[1]][6]
                if v is not None:
                    numbers.append(float(v))
            if not numbers:
                continue  # e.g. Isotype ABS is not applicable.
            mean = statistics.mean(numbers)
            sd = statistics.stdev(numbers) if len(numbers) >= 2 else None
            cv = None
            if not m['Run number']:
                result = '缺少Run，无法确认复管批次，不计算CV'
            elif len(numbers) != len(group):
                result = '有复管缺值，不计算CV'
            elif len(numbers) < 2:
                result = '不足2管，不计算CV'
            elif mean <= 0:
                result = '均值≤0，不计算CV'
            else:
                cv = sd / mean * 100
                result = '未设阈值' if metric_threshold is None else ('达标' if cv <= metric_threshold else '超阈值')
            if any(v is not None and not math.isfinite(v) for v in [mean, sd, cv]):
                raise ValueError('复管统计结果不是有限数值')
            output.append(prefix + [c['name'], len(numbers), ', '.join('rpt'+str(x) for x in indices),
                ', '.join(str(g[1]['id']) for g in group), mean, sd, cv, metric_threshold, result, pf_note, m.get('Cycle',''), group[0][2].get('sample_name','')])
    return output


def digest_payload(p):
    # Labels differ only cosmetically; keep all raw fields/values in the content identity.
    raw = dict(p['raw'])
    key = p.get('input_fields',{}).get('Specimen') or norm('Specimen')
    m = re.fullmatch(SPECIMEN_PATTERN, txt(raw.get(key)), re.I)
    if m:
        raw[key] = ' '.join([txt(m[1]), 'Run'+str(run_number(m[2])), bas_id(m[3])]).strip()
    return hashlib.sha256(dumps({'meta':normalized_meta(p['meta']), 'raw':raw}).encode('utf-8')).hexdigest()


def natural_sort_key(value):
    text=txt(value).casefold()
    try: return (0,0,float(text)) if text else (1,0,0)
    except ValueError: return (0,1,tuple((0,int(x)) if x.isdigit() else (1,x) for x in re.split(r'(\d+)',text)))


def meta_differences(left,right):
    a,b=normalized_meta(left),normalized_meta(right); differences={}
    for field in ('Subject ID','Cycle','Visit','Timepoint'):
        x,y=txt(a.get(field)),txt(b.get(field))
        if not x or not y: continue
        different=not compatible_time(x,y) if field=='Timepoint' else norm(x)!=norm(y)
        if different: differences[field]=(x,y)
    return differences


def merge_import_meta(left,right,source,alternatives):
    a,b=normalized_meta(left),normalized_meta(right)
    for field,(x,y) in meta_differences(a,b).items():
        alternatives.append({'field':field,'value':y,'source':source})
    for field in ('Subject ID','Cycle','Visit'):
        if not txt(a.get(field)): a[field]=b.get(field,'')
    if compatible_time(a.get('Timepoint',''),b.get('Timepoint','')):
        a['Timepoint']=merge_time(a.get('Timepoint',''),b.get('Timepoint',''))
    return a


VALIDATION_CV_HEADERS=['比较','验证材料','管组','Run','分析员','批次用途','指标','数量','CV%','阈值%','判定','记录ID','计算值']


def validation_tube(tube):
    tag=replicate_tag(tube)
    if tag: return tag
    match=re.fullmatch(r'\s*(.*?)\s*(\d+)\s*',txt(tube))
    return (norm(match[1]),int(match[2])) if match else (norm(tube),None)

class ReportReadContext:
    """Read reuse for one synchronous calculation only; never retained by Store/UI.

    A new context for each public operation prevents stale data after edits,
    parameter confirmation, reanalysis, or reference-list imports.
    """
    def __init__(self, store):
        self.store = store
        self._subjects = {}; self._beads = {}; self._conflicts = {}
        self._by_bas = {}; self._conflict_bas = {}

    def subjects(self, pid):
        if pid not in self._subjects:
            self._subjects[pid] = {r['样品编号']:r for r in self.store.subject_rows(pid)}
        return self._subjects[pid]

    def beads(self, panel):
        if panel not in self._beads: self._beads[panel] = self.store.beads_reviews(panel)
        return self._beads[panel]

    def records_for_bas(self, panel, state, bas):
        key = (panel, state)
        if key not in self._by_bas:
            groups = {}
            for r in self.store.records(panel, [state]):
                groups.setdefault(json.loads(r['payload'])['meta'].get('BAS ID'), []).append(r)
            self._by_bas[key] = groups
        return self._by_bas[key].get(bas, [])

    def conflicts(self, pid):
        if pid not in self._conflicts:
            rows = list(self.store.db.execute("SELECT r.*,p.name AS panel_name FROM records r JOIN panels p ON p.id=r.panel_id WHERE p.project_id=? AND p.archived=0 AND r.state='current'", (pid,)))
            result = self.store.metadata_conflicts(pid, _refs=self.subjects(pid), _records=rows)
            self._conflicts[pid] = result
            self._conflict_bas[pid] = {json.loads(r['payload'])['meta'].get('BAS ID') for r in rows if r['id'] in result}
        return self._conflicts[pid]

    def bas_has_conflict(self, pid, bas):
        self.conflicts(pid)
        return bas in self._conflict_bas[pid]


class Store:
    def __init__(self, path):
        self.path = Path(path).resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(str(self.path), timeout=15)
        self.db.row_factory = sqlite3.Row
        self.db.execute('PRAGMA foreign_keys=ON')
        self.db.execute('PRAGMA busy_timeout=15000')
        ver = self.db.execute('PRAGMA user_version').fetchone()[0]
        tables = {r[0] for r in self.db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if ver not in (0, 200, 203, 206, 207, 310) or (ver == 0 and tables):
            self.db.close()
            raise ValueError('不是本工具支持的数据库；请选择原V2数据库或V3.1数据库')
        if ver and not {'validation_settings','beads_review'}.issubset(tables):self.backup()
        self.db.executescript('''
        CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,experiment TEXT NOT NULL DEFAULT '');
        CREATE TABLE IF NOT EXISTS panels(id INTEGER PRIMARY KEY,project_id INTEGER REFERENCES projects(id),name TEXT NOT NULL,config TEXT NOT NULL,UNIQUE(project_id,name));
        CREATE TABLE IF NOT EXISTS subjects(project_id INTEGER REFERENCES projects(id),bas TEXT NOT NULL,payload TEXT NOT NULL,PRIMARY KEY(project_id,bas));
        CREATE TABLE IF NOT EXISTS beads(panel_id INTEGER REFERENCES panels(id),run INTEGER NOT NULL,count INTEGER NOT NULL,volume REAL NOT NULL,lot TEXT NOT NULL,PRIMARY KEY(panel_id,run));
        CREATE TABLE IF NOT EXISTS records(id INTEGER PRIMARY KEY,panel_id INTEGER REFERENCES panels(id),logical TEXT NOT NULL,fingerprint TEXT NOT NULL,payload TEXT NOT NULL,state TEXT NOT NULL CHECK(state IN ('current','pending','superseded','rejected')),pf TEXT NOT NULL DEFAULT '' CHECK(pf IN ('','P','F')),beads INTEGER,volume REAL,lot TEXT,created TEXT NOT NULL,reason TEXT NOT NULL,supersedes INTEGER REFERENCES records(id));
        CREATE UNIQUE INDEX IF NOT EXISTS one_current ON records(panel_id,logical) WHERE state='current';
        CREATE INDEX IF NOT EXISTS content_index ON records(panel_id,fingerprint);
        CREATE TABLE IF NOT EXISTS slots(panel_id INTEGER REFERENCES panels(id),source_key TEXT NOT NULL,logical TEXT NOT NULL,PRIMARY KEY(panel_id,source_key));
        CREATE TABLE IF NOT EXISTS sources(id INTEGER PRIMARY KEY,record_id INTEGER REFERENCES records(id),file TEXT NOT NULL,sha TEXT NOT NULL,sheet TEXT NOT NULL,line INTEGER NOT NULL,UNIQUE(record_id,file,sha,sheet,line));
        CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,project_id INTEGER REFERENCES projects(id),stamp TEXT NOT NULL,action TEXT NOT NULL,detail TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS exclusions(path TEXT PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS validation_settings(panel_id INTEGER PRIMARY KEY,payload TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS beads_review(panel_id INTEGER,run INTEGER,payload TEXT NOT NULL,PRIMARY KEY(panel_id,run));
        ''')
        try:
            missing=[]
            for table in ('projects','panels'):
                fields={row[1] for row in self.db.execute('PRAGMA table_info('+table+')')}
                if 'archived' not in fields: missing.append((table,'archived','INTEGER NOT NULL DEFAULT 0'))
                if 'original_name' not in fields: missing.append((table,'original_name',"TEXT NOT NULL DEFAULT ''"))
            if missing:
                if ver: self.backup()
                with self.db:
                    for table,field,definition in missing:self.db.execute('ALTER TABLE '+table+' ADD COLUMN '+field+' '+definition)
            if ver == 200:
                self.migrate_cycle()
            if ver in (200,203):
                self.migrate_sample_identity()
            elif ver == 0:
                self.db.execute('PRAGMA user_version=206')
            if ver not in (207,310):
                if ver == 206: self.backup()
                with self.db:
                    self.db.execute('PRAGMA user_version=207')
                    for project in self.projects():
                        self.audit(project['id'],'升级验证样品与逐文件导入',{'version':VERSION,'existing_records':'已有记录、指纹、P/F与历史保持不变'})
            if not self.db.execute("SELECT 1 FROM audit WHERE action='启动V2.1' LIMIT 1").fetchone():
                if ver: self.backup()
                with self.db:
                    self.audit(None,'启动V2.1',{'version':VERSION,'migration':'保留已有Panel名称、配置、记录ID、P/F和历史；数据库结构不变'})
            self.normalize_visit_data()
            self.normalize_date_data()
            if ver!=310:
                with self.db:
                    for row in self.db.execute('SELECT id,config FROM panels').fetchall():
                        cfg=json.loads(row['config']);cfg.pop('_setup_pending',None);cfg.pop('_setup_step',None)
                        self.db.execute('UPDATE panels SET config=? WHERE id=?',(dumps(cfg),row['id']))
                    self.db.execute('PRAGMA user_version=310')
                    self.audit(None,'升级V3.1',{'version':VERSION,'changes':'直接导航、数据库选择、可恢复删除及日期整理；数据与P/F保留'})
            self.remove_retired_config()
        except Exception:
            self.db.close()
            raise

    def remove_retired_config(self):
        """Remove retired export configuration, including copies in audit snapshots."""
        if self.db.execute("SELECT 1 FROM audit WHERE action='移除旧传输导出配置' LIMIT 1").fetchone():return
        changes=[]; deleted=[]
        for row in self.db.execute("SELECT id,config FROM panels"):
            old=json.loads(row["config"]);new=strip_retired_config(old)
            if old!=new:changes.append(("panels","config",dumps(new),row["id"]))
        for row in self.db.execute("SELECT id,action,detail FROM audit"):
            if "DTS" in row["action"].upper():deleted.append((row["id"],));continue
            old=json.loads(row["detail"]);new=strip_retired_config(old)
            if old!=new:changes.append(("audit","detail",dumps(new),row["id"]))
        if changes or deleted:self.backup()
        with self.db:
            for table,col,value,ident in changes:
                self.db.execute(f"UPDATE {table} SET {col}=? WHERE id=?",(value,ident))
            self.db.executemany("DELETE FROM audit WHERE id=?",deleted)
            self.audit(None,"移除旧传输导出配置",{"version":VERSION,"configuration_changes":len(changes),"removed_actions":len(deleted)})

    def close(self):
        self.db.close()

    def normalize_date_data(self):
        action='整理日期格式V3.1'
        if self.db.execute('SELECT 1 FROM audit WHERE action=?',(action,)).fetchone():return
        records=[];subjects=[]
        for row in self.db.execute('SELECT * FROM records'):
            p=json.loads(row['payload']);before=p['meta'].get('Date','');after=clean_date(before)
            if before!=after:
                p.setdefault('date_original',before);p['meta']['Date']=after
                records.append((dumps(p),digest_payload(p),row['id'],before,after))
        for row in self.db.execute('SELECT * FROM subjects'):
            p=json.loads(row['payload']);before=p.get('Collection Date','');after=clean_date(before)
            if before!=after:
                p['Collection Date']=after;subjects.append((dumps(p),row['project_id'],row['bas'],before,after))
        if records or subjects:self.backup()
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            if self.db.execute('SELECT 1 FROM audit WHERE action=?',(action,)).fetchone():return
            self.db.executemany('UPDATE records SET payload=?,fingerprint=? WHERE id=?',[r[:3] for r in records])
            self.db.executemany('UPDATE subjects SET payload=? WHERE project_id=? AND bas=?',[r[:3] for r in subjects])
            self.audit(None,action,{'records':[{'id':r[2],'before':r[3],'after':r[4]} for r in records],
                'subjects':[{'project':r[1],'bas':r[2],'before':r[3],'after':r[4]} for r in subjects],'raw_source':'原始值和来源保留'})
            owners={r['id']:r['project_id'] for r in self.db.execute('SELECT r.id,p.project_id FROM records r JOIN panels p ON p.id=r.panel_id')}
            for project in self.projects():
                rc=[{'id':r[2],'before':r[3],'after':r[4]} for r in records if owners.get(r[2])==project['id']]
                sc=[{'bas':r[2],'before':r[3],'after':r[4]} for r in subjects if r[1]==project['id']]
                if rc or sc:self.audit(project['id'],'整理显示日期',{'records':rc,'subjects':sc})

    def normalize_visit_data(self):
        action='统一访视采血格式V2.1.8'
        if self.db.execute('SELECT 1 FROM audit WHERE action=?',(action,)).fetchone(): return
        changes=[]; subjects=[]; failures=[]
        for row in self.db.execute('SELECT * FROM records ORDER BY id'):
            p=json.loads(row['payload']); before=dict(p['meta'])
            try: after=normalized_meta(before)
            except ValueError as exc:
                failures.append({'record':row['id'],'reason':str(exc)}); continue
            if before!=after:
                p.setdefault('visit_normalization_original',before)
                p['meta']=after
                changes.append((dumps(p),digest_payload(p),row['id'],before,after))
        for row in self.db.execute('SELECT * FROM subjects'):
            before=json.loads(row['payload'])
            try: after=normalized_reference(before)
            except ValueError as exc:
                failures.append({'bas':row['bas'],'reason':str(exc)}); continue
            if before!=after: subjects.append((dumps(after),row['project_id'],row['bas'],before,after))
        if changes or subjects: self.backup()
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            if self.db.execute('SELECT 1 FROM audit WHERE action=?',(action,)).fetchone(): return
            self.db.executemany('UPDATE records SET payload=?,fingerprint=? WHERE id=?',[r[:3] for r in changes])
            self.db.executemany('UPDATE subjects SET payload=? WHERE project_id=? AND bas=?',[r[:3] for r in subjects])
            self.audit(None,action,{'version':VERSION,'records':[{'id':r[2],'before':r[3],'after':r[4]} for r in changes],
                'subjects':[{'project':r[1],'bas':r[2],'before':r[3],'after':r[4]} for r in subjects],
                'unresolved':failures,'preserved':'记录ID、P/F、复管、原始单元格及来源不变；格式更新后重新导出Summary'})
            ownership={r['id']:r['project_id'] for r in self.db.execute('SELECT r.id,p.project_id FROM records r JOIN panels p ON p.id=r.panel_id')}
            for project in self.projects():
                rc=[{'id':r[2],'before':r[3],'after':r[4]} for r in changes if ownership.get(r[2])==project['id']]
                sc=[{'bas':r[2],'before':r[3],'after':r[4]} for r in subjects if r[1]==project['id']]
                if rc or sc: self.audit(project['id'],'整理访视采血字段',{'version':VERSION,'records':rc,'subjects':sc,'source_values':'原始值及P/F保留'})
        # Retry only old formatting conflicts, not IDs, formula cells or real duplicate conflicts.
        for project in self.projects():
            for issue in list(self.subject_import_issues(project['id'])):
                if issue.get('header_error') or not set(issue['errors']).issubset({'Cycle','Visit-Cycle','Time Point'}): continue
                try:
                    self.resolve_subject_issue(project['id'],issue['audit_id'],issue['index'],issue['row'],
                        '统一格式后重新核对通过；自动修复旧访视格式误报',_backup=False)
                except ValueError: pass

    def migrate_cycle(self):
        """Atomic v200 -> v203 migration; preserve IDs, P/F, source values and revisions."""
        self.backup()
        def key_v3(value):
            parts = json.loads(value)
            if len(parts) == 6:
                return dumps(parts[:3] + [parts[-1]])
            if len(parts) == 4:
                return value
            raise ValueError('旧数据库结果标识格式不支持，升级已停止')
        records, current, slots = [], set(), {}
        for r in self.db.execute('SELECT * FROM records ORDER BY id'):
            p = json.loads(r['payload']); m = p['meta']
            m['Cycle'],m['Visit'],m['Timepoint'] = normalized_visit(m.get('Cycle',''),m.get('Visit',''),m.get('Timepoint',''))
            key = key_v3(r['logical'])
            if r['state'] == 'current':
                pair=(r['panel_id'],key)
                if pair in current:
                    raise ValueError('旧数据库同一样品同一管存在多个有效版本，需核对后再升级；原数据保留')
                current.add(pair)
            records.append((dumps(p),digest_payload(p),key,r['id']))
        for r in self.db.execute('SELECT * FROM slots'):
            source=(r['panel_id'],key_v3(r['source_key']));target=key_v3(r['logical'])
            if source in slots and slots[source] != target:
                raise ValueError('旧数据库管名关联存在歧义，升级已停止')
            slots[source]=target
        subjects=[(dumps(normalized_reference(json.loads(r['payload']))),r['project_id'],r['bas'])
                  for r in self.db.execute('SELECT * FROM subjects')]
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            self.db.execute('DROP INDEX one_current')
            self.db.executemany('UPDATE records SET payload=?,fingerprint=?,logical=? WHERE id=?',records)
            self.db.execute('DELETE FROM slots')
            self.db.executemany('INSERT INTO slots VALUES(?,?,?)',[(p,k,v) for (p,k),v in slots.items()])
            self.db.executemany('UPDATE subjects SET payload=? WHERE project_id=? AND bas=?',subjects)
            self.db.execute("CREATE UNIQUE INDEX one_current ON records(panel_id,logical) WHERE state='current'")
            for p in self.projects():
                self.audit(p['id'],'升级Cycle与时间格式',{'version':VERSION,'source_values':'原始值不变','identity':'唯一样品编号；可缺省的访视信息另行核对'})
            self.db.execute('PRAGMA user_version=203')

    def audit(self, project, action, detail):
        self.db.execute('INSERT INTO audit(project_id,stamp,action,detail) VALUES(?,?,?,?)', (project, now(), action, dumps(detail)))

    def migrate_sample_identity(self):
        """Only replace relationship keys; preserve source payloads, hashes, IDs and P/F."""
        self.backup()
        def convert(key):
            parts = json.loads(key)
            if len(parts) == 4: return dumps([bas_id(parts[1]),parts[-1]])
            if len(parts) == 2: return dumps([bas_id(parts[0]),parts[-1]])
            raise ValueError('旧结果标识无法迁移，原数据已保留')
        updates=[];current=set();slots={}
        for row in self.db.execute('SELECT id,panel_id,logical,state FROM records'):
            key=convert(row['logical']);pair=(row['panel_id'],key)
            if row['state']=='current':
                if pair in current: raise ValueError('同一样品同一管存在多个有效结果，需核对后再升级；已备份，未合并数据')
                current.add(pair)
            updates.append((key,row['id']))
        for row in self.db.execute('SELECT * FROM slots'):
            key=(row['panel_id'],convert(row['source_key']));target=convert(row['logical'])
            if key in slots and slots[key]!=target: raise ValueError('旧管名关联有冲突，升级停止，原数据保留')
            slots[key]=target
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            self.db.execute('DROP INDEX one_current')
            self.db.executemany('UPDATE records SET logical=? WHERE id=?',updates)
            self.db.execute('DELETE FROM slots')
            self.db.executemany('INSERT INTO slots VALUES(?,?,?)',[(p,k,v) for (p,k),v in slots.items()])
            self.db.execute("CREATE UNIQUE INDEX one_current ON records(panel_id,logical) WHERE state='current'")
            for project in self.projects(): self.audit(project['id'],'升级样品唯一标识',{'version':VERSION,'identity':'BAS ID；Panel和管名分别保存，原始值和历史不变'})
            self.db.execute('PRAGMA user_version=206')

    def backup(self, destination=None):
        if destination is None:
            d = self.path.parent / (self.path.stem + '_backups')
            d.mkdir(exist_ok=True)
            destination = d / (datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.sqlite3')
        dest = Path(destination).resolve()
        if dest == self.path:
            raise ValueError('备份不能覆盖当前数据库')
        con = sqlite3.connect(str(dest))
        try:
            self.db.backup(con)
        finally:
            con.close()
        return dest

    def projects(self, include_deleted=False):
        return self.db.execute('SELECT * FROM projects'+('' if include_deleted else ' WHERE archived=0')+' ORDER BY id').fetchall()

    def panels(self, pid, include_deleted=False):
        return self.db.execute('SELECT p.* FROM panels p JOIN projects j ON j.id=p.project_id WHERE p.project_id=?'+('' if include_deleted else ' AND p.archived=0 AND j.archived=0')+' ORDER BY p.id', (pid,)).fetchall()

    def panel(self, panel):
        row = self.db.execute('SELECT * FROM panels WHERE id=?', (panel,)).fetchone()
        if row is None:
            raise ValueError('请先选择项目和Panel')
        if row['archived'] or self.db.execute('SELECT archived FROM projects WHERE id=?',(row['project_id'],)).fetchone()[0]:
            raise ValueError('该项目或Panel已删除；请先恢复后使用')
        return row

    def add_project(self, name, use_default=True):
        name = txt(name)
        if not name:
            raise ValueError('项目名不能为空')
        with self.db:
            if self.db.execute('SELECT 1 FROM projects WHERE name=?',(name,)).fetchone():raise ValueError('已有同名项目，请使用其他名称')
            ident = self.db.execute('INSERT INTO projects(name) VALUES(?)', (name,)).lastrowid
            self.audit(ident, '新建项目', {'name': name})
        self.add_panel(ident, 'Panel1', use_default)
        return ident

    def add_panel(self, pid, name, use_default=False):
        if not txt(name):
            raise ValueError('Panel名称不能为空')
        if not self.db.execute('SELECT 1 FROM projects WHERE id=? AND archived=0',(pid,)).fetchone():
            raise ValueError('请先选择未删除的项目')
        if self.db.execute('SELECT 1 FROM panels WHERE project_id=? AND name=?',(pid,txt(name))).fetchone():
            raise ValueError('已有同名Panel，请使用其他名称')
        with self.db:
            ident = self.db.execute('INSERT INTO panels(project_id,name,config) VALUES(?,?,?)',
                                    (pid, txt(name), dumps(default_config() if use_default else empty_config()))).lastrowid
            self.audit(pid, '新建Panel', {'panel': ident, 'name': name})
        return ident

    def panel_dependents(self, ident):
        parent=self.panel(ident);out=[]
        for p in self.panels(parent['project_id']):
            if p['id']==ident:continue
            fields=[c['name'] for c in self.config(p['id'])['columns'] if c['mode']=='parent_abs' and c['parent_panel']==ident]
            if fields:out.append(p['name']+'：'+', '.join(fields))
        return out

    def delete_panel(self, ident):
        row=self.panel(ident);dependents=self.panel_dependents(ident)
        if dependents:raise ValueError('该Panel仍被跨Panel汇报项引用，请先修改这些来源：\n'+'\n'.join(dependents))
        backup=self.backup()
        with self.db:
            self.db.execute('UPDATE panels SET archived=1,original_name=name,name=? WHERE id=?',('__deleted_panel_'+str(ident),ident))
            self.audit(row['project_id'],'删除Panel（可恢复）',{'panel':ident,'name':row['name'],'backup':str(backup)})
        return backup

    def delete_project(self, ident):
        row=self.db.execute('SELECT * FROM projects WHERE id=? AND archived=0',(ident,)).fetchone()
        if row is None:raise ValueError('项目不存在或已经删除')
        backup=self.backup()
        with self.db:
            self.db.execute('UPDATE projects SET archived=1,original_name=name,name=? WHERE id=?',('__deleted_project_'+str(ident),ident))
            self.audit(ident,'删除项目（可恢复）',{'project':ident,'name':row['name'],'backup':str(backup)})
        return backup

    def deleted_items(self):
        out=[{'kind':'project','id':r['id'],'name':r['original_name'],'project':'','project_id':r['id']} for r in self.projects(True) if r['archived']]
        for row in self.db.execute('SELECT p.*,j.name AS project_name FROM panels p JOIN projects j ON j.id=p.project_id WHERE p.archived=1 AND j.archived=0 ORDER BY p.id'):
            out.append({'kind':'panel','id':row['id'],'name':row['original_name'],'project':row['project_name'],'project_id':row['project_id']})
        return out

    def restore_item(self, kind, ident, name=''):
        if kind not in ('project','panel'):raise ValueError('恢复类型错误')
        table='projects' if kind=='project' else 'panels'
        row=self.db.execute('SELECT * FROM '+table+' WHERE id=? AND archived=1',(ident,)).fetchone()
        if row is None:raise ValueError('删除项不存在或已经恢复')
        name=txt(name) or row['original_name'];pid=ident if kind=='project' else row['project_id']
        if kind=='project':conflict=self.db.execute('SELECT 1 FROM projects WHERE name=? AND id<>?',(name,ident)).fetchone()
        else:
            if self.db.execute('SELECT archived FROM projects WHERE id=?',(pid,)).fetchone()[0]:raise ValueError('请先恢复所属项目')
            conflict=self.db.execute('SELECT 1 FROM panels WHERE project_id=? AND name=? AND id<>?',(pid,name,ident)).fetchone()
        if conflict:raise ValueError('已有同名项目/Panel，请在恢复名称里填写其他名称')
        self.backup()
        with self.db:
            self.db.execute('UPDATE '+table+" SET archived=0,name=?,original_name='' WHERE id=?",(name,ident))
            self.audit(pid,'恢复删除项',{'type':kind,'id':ident,'name':name})
        return pid

    def proposed_config(self, panel):
        """Suggest raw metrics from a stored example; no guessing of ABS or gates."""
        cfg=self.config(panel)
        if cfg.get('columns'): return json.loads(dumps(cfg))
        records=self.records(panel,['current','pending'])
        if not records: raise ValueError('请先导入一份该Panel的示例数据')
        p=json.loads(records[0]['payload']); fields=set(p.get('input_fields',{}).values())
        result={**cfg,'columns':[]}
        for key in p.get('column_order',list(p['labels'])):
            label=p['labels'][key]
            if key in fields or any(header_token(label) in set(map(norm,v)) for v in META_ALIASES.values()): continue
            kind=infer_kind(label)
            if kind=='number' and isinstance(p['raw'].get(key),str):
                try: metric_number(p['raw'][key],label)
                except ValueError: kind='text'
            result['columns'].append({'source':label,'name':label,'mode':'raw','kind':kind})
        validate_config(result)
        return result

    def panel_number(self, panel):
        row = self.panel(panel)
        return numbered_panel(self.config(panel).get('import_options',{}).get('source_panel','')) or numbered_panel(row['name'])

    def next_panel(self, pid):
        used = [self.panel_number(p['id']) or 0 for p in self.panels(pid)]
        n = max(used,default=0)+1
        while any(p['name'].casefold() == f'panel{n}' for p in self.panels(pid)): n += 1
        return self.add_panel(pid,f'Panel{n}')

    def project_routes(self, pid, candidates):
        """Read-only preview. Existing numbered/explicitly mapped Panels win."""
        panels = self.panels(pid); result=[]
        for candidate in candidates:
            row = dict(candidate); n = row['number']
            matches = [p for p in panels if self.panel_number(p['id']) == n] if n else []
            row['panel_id'] = matches[0]['id'] if len(matches)==1 and not row['skip'] else None
            row['create_name'] = f'Panel{n}' if n and not matches and not row['skip'] else ''
            if len(matches)>1: row['detail'] += '；多个已有Panel对应此编号，需手动选择'
            result.append(row)
        return result

    def import_project(self, pid, assignments):
        """One batch per target Panel: updates in the same batch remain pending."""
        self.backup()
        groups = {}; outcomes = []; seen = set()
        for item in assignments:
            path = str(Path(item['file']).resolve())
            if path in seen: continue
            seen.add(path)
            ident = item.get('panel_id'); create = item.get('create_name','')
            if item.get('skip') or (ident is None and not create):
                outcomes.append({'file':path,'status':'跳过','detail':item.get('detail') or '未分配Panel',
                                 'ids':[],'current':0,'pending':0,'duplicate':0,'panel_id':None,'panel_name':'未分配'})
                continue
            try:
                if ident is None:
                    n = numbered_panel(create)
                    if not n: raise ValueError('新增Panel必须是Panel加编号')
                    matches = [p for p in self.panels(pid) if self.panel_number(p['id'])==n]
                    if len(matches)>1: raise ValueError('多个Panel对应此编号，请手动分配')
                    ident = matches[0]['id'] if matches else self.add_panel(pid,f'Panel{n}')
                target = self.panel(ident)
                if target['project_id'] != pid: raise ValueError('目标Panel不属于当前项目')
                n = self.panel_number(ident)
                if item.get('number') and n is None:
                    n = item['number']
                    if any(p['id']!=ident and self.panel_number(p['id'])==n for p in self.panels(pid)):
                        raise ValueError(f'Panel{n}已关联其他Panel；请选择该Panel或在导入字段配置中调整来源Panel')
                    options=dict(self.config(ident).get('import_options',{})); options['source_panel']=f'Panel{n}'
                    self.save_import_options(ident,options)
                if item.get('number') and n and item['number'] != n:
                    raise ValueError(f'文件名为Panel{item["number"]}，目标为Panel{n}；请检查分配或修正文件名')
                groups.setdefault(ident,[]).append(path)
            except Exception as exc:
                outcomes.append({'file':path,'status':'失败','detail':str(exc),'ids':[],
                                 'current':0,'pending':0,'duplicate':0,'panel_id':ident,'panel_name':create or '分配失败'})
        warnings=[]
        for ident,paths in groups.items():
            try:
                stats = self.import_batch(ident,paths)
                warnings.extend(stats['warnings'])
                for f in stats['files']: outcomes.append({**f,'panel_id':ident,'panel_name':self.panel(ident)['name']})
            except Exception as exc:
                for path in paths:
                    outcomes.append({'file':path,'status':'失败','detail':str(exc),'ids':[],
                                     'current':0,'pending':0,'duplicate':0,'panel_id':ident,'panel_name':self.panel(ident)['name']})
        result = {k:sum(f[k] for f in outcomes) for k in ('current','pending','duplicate')}
        result.update(files=outcomes,errors=[f for f in outcomes if f['status']=='失败'],warnings=warnings,project=pid)
        with self.db: self.audit(pid,'项目自动分Panel导入',result)
        return result

    def summary_token(self, panel):
        cfg = dict(self.config(panel)); cfg.pop('_setup_pending',None); cfg.pop('_setup_step',None)
        data = {'config':cfg,'records':[dict(r) for r in self.records(panel)],
                'subjects':self.subject_rows(self.panel(panel)['project_id']),
                'subject_import_issues':[{k:r[k] for k in ('file','sheet','line','row','errors')} for r in self.subject_import_issues(self.panel(panel)['project_id'])]}
        data['parent_dependencies']=[]
        data['metric_hash_policy']='literal-#-to-zero-v1;Beads-denominator-strict'
        data['summary_policy']='current-only;partial-fields;validation-separate;inline-optional-five-visits-v342'
        for ident in sorted({c['parent_panel'] for c in cfg['columns'] if c['mode']=='parent_abs'}):
            parent=self.db.execute('SELECT * FROM panels WHERE id=?',(ident,)).fetchone()
            if parent and not parent['archived'] and parent['project_id']==self.panel(panel)['project_id']:
                pcfg=dict(json.loads(parent['config']))
                data['parent_dependencies'].append({'panel':ident,'config':pcfg,'records':[dict(r) for r in self.records(ident)]})
            else: data['parent_dependencies'].append({'panel':ident,'missing_or_other_project':True})
        return hashlib.sha256(dumps(data).encode('utf-8')).hexdigest()

    def summary_status(self, panel, exclude_fail=False):
        token = self.summary_token(panel)
        for row in self.db.execute("SELECT detail,stamp FROM audit WHERE project_id=? AND action='导出Excel' ORDER BY id DESC",(self.panel(panel)['project_id'],)):
            info=json.loads(row[0])
            if info.get('panel')==panel:
                return {'ready':info.get('summary_token')==token and info.get('exclude_fail',False)==exclude_fail,
                        'token':token,'destination':info.get('destination',''),
                        'exclude_fail':info.get('exclude_fail',False),'rows':info.get('rows'), 'stamp':row['stamp']}
        return {'ready':False,'token':token,'destination':''}


    def check_project(self, pid, exclude_fail=False):
        checks=[]
        context=ReportReadContext(self)
        for panel in self.panels(pid):
            ident=panel['id']; cfg=self.config(ident); records=self.records(ident,['current'])
            pending=self.records(ident,['pending']); issues=[]; ready=0; excluded=0
            try: validate_config(cfg)
            except ValueError as exc: issues.append(str(exc))
            if pending: issues.append(f'{len(pending)}条待审核：Summary可导出当前有效版本，候选不自动采用')
            if not records: issues.append('没有当前有效结果')
            for rec in records:
                if exclude_fail and rec['pf']=='F':
                    excluded+=1
                    if not cfg.get('replicate_cv',True): continue
                try:
                    self.prepare_report(ident,_rows=[rec],_check_pending=False,_context=context)
                    if not (exclude_fail and rec['pf']=='F'): ready+=1
                except ValueError as exc: issues.append(str(exc))
            if records and excluded==len(records): issues.append('排除F后没有可汇报结果')
            checks.append({'panel_id':ident,'panel_name':panel['name'],'current':len(records),'pending':len(pending),
                           'ready':ready,'excluded':excluded,'issues':list(dict.fromkeys(issues)),
                            'summary_ready':self.summary_status(ident,exclude_fail)['ready']})
        return {'panels':checks,'files':self.unresolved_imports(pid)}

    def unresolved_imports(self, pid):
        latest={}
        active={p['id'] for p in self.panels(pid)}
        for row in self.db.execute("SELECT detail FROM audit WHERE project_id=? AND action IN ('批量导入结果','项目自动分Panel导入') ORDER BY id",(pid,)):
            info=json.loads(row[0])
            for f in info.get('files',[]): latest[f['file']]={**f,'panel_id':f.get('panel_id',info.get('panel'))}
        return [f for f in latest.values() if (f.get('panel_id') is None or f['panel_id'] in active) and f['status']!='成功' and f.get('detail') not in (
            '已登记的汇总文件或受试者资料','程序导出的汇总文件','已登记为受试者资料或程序导出文件')]

    def records(self, panel, states=None):
        sql, args = 'SELECT * FROM records WHERE panel_id=?', [panel]
        if states:
            sql += ' AND state IN (' + ','.join('?' for _ in states) + ')'
            args.extend(states)
        return self.db.execute(sql + ' ORDER BY id', args).fetchall()

    def record(self, ident):
        r = self.db.execute('SELECT * FROM records WHERE id=?', (ident,)).fetchone()
        if r is None:
            raise ValueError('记录不存在')
        return r

    def config(self, panel):
        return json.loads(self.panel(panel)['config'])

    def save_config(self, panel, cfg):
        cfg=strip_retired_config(cfg)
        validate_config(cfg)
        for column in cfg['columns']:
            if column['mode']=='parent_abs': self.parent_source_config(panel,column)
        self.backup()
        with self.db:
            old = self.config(panel)
            self.db.execute('UPDATE panels SET config=? WHERE id=?', (dumps(cfg), panel))
            self.audit(self.panel(panel)['project_id'], '修改报告配置', {'panel': panel, 'before': old, 'after': cfg})

    def save_import_options(self, panel, options):
        validate_import_options(options); self.backup()
        with self.db:
            cfg = self.config(panel); old = cfg.get('import_options',{})
            cfg['import_options'] = options
            self.db.execute('UPDATE panels SET config=? WHERE id=?',(dumps(cfg),panel))
            self.audit(self.panel(panel)['project_id'],'修改导入字段配置',{'panel':panel,'before':old,'after':options})

    def configure_validation_runs(self, updates):
        prepared=[]
        for panel,runs in updates.items():
            self.panel(panel)
            cfg=self.config(panel); options=dict(cfg.get('import_options',{}))
            before=set(options.get('validation_runs',[])); after=set(runs)
            options['validation_runs']=sorted(after); validate_import_options(options)
            for rec in self.records(panel):
                payload=json.loads(rec['payload']); run=payload['meta'].get('Run number')
                if not txt(run): continue
                n=run_number(run); validation=payload.get('sample_kind')=='validation'
                if (n in after-before and not validation) or (n in before-after and validation):
                    raise ValueError(f'该Panel的Run{n}已有不同类型的入库记录，不能直接更改分类；请先核对，已有记录未修改')
            prepared.append((panel,cfg,options,sorted(before)))
        self.backup()
        with self.db:
            for panel,cfg,options,before in prepared:
                cfg['import_options']=options
                self.db.execute('UPDATE panels SET config=? WHERE id=?',(dumps(cfg),panel))
                self.audit(self.panel(panel)['project_id'],'人工设置验证Run',{'panel':panel,'before':before,'after':options['validation_runs']})

    def subject_rows(self, pid):
        return [normalized_reference(json.loads(r[0])) for r in self.db.execute('SELECT payload FROM subjects WHERE project_id=? ORDER BY bas', (pid,))]

    def check_subject(self, pid, payload, must_exist=False, _refs=None):
        if payload.get('sample_kind') == 'validation': return False
        m = payload['meta']
        if _refs is None:
            row = self.db.execute('SELECT payload FROM subjects WHERE project_id=? AND bas=?', (pid, m['BAS ID'])).fetchone()
            ref = normalized_reference(json.loads(row[0])) if row else None
        else:
            ref = _refs.get(m['BAS ID'])
        if not ref:
            if must_exist:
                raise ValueError(f'受试者库缺少 {m["BAS ID"]}')
            return False
        if txt(m['Subject ID']) and txt(ref.get('Patient ID')) != txt(m['Subject ID']):
            raise ValueError(f'{m["BAS ID"]} 的Patient ID与受试者库不一致')
        check_optional_fields(m,{'Cycle':ref['Cycle'],'Visit':ref['Visit-Cycle'],'Timepoint':ref['Time Point']},m['BAS ID'])
        return True

    def put_subjects(self, pid, rows, _backup=True, _sync_visit=False, reason=''):
        if _sync_visit and not txt(reason): raise ValueError('同步修正检测记录时请填写原因')
        cleaned, keys = [], set()
        for r in rows:
            item = normalized_reference(r)
            item['样品编号'] = bas_id(item['样品编号'])
            if item['样品编号'] in keys:
                raise ValueError('输入中样品编号重复：' + item['样品编号'])
            keys.add(item['样品编号'])
            if not txt(item['Patient ID']):
                raise ValueError('Patient ID不能为空；Cycle、Visit、Timepoint可留空')
            cleaned.append(item)
        if not cleaned:
            raise ValueError('没有找到受试者数据')
        if _backup: self.backup()
        with self.db:
            changes = []
            for item in cleaned:
                k = item['样品编号']
                old = self.db.execute('SELECT payload FROM subjects WHERE project_id=? AND bas=?', (pid, k)).fetchone()
                if old and old[0] == dumps(item) and not _sync_visit:
                    continue
                matches=list(self.db.execute("SELECT r.* FROM records r JOIN panels p ON r.panel_id=p.id WHERE p.project_id=? AND p.archived=0 AND r.state IN ('current','pending')", (pid,)))
                if _sync_visit:
                    for row in matches:
                        payload=json.loads(row['payload']); before=dict(payload['meta'])
                        if before['BAS ID']!=k: continue
                        after=dict(before)
                        for field,source in [('Cycle','Cycle'),('Visit','Visit-Cycle'),('Timepoint','Time Point')]:
                            if txt(item[source]): after[field]=item[source]
                        if before==after: continue
                        payload['meta']=after
                        payload.setdefault('manual_edits',[]).append({'area':'meta','field':'访视采血资料','label':'访视采血资料',
                            'before':{f:before.get(f,'') for f in ('Cycle','Visit','Timepoint')},
                            'after':{f:after.get(f,'') for f in ('Cycle','Visit','Timepoint')},
                            'reason':txt(reason),'time':now(),'from_record':row['id'],
                            'operator':os.environ.get('USERNAME') or os.environ.get('USER') or ''})
                        self.db.execute("UPDATE records SET state='superseded' WHERE id=?",(row['id'],))
                        new_id=self._insert_record(row['panel_id'],payload,row['state'],'人工同步访视资料：'+txt(reason),
                            row['logical'],row['id'],'',(row['beads'],row['volume'],row['lot']))
                        self.db.execute('INSERT INTO sources(record_id,file,sha,sheet,line) SELECT ?,file,sha,sheet,line FROM sources WHERE record_id=?',(new_id,row['id']))
                        self.audit(pid,'人工同步访视资料',{'old_record':row['id'],'new_record':new_id,'before':before,'after':after,
                            'reason':txt(reason),'previous_pf':row['pf'],'pf_reset':True,'scope':'同BAS所有Panel的当前/待审核记录；不修改样品编号和受试者编号'})
                self.db.execute('INSERT INTO subjects VALUES(?,?,?) ON CONFLICT(project_id,bas) DO UPDATE SET payload=excluded.payload', (pid, k, dumps(item)))
                changes.append({'before': json.loads(old[0]) if old else None, 'after': item})
            if changes:
                self.audit(pid, '更新受试者库', changes)
        return len(changes)

    def import_subjects(self, pid, path):
        path=Path(path).resolve(); candidates=list(iter_subject_candidates(path))
        for candidate in candidates:
            item,errors=subject_candidate(candidate['row']); candidate['item']=item
            candidate['errors'].update(errors); candidate['status']='待检查'
        # Check all conflicting duplicate IDs before writing any of their rows.
        groups={}
        for i,c in enumerate(candidates):
            if '样品编号' not in c['errors']:
                groups.setdefault(c['item']['样品编号'],[]).append(i)
        for bas,indices in groups.items():
            if len(indices)>1:
                first=candidates[indices[0]]['item']
                if any(candidates[i]['errors'] for i in indices) or any(candidates[i]['item']!=first for i in indices):
                    for i in indices: candidates[i]['errors']['样品编号']='同一清单中样品编号重复且资料不一致：'+bas
                else:
                    for i in indices[1:]: candidates[i]['duplicate']=True
        self.backup(); updated=0
        for c in candidates:
            if c['errors']: c['status']='待修正'
            elif c.get('duplicate'): c['status']='重复'
            else:
                try:
                    count=self.put_subjects(pid,[c['item']],_backup=False); updated+=count
                    c['status']='已更新' if count else '无变化'
                except ValueError as exc:
                    c['status']='待修正'; message=str(exc)
                    field=next((h for h in ('Subject ID','Cycle','Visit','Timepoint') if h in message),'Patient ID')
                    field={'Subject ID':'Patient ID','Visit':'Visit-Cycle','Timepoint':'Time Point'}.get(field,field)
                    c['errors'][field]=message
            c['detail']='；'.join(f'{SUBJECT_LABELS.get(k,k)}：{v}' for k,v in c['errors'].items()) or ('相同资料重复，未重复写入' if c['status']=='重复' else c['status'])
        result={'file':str(path),'updated':updated,'accepted':sum(c['status'] in ('已更新','无变化','重复') for c in candidates),
                'failed':sum(c['status']=='待修正' for c in candidates),'rows':candidates}
        with self.db:
            self.db.execute('INSERT OR IGNORE INTO exclusions VALUES(?)',(str(path),))
            self.audit(pid,'受试者清单导入结果',result)
        return result

    def subject_import_issues(self, pid):
        latest={}; resolved=set()
        for audit in self.db.execute("SELECT id,action,detail FROM audit WHERE project_id=? AND action IN ('受试者清单导入结果','修正受试者问题行','不采用受试者问题行') ORDER BY id",(pid,)):
            data=json.loads(audit['detail'])
            if audit['action']=='受试者清单导入结果':
                latest[data['file']]={'audit_id':audit['id'],**data}
            else: resolved.add((data['audit_id'],data['index']))
        issues=[]
        for batch in latest.values():
            for i,c in enumerate(batch['rows']):
                if c['status']=='待修正' and (batch['audit_id'],i) not in resolved:
                    issues.append({**c,'file':batch['file'],'audit_id':batch['audit_id'],'index':i})
        return issues

    def resolve_subject_issue(self, pid, audit_id, index, row=None, reason='', _backup=True, _sync_visit=False):
        issue=next((r for r in self.subject_import_issues(pid) if r['audit_id']==audit_id and r['index']==index),None)
        if not issue: raise ValueError('该问题行已处理或来源清单已重新导入，请刷新后重选')
        if row is None:
            if not txt(reason): raise ValueError('请填写不采用原因')
            action='不采用受试者问题行'
        else:
            item,errors=subject_candidate(row)
            if errors: raise ValueError('；'.join(f'{k}：{v}' for k,v in errors.items()))
            self.put_subjects(pid,[item],_backup=_backup,_sync_visit=_sync_visit,reason=reason); row=item; action='修正受试者问题行'
        with self.db:
            self.audit(pid,action,{'audit_id':audit_id,'index':index,'source':issue,'corrected':row,'reason':txt(reason) or '界面修正并保存'})

    def metadata_conflicts(self,pid,_refs=None,_records=None):
        groups={}; result={}
        def add(ident,field,text):
            items=result.setdefault(ident,{}).setdefault(field,[])
            if text not in items: items.append(text)
        refs={r['样品编号']:r for r in self.subject_rows(pid)} if _refs is None else _refs
        records = _records if _records is not None else self.db.execute("SELECT r.*,p.name AS panel_name FROM records r JOIN panels p ON p.id=r.panel_id WHERE p.project_id=? AND p.archived=0 AND r.state='current'",(pid,))
        for r in records:
            p=json.loads(r['payload']);m=p['meta'];bas=m.get('BAS ID')
            if not bas: continue
            groups.setdefault(bas,[]).append((r,p))
            for alt in p.get('metadata_alternatives',[]):
                for f,(x,y) in meta_differences(m,{alt['field']:alt['value']}).items():
                    add(r['id'],f,f'当前：{x}；{alt["source"]}：{y}')
            ref=refs.get(bas)
            if ref:
                rm={'Subject ID':ref['Patient ID'],'Cycle':ref['Cycle'],'Visit':ref['Visit-Cycle'],'Timepoint':ref['Time Point']}
                for f,(x,y) in meta_differences(m,rm).items(): add(r['id'],f,f'当前：{x}；受试者库：{y}')
        for items in groups.values():
            for i,(r,p) in enumerate(items):
                for other,op in items[i+1:]:
                    for f,(x,y) in meta_differences(p['meta'],op['meta']).items():
                        add(r['id'],f,f'当前：{x}；{other["panel_name"]} 记录{other["id"]}：{y}')
                        add(other['id'],f,f'当前：{y}；{r["panel_name"]} 记录{r["id"]}：{x}')
        return result

    def beads_reviews(self,panel):
        return {r['run']:json.loads(r['payload']) for r in self.db.execute('SELECT * FROM beads_review WHERE panel_id=?',(panel,))}

    def detect_beads_names(self,panel,payloads):
        candidates={}
        cfg=self.config(panel)
        existing=self.beads_reviews(panel)
        for p in payloads:
            m=p['meta'];count=p.get('beads_from_name')
            if not count or not m.get('BAS ID') or not m.get('Run number'): continue
            if not re.fullmatch(cfg['sample_regex'],txt(m['Tube'])): continue
            candidates.setdefault(run_number(m['Run number']),set()).add(count)
        for run,counts in candidates.items():
            counts.update(existing.get(run,{}).get('counts',[]))
            # Compare other active source names as well, including earlier imported files.
            for r in self.records(panel,['current','pending']):
                p=json.loads(r['payload']);m=p['meta']
                if m.get('Run number') and run_number(m['Run number'])==run and p.get('beads_from_name'):
                    counts.add(p['beads_from_name'])
            stored=self.db.execute('SELECT count,volume,lot FROM beads WHERE panel_id=? AND run=?',(panel,run)).fetchone()
            conflict=len(counts)>1 or (stored is not None and stored['count'] not in counts)
            if stored and not conflict and stored['lot'] and run not in existing: continue
            detail={'counts':sorted(counts),'conflict':conflict,'source':'样品名尾缀','confirmed':False}
            self.db.execute('INSERT OR REPLACE INTO beads_review VALUES(?,?,?)',(panel,run,dumps(detail)))
            if not stored and len(counts)==1:
                count=next(iter(counts))
                for rec in self.records(panel,['current','pending']):
                    p=json.loads(rec['payload']);m=p['meta']
                    if m.get('Run number') and run_number(m['Run number'])==run and rec['beads'] is None:
                        self.db.execute("UPDATE records SET state='superseded' WHERE id=?",(rec['id'],))
                        new=self._insert_record(panel,p,rec['state'],'样品名自动提取Beads，待确认',rec['logical'],rec['id'],'',(count,rec['volume'],rec['lot']))
                        self.db.execute('INSERT INTO sources(record_id,file,sha,sheet,line) SELECT ?,file,sha,sheet,line FROM sources WHERE record_id=?',(new,rec['id']))
            self.audit(self.panel(panel)['project_id'],'样品名识别Beads',{'panel':panel,'run':run,**detail})

    def validation_config(self,panel):
        r=self.db.execute('SELECT payload FROM validation_settings WHERE panel_id=?',(panel,)).fetchone()
        return json.loads(r[0]) if r else {'records':{},'runs':{},'threshold':20,'metrics':[]}

    def save_validation_config(self,panel,cfg):
        validate_validation_batches(cfg)
        threshold=numeric(cfg.get('threshold'),'CV阈值')
        if threshold<0: raise ValueError('CV阈值不能小于0')
        self.backup()
        with self.db:
            self.db.execute('INSERT OR REPLACE INTO validation_settings VALUES(?,?)',(panel,dumps(cfg)))
            self.audit(self.panel(panel)['project_id'],'验证设置',{'panel':panel,'config':cfg})

    def validation_items(self,panel):
        settings=self.validation_config(panel);items=[]
        for rec in self.records(panel,['current']):
            p=json.loads(rec['payload']);m=p['meta']
            if m.get('BAS ID'): continue
            group,rep=validation_tube(m['Tube'])
            control=bool(re.search(r'(?i)control|fmo|isotype|fm[xo]',txt(m['Tube'])))
            default={'material':p.get('sample_name',''),'tube_group':group,'replicate':rep,'role':'Control' if control else '样本管'}
            config={**default,**settings.get('records',{}).get(rec['logical'],{})}
            batch=next((b for b in settings.get('batches',[]) if m['Run number'] in b['runs']),None)
            role=('Baseline' if batch['baseline']==m['Run number'] else 'Qualification') if batch else '未分批'
            items.append({'rec':rec,'payload':p,**config,'run_role':role,'batch':batch['name'] if batch else ''})
        return items

    def validation_report(self,panel,ids,metrics,threshold):
        threshold=numeric(threshold,'CV阈值')
        if threshold<0:raise ValueError('CV阈值不能小于0')
        items=[x for x in self.validation_items(panel) if x['rec']['id'] in set(ids) and x['role']=='样本管']
        if not items:raise ValueError('请选择验证记录')
        metrics=validation_metrics(self.validation_items(panel))
        if not metrics:raise ValueError('没有可计算的数值指标')
        allowed={norm(label):label for x in self.validation_items(panel) for label in x['payload']['labels'].values()}
        if any(norm(metric) not in allowed for metric in metrics):raise ValueError('指标来源不存在，请重新选择')
        groups={};out=[];between={}
        for x in items:
            m=x['payload']['meta'];run=m['Run number']
            key=(x['material'],x['tube_group'],x['role'],run,txt(m['Analyst']))
            groups.setdefault(key,[]).append(x)
        for (material,group,role,run,analyst),rows in sorted(groups.items(),key=lambda pair:str(pair[0])):
            reps=[x['replicate'] for x in rows];ids_text=','.join(str(x['rec']['id']) for x in rows)
            issue=''
            if not run:issue='Run缺失'
            elif not material or not group:issue='请设置验证材料和管组'
            elif len(rows)!=3:issue='需3支复管，当前'+str(len(rows))+'支'
            elif None in reps or len(set(reps))!=3:issue='请确认三个不同的复管编号'
            roles=','.join(sorted({x['run_role'] for x in rows}))
            for metric in metrics:
                values=[];problem=issue
                try:
                    values=[metric_number(x['payload']['raw'].get(norm(metric)),metric,infer_kind(metric)) for x in rows]
                except ValueError as exc:problem=str(exc)
                cv=None
                if not problem:
                    mean=statistics.mean(values)
                    if mean==0:problem='均值为0，CV未定义'
                    else:cv=statistics.stdev(values)/abs(mean)*100
                decision=problem or ('超阈值' if cv>threshold else '达标')
                out.append(['批内CV',material,group+' / '+role,run,analyst,roles,metric,len(rows),cv,threshold,decision,ids_text,dumps(values)])
                if not issue and len(values)==3:
                    between.setdefault((material,group,role,metric),[]).append((run,analyst,roles,statistics.mean(values),ids_text))
        settings=self.validation_config(panel)
        batches=settings.get('batches',[])
        for (material,group,role,metric),values in between.items():
            for batch in batches:
                selected=[v for v in values if v[0] in batch['runs']]
                baseline=[v for v in selected if v[0]==batch['baseline']]
                for current in selected:
                    if current[0]==batch['baseline']:continue
                    problem='';cv=None
                    if len(baseline)!=1:problem='Baseline缺少完整三复管或存在多个分析员组，无法唯一匹配'
                    elif sum(v[0]==current[0] for v in selected)!=1:problem='当前Run有多个分析员组，无法唯一匹配'
                    means=[baseline[0][3],current[3]] if len(baseline)==1 else []
                    if not problem:
                        mean=statistics.mean(means)
                        if mean==0:problem='两批均值的平均为0，CV未定义'
                        else:cv=statistics.stdev(means)/abs(mean)*100
                    ids_text=(','.join([baseline[0][4],current[4]]) if len(baseline)==1 else current[4])
                    out.append(['与Baseline均值CV',material,group+' / '+role,current[0],current[1],batch['name']+' / Baseline '+batch['baseline'],metric,2,cv,threshold,problem or ('超阈值' if cv>threshold else '达标'),ids_text,dumps(means)])
        # Incomplete qualification groups must remain visible, rather than disappear.
        existing={(r[1],r[2],r[3],r[4],r[6]) for r in out if r[0]=='与Baseline均值CV'}
        for row in list(out):
            if row[0]!='批内CV':continue
            batch=next((b for b in batches if row[3] in b['runs']),None)
            if not batch or row[3]==batch['baseline']:continue
            key=(row[1],row[2],row[3],row[4],row[6])
            if key not in existing:
                out.append(['与Baseline均值CV',row[1],row[2],row[3],row[4],batch['name']+' / Baseline '+batch['baseline'],row[6],2,None,threshold,'当前Run缺少完整三复管，未计算',row[11],'[]'])
        return items,out

    def export_validation(self,panel,ids,metrics,threshold,destination):
        items=self.validation_items(panel)
        ids=[x['rec']['id'] for x in items if x['role']=='样本管']
        metrics=validation_metrics(items)
        _,cv=self.validation_report(panel,ids,metrics,threshold) if ids and metrics else ([],[])
        if not items:raise ValueError('没有验证数据可留档')
        selected_metrics=list(metrics)
        archived=[(r,json.loads(r['payload'])) for r in self.records(panel) if not json.loads(r['payload'])['meta'].get('BAS ID')]
        metrics=list(dict.fromkeys(label for r,p in archived for label in p['labels'].values()))
        path=Path(destination).resolve()
        if path.exists():raise ValueError('请使用新的文件名，避免覆盖已有数据')
        wb=Workbook();ws=wb.active;ws.title='验证CV';literal_row(ws,VALIDATION_CV_HEADERS)
        for row in cv:literal_row(ws,row)
        raw=wb.create_sheet('验证原始数据')
        literal_row(raw,['记录ID','验证材料','管组','类型','复管编号','Run','分析员','批次用途','P/F','管名']+metrics)
        sources=wb.create_sheet('来源');literal_row(sources,['记录ID','原文件','Sheet','原始行'])
        for x in items:
            r=x['rec'];m=x['payload']['meta']
            literal_row(raw,[r['id'],x['material'],x['tube_group'],x['role'],x['replicate'],m['Run number'],m['Analyst'],x['run_role'],r['pf'],m['Tube']]+[x['payload']['raw'].get(norm(c)) for c in metrics])
            for source in self.db.execute('SELECT file,sheet,line FROM sources WHERE record_id=?',(r['id'],)):literal_row(sources,[r['id'],*source])
        history=wb.create_sheet('验证历史留档')
        literal_row(history,['记录ID','状态','P/F','样品名','元数据JSON','原始值JSON','人工修改JSON'])
        for rec,p in archived:
            original={**p['raw'],**p.get('hash_originals',{})}
            literal_row(history,[rec['id'],STATE_LABEL[rec['state']],rec['pf'],p.get('sample_name',''),dumps(p['meta']),dumps(original),dumps(p.get('manual_edits',[]))])
            if rec['state']!='current':
                for source in self.db.execute('SELECT file,sheet,line FROM sources WHERE record_id=?',(rec['id'],)):literal_row(sources,[rec['id'],*source])
        note=wb.create_sheet('计算说明');literal_row(note,['项目','规则'])
        for pair in [('版本',VERSION),('批内CV','同Run、同分析员、同材料和管组的3支不同编号复管；样本SD / |均值| ×100'),('与Baseline均值CV','手动验证批次内，Baseline三复管均值与当前Run三复管均值两个值的样本SD / |两者平均| ×100；Baseline自身不计算。不同材料、管组不混合；不等同于方差分析批间精密度。'),('阈值%',threshold),('选择','自动计算全部当前样本管的所有原始数值指标（含Events）；全部管保留，Control不参与；未分批只算批内；不做ABS回算。')]:literal_row(note,pair)
        for sheet in wb:
            sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
            for c in sheet[1]:c.font=Font(bold=True);c.fill=PatternFill('solid',fgColor='E9EFF5')
            for i in range(1,sheet.max_column+1):sheet.column_dimensions[get_column_letter(i)].width=22
        for row in ws.iter_rows(min_row=2):
            row[8].number_format='0.00'
            if row[10].value!='达标':row[10].font=Font(color='B42318')
        add_validation_summary(wb,items,cv,selected_metrics,threshold,self.validation_config(panel))
        style_export_workbook(wb);wb.save(path);wb.close()
        with self.db:
            self.db.execute('INSERT OR IGNORE INTO exclusions VALUES(?)',(str(path),))
            self.audit(self.panel(panel)['project_id'],'导出验证CV',{'panel':panel,'ids':ids,'metrics':selected_metrics,'threshold':threshold,'destination':str(path)})
        return len(cv)

    def parameters(self, panel):
        return self.db.execute('SELECT * FROM beads WHERE panel_id=? ORDER BY run', (panel,)).fetchall()

    def snapshot(self, panel, run):
        if not txt(run): return (None,None,None)
        r = self.db.execute('SELECT count,volume,lot FROM beads WHERE panel_id=? AND run=?', (panel, run_number(run))).fetchone()
        if r:return tuple(r)
        review=self.beads_reviews(panel).get(run_number(run),{})
        counts=review.get('counts',[])
        return (counts[0],None,'') if len(counts)==1 and not review.get('conflict') else (None,None,None)

    def result_slot(self, panel, payload):
        source_key = logical_key(payload)
        row = self.db.execute('SELECT logical FROM slots WHERE panel_id=? AND source_key=?', (panel, source_key)).fetchone()
        return row[0] if row else source_key

    def _insert_record(self, panel, payload, state, reason, logical=None, supersedes=None, pf='', snap=None):
        snap = self.snapshot(panel, payload['meta']['Run number']) if snap is None else snap
        return self.db.execute('''INSERT INTO records(panel_id,logical,fingerprint,payload,state,pf,beads,volume,lot,created,reason,supersedes)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''', (panel, logical or logical_key(payload), digest_payload(payload), dumps(payload), state, pf,
            *snap, now(), reason, supersedes)).lastrowid

    def editable_record_fields(self, ident):
        rec=self.record(ident); p=json.loads(rec['payload']); cfg=self.config(rec['panel_id'])
        fields=[]
        for field in ['Subject ID','Cycle','Visit','Timepoint','Analyst','Date','Time']:
            if p.get('sample_kind')=='validation' and field=='Subject ID': continue
            fields.append({'area':'meta','key':field,'name':field,'kind':'text','value':p['meta'].get(field,'')})
        mapped=set(v for v in p.get('input_fields',{}).values() if v)
        aliases={header_token(x) for values in META_ALIASES.values() for x in values}
        by_source={}
        for c in cfg['columns']:
            by_source.setdefault(norm(c['source']),[]).append('events' if c['mode']=='abs' else c['kind'])
        bead=norm(cfg.get('beads_field',''))
        if bead: by_source.setdefault(bead,[]).append('events')
        for key,label in p['labels'].items():
            if key in mapped or header_token(label) in aliases or key not in p['raw']: continue
            kinds=by_source.get(key)
            if not kinds:
                try: numeric(p['raw'][key],label,infer_kind(label))
                except ValueError: continue
                kinds=[infer_kind(label)]
            fields.append({'area':'raw','key':key,'name':label,'kind':kinds[0],'kinds':kinds,'value':p['raw'][key]})
        return fields

    def correct_record(self, ident, area, key, value, reason):
        if not txt(reason): raise ValueError('请填写修改原因')
        self.backup()
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            rec=self.record(ident)
            if rec['state']!='current': raise ValueError('该记录已不是当前有效版本，请刷新后重选；待审核记录请先处理审核')
            descriptor=next((f for f in self.editable_record_fields(ident) if f['area']==area and f['key']==key),None)
            if descriptor is None: raise ValueError('此字段不能直接修改；样品编号、Run、管名等身份字段需核对来源后重新导入')
            payload=json.loads(rec['payload']); before=payload[area].get(key,''); pid=self.panel(rec['panel_id'])['project_id']
            if area=='raw':
                kinds=descriptor['kinds']
                if all(k=='text' for k in kinds): value=txt(value)
                else:
                    for kind in kinds:
                        if kind!='text': value=numeric(value,descriptor['name'],kind)
                    cfg=self.config(rec['panel_id'])
                    if key==norm(cfg.get('beads_field','')) and any(c['mode']=='abs' for c in cfg['columns']) and value<=0:
                        raise ValueError('用于ABS回算的Beads Events必须大于0')
            else:
                value=txt(value)
                if key=='Date': value=clean_date(value)
                if key=='Timepoint': value=canonical_time(value)
                if key=='Cycle': value=visit_tokens(value,'C')
                if key=='Visit': value=visit_tokens(value,'V')
                if value and key in ('Date','Time'):
                    formats=['%Y-%m-%d','%m/%d/%Y','%d/%m/%Y','%Y/%m/%d','%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M:%S'] if key=='Date' else ['%H:%M','%H:%M:%S','%H:%M:%S.%f']
                    for fmt in formats:
                        try: datetime.strptime(value,fmt); break
                        except ValueError: pass
                    else: raise ValueError('请输入有效日期（如2026-10-08）或时间（如10:30:00）')
            if before==value: return ident
            payload[area][key]=value
            if area=='meta':
                payload['metadata_alternatives']=[x for x in payload.get('metadata_alternatives',[]) if x['field']!=key]
                raw_key=payload.get('input_fields',{}).get(key)
                if raw_key is None:
                    matches=[k for k,label in payload['labels'].items() if header_token(label) in set(map(norm,META_ALIASES.get(key,[])))]
                    if len(matches)==1: raw_key=matches[0]
                if raw_key and raw_key in payload['raw']: payload['raw'][raw_key]=value
            edit={'area':area,'field':key,'label':descriptor['name'],'before':before,'after':value,'reason':txt(reason),'time':now(),'from_record':ident,
                  'operator':os.environ.get('USERNAME') or os.environ.get('USER') or ''}
            payload.setdefault('manual_edits',[]).append(edit)
            self.db.execute("UPDATE records SET state='superseded' WHERE id=?",(ident,))
            new_id=self._insert_record(rec['panel_id'],payload,'current','人工修改：'+txt(reason),rec['logical'],ident,'',(rec['beads'],rec['volume'],rec['lot']))
            self.db.execute('INSERT INTO sources(record_id,file,sha,sheet,line) SELECT ?,file,sha,sheet,line FROM sources WHERE record_id=?',(new_id,ident))
            self.audit(pid,'人工修改记录',{'panel':rec['panel_id'],'old_record':ident,'new_record':new_id,**edit,'previous_pf':rec['pf'],'pf_reset':True})
        return new_id

    def set_parameters(self, panel, runs, count, volume, lot, reason):
        runs = sorted(set(run_number(r) for r in runs))
        if not runs or not txt(reason):
            raise ValueError('Run和修改原因不能为空')
        count = numeric(count, '每管Beads总数', 'events')
        volume = numeric(volume, '血量µL')
        if count <= 0 or volume <= 0:
            raise ValueError('Beads总数和血量必须大于0')
        snap = (count, volume, txt(lot))
        self.backup()
        modified = 0
        with self.db:
            self.db.executemany('DELETE FROM beads_review WHERE panel_id=? AND run=?',[(panel,r) for r in runs])
            changes = []
            for run in runs:
                before = self.snapshot(panel, run)
                self.db.execute('INSERT INTO beads VALUES(?,?,?,?,?) ON CONFLICT(panel_id,run) DO UPDATE SET count=excluded.count,volume=excluded.volume,lot=excluded.lot', (panel, run, *snap))
                if before != snap:
                    changes.append({'run': run, 'before': before, 'after': snap})
            for rec in self.records(panel, ['current', 'pending']):
                payload = json.loads(rec['payload'])
                if not txt(payload['meta']['Run number']) or run_number(payload['meta']['Run number']) not in runs or (rec['beads'], rec['volume'], rec['lot']) == snap:
                    continue
                self.db.execute("UPDATE records SET state='superseded' WHERE id=?", (rec['id'],))
                new_id = self._insert_record(panel, payload, rec['state'], '参数修改：' + reason,
                    rec['logical'], rec['id'], rec['pf'], snap)
                self.db.execute('INSERT INTO sources(record_id,file,sha,sheet,line) SELECT ?,file,sha,sheet,line FROM sources WHERE record_id=?', (new_id, rec['id']))
                modified += 1
            self.audit(self.panel(panel)['project_id'], '批量更新Beads及参数版本', {'panel': panel, 'changes': changes, 'new_versions': modified, 'reason': reason})
        return modified

    def import_batch(self, panel, paths):
        """Commit each complete file; keep batch-new competing versions pending.

        Temporary pending states are safe if the application is interrupted.
        A bad file never commits a subset of its rows.
        """
        pid = self.panel(panel)['project_id']
        paths = list(dict.fromkeys(str(Path(p).resolve()) for p in paths))
        prior_keys = {r['logical'] for r in self.records(panel)}
        outcomes = []; new_ids = []; warnings = []
        excluded = {r[0] for r in self.db.execute('SELECT path FROM exclusions')}
        self.backup()
        progress=getattr(self,'progress',lambda message:None)
        panel_name=self.panel(panel)['name']
        for index,path in enumerate(paths,1):
            progress(f'{panel_name} · 文件 {index}/{len(paths)}：{Path(path).name}')
            if path in excluded:
                outcomes.append({'file':path,'status':'跳过','detail':'已登记为受试者资料或程序导出文件','ids':[],'duplicate':0})
                continue
            before = self.db.execute('SELECT COALESCE(MAX(id),0) FROM records').fetchone()[0]
            try:
                result = self.import_raw(panel, [path], _defer_new=True, _backup=False)
            except Exception as exc:
                outcomes.append({'file':path,'status':'失败','detail':str(exc),'ids':[],'duplicate':0})
                continue
            ids = [r[0] for r in self.db.execute('SELECT id FROM records WHERE panel_id=? AND id>?',(panel,before))]
            new_ids.extend(ids); warnings.extend(result['warnings'])
            detail='原文件未变化，已跳过重复导入' if result.get('unchanged_files') else ('结果重复，未新增记录' if result['duplicate'] and not ids else '整份文件已处理')
            outcomes.append({'file':path,'status':'成功','detail':detail,'ids':ids,'duplicate':result['duplicate']})
        progress(f'{panel_name} · 文件读取完成，正在整理重复记录和待审核项…')
        with self.db:
            groups = {}
            for ident in new_ids:
                rec = self.record(ident); groups.setdefault(rec['logical'],[]).append(rec)
            for key, records in groups.items():
                if key not in prior_keys and len(records) == 1:
                    self.db.execute("UPDATE records SET state='current' WHERE id=? AND state='pending'",(records[0]['id'],))
            states = {ident:self.record(ident)['state'] for ident in new_ids}
            for outcome in outcomes:
                for state in ('current','pending'):
                    outcome[state] = sum(states.get(i)==state for i in outcome['ids'])
            stats = {'current':sum(s=='current' for s in states.values()),
                     'pending':sum(s=='pending' for s in states.values()),
                     'duplicate':sum(o['duplicate'] for o in outcomes),
                     'warnings':warnings, 'files':outcomes,
                     'errors':[o for o in outcomes if o['status']=='失败']}
            self.audit(pid,'批量导入结果',{'panel':panel,**stats})
        return stats

    def unchanged_source_rows(self, panel, sha, options):
        rows=self.db.execute('SELECT r.*,s.sheet,s.line FROM sources s JOIN records r ON r.id=s.record_id WHERE r.panel_id=? AND s.sha=? ORDER BY r.id DESC,s.id DESC',(panel,sha)).fetchall()
        if not rows:return []
        selected={}
        for row in rows:
            p=json.loads(row['payload']);key=(row['sheet'],row['line'])
            if key in selected:continue
            original=p.get('import_options_original')
            if original is not None:
                for field in ('mapping','tube_default','source_panel'):
                    default={} if field=='mapping' else ''
                    if original.get(field,default)!=options.get(field,default):return []
            else:
                # Existing V2/V3 records: verify effective field mapping, not reference-derived metadata.
                try:fields=resolve_metadata(list(p['labels'].values()),options)
                except ValueError:return []
                if fields!=p.get('input_fields'):return []
                if p.get('manual_tube') and txt(options.get('tube_default'))!=txt(p['meta']['Tube']):return []
            run=p['meta'].get('Run number')
            fields=p.get('input_fields',{})
            specimen=p['raw'].get(fields.get('Specimen') or norm('Specimen'))
            explicit_run=p['raw'].get(fields.get('Run number') or norm('Run number'))
            try:confirmed=confirmed_validation_run(txt(specimen),explicit_run,options)
            except ValueError:return []
            if confirmed is not None and (not run or run_number(run)!=confirmed):return []
            registered=bool(run and run_number(run) in options.get('validation_runs',[]))
            if registered and p.get('sample_kind')!='validation':return []
            if original is not None and run:
                was_registered=run_number(run) in original.get('validation_runs',[])
                if was_registered!=registered:return []
            expected=numbered_panel(options.get('source_panel') or self.panel(panel)['name'])
            actual=numbered_panel(p['meta'].get('Source Panel'))
            if expected and actual and expected!=actual:return []
            selected[key]=row
        return list(selected.values())

    def import_raw(self, panel, paths, _defer_new=False, _backup=True):
        pid = self.panel(panel)['project_id']
        cfg=self.config(panel);options = cfg.get('import_options', {})
        expected_panel = options.get('source_panel') or self.panel(panel)['name']
        expected_number = numbered_panel(expected_panel)
        expected_panel = f'Panel{expected_number}' if expected_number else ''
        references = {r['样品编号']: r for r in self.subject_rows(pid)}
        prepared = [];unchanged=[];unchanged_files=[]
        excluded = {r[0] for r in self.db.execute('SELECT path FROM exclusions')}
        for path in paths:
            path = Path(path).resolve()
            if str(path) in excluded:
                continue
            file_panels = filename_panels(path.name)
            if len(file_panels)==1 and expected_number and file_panels[0] != expected_number:
                raise ValueError(f'{path.name}文件名来源是Panel{file_panels[0]}，与目标{expected_panel}不一致')
            before = hashlib.sha256(path.read_bytes()).hexdigest()
            known=self.unchanged_source_rows(panel,before,{**options,'_file_runs':runs_in_name(path.stem)})
            if known:
                if hashlib.sha256(path.read_bytes()).hexdigest()!=before:raise ValueError(f'{path.name}读取期间发生变化，请关闭Excel重试')
                unchanged.extend((r['id'],str(path),before,r['sheet'],r['line']) for r in known)
                unchanged_files.append(str(path));continue
            found = False
            for sheet, line, labels, raw, fields in iter_raw_tables(path, options):
                found = True
                try:
                    payload = parse_raw(raw, labels, references, fields, {**options,'_file_runs':runs_in_name(path.stem),
                        '_beads_field':cfg['beads_field'],'_text_fields':[norm(c['source']) for c in cfg['columns'] if c['mode']=='raw' and c['kind']=='text']})
                    source_panel = payload['meta'].get('Source Panel')
                    if source_panel and len(file_panels)==1 and numbered_panel(source_panel)!=file_panels[0]:
                        raise ValueError('文件名Panel与表内Experiment的Panel不一致，请核对源文件')
                    if source_panel and expected_panel and norm(source_panel) != norm(expected_panel):
                        raise ValueError(f'来源是{source_panel}，当前配置为{expected_panel}；请切换到正确Panel再导入')
                    prepared.append((payload, str(path), before, sheet, line))
                except ValueError as exc:
                    raise ValueError(f'{path.name}/{sheet}/行{line}：{exc}') from exc
            if not found:
                raise ValueError(f'{path.name}未找到可导入表头；请检查“导入字段配置”或确认未选到汇总表')
            if hashlib.sha256(path.read_bytes()).hexdigest() != before:
                raise ValueError(f'{path.name}读取期间发生变化，请关闭Excel重试')
        if not prepared and not unchanged:
            raise ValueError('所选文件没有可导入原始行')
        if _backup and prepared: self.backup()
        stats = {'current': 0, 'pending': 0, 'duplicate': len(unchanged), 'warnings': [],'unchanged_files':unchanged_files}
        if not prepared:
            with self.db:
                self.db.executemany('INSERT OR IGNORE INTO sources(record_id,file,sha,sheet,line) VALUES(?,?,?,?,?)',unchanged)
                self.audit(pid,'原文件重复导入跳过',{'panel':panel,'files':unchanged_files,'duplicate':len(unchanged)})
            return stats
        with self.db:
            expected = self.db.execute('SELECT experiment FROM projects WHERE id=?', (pid,)).fetchone()[0]
            prefixes = {project_code(p) for p, *_ in prepared if project_code(p)}
            project_name=self.db.execute('SELECT name FROM projects WHERE id=?',(pid,)).fetchone()[0]
            named_codes=set(re.findall(r'(?<!\d)\d{6}(?!\d)',project_name))
            if len(named_codes)==1: prefixes.update(named_codes)
            for old in self.db.execute('SELECT r.payload FROM records r JOIN panels p ON p.id=r.panel_id WHERE p.project_id=? AND p.archived=0',(pid,)):
                code = project_code(json.loads(old[0]))
                if code: prefixes.add(code)
            if len(prefixes) > 1: raise ValueError('样品编号的六位项目号不一致；不同项目请分别新建')
            expected = expected or prepared[0][0]['meta']['Experiment']
            self.db.execute('UPDATE projects SET experiment=? WHERE id=?', (expected, pid))
            grouped = {}
            for p, *_ in prepared:
                grouped.setdefault(self.result_slot(panel,p),set()).add(digest_payload(p))
            self.detect_beads_names(panel,[p for p,*_ in prepared])
            self.db.executemany('INSERT OR IGNORE INTO sources(record_id,file,sha,sheet,line) VALUES(?,?,?,?,?)',unchanged)
            for p, path, sha, sheet, line in prepared:
                stats['warnings'].extend(f'{Path(path).name}/{sheet}/行{line}：{w}' for w in import_warnings(p))
                fp = digest_payload(p)
                same = self.db.execute('SELECT id FROM records WHERE panel_id=? AND fingerprint=? ORDER BY id DESC LIMIT 1', (panel, fp)).fetchone()
                if same:
                    ident = same[0]; stats['duplicate'] += 1
                else:
                    key = self.result_slot(panel, p)
                    old = self.db.execute('SELECT id FROM records WHERE panel_id=? AND logical=? LIMIT 1', (panel, key)).fetchone()
                    state = 'pending' if _defer_new or old or len(grouped[key]) > 1 else 'current'
                    ident = self._insert_record(panel, p, state, '原始Excel导入', logical=key)
                    self.db.execute('INSERT OR IGNORE INTO slots VALUES(?,?,?)', (panel, logical_key(p), key))
                    stats[state] += 1
                self.db.execute('INSERT OR IGNORE INTO sources(record_id,file,sha,sheet,line) VALUES(?,?,?,?,?)', (ident, path, sha, sheet, line))
            self.audit(pid, '导入原始数据', {'panel': panel, 'files': [str(p) for p in paths], **stats})
        return stats

    def batch_review_plan(self, panel, ids, action):
        self.panel(panel)
        if action not in ('adopt','reject'):raise ValueError('批量处理方式错误')
        ids=list(dict.fromkeys(ids))
        if not ids:raise ValueError('请先选择待审核记录；可点“全选待审核”')
        records=[self.record(i) for i in ids]
        if any(r['panel_id']!=panel for r in records):raise ValueError('批量审核仅限当前Panel')
        groups={}
        for r in records:
            if r['state']=='pending':groups.setdefault(r['logical'],[]).append(r['id'])
        plan=[]
        for rec in records:
            p=json.loads(rec['payload']);m=p['meta'];old=None;detail='';changes=[]
            eligible=rec['state']=='pending'
            if not eligible:detail='状态已变化，跳过'
            elif action=='reject':detail='保留原有效结果，不采用此候选'
            elif len(groups[rec['logical']])>1:
                eligible=False;detail='同一样品同一管选了多个候选，请仅选一个；可逐条比较'
            else:
                old=self.db.execute("SELECT * FROM records WHERE panel_id=? AND logical=? AND state='current'",(panel,rec['logical'])).fetchone()
                if old:
                    op=json.loads(old['payload'])
                    try:
                        if distinct_replicates(p,op):raise ValueError('同Run不同rpt为复管，不允许替代')
                    except ValueError as exc:eligible=False;detail=str(exc)
                    for area,prefix in [('meta','信息：'),('raw','')]:
                        for key in dict.fromkeys(list(p.get(area,{}))+list(op.get(area,{}))):
                            left=p.get(area,{}).get(key);right=op.get(area,{}).get(key)
                            if left!=right:changes.append({'field':prefix+p['labels'].get(key,key),'old':right,'new':left})
                    for key,name in [('beads','每管Beads总数'),('volume','血量µL'),('lot','Beads批号')]:
                        if rec[key]!=old[key]:changes.append({'field':name,'old':old[key],'new':rec[key]})
                    if eligible:detail=f'替代有效记录{old["id"]}；{len(changes)}处差异'
                else:detail='采用为有效结果；此前无同管有效记录'
            sources=self.db.execute('SELECT DISTINCT file FROM sources WHERE record_id=?',(rec['id'],)).fetchall()
            plan.append({'id':rec['id'],'old_id':old['id'] if old else None,'eligible':eligible,'detail':detail,
                'bas':sample_display(p),'tube':m['Tube'],'run':m['Run number'],'changes':changes,
                'files':[r[0] for r in sources]})
        return plan

    def batch_review(self, panel, ids, action, reason, old_fail=False, expected=None):
        if not txt(reason):raise ValueError('请填写批量处理原因')
        self.batch_review_plan(panel,ids,action)
        self.backup()
        with self.db:
            self.db.execute('BEGIN IMMEDIATE')
            plan=self.batch_review_plan(panel,ids,action)
            for item in plan:
                if expected is not None and item['old_id']!=expected.get(item['id']):
                    item['eligible']=False;item['detail']='预览后有效记录已变化，请刷新比较后处理'
                if not item['eligible']:continue
                ident=item['id'];old_id=item['old_id']
                if action=='reject':
                    self.db.execute("UPDATE records SET state='rejected',reason=? WHERE id=?",(reason,ident))
                    self.audit(self.panel(panel)['project_id'],'不采用候选结果',{'record':ident,'reason':reason,'batch':True})
                else:
                    if old_id:
                        if old_fail:self.db.execute("UPDATE records SET state='superseded',pf='F' WHERE id=?",(old_id,))
                        else:self.db.execute("UPDATE records SET state='superseded' WHERE id=?",(old_id,))
                    self.db.execute("UPDATE records SET state='current',reason=?,supersedes=COALESCE(?,supersedes) WHERE id=?",(reason,old_id,ident))
                    self.audit(self.panel(panel)['project_id'],'采用重分析结果',{'record':ident,'target':old_id,'old_set_fail':old_fail,'reason':reason,'batch':True})
            result={'processed':sum(r['eligible'] for r in plan),'skipped':sum(not r['eligible'] for r in plan),'rows':plan,'action':action}
            self.audit(self.panel(panel)['project_id'],'批量重分析审核',{'panel':panel,'reason':reason,**result})
        return result

    def activate(self, ident, reason, target=None, old_fail=False):
        if not txt(reason):
            raise ValueError('请填写采用/替换原因')
        self.backup()
        with self.db:
            rec = self.record(ident)
            if rec['state'] not in ('pending', 'current'):
                raise ValueError('只能采用待审核记录，或把当前记录明确关联到另一旧管')
            payload = json.loads(rec['payload'])
            pid = self.panel(rec['panel_id'])['project_id']
            old = None
            if target is not None:
                old = self.record(target)
                if old['id'] == ident or old['state'] != 'current' or old['panel_id'] != rec['panel_id']:
                    raise ValueError('替代目标必须是本Panel另一条当前有效记录')
                if sample_identity(json.loads(old['payload'])) != sample_identity(payload):
                    raise ValueError('替代目标的完整样品编号必须一致')
                if distinct_replicates(payload, json.loads(old['payload'])):
                    raise ValueError('同Run的rpt1、rpt2等为复管，不能相互替代；应保留并计算CV')
                if (norm(json.loads(old['payload'])['meta']['Tube']) == 'isotype') != (norm(payload['meta']['Tube']) == 'isotype'):
                    raise ValueError('Isotype不能与样本管相互替代')
                key = old['logical']
            else:
                key = rec['logical']
                old = self.db.execute("SELECT * FROM records WHERE panel_id=? AND logical=? AND state='current' AND id<>?", (rec['panel_id'], key, ident)).fetchone()
                if rec['state'] == 'current' and old is None:
                    raise ValueError('该记录已为当前有效，请指定它需要替代的旧记录')
            if old:
                self.db.execute("UPDATE records SET state='superseded',pf=? WHERE id=?", ('F' if old_fail else old['pf'], old['id']))
            if rec['logical'] != key:
                # Remember explicit tube equivalence for future reanalyses and queued candidates.
                self.db.execute('UPDATE slots SET logical=? WHERE panel_id=? AND logical=?', (key, rec['panel_id'], rec['logical']))
                self.db.execute("UPDATE records SET logical=? WHERE panel_id=? AND logical=? AND state='pending'", (key, rec['panel_id'], rec['logical']))
                self.db.execute('INSERT INTO slots VALUES(?,?,?) ON CONFLICT(panel_id,source_key) DO UPDATE SET logical=excluded.logical', (rec['panel_id'], logical_key(payload), key))
            self.db.execute("UPDATE records SET state='current',logical=?,reason=?,supersedes=COALESCE(?,supersedes) WHERE id=?", (key, reason, old['id'] if old else None, ident))
            self.audit(pid, '采用重分析结果', {'record': ident, 'target': old['id'] if old else None, 'old_set_fail': old_fail, 'reason': reason})

    def reject(self, ident, reason):
        if not txt(reason):
            raise ValueError('请填写不采用原因')
        self.backup()
        with self.db:
            r = self.record(ident)
            if r['state'] != 'pending':
                raise ValueError('仅待审核记录可不采用')
            self.db.execute("UPDATE records SET state='rejected',reason=? WHERE id=?", (reason, ident))
            self.audit(self.panel(r['panel_id'])['project_id'], '不采用候选结果', {'record': ident, 'reason': reason})

    def set_pf(self, ident, pf, reason):
        if pf not in ('', 'P', 'F') or not txt(reason):
            raise ValueError('P/F只能为空、P、F，并填写原因')
        self.backup()
        with self.db:
            r = self.record(ident)
            self.db.execute('UPDATE records SET pf=? WHERE id=?', (pf, ident))
            self.audit(self.panel(r['panel_id'])['project_id'], '修改P/F', {'record': ident, 'before': r['pf'], 'after': pf, 'reason': reason})

    def export_subjects(self, pid, destination):
        destination=Path(destination).resolve()
        if destination.suffix.lower()!='.xlsx': raise ValueError('导出文件必须为.xlsx')
        if self.db.execute('SELECT 1 FROM sources WHERE file=? LIMIT 1',(str(destination),)).fetchone():
            raise ValueError('不能覆盖已导入的原始Excel')
        if self.db.execute('SELECT 1 FROM exclusions WHERE path=?',(str(destination),)).fetchone() and destination.exists():
            book=load_workbook(destination,read_only=True)
            try:
                if '_受试者库导出' not in book.sheetnames: raise ValueError('不能覆盖受试者来源Excel或其他已排除文件')
            finally: book.close()
        rows=self.subject_rows(pid); wb=Workbook();ws=wb.active;ws.title='受试者库'
        literal_row(ws,[SUBJECT_SPLIT_LABELS.get(h,h) for h in SUBJECT_SPLIT_VIEW])
        for row in rows:
            split=split_subject(row);literal_row(ws,[split.get(h,'') for h in SUBJECT_SPLIT_VIEW])
        issues=self.subject_import_issues(pid)
        if issues:
            problem=wb.create_sheet('待修正');literal_row(problem,[SUBJECT_SPLIT_LABELS.get(h,h) for h in SUBJECT_SPLIT_VIEW]+['问题说明'])
            for issue in issues:
                split=split_subject(issue['row']);literal_row(problem,[split.get(h,'') for h in SUBJECT_SPLIT_VIEW]+[issue['detail']])
        for sheet in wb:
            sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
            for cell in sheet[1]: cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='007AFF')
            for i in range(1,sheet.max_column+1): sheet.column_dimensions[get_column_letter(i)].width=20 if i!=sheet.max_column or sheet.title!='待修正' else 70
        marker=wb.create_sheet('_受试者库导出');marker.append(['版本',VERSION]);marker.sheet_state='hidden'
        fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=destination.parent);os.close(fd)
        try: style_export_workbook(wb);wb.save(tmp);os.replace(tmp,destination)
        finally:
            wb.close()
            if Path(tmp).exists():Path(tmp).unlink()
        with self.db:
            self.db.execute('INSERT OR IGNORE INTO exclusions(path) VALUES(?)',(str(destination),))
            self.audit(pid,'导出受试者库',{'file':str(destination),'rows':len(rows),'issues':len(issues)})
        return len(rows)

    def available_fields(self, panel, _records=None):
        fields = {}
        for r in (self.records(panel) if _records is None else _records):
            p = json.loads(r['payload'])
            for key, label in p['labels'].items():
                if key and key not in map(norm, RAW_META):
                    fields.setdefault(key, label)
        return list(fields.values())

    def parent_abs_options(self, panel):
        options=[]
        for parent in self.panels(self.panel(panel)['project_id']):
            if parent['id']==panel: continue
            for c in self.config(parent['id']).get('columns',[]):
                if c['mode']=='abs': options.append({'panel':parent['id'],'panel_name':parent['name'],'source':c['name'],
                    'label':parent['name']+' / '+c['name']})
        return options

    def parent_source_config(self, panel, column):
        parent=self.db.execute('SELECT * FROM panels WHERE id=?',(column['parent_panel'],)).fetchone()
        if parent is None: raise ValueError('父群Panel不存在，请在本项目重新选择ABS来源')
        if parent['archived']: raise ValueError('父群来源Panel已删除，请先恢复或更换来源')
        if parent['id']==panel or parent['project_id']!=self.panel(panel)['project_id']:
            raise ValueError('父群ABS必须来自同一项目的另一个Panel')
        cfg=self.config(parent['id'])
        source=next((c for c in cfg['columns'] if norm(c['name'])==norm(column['parent_source'])),None)
        if source is None or source['mode']!='abs':
            raise ValueError(parent['name']+'没有选定的Beads回算ABS汇报项：'+column['parent_source'])
        return parent,cfg,source

    def parent_candidates(self, panel, column, bas, _context=None):
        parent,cfg,source=self.parent_source_config(panel,column)
        result=[]
        candidates = self.records(parent['id'],['current']) if _context is None else _context.records_for_bas(parent['id'],'current',bas)
        for rec in candidates:
            p=json.loads(rec['payload']);m=p['meta']
            if m['BAS ID']!=bas or p.get('sample_kind')=='validation' or rec['pf']=='F': continue
            if re.match(r'(?i)^(FMO|Isotype)(?:\b|\s|$)',txt(m['Tube'])): continue
            if re.fullmatch(cfg['sample_regex'],txt(m['Tube'])): result.append(rec)
        return result

    def parent_abs_value(self, panel, column, child, _context=None):
        parent,cfg,source=self.parent_source_config(panel,column)
        m=child['meta'];bas=m['BAS ID']
        context = _context if _context is not None else ReportReadContext(self)
        if context.bas_has_conflict(self.panel(panel)['project_id'],bas):
            raise ValueError('该BAS资料冲突待修正，暂不进行跨Panel回算')
        if not bas or child.get('sample_kind')=='validation': raise ValueError('跨Panel回算需要正式样品的完整BAS ID，验证样品暂不跨Panel关联')
        for rec in context.records_for_bas(parent['id'],'pending',bas):
            p=json.loads(rec['payload'])
            if p['meta']['BAS ID']==bas and re.fullmatch(cfg['sample_regex'],txt(p['meta']['Tube'])):
                raise ValueError(parent['name']+'该样品有待审核的父群结果，请先处理重分析')
        candidates=self.parent_candidates(panel,column,bas,_context=context)
        by_id={r['id']:r for r in candidates}
        choice_key=bas+'|'+txt(m['Run number'])
        choices=column.get('parent_choices',{}).get(choice_key)
        if choices is not None:
            if any(i not in by_id for i in choices): raise ValueError('先前指定的父群来源已失效或为F，请重新选择来源管')
            selected=[by_id[i] for i in choices]
            if len({r['id'] for r in selected})!=len(selected): raise ValueError('父群来源记录重复')
        else:
            same_run=[r for r in candidates if txt(m['Run number']) and txt(json.loads(r['payload'])['meta']['Run number'])==txt(m['Run number'])]
            selected=same_run or candidates
        if not selected: raise ValueError(parent['name']+'未找到同BAS的有效父群ABS；请先导入计数Panel并配置Beads')
        if len(selected)>1:
            if column.get('parent_replicates','unique')!='mean':
                raise ValueError(parent['name']+'存在多条父群来源，请指定一管，或选择同Run rpt复管均值')
            payloads=[json.loads(r['payload']) for r in selected]
            runs={txt(p['meta']['Run number']) for p in payloads};tags=[replicate_tag(p['meta']['Tube']) for p in payloads]
            if len(runs)!=1 or '' in runs: raise ValueError('父群ABS均值只能用于同Run复管；不能跨Run合并，请指定来源管')
            if any(t is None for t in tags) or len({t[0] for t in tags})!=1 or len({t[1] for t in tags})!=len(tags):
                raise ValueError('父群均值要求同管组且编号不同的rpt复管')
        prepared=self.prepare_report(parent['id'],_rows=selected,_cfg={**cfg,'columns':[source]},_check_pending=False,_context=context)
        values=[];lineage=[]
        for rec,p,nums,calc,_ in prepared:
            check_optional_fields(m,p['meta'],bas)
            value=nums[0]
            if isinstance(value,tuple): value=calc[value[1]][6]
            value=numeric(value,parent['name']+' / '+source['name'])
            if value<0: raise ValueError('父群ABS不能为负数')
            values.append(value)
            lineage.append({'record_id':rec['id'],'run':p['meta']['Run number'],'tube':p['meta']['Tube'],'pf':rec['pf'],
                'abs':value,'beads':rec['beads'],'volume':rec['volume'],'lot':rec['lot'],'calculation':calc,
                'source_hash_values':p.get('hash_originals',{}),
                'sources':[dict(s) for s in self.db.execute('SELECT file,sha,sheet,line FROM sources WHERE record_id=?',(rec['id'],))]})
        return statistics.mean(values),{'mode':'parent_abs','panel_id':parent['id'],'panel':parent['name'],'source':source['name'],
            'method':'同Run复管均值' if len(values)>1 else '指定单管' if choices is not None else '同BAS匹配单管',
            'records':lineage,'child_bas':bas,'child_run':m['Run number']}

    def prepare_report(self, panel, exclude_fail=False, _rows=None, _cfg=None, _check_pending=True, _allow_partial=False, _context=None):
        cfg = self.config(panel) if _cfg is None else _cfg
        validate_config(cfg)
        rows = self.records(panel, ['current']) if _rows is None else list(_rows)
        pending = self.records(panel, ['pending'])
        if pending and _check_pending:
            raise ValueError(f'还有{len(pending)}条待审核结果。请先采用或不采用，再导出正式汇总')
        if exclude_fail: rows = [r for r in rows if r['pf'] != 'F']
        if not rows: raise ValueError('没有符合条件的当前有效结果')
        pid = self.panel(panel)['project_id']; prepared=[]; parent_cache={}
        context = _context if _context is not None else ReportReadContext(self)
        conflicts=context.conflicts(pid) if _allow_partial else {}
        refs=context.subjects(pid)
        beads_reviews=context.beads(panel)
        for rec in rows:
            p=json.loads(rec['payload']);m=p['meta'];raw=p['raw']
            validation=p.get('sample_kind')=='validation'
            sample=bool(re.fullmatch(cfg['sample_regex'],txt(m['Tube'])))
            control=bool(re.match(r'(?i)^(FMO|Isotype)(?:\b|\s|$)',txt(m['Tube'])))
            unknown=not validation and not sample and not control
            values=[];calc=[];p['report_issues']=[{'name':field+'资料核对','source':'','detail':'；'.join(messages)} for field,messages in conflicts.get(rec['id'],{}).items()]
            for c in cfg['columns']:
                # Validation runs are raw-only and never need a clinical BAS/parent source.
                if c['mode'] in ('abs','parent_abs') and (validation or control and not sample):
                    values.append(None);continue
                try:
                    if c['mode'] in ('abs','parent_abs') and unknown:
                        raise ValueError(f'管名 {m["Tube"]!r} 未列入参与ABS回算的管名；原始指标仍可导出')
                    key = norm(c['source'])
                    optional = any(header_token(c['source']) in set(map(norm,META_ALIASES[f])) for f in ['Experiment','Date','Time','Analyst'])
                    if c['mode'] == 'raw' and optional and not txt(raw.get(key)):
                        values.append(None); continue
                    if key not in raw:
                        raise ValueError(f'缺少来源表头 {c["source"]}')
                    v = raw[key]
                    if txt(v)=='#' and not (c['mode']=='raw' and c['kind']=='text') and key!=norm(cfg['beads_field']):
                        p.setdefault('hash_originals',{})[key]=v
                    if c['mode'] == 'raw':
                        parser=numeric if key==norm(cfg['beads_field']) else metric_number
                        values.append(txt(v) if c['kind'] == 'text' else parser(v, c['source'], c['kind']))
                    elif c['mode']=='parent_abs':
                        percentage=metric_number(v,c['source'],c['kind']); divisor=100 if c['kind']=='pct100' else 1
                        cache_key=dumps({'source':{k:c.get(k) for k in ('parent_panel','parent_source','parent_replicates','parent_choices')},
                            'sample':{k:m.get(k) for k in ('BAS ID','Run number','Subject ID','Cycle','Visit','Timepoint')}})
                        if cache_key not in parent_cache: parent_cache[cache_key]=self.parent_abs_value(panel,c,p,_context=context)
                        parent_value,detail=parent_cache[cache_key];detail=dict(detail)
                        answer=parent_value*percentage/divisor
                        if not math.isfinite(answer): raise ValueError('跨Panel ABS不是有限数值')
                        detail.update(percentage_source=c['source'],percentage_kind=c['kind'])
                        values.append(('abs',len(calc)))
                        calc.append([rec['id'],c['name'],percentage,divisor,parent_value,1,answer,detail['panel']+' / '+detail['source'],m['Run number'],detail])
                    else:
                        if not txt(m['Run number']): raise ValueError('缺少Run，无法选定Beads参数进行ABS回算；可先汇总原始指标')
                        review=beads_reviews.get(run_number(m['Run number']),{})
                        if review.get('conflict'):raise ValueError('样品名Beads总数冲突，请在Beads设置确认')
                        events = metric_number(v, c['source'], 'events')
                        denominator = numeric(raw.get(norm(cfg['beads_field'])), cfg['beads_field'], 'events')
                        if denominator <= 0:
                            raise ValueError('Beads Events必须大于0')
                        if rec['beads'] is None or rec['volume'] is None:
                            raise ValueError(f'{m["Run number"]}缺少Beads/血量设置')
                        answer = events / denominator * rec['beads'] / rec['volume']
                        if not math.isfinite(answer):
                            raise ValueError('ABS结果不是有限数值')
                        values.append(('abs', len(calc)))
                        calc.append([rec['id'], c['name'], events, denominator, rec['beads'], rec['volume'], answer, rec['lot'] or '', m['Run number']])
                except ValueError as exc:
                    if not _allow_partial:
                        raise ValueError(f'记录{rec["id"]} / {m["BAS ID"]} / {m["Tube"]}：{exc}') from exc
                    values.append(None)
                    p['report_issues'].append({'name':c['name'],'source':c['source'],'detail':str(exc)})
            try: verified=self.check_subject(pid,p,False,_refs=refs)
            except ValueError as exc:
                if not _allow_partial: raise ValueError(f'记录{rec["id"]} / {m["BAS ID"]} / {m["Tube"]}：{exc}') from exc
                verified=False;p['report_issues'].append({'name':'受试者资料核对','source':'','detail':str(exc)})
            prepared.append((rec,p,values,calc,verified))
        return prepared


    def export(self, panel, destination, exclude_fail=False):
        destination = Path(destination).resolve()
        if destination.suffix.lower() != '.xlsx':
            raise ValueError('导出文件必须为.xlsx')
        cfg = self.config(panel)
        rows = self.records(panel, ['current'])
        if exclude_fail:
            rows = [r for r in rows if r['pf'] != 'F']
        if self.db.execute('SELECT 1 FROM sources WHERE file=? LIMIT 1', (str(destination),)).fetchone():
            raise ValueError('不能覆盖已导入的原始Excel')
        if self.db.execute('SELECT 1 FROM exclusions WHERE path=?', (str(destination),)).fetchone() and destination.exists():
            # Distinguish prior exports from reference input by marker.
            book = load_workbook(destination, read_only=True)
            try:
                if '_V2说明' not in book.sheetnames:
                    raise ValueError('不能覆盖受试者来源Excel')
            finally:
                book.close()
        panel_row = self.panel(panel); pid = panel_row['project_id']
        project = self.db.execute('SELECT * FROM projects WHERE id=?', (pid,)).fetchone()
        all_records = self.records(panel)
        sources = {r['id']: [] for r in all_records}
        for source in self.db.execute('SELECT s.* FROM sources s JOIN records r ON r.id=s.record_id WHERE r.panel_id=? ORDER BY s.record_id,s.file,s.sha,s.sheet,s.line', (panel,)):
            sources[source['record_id']].append(source)
        context = ReportReadContext(self)
        prepared = self.prepare_report(panel, exclude_fail,_check_pending=False,_allow_partial=True,_context=context)
        prepared = [entry for entry in prepared if entry[1]['meta'].get('BAS ID')]
        if not prepared: raise ValueError('没有可导出的正式样品，请到验证样品页导出验证留档')
        cv_prepared = self.prepare_report(panel, False,_check_pending=False,_allow_partial=True,_context=context) if cfg.get('replicate_cv', True) else []
        conflict_ids=set(context.conflicts(pid))
        cv_prepared=[r for r in cv_prepared if r[0]['id'] not in conflict_ids]
        cv_data = replicate_cv_rows([r for r in cv_prepared if r[1].get('sample_kind') != 'validation'], cfg, panel_row['name'],_allow_partial=True)
        validation_cv = []  # Validation archive is exported from its dedicated page.
        wb = Workbook(); ws = wb.active; ws.title = 'Summary'
        ws.append(META + [c['name'] for c in cfg['columns']])
        summary_ws = ws
        validation_ws = None
        validation_columns=[i for i,c in enumerate(cfg['columns']) if c['mode']=='raw']
        if any(r[1].get('sample_kind') == 'validation' for r in prepared):
            validation_ws = wb.create_sheet('验证样品')
            validation_ws.append(META + [cfg['columns'][i]['name'] for i in validation_columns] + ['样品类型','验证样品名称','原始Specimen'])
        report_issues=wb.create_sheet('汇总问题')
        report_issues.append(['记录ID','样品编号或验证名称','Tube','Run','汇报项','来源表头','处理','问题说明'])
        for rec,p,_,_,_ in prepared:
            for item in p.get('report_issues',[]):
                literal_row(report_issues,[rec['id'],sample_display(p),p['meta']['Tube'],p['meta']['Run number'],item['name'],item['source'],'本BAS待修正，未纳入Summary' if rec['id'] in conflict_ids else '该项留空，其他项正常导出',item['detail']])
        for rec in self.records(panel,['pending']):
            p=json.loads(rec['payload']);m=p['meta']
            literal_row(report_issues,[rec['id'],sample_display(p),m['Tube'],m['Run number'],'待审核候选','','未采用，保留当前有效版本','候选仍在待审核；本次Summary不自动替换旧结果'])
        for row in cv_data+validation_cv:
            if row[16]=='复管编号重复，未计算CV':literal_row(report_issues,[row[11],row[6] or row[19],row[7],row[5],'复管CV','','仅跳过此组CV，原始指标照常导出',row[16]])
        missing_ws=wb.create_sheet('批次文件提示')
        literal_row(missing_ws,['文件状态','目标Panel','文件路径','未纳入原因（项目范围）'])
        for f in self.unresolved_imports(pid):
            literal_row(missing_ws,[f['status'],f.get('panel_name') or f.get('panel_id') or '待分配',f['file'],f['detail']])
        calc_ws = wb.create_sheet('计算明细')
        calc_ws.append(['记录ID', '汇报项', '细胞Events', 'Beads Events', '每管Beads总数', '血量µL', 'ABS（cells/µL）', 'Beads批号', 'Run number'])
        parent_ws=wb.create_sheet('跨Panel回算')
        parent_ws.append(['子群记录ID','子群汇报项','百分比来源','百分比原值','比例类型','父群Panel','父群ABS项','父群ABS值','百分比分母','子群ABS（cells/µL）','父群记录ID','父群取管方式','父群来源与计算JSON','BAS ID','子群Run'])
        source_ws = wb.create_sheet('来源')
        source_ws.append(['输出行号', '记录ID', 'Panel', '来源文件', '文件SHA256', 'Sheet', '源行号', '受试者核对', '替代记录ID', '版本状态','输出工作表'])
        exported_ids = set()
        metadata_conflicts=context.conflicts(pid)
        for rec, p, values, calc, verified in prepared:
            if rec['id'] in metadata_conflicts:
                literal_row(report_issues,[rec['id'],sample_display(p),p['meta']['Tube'],p['meta']['Run number'],'资料冲突','','本BAS待修正，未纳入Summary','；'.join(f+':'+ '；'.join(v) for f,v in metadata_conflicts[rec['id']].items())])
                continue
            exported_ids.add(rec['id'])
            refs = []
            for line in calc:
                if len(line)>9 and line[9].get('mode')=='parent_abs':
                    info=line[9];n=parent_ws.max_row+1
                    literal_row(parent_ws,[line[0],line[1],info['percentage_source'],line[2],info['percentage_kind'],info['panel'],info['source'],line[4],line[3],line[6],
                        ', '.join(str(r['record_id']) for r in info['records']),info['method'],dumps(info['records']),info['child_bas'],info['child_run']])
                    parent_ws.cell(n,10,f'=D{n}/I{n}*H{n}').number_format='0.0'
                    refs.append("='跨Panel回算'!J"+str(n))
                else:
                    n = calc_ws.max_row + 1
                    literal_row(calc_ws, line)
                    calc_ws.cell(n, 7, f'=C{n}/D{n}*E{n}/F{n}').number_format = '0.0'
                    refs.append("='计算明细'!G"+str(n))
            is_validation = p.get('sample_kind') == 'validation'
            ws = validation_ws if is_validation else summary_ws
            m = p['meta']; n = ws.max_row + 1
            extra = ['验证样品',p['sample_name'],p['specimen_original']] if is_validation else []
            indices=validation_columns if is_validation else list(range(len(cfg['columns'])))
            display_columns=[cfg['columns'][i] for i in indices];display_values=[values[i] for i in indices]
            literal_row(ws, [rec['pf'] if h == 'P/F' else m[h] for h in META] + [None if isinstance(v, tuple) else v for v in display_values] + extra)
            for i, (c, v) in enumerate(zip(display_columns, display_values), len(META) + 1):
                cell = ws.cell(n, i)
                if isinstance(v, tuple):
                    cell.value = refs[v[1]]
                cell.number_format = '0.0%' if c['mode'] == 'raw' and c['kind'] == 'pct1' else ('0' if c['kind'] == 'events' and c['mode'] == 'raw' else '0.0')
                if c['kind'] == 'text' and c['mode'] == 'raw':
                    cell.number_format = '@'
            for s in sources[rec['id']]:
                checked='验证样品，只导原始指标' if is_validation else ('已核对' if verified else '未通过核对，见汇总问题' if any(i['name']=='受试者资料核对' for i in p['report_issues']) else '未提供受试者资料')
                literal_row(source_ws, [n, rec['id'], panel_row['name'], s['file'], s['sha'], s['sheet'], s['line'], checked+('；存在留空项，见汇总问题' if p['report_issues'] else '')+('；含人工修改，详见操作记录' if p.get('manual_edits') else ''), rec['supersedes'], STATE_LABEL[rec['state']],ws.title])
        hist = wb.create_sheet('历史版本')
        hist.append(['记录ID', '状态', 'P/F', 'BAS ID', 'Subject ID', 'Cycle', 'Visit', 'Timepoint', 'Tube', 'Run number', 'Date', 'Time', '每管Beads', '血量µL', 'Beads批号', '创建UTC', '原因', '替代记录ID','样品类型','验证样品名称'])
        raw_ws = wb.create_sheet('原始值')
        raw_ws.append(['记录ID', '来源表头', '原始值'])
        checks_ws = wb.create_sheet('导入检查'); checks_ws.append(['记录ID','状态','样品编号或验证名称','来源Panel','提示'])
        mapping_ws = wb.create_sheet('导入字段'); mapping_ws.append(['记录ID','标准字段','来源表头'])
        for rec in all_records:
            p = json.loads(rec['payload']); m = p['meta']
            for warning in import_warnings(p):
                literal_row(checks_ws,[rec['id'],STATE_LABEL[rec['state']],sample_display(p),m.get('Source Panel',''),warning])
            for field,key in p.get('input_fields',{}).items():
                literal_row(mapping_ws,[rec['id'],field,p['labels'].get(key,'未提供')])
            if rec['id'] not in exported_ids:
                for source in sources[rec['id']]:
                    literal_row(source_ws, [None, rec['id'], panel_row['name'], source['file'], source['sha'], source['sheet'], source['line'], '历史/未纳入Summary', rec['supersedes'], STATE_LABEL[rec['state']]])
            literal_row(hist, [rec['id'], STATE_LABEL[rec['state']], rec['pf']] + [m[h] for h in ['BAS ID','Subject ID','Cycle','Visit','Timepoint','Tube','Run number','Date','Time']] + [rec['beads'], rec['volume'], rec['lot'], rec['created'], rec['reason'], rec['supersedes'], '验证样品' if p.get('sample_kind') == 'validation' else '正式样品',p.get('sample_name','')])
            for key, value in p['raw'].items():
                original=p.get('hash_originals',{}).get(key) if value==0 else None
                literal_row(raw_ws, [rec['id'], p['labels'].get(key, key), original if original is not None else value])
        cv_ws = wb.create_sheet('复管CV')
        cv_ws.append(CV_HEADERS)
        for row in cv_data:
            literal_row(cv_ws, row)
            for column in [13, 14, 15, 16]:
                cv_ws.cell(cv_ws.max_row, column).number_format = '0.000' if column in (13, 14) else '0.0'
        if validation_cv:
            vcv_ws = wb.create_sheet('验证复管CV'); vcv_ws.append(CV_HEADERS)
            for row in validation_cv:
                literal_row(vcv_ws,row)
                for column in [13,14,15,16]: vcv_ws.cell(vcv_ws.max_row,column).number_format = '0.000' if column in (13,14) else '0.0'
        config_ws = wb.create_sheet('报告配置')
        config_ws.append(['输出名称', '来源表头', '方式', '数据类型','ABS回算来源','父群复管规则','父群取管配置'])
        for c in cfg['columns']:
            literal_row(config_ws, [c['name'], c['source'], c['mode'], c['kind'],abs_source_label(c),c.get('parent_replicates',''),dumps(c.get('parent_choices',{}))])
        param_ws = wb.create_sheet('Beads设置')
        param_ws.append(['Run number', '每管Beads', '血量µL', '批号'])
        for r in self.parameters(panel):
            literal_row(param_ws, ['Run'+str(r['run']), r['count'], r['volume'], r['lot']])
        reference_issues=self.subject_import_issues(pid)
        if reference_issues:
            issue_ws=wb.create_sheet('受试者清单问题'); issue_ws.append(['来源文件','Sheet','行号']+SUBJECT_VIEW+['问题说明'])
            for r in reference_issues: literal_row(issue_ws,[r['file'],r['sheet'],r['line']]+[r['row'].get(h,'') for h in SUBJECT_VIEW]+[r['detail']])
        subject_ws = wb.create_sheet('受试者信息'); subject_ws.append(SUBJECT_VIEW)
        for r in self.subject_rows(pid):
            literal_row(subject_ws, [r[h] for h in SUBJECT_VIEW])
        audit_ws = wb.create_sheet('操作记录'); audit_ws.append(['操作ID', 'UTC时间', '操作', '详情JSON'])
        files_ws = wb.create_sheet('文件导入结果')
        files_ws.append(['操作ID','UTC时间','文件','状态','新增有效','待审核','重复跳过','说明'])
        for r in self.db.execute('SELECT * FROM audit WHERE project_id=? ORDER BY id', (pid,)):
            literal_row(audit_ws, [r['id'], r['stamp'], r['action'], r['detail']])
            if r['action'] == '批量导入结果':
                detail = json.loads(r['detail'])
                if detail.get('panel') == panel:
                    for f in detail['files']:
                        literal_row(files_ws,[r['id'],r['stamp'],f['file'],f['status'],f['current'],f['pending'],f['duplicate'],f['detail']])
        note = wb.create_sheet('_V2说明'); note.append(['项目', '值'])
        for line in [('软件版本', VERSION), ('项目', project['name']), ('Experiment', project['experiment']), ('Panel', panel_row['name']),
                     ('导出UTC', now()), ('计数公式', '细胞Events / Beads Events × 每管Beads总数 / 血量µL'),
                     ('检测#约定','检测指标中的字面#在导入时自动按0保存并参与汇总、ABS、CV；原符号保留在原始值/追溯中。不是对分母为0的数学求值；Beads分母仍须为正数。'),
                     ('Summary导出','当前有效版本正常导出；缺表头、无父群来源、无Beads等只使对应项留空，原因记录在汇总问题；不填0、不改P/F。待审核候选不自动采用。'),
                     ('Beads Events字段', cfg['beads_field']), ('参与ABS回算的管名', cfg.get('abs_tube_names',tube_names_display(cfg))), ('排除F', str(exclude_fail)),
                     ('复管CV启用', str(cfg.get('replicate_cv', True))), ('CV阈值%', cfg.get('cv_threshold')),
                     ('低表达%parent界限',cfg.get('cv_low_pct_boundary')),('低表达CV限值%',cfg.get('cv_low_pct_threshold')),
                     ('低表达CV规则','按复管来源百分比均值判断，小于界限使用低表达CV限值；等于或高于界限用常规CV阈值。比例0–1先换成百分比。Events和MFI不套用百分比界限。'),
                     ('CV规则', '同Run/样品/Panel/管组的rpt复管；样本SD(n-1)/均值×100。达标定义CV≤阈值；不回写P/F。'),
                     ('CV取数', '使用所有当前有效复管，包括F；Summary排除F不改变CV取数。均值≤0或不足2管不计算CV。'),
                     ('P/F说明', '人工判断；已被替代不等于Fail。空白为尚未判定。'),
                     ('样品名解析', 'Cycle只提取C编号；Visit只提取V编号；W/D阶段归入Timepoint；前后统一Pre/Post，如D0 Post 7D；尾随Beads不参与计算。'),
                     ('EOI与pre', 'EOI保留结束参照；单独给药前统一Pre，pre4h或-4h统一Pre 4h；post24h或24h统一Post 24h。'),
                     ('时间点核对', '先统一关键信息再核对；缺信息不算冲突；相同Pre/Post但仅一边指定时长可合并。只有明确编号、阶段、方向或时长不一致才拦截。源文字保留。'),
                     ('样品身份', 'BAS ID为唯一样品编号；Panel和管名区分检测结果；重分析进入审核，不按分析人另建样品。'),
                     ('验证样品', '无完整样品编号且无显式BAS的任意非空样品名，按验证样品导入；也保留人工验证Run设置。BAS ID留空，单独导出原始指标，不导出ABS、不做跨Panel关联；同Run原始rpt指标可计算验证CV。'),
                     ('批量导入', '逐文件处理；失败文件整份不导入，其他文件继续。文件导入结果中保留历次成功、失败与跳过清单。'),
                     ('导入配置', dumps(cfg.get('import_options',{}))),
                     ('公式', 'Summary和计算明细含Excel公式；请用Excel打开重算。'),
                     ('跨Panel回算', '同项目按完整BAS ID关联：父群ABS × 子群占该父群的比例。优先同Run；多来源需指定管或明确同Run复管均值。排除F及对照/验证管；来源变更需重导子Panel Summary。'),
                     ('数据来源', '本机数据库；修改导出Excel不会回写数据库。')]:
            literal_row(note, line)
        for sheet in wb:
            sheet.freeze_panes = 'A2'; sheet.auto_filter.ref = sheet.dimensions
            sheet.row_dimensions[1].height = 36
            for cell in sheet[1]:
                cell.font = Font(name='Calibri', bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='17324D')
                cell.alignment = Alignment(wrap_text=True, vertical='center')
            for i in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(i)].width = min(42, max(18, len(txt(sheet.cell(1, i).value)) + 3))
        if reference_issues:
            for row in wb['受试者清单问题'].iter_rows(min_row=2):
                for cell in row:
                    cell.fill=PatternFill('solid',fgColor='FFF0F0'); cell.font=Font(color='A2373E')
        ws.column_dimensions['A'].width = 32
        source_ws.column_dimensions['D'].width = 55
        audit_ws.column_dimensions['D'].width = 90
        note.column_dimensions['B'].width = 95
        compact_qc_summary(self,wb,panel,prepared,exported_ids,cfg,exclude_fail,cv_prepared)
        apply_summary_precision(wb,cfg,prepared)
        capture_summary_results(wb,cfg,prepared)
        if cfg.get('replicate_cv',True):inline_sample_statistics(wb,cfg,prepared,exclude_fail)
        split_summary_visits(wb['Summary'])
        materialize_summary(wb)
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        destination.parent.mkdir(parents=True, exist_ok=True)
        original_hash = hashlib.sha256(destination.read_bytes()).hexdigest() if destination.exists() else None
        fd, path = tempfile.mkstemp(prefix='.express7_v2_', suffix='.xlsx', dir=destination.parent)
        os.close(fd)
        try:
            style_export_workbook(wb);wb.save(path)
            current_hash = hashlib.sha256(destination.read_bytes()).hexdigest() if destination.exists() else None
            if current_hash != original_hash:
                raise ValueError('输出文件在导出期间被修改，请关闭Excel重试')
            if destination.exists():
                import shutil
                back = destination.parent / (destination.stem + '_backups'); back.mkdir(exist_ok=True)
                shutil.copy2(destination, back / (datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.xlsx'))
            os.replace(path, destination)
        finally:
            Path(path).unlink(missing_ok=True); wb.close()
        with self.db:
            self.db.execute('INSERT OR IGNORE INTO exclusions VALUES(?)', (str(destination),))
            self.audit(pid, '导出Excel', {'panel': panel, 'destination': str(destination), 'rows': len(exported_ids), 'metadata_conflict_excluded':len(conflict_ids.intersection(r['id'] for r in rows)), 'exclude_fail': exclude_fail, 'config': cfg, 'summary_token':self.summary_token(panel)})
        return len(exported_ids)


def style_export_workbook(wb):
    """Apply typography without changing numeric formats, formulas or warning colors."""
    from copy import copy
    for sheet in wb:
        sheet.sheet_view.showGridLines = False
        for row in sheet:
            for cell in row:
                font=copy(cell.font);font.name='Arial';font.sz=10;cell.font=font
        sheet.row_dimensions[1].height=32
        for cell in sheet[1]:
            cell.font=Font(name='Arial',size=10,bold=True,color='24344D')
            cell.fill=PatternFill('solid',fgColor='E8EEF7')
            cell.alignment=Alignment(vertical='center',wrap_text=True)


def compact_qc_summary(store, wb, panel, prepared, exported_ids, cfg, exclude_fail, cv_prepared):
    """Project the full export into a compact QC workbook; database remains intact."""
    record_cache={};source_cache={}
    def record(ident):
        if ident not in record_cache:record_cache[ident]=store.record(ident)
        return record_cache[ident]
    def source_text(ident):
        if ident not in source_cache:
            source_cache[ident]='\n'.join(f'{r[0]} | {r[1]} | 行{r[2]}' for r in store.db.execute('SELECT file,sheet,line FROM sources WHERE record_id=? ORDER BY file,sheet,line',(ident,)))
        return source_cache[ident]
    # Preserve provenance on the result row instead of a separate source sheet.
    output_sources={}
    for row in list(wb['来源'].values)[1:]:
        if row[0] is not None and len(row)>10:
            output_sources[(row[10],row[0])]=row[1]
    for name in ('Summary','验证样品'):
        if name not in wb.sheetnames:continue
        sheet=wb[name];start=sheet.max_column+1
        for j,label in enumerate(['记录ID','来源文件 / Sheet / 行']):sheet.cell(1,start+j,label)
        for n in range(2,sheet.max_row+1):
            ident=output_sources.get((name,n))
            if ident:
                sheet.cell(n,start,ident);sheet.cell(n,start+1,source_text(ident)).data_type='s'
        sheet.column_dimensions[get_column_letter(start+1)].width=55
    old_direct={(r[0],r[1]):i for i,r in enumerate(list(wb['计算明细'].values)[1:],2)}
    old_parent={(r[0],r[1]):i for i,r in enumerate(list(wb['跨Panel回算'].values)[1:],2)}
    wb.remove(wb['计算明细']);calc=wb.create_sheet('计算明细',1)
    literal_row(calc,['记录ID','汇报项','细胞Events','Beads Events','每管Beads总数','血量µL','ABS（cells/µL）','Beads批号','Run','BAS ID','Tube','Panel','计算方式','子群比例原值','比例除数','父群ABS','父群记录ID','来源文件 / Sheet / 行','公式说明'])
    direct_rows={};new_rows={};related=set(exported_ids)
    def direct(line):
        key=(line[0],line[1])
        if key in direct_rows:return direct_rows[key]
        r=record(line[0]);p=json.loads(r['payload']);m=p['meta'];n=calc.max_row+1
        literal_row(calc,line[:9]+[m.get('BAS ID'),m.get('Tube'),store.panel(r['panel_id'])['name'],'Beads回算',None,None,None,None,source_text(r['id']),'细胞Events / Beads Events × Beads总数 / 血量µL'])
        calc.cell(n,7,f'=C{n}/D{n}*E{n}/F{n}');direct_rows[key]=n;related.add(r['id']);return n
    for rec,p,values,lines,verified in prepared:
        if rec['id'] not in exported_ids:continue
        for line in lines:
            key=(line[0],line[1])
            if len(line)<=9 or line[9].get('mode')!='parent_abs':new_rows[key]=direct(line);continue
            info=line[9];parents=[]
            for parent in info['records']:
                related.add(parent['record_id'])
                for parent_line in parent['calculation']:
                    if parent_line[1]==info['source']:parents.append(direct(parent_line))
            if len(parents)!=len(info['records']):raise ValueError('父群计算来源不完整，无法生成可核查的Excel公式')
            m=p['meta'];n=calc.max_row+1
            literal_row(calc,[rec['id'],line[1],None,None,None,None,line[6],None,m['Run number'],m['BAS ID'],m['Tube'],store.panel(panel)['name'],'跨Panel：'+info['method'],line[2],line[3],line[4],','.join(str(x['record_id']) for x in info['records']),source_text(rec['id']),f"{info['panel']} / {info['source']} × {info['percentage_source']} / 比例除数；百分比分母须为该父群"])
            calc.cell(n,16,'=AVERAGE('+','.join('G'+str(i) for i in parents)+')')
            calc.cell(n,7,f'=P{n}*N{n}/O{n}');new_rows[key]=n
    formula_map={}
    for key,n in old_direct.items():formula_map["='计算明细'!G"+str(n)]="='计算明细'!G"+str(new_rows[key])
    for key,n in old_parent.items():formula_map["='跨Panel回算'!J"+str(n)]="='计算明细'!G"+str(new_rows[key])
    for name in ('Summary','验证样品'):
        if name in wb.sheetnames:
            for row in wb[name]:
                for cell in row:
                    if cell.data_type=='f' and cell.value in formula_map:cell.value=formula_map[cell.value]
    for row in calc.iter_rows(min_row=2):
        for c in (row[6],row[15]):c.number_format='0.000'
    # Include only ancestors of the exported results and selected parent sources.
    history=wb.create_sheet('变更记录',2)
    literal_row(history,['时间（UTC）','类型','Panel','BAS ID','Run','Tube','旧记录ID','新记录ID','字段','修改前','修改后','原因','操作人（已记录时）'])
    audits=list(store.db.execute('SELECT * FROM audit WHERE project_id=? ORDER BY id',(store.panel(panel)['project_id'],)))
    adoptions={json.loads(a['detail']).get('record'):a for a in audits if a['action']=='采用重分析结果'}
    walked=set();todo=list(related)
    def value(v):return dumps(v) if isinstance(v,(dict,list)) else v
    while todo:
        ident=todo.pop()
        if ident in walked:continue
        walked.add(ident);r=record(ident);p=json.loads(r['payload']);m=p['meta']
        if not r['supersedes']:continue
        previous=record(r['supersedes']);before=json.loads(previous['payload']);todo.append(previous['id'])
        stamp=r['created'];reason=r['reason'];kind='版本变更';operator=''
        if ident in adoptions:stamp=adoptions[ident]['stamp'];kind='重分析采用'
        edits=[e for e in p.get('manual_edits',[]) if e.get('from_record')==previous['id']]
        if edits:stamp=edits[-1].get('time',stamp);operator=edits[-1].get('operator','');kind='人工修改'
        differences=[]
        for area in ('meta','raw'):
            left=before.get(area,{});right=p.get(area,{})
            for key in sorted(set(left)|set(right)):
                if left.get(key)!=right.get(key):differences.append((p.get('labels',{}).get(key,key) if area=='raw' else key,left.get(key),right.get(key)))
        for key,label in [('beads','Beads总数'),('volume','血量µL'),('lot','Beads批号')]:
            if previous[key]!=r[key]:differences.append((label,previous[key],r[key]))
        if not differences:differences=[('采用版本',previous['id'],ident)]
        for field,left,right in differences:literal_row(history,[stamp,kind,store.panel(r['panel_id'])['name'],m.get('BAS ID'),m.get('Run number'),m.get('Tube'),previous['id'],ident,field,value(left),value(right),reason,operator])
    for a in audits:
        if a['action']!='修改P/F':continue
        d=json.loads(a['detail']);ident=d.get('record')
        if ident not in walked:continue
        r=record(ident);m=json.loads(r['payload'])['meta']
        literal_row(history,[a['stamp'],'P/F修改',store.panel(r['panel_id'])['name'],m.get('BAS ID'),m.get('Run number'),m.get('Tube'),ident,ident,'P/F',d.get('before'),d.get('after'),d.get('reason'),''])
    changes=sorted(list(history.values)[1:],key=lambda row:(str(row[0]),str(row[7]),str(row[8])))
    history.delete_rows(2,history.max_row)
    for row in changes:
        row=list(row)
        try:row[0]=datetime.fromisoformat(str(row[0]).replace('Z','+00:00')).strftime('%Y-%m-%d %H:%M:%S')
        except ValueError:pass
        literal_row(history,row)
    # One actionable issue sheet, only when populated.
    issues=wb.create_sheet('待处理问题');literal_row(issues,['类别','记录ID','BAS ID或文件','Tube / Run','汇报项','处理','问题说明'])
    for row in list(wb['汇总问题'].values)[1:]:literal_row(issues,['结果核对',row[0],row[1],f'{row[2]} / {row[3]}',row[4],row[6],row[7]])
    for row in list(wb['批次文件提示'].values)[1:]:literal_row(issues,['批次文件',None,row[2],row[1],None,row[0],row[3]])
    if '受试者清单问题' in wb.sheetnames:
        for row in list(wb['受试者清单问题'].values)[1:]:literal_row(issues,['受试者清单',None,row[0],f'{row[1]} / 行{row[2]}',None,'待修正',row[-1]])
    for r in store.records(panel,['current']):
        if exclude_fail and r['pf']=='F':
            m=json.loads(r['payload'])['meta'];literal_row(issues,['排除F',r['id'],m.get('BAS ID'),f"{m.get('Tube')} / {m.get('Run number')}",'','未纳入Summary','用户选择排除F；复管CV仍沿用全部当前复管规则'])
    # Expose replicate inputs beside the existing SD/CV results.
    by_id={r['id']:(p,v,lines) for r,p,v,lines,ok in cv_prepared}
    for name in ('复管CV','验证复管CV'):
        if name not in wb.sheetnames:continue
        sheet=wb[name]
        for n in range(sheet.max_row,1,-1):
            if (sheet.cell(n,10).value or 0)<2:sheet.delete_rows(n)
        if sheet.max_row==1:wb.remove(sheet);continue
        start=max(21,sheet.max_column+1)
        for n in range(2,sheet.max_row+1):
            ids=[int(x) for x in str(sheet.cell(n,12).value or '').split(',') if x.strip().isdigit()]
            column=next((i for i,c in enumerate(cfg['columns']) if c['name']==sheet.cell(n,9).value),None)
            inputs=[]
            if column is not None:
                for ident in ids:
                    if ident not in by_id:break
                    p,values,lines=by_id[ident];v=values[column]
                    if isinstance(v,tuple):v=lines[v[1]][6]
                    if not isinstance(v,(int,float)):break
                    inputs.append(v)
            if len(inputs)!=len(ids):continue
            for j,v in enumerate(inputs):sheet.cell(1,start+j,f'复管值{j+1}');sheet.cell(n,start+j,v)
            if len(inputs)>=2 and sheet.cell(n,15).value is not None and statistics.mean(inputs)>0:
                span=f'{get_column_letter(start)}{n}:{get_column_letter(start+len(inputs)-1)}{n}'
                sheet.cell(n,13,'=AVERAGE('+span+')');sheet.cell(n,14,'=STDEV('+span+')');sheet.cell(n,15,f'=N{n}/M{n}*100')
    if issues.max_row==1:wb.remove(issues)
    keep={'Summary','计算明细','变更记录','复管CV','待处理问题','验证样品','验证复管CV','_V2说明'}
    for sheet in list(wb):
        if sheet.title not in keep:wb.remove(sheet)
    note=wb['_V2说明'];note.delete_rows(1,note.max_row)
    for row in [('项目','值'),('软件版本',VERSION),('导出UTC',now()),('说明','QC精简导出；完整原始数据及历史保存在数据库中。'),('公式','请用Excel打开以重新计算公式。'),('变更范围','本次汇报及引用父群的版本链；无操作人记录时留空。')]:literal_row(note,row)
    note.sheet_state='hidden'
    for sheet in wb:
        sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
        for i in range(1,sheet.max_column+1):sheet.column_dimensions[get_column_letter(i)].width=20
    for name in ('Summary','验证样品'):
        if name in wb.sheetnames:
            sheet=wb[name];sheet.column_dimensions['A'].width=32
            sheet.column_dimensions[get_column_letter(sheet.max_column)].width=55
            for column in sheet.iter_cols():
                if column[0].value in ('BAS ID','Subject ID'):
                    for cell in column[1:]:cell.number_format='@'
    calc.column_dimensions['R'].width=55;calc.column_dimensions['S'].width=60
    history.column_dimensions['A'].width=24;history.column_dimensions['J'].width=28;history.column_dimensions['K'].width=28;history.column_dimensions['L'].width=45
    if '待处理问题' in wb.sheetnames:wb['待处理问题'].column_dimensions['G'].width=65


def literal_row(ws, values):
    ws.append(values)
    for c in ws[ws.max_row]:
        if isinstance(c.value, str):
            # Only deliberately constructed formulas are later assigned as formulas.
            c.data_type = 's'


class ResultTree(ttk.Treeview):
    """Presentation only: alternating rows; preserve explicitly supplied tags."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._stripe_number=0
        self.bind('<Configure>',self.schedule_redraw,add='+')
        self.bind('<Map>',self.schedule_redraw,add='+')
        self.tag_configure('stripe_even',background='#F1F3F4')
        self.tag_configure('stripe_odd',background='#F1F3F4')
        self.tag_configure('warning',background='#FFF5DE',foreground='#825410')
        self.tag_configure('error',background='#FFF0F0',foreground='#A2373E')
        self.tag_configure('muted',foreground='#758491')

    def set_red_cells(self,cells):
        self.red_cells=cells;self.schedule_redraw()

    def schedule_redraw(self,*_):
        if getattr(self,'redraw_pending',None):return
        self.redraw_pending=self.after_idle(self.draw_red_cells)

    def draw_red_cells(self):
        self.redraw_pending=None
        for label in getattr(self,'red_labels',[]):label.destroy()
        self.red_labels=[]
        for (iid,col),message in getattr(self,'red_cells',{}).items():
            if not self.exists(iid):continue
            box=self.bbox(iid,str(col))
            if not box:continue
            x,y,w,h=box
            if x+w<=0 or x>=self.winfo_width():continue
            label=tk.Label(self,text=self.set(iid,str(col)),fg='#B42318',bg='#FEE4E2',anchor='w',padx=3,font='TkDefaultFont',cursor='hand2')
            label.place(x=x+1,y=y+1,width=max(1,w-2),height=max(1,h-2));self.red_labels.append(label)
            label.bind('<Button-1>',lambda e,item=iid:self.selection_set(item))
            label.bind('<Double-1>',lambda e,item=iid,c=col:self.red_cell_edit(item,c))
            label.bind('<MouseWheel>',lambda e:self.event_generate('<MouseWheel>',delta=e.delta))
            label.bind('<Button-4>',lambda e:self.yview_scroll(-3,'units'))
            label.bind('<Button-5>',lambda e:self.yview_scroll(3,'units'))

    def red_cell_edit(self,iid,col):
        callback=getattr(self,'red_edit',None)
        if callback:callback(iid,col)

    def insert(self, parent, index, iid=None, **kwargs):
        if not kwargs.get('tags'):
            values=kwargs.get('values',[])
            state=str(values[0]) if values else ''
            tag={'失败':'error','待审核':'warning','跳过':'muted'}.get(state)
            kwargs['tags']=(tag or ('stripe_even' if self._stripe_number%2==0 else 'stripe_odd'),)
        result=super().insert(parent,index,iid=iid,**kwargs)
        self._stripe_number+=1
        return result

    def delete(self, *items):
        super().delete(*items)
        if not self.get_children(): self._stripe_number=0


class PagedTable:
    """Keep all display rows in memory, but only mount one page in Tk."""
    def __init__(self, app, tree, run_column=None):
        self.app=app;self.tree=tree;self.run_column=run_column
        self.rows=[];self.filtered=[];self.page=0;self.context=None;self.red_cells={};self.pending=None
        self.query=tk.StringVar();self.run=tk.StringVar();self.size=tk.StringVar(value='200')
        self.only_issues=tk.BooleanVar(value=False);self.caption=tk.StringVar()
        bar=ttk.Frame(tree.master);bar.grid(row=2,column=0,columnspan=2,sticky='ew',pady=(6,0))
        ttk.Label(bar,text='搜索').pack(side='left')
        ttk.Entry(bar,textvariable=self.query,width=20).pack(side='left',padx=5)
        if run_column is not None:
            ttk.Label(bar,text='Run').pack(side='left')
            ttk.Entry(bar,textvariable=self.run,width=7).pack(side='left',padx=5)
        ttk.Button(bar,text='显示全部',command=self.reset).pack(side='left',padx=3)
        ttk.Button(bar,text='上一页',command=lambda:self.turn(-1)).pack(side='left',padx=3)
        ttk.Button(bar,text='下一页',command=lambda:self.turn(1)).pack(side='left',padx=3)
        ttk.Combobox(bar,textvariable=self.size,values=['100','200','500'],state='readonly',width=5).pack(side='left',padx=5)
        ttk.Label(bar,textvariable=self.caption).pack(side='left',padx=5)
        for v in (self.query,self.run,self.size):v.trace_add('write',self.schedule_filter)
        tree.page_view=self
        tree.bind('<Destroy>',self.close,add='+')

    def close(self,event):
        if event.widget==self.tree and self.pending:
            self.tree.after_cancel(self.pending);self.pending=None

    def schedule_filter(self,*_):
        if self.pending:self.tree.after_cancel(self.pending)
        self.pending=self.tree.after(250,self.apply_filter)

    def reset(self):
        self.query.set('');self.run.set('');self.only_issues.set(False);self.apply_filter()

    def show_issues(self):
        self.query.set('');self.run.set('');self.only_issues.set(True);self.apply_filter()

    def set_rows(self,rows,red_cells,context):
        selected=self.tree.selection()
        if self.context!=context:
            self.query.set('');self.run.set('');self.only_issues.set(False);self.page=0
        self.context=context;self.rows=rows;self.red_cells=red_cells
        # Precompute search text once per refresh, not for each key stroke.
        self.search_text={r['iid']:' '.join(txt(v) for v in r['values']).casefold() for r in rows}
        self.apply_filter(reset_page=False,selected=selected)

    def apply_filter(self,reset_page=True,selected=()):
        if self.pending:self.tree.after_cancel(self.pending);self.pending=None
        if reset_page:self.page=0
        words=self.query.get().casefold().split();run=re.sub(r'(?i)^run\s*','',self.run.get().strip())
        issues={iid for iid,col in self.red_cells}
        self.filtered=[r for r in self.rows if
            all(w in self.search_text.get(r['iid'],'') for w in words)
            and (not self.only_issues.get() or r['iid'] in issues)
            and (not run or self.run_column is None or re.sub(r'(?i)^run\s*','',txt(r['values'][self.run_column]))==run)]
        self.sort_rows();self.render(selected)

    def sort_rows(self):
        if not hasattr(self.tree,'sort_state'):return
        col,descending=self.tree.sort_state
        filled=[r for r in self.filtered if txt(r['values'][col])]
        empty=[r for r in self.filtered if not txt(r['values'][col])]
        filled.sort(key=lambda r:natural_sort_key(r['values'][col]),reverse=descending)
        self.filtered=filled+empty

    def turn(self,delta):
        self.page+=delta;self.render()

    def render(self,selected=()):
        self.app.close_pf_dropdown()
        size=int(self.size.get());pages=max(1,(len(self.filtered)+size-1)//size)
        self.page=max(0,min(self.page,pages-1))
        self.tree.delete(*self.tree.get_children())
        shown=self.filtered[self.page*size:(self.page+1)*size]
        for row in shown:self.tree.insert('','end',**row)
        visible={r['iid'] for r in shown}
        self.tree.set_red_cells({k:v for k,v in self.red_cells.items() if k[0] in visible})
        retained=[iid for iid in selected if iid in visible]
        if retained:self.tree.selection_set(retained)
        scope='异常筛选 · ' if self.only_issues.get() else ''
        self.caption.set(f'{scope}{len(self.filtered)}/{len(self.rows)} 条 · 第 {self.page+1}/{pages} 页')


class StartupCancelled(Exception):
    pass


class App(tk.Tk):
    def __init__(self, db_path=None):
        super().__init__()
        self.withdraw()
        self.title('流式数据汇总工具 V' + VERSION)
        screen_w=self.winfo_screenwidth(); screen_h=self.winfo_screenheight()
        width=min(1420,max(1000,screen_w-48)); height=min(900,max(640,screen_h-90))
        self.geometry(f'{width}x{height}')
        self.minsize(min(1180,width),min(680,height))
        self.configure(bg='#F5F5F5')
        self.apply_theme()
        self.pid = None; self.panel_id = None; self.cfg = empty_config(); self.mark_config_dirty(False)
        chosen=db_path or self.choose_database_path()
        if not chosen:
            self.destroy();raise StartupCancelled()
        self.store = Store(chosen)
        self.protocol('WM_DELETE_WINDOW', self.close_app)
        self.build()
        self.bind('<Control-s>',self.save_config_shortcut)
        self.bind('<Control-S>',self.save_config_shortcut)
        self.refresh_projects()
        self.deiconify()
        self.report_callback_exception = self.on_error

    def choose_database_path(self):
        choice={'path':None}
        win=tk.Toplevel(self);win.title('选择数据库 · V'+VERSION);win.geometry('520x250');win.resizable(False,False)
        ttk.Label(win,text='先选择本次使用的数据库',style='Section.TLabel').pack(anchor='w',padx=22,pady=(22,10))
        ttk.Label(win,text='打开原V2数据库可保留项目、数据和配置；\n新建数据库由你选择保存位置。',style='Muted.TLabel').pack(anchor='w',padx=22,pady=(0,16))
        def pick(new=False):
            options={'parent':win,'filetypes':[('SQLite数据库','*.sqlite3 *.db'),('所有文件','*.*')]}
            if hasattr(self,'store'):options['initialdir']=str(self.store.path.parent)
            path=filedialog.asksaveasfilename(title='选择新数据库的保存位置',defaultextension='.sqlite3',initialfile='express7_databank.sqlite3',confirmoverwrite=False,**options) if new else filedialog.askopenfilename(title='打开已有数据库（可选择原V2数据库）',**options)
            if not path:return
            if new and Path(path).exists():
                messagebox.showinfo('文件已存在','这个数据库已经存在，请用“打开已有数据库”，或输入新的文件名。',parent=win);return
            choice['path']=path;win.destroy()
        bar=ttk.Frame(win);bar.pack(fill='x',padx=18,pady=4)
        ttk.Button(bar,text='打开已有数据库',style='Primary.TButton',command=lambda:pick(False)).pack(side='left',padx=4)
        ttk.Button(bar,text='新建数据库',command=lambda:pick(True)).pack(side='left',padx=4)
        ttk.Button(bar,text='取消',command=win.destroy).pack(side='left',padx=4)
        ttk.Label(win,text='每次启动都会选择，不自动保存到AppData。',style='Muted.TLabel').pack(anchor='w',padx=22,pady=14)
        win.update_idletasks();win.geometry('+%d+%d'%(max(0,(self.winfo_screenwidth()-520)//2),max(20,(self.winfo_screenheight()-250)//2)))
        win.grab_set();self.wait_window(win)
        return choice['path']

    def apply_theme(self):
        ctk.set_appearance_mode('light')
        ctk.set_default_color_theme('blue')
        families=set(tkfont.families(self))
        if os.name!='nt' and 'CustomTkinter_shapes_font' not in families:
            ctk.DrawEngine.preferred_drawing_method='polygon_shapes'
        self.ui_font=next((f for f in ('Microsoft YaHei UI','Microsoft YaHei','Noto Sans CJK SC','PingFang SC','WenQuanYi Micro Hei','DejaVu Sans') if f in families),'TkDefaultFont')
        for name in ('TkDefaultFont','TkTextFont','TkMenuFont','TkHeadingFont'):
            tkfont.nametofont(name).configure(family=self.ui_font,size=10)
        self.option_add('*Toplevel.background','#FFFFFF')
        self.option_add('*Text.background','#FFFFFF'); self.option_add('*Text.foreground','#263D4A')
        self.option_add('*Text.font',(self.ui_font,10)); self.option_add('*Text.relief','flat')
        self.option_add('*Text.highlightThickness',1); self.option_add('*Text.highlightBackground','#E5E5EA')
        self.option_add('*Text.selectBackground','#DCEAFF')
        self.option_add('*TCombobox*Listbox.font',(self.ui_font,10))
        self.option_add('*TCombobox*Listbox.background','#FFFFFF')
        self.option_add('*TCombobox*Listbox.selectBackground','#DCEAFF')
        self.option_add('*TCombobox*Listbox.selectForeground','#005AC8')
        style=ttk.Style(self); style.theme_use('clam'); self.ui_style=style
        style.configure('.',font=(self.ui_font,10),background='#FFFFFF',foreground='#1C1C1E')
        style.configure('TFrame',background='#FFFFFF')
        style.configure('Workspace.TFrame',background='#F5F5F5')
        style.configure('Sidebar.TFrame',background='#F5F5F5')
        style.configure('Toolbar.TFrame',background='#E3E3E3')
        style.configure('Toolbar.TLabel',background='#E3E3E3',foreground='#5F6368')
        style.configure('Mac.TSeparator',background='#E8EAED')
        style.configure('Notice.TFrame',background='#EDF6F4')
        style.configure('Card.TFrame',background='#FFFFFF',relief='solid',borderwidth=1,bordercolor='#E5E5EA')
        style.configure('TLabel',background='#FFFFFF',foreground='#1C1C1E')
        for name,size,color,background,weight in [
            ('Title',14,'#202124','#F5F5F5','normal'),('Subtitle',10,'#8E8E93','#FFFFFF','normal'),
            ('Muted',9,'#636366','#FFFFFF','normal'),('Section',11,'#1C1C1E','#FFFFFF','bold'),
            ('Notice',10,'#38615F','#EDF6F4','normal'),('Brand',12,'#202124','#F5F5F5','normal'),
            ('Sidebar',9,'#8E8E93','#F5F5F5','normal'),('Stat',20,'#1D1D1F','#FFFFFF','normal'),
            ('StatWarn',25,'#A46C18','#FFFFFF','bold'),('StatError',25,'#A44650','#FFFFFF','bold'),
            ('StatText',17,'#3A3A3C','#FFFFFF','normal'),('Footer',9,'#8E8E93','#F5F5F5','normal')]:
            style.configure(name+'.TLabel',font=(self.ui_font,size,weight),foreground=color,background=background)
        style.configure('TButton',padding=(12,7),background='#FFFFFF',foreground='#36515E',bordercolor='#D9E3E9',
                        lightcolor='#FFFFFF',darkcolor='#D9E3E9',borderwidth=1,relief='flat',focusthickness=1,focuscolor='#007AFF')
        style.map('TButton',background=[('disabled','#F3F6F8'),('pressed','#E0EFEE'),('active','#EFF5F7')],
                  foreground=[('disabled','#9CAAB4')],bordercolor=[('focus','#007AFF'),('active','#B4D2D5')])
        style.configure('Primary.TButton',background='#2B74CF',foreground='#FFFFFF',font=(self.ui_font,10,'bold'),bordercolor='#2B74CF',padding=(15,8))
        style.map('Primary.TButton',background=[('disabled','#E1E9EC'),('pressed','#205DAA'),('active','#2468BB')],
                  foreground=[('disabled','#9AAAB3'),('!disabled','#FFFFFF')],bordercolor=[('!disabled','#2B74CF')])
        style.configure('Nav.TButton',background='#F2F2F7',foreground='#B5CAD3',borderwidth=0,padding=(14,7),anchor='w',font=(self.ui_font,10))
        style.map('Nav.TButton',background=[('pressed','#244654'),('active','#1B3C49')],foreground=[('!disabled','#D0DFE5')])
        style.configure('ActiveNav.TButton',background='#24505A',foreground='#FFFFFF',borderwidth=0,padding=(14,7),anchor='w',font=(self.ui_font,10,'bold'))
        style.map('ActiveNav.TButton',background=[('!disabled','#24505A')],foreground=[('!disabled','#FFFFFF')])
        style.configure('TEntry',padding=5,fieldbackground='#FFFFFF',bordercolor='#D3D3D6',lightcolor='#FFFFFF',darkcolor='#D3D3D6')
        style.map('TEntry',bordercolor=[('focus','#007AFF')])
        style.configure('TCombobox',padding=5,fieldbackground='#FFFFFF',background='#F9F9FA',arrowcolor='#6E6E73',bordercolor='#D3D3D6')
        style.map('TCombobox',fieldbackground=[('readonly','#FFFFFF'),('disabled','#F3F6F8')],foreground=[('disabled','#9AAAB3'),('readonly','#1C1C1E')],bordercolor=[('focus','#007AFF')])
        style.configure('Step.TLabel',background='#EFF4F7',foreground='#738895',padding=(12,9),font=(self.ui_font,9))
        style.configure('StepActive.TLabel',background='#007AFF',foreground='#FFFFFF',padding=(12,9),font=(self.ui_font,9,'bold'))
        style.configure('StepDone.TLabel',background='#E6F3F0',foreground='#397D73',padding=(12,9),font=(self.ui_font,9))
        style.configure('TCheckbutton',background='#FFFFFF',foreground='#3A3A3C',padding=(2,4))
        style.map('TCheckbutton',background=[('active','#FFFFFF')])
        style.configure('Treeview',rowheight=34,background='#FFFFFF',fieldbackground='#FFFFFF',foreground='#202124',bordercolor='#E8EAED',borderwidth=0,relief='flat')
        style.configure('Treeview.Heading',font=(self.ui_font,9),padding=(10,9),background='#FFFFFF',foreground='#5F6368',relief='flat',borderwidth=0)
        style.map('Treeview.Heading',background=[('active','#E9E9EB')])
        style.map('Treeview',background=[('selected','#C2DBFF')],foreground=[('selected','#202124')])
        style.configure('TNotebook',background='#FFFFFF',borderwidth=0,tabmargins=(0,4,0,0))
        style.configure('TNotebook.Tab',padding=(16,9),background='#F0F4F7',foreground='#6C808D')
        style.map('TNotebook.Tab',background=[('selected','#FFFFFF'),('active','#E7F0F3')],foreground=[('selected','#007AFF')])
        style.layout('Workspace.TNotebook.Tab',[])
        style.configure('Workspace.TNotebook',background='#FFFFFF',borderwidth=0,tabmargins=0,bordercolor='#FFFFFF',lightcolor='#FFFFFF',darkcolor='#FFFFFF')
        style.configure('TScrollbar',background='#C7C7CC',troughcolor='#F8F8F8',bordercolor='#F8F8F8',arrowcolor='#8E8E93',arrowsize=10)
        style.configure('Horizontal.TProgressbar',background='#007AFF',troughcolor='#E8F0F3',borderwidth=0)
        self.bind_class('Toplevel','<Map>',self.fit_dialog,add='+')

    def fit_dialog(self, event):
        window=event.widget
        if not isinstance(window,tk.Toplevel): return
        def fit():
            if not window.winfo_exists(): return
            window.configure(background='#FFFFFF')
            width=min(window.winfo_width(),max(600,self.winfo_screenwidth()-48))
            height=min(window.winfo_height(),max(480,self.winfo_screenheight()-90))
            x=max(0,min(self.winfo_rootx()+(self.winfo_width()-width)//2,self.winfo_screenwidth()-width-20))
            y=max(24,min(self.winfo_rooty()+(self.winfo_height()-height)//2,self.winfo_screenheight()-height-65))
            window.geometry(f'{width}x{height}+{x}+{y}')
        self.after_idle(fit)

    def section_changed(self, _event=None):
        self.close_pf_dropdown()
        selected=self.notebook.select()
        for name,widget in self.nav_buttons.items():
            active=selected==str(self.nav_frames[name])
            widget.configure(fg_color='#DADCE0' if active else 'transparent',text_color='#202124',font=(self.ui_font,14,'bold' if active else 'normal'))
            if hasattr(widget,'mac_icon'): widget.mac_icon.configure(bg='#DADCE0' if active else '#F5F5F5')
        self.schedule_header_layout()

    def schedule_header_layout(self, _event=None):
        if not hasattr(self,'toolbar_bar') or getattr(self,'header_layout_busy',False): return
        pending=getattr(self,'header_layout_job',None)
        if pending is not None:self.after_cancel(pending)
        self.header_layout_job=self.after_idle(self.layout_header)

    def layout_header(self):
        self.header_layout_job=None
        if not hasattr(self,'nav_frames') or not hasattr(self,'notebook'):return
        bar=self.toolbar_bar
        if not bar.winfo_exists() or bar.winfo_width()<20:return
        active=next((name for name,frame in self.nav_frames.items() if str(frame)==self.notebook.select()),None)
        for name,buttons in self.header_actions.items():
            if name!=active:
                for button in buttons:button.place_forget()
        items=self.toolbar_selectors+self.header_actions.get(active,[])
        available=bar.winfo_width()-12
        # Leave the selectors compact and give page actions about 0.3 cm of air.
        gaps=[4 if i<len(self.toolbar_selectors)-1 else 12 for i in range(len(items)-1)]
        self.header_layout_busy=True
        try:
            # Keep one row: trim selector space first, then match the toolbar font
            # to the actual window/DPI. Never hide an action in a second row.
            fits=False
            for size in (13,12,11,10):
                for button in self.header_actions.get(active,[]):button.configure(font=(self.ui_font,size))
                for pw,qw in ((20,9),(16,7),(12,6),(10,5),(8,4)):
                    self.project_box.configure(width=pw);self.panel_box.configure(width=qw)
                    bar.update_idletasks()
                    if sum(w.winfo_reqwidth() for w in items)+sum(gaps)<=available:
                        fits=True;break
                if fits:break
            rowheight=max(32,max((w.winfo_reqheight() for w in items),default=32))+8
            x=6
            for index,widget in enumerate(items):
                width=widget.winfo_reqwidth();height=widget.winfo_reqheight()
                # Place in Tk's measured pixels; CTk's public place method applies
                # scaling again and does not accept measured width/height.
                tk.Place.place_configure(widget,x=x,y=(rowheight-height)//2,width=width,height=height)
                x+=width+(gaps[index] if index<len(gaps) else 0)
            # Tk may return a Tcl_Obj for a screen-distance option. Resolve it
            # through Tk instead of Python int(), including Tk 9/Windows.
            if bar.winfo_pixels(str(bar.cget('height')))!=rowheight:bar.configure(height=rowheight)
        finally:self.header_layout_busy=False

    def add_header_actions(self,section,actions):
        self.header_actions[section]=[self.button(self.toolbar_bar,label,callback,style=style) for label,callback,style in actions]
        self.schedule_header_layout()

    def on_error(self, kind, value, tb):
        messagebox.showerror('未完成操作', str(value), parent=self)

    def guard(self, callback):
        try:
            return callback()
        except Exception as e:
            messagebox.showerror('未完成操作', str(e), parent=self)

    def button(self, parent, text, command, **kw):
        style=kw.pop('style','TButton')
        primary=style=='Primary.TButton'
        danger=style=='Danger.TButton'
        compose=style=='Compose.TButton'
        arrow=text in ('↑','↓')
        b=ctk.CTkButton(parent,text=text,command=lambda:self.guard(command),width=152 if compose else (36 if arrow else 0),height=52 if compose else 32,
            corner_radius=10 if compose else 8,font=(self.ui_font,18 if arrow else (14 if compose else 13)),
            bg_color='#F5F5F5' if compose else ('#E3E3E3' if getattr(parent,'toolbar_host',False) else '#FFFFFF'),fg_color='#2B74CF' if primary else '#FFFFFF',
            hover_color='#2468BB' if primary else '#F1F3F4',
            text_color='#FFFFFF' if primary else ('#B3261E' if danger else '#202124'),
            border_color='#2B74CF' if primary else '#DADCE0',border_width=0 if primary else 1)
        if not getattr(parent,'toolbar_host',False):b.pack(**{'side':'left','padx':6,'pady':3,**kw})
        return b

    def action_menu(self, parent, text, actions):
        button=ctk.CTkButton(parent,text=text+' ▾',width=90,height=30,corner_radius=8,font=(self.ui_font,13),bg_color='#E3E3E3',fg_color='#E3E3E3',hover_color='#DADCE0',text_color='#3C4043',border_width=0)
        menu=tk.Menu(button,tearoff=False)
        for label,callback in actions:menu.add_command(label=label,command=lambda fn=callback:self.guard(fn))
        def show():
            try: menu.tk_popup(button.winfo_rootx(),button.winfo_rooty()+button.winfo_height()+3)
            finally: menu.grab_release()
        button.configure(command=show);button.pack(side='left',padx=4,pady=4)
        return button

    def sidebar_symbol(self,button,name):
        icon=tk.Canvas(button,width=20,height=20,background='#F5F5F5',highlightthickness=0,bd=0,cursor='hand2')
        icon.place(x=22,rely=.5,anchor='w');button.mac_icon=icon
        color='#5D646D'
        def line(*points):icon.create_line(*points,fill=color,width=1.5,capstyle='round',joinstyle='round')
        def oval(x1,y1,x2,y2):icon.create_oval(x1,y1,x2,y2,outline=color,width=1.5)
        def box(x1,y1,x2,y2):icon.create_rectangle(x1,y1,x2,y2,outline=color,width=1.5)
        if name=='数据与导出':
            box(2,3,18,17);line(2,8,18,8);line(7,8,7,17);line(12,8,12,17)
        elif name=='汇报项配置':
            line(2,5,18,5);line(2,15,18,15);oval(5,2,11,8);oval(11,12,17,18)
        elif name=='Beads设置':
            oval(2,2,9,9);oval(11,4,18,11);oval(5,12,12,19)
        elif name=='受试者库':
            oval(6,2,13,9);line(2,18,2,15,6,12,13,12,17,15,17,18)
        elif name=='重分析审核':
            line(3,5,16,5,13,2);line(16,5,13,8);line(17,15,4,15,7,18);line(4,15,7,12)
        elif name=='导入字段配置':
            line(2,12,2,18,18,18,18,12);line(10,1,10,13);line(6,9,10,13,14,9)
        elif name=='一致性核查':
            box(2,4,13,18);line(5,9,10,9);line(5,13,10,13);line(10,5,18,5,15,2);line(18,5,15,8)
        else:
            oval(2,2,18,18);line(10,5,10,10,14,12)
        icon.bind('<Button-1>',lambda _:self.notebook.select(self.nav_frames[name]))
        button.bind('<Enter>',lambda _:icon.configure(bg='#DADCE0' if self.notebook.select()==str(self.nav_frames[name]) else '#E8EAED'),add='+')
        button.bind('<Leave>',lambda _:icon.configure(bg='#DADCE0' if self.notebook.select()==str(self.nav_frames[name]) else '#F5F5F5'),add='+')

    def label(self, parent, text, **kw):
        kw.setdefault('style','Muted.TLabel'); kw.setdefault('wraplength',1000)
        widget=ttk.Label(parent,text=text,**kw); widget.pack(anchor='w',padx=12,pady=6)
        def reflow(event):
            if event.widget==parent and widget.winfo_exists():
                width=max(220,event.width-32)
                if int(str(widget.cget('wraplength')))!=width: widget.configure(wraplength=width)
        parent.bind('<Configure>',reflow,add='+')
        return widget

    def tree(self, parent, columns, widths=None, height=12):
        wrap=ttk.Frame(parent); wrap.pack(fill='both',expand=True,padx=12,pady=8)
        height=min(height,max(5,int((self.winfo_screenheight()-420)/34)))
        tree=ResultTree(wrap,columns=list(range(len(columns))),show='headings',height=height,selectmode='extended')
        for i,name in enumerate(columns):
            tree.heading(i,text=name,anchor='w'); tree.column(i,width=(widths[i] if widths else 125),minwidth=65,stretch=False,anchor='w')
        def scroll(axis,*args):
            self.close_pf_dropdown()
            (tree.yview if axis=='y' else tree.xview)(*args)
        y=ttk.Scrollbar(wrap,orient='vertical',command=lambda *args:scroll('y',*args))
        x=ttk.Scrollbar(wrap,orient='horizontal',command=lambda *args:scroll('x',*args))
        def view_changed(setter,*args):
            setter(*args);tree.schedule_redraw()
        tree.configure(yscrollcommand=lambda *a:view_changed(y.set,*a),xscrollcommand=lambda *a:view_changed(x.set,*a))
        tree.grid(row=0,column=0,sticky='nsew'); y.grid(row=0,column=1,sticky='ns'); x.grid(row=1,column=0,sticky='ew')
        wrap.rowconfigure(0,weight=1); wrap.columnconfigure(0,weight=1)
        return tree

    def build_validation_tab(self):
        t=self.tabs['验证样品']
        self.add_header_actions('验证样品',[
            ('补充文件',self.select_files,'TButton'),
            ('修改分组',self.validation_edit_group,'TButton'),
            ('导出留档',self.validation_export_ui,'Primary.TButton')])
        bar=ttk.Frame(t);bar.pack(fill='x',padx=12,pady=8)
        self.button(bar,'选择样本管',self.validation_select_samples)
        self.button(bar,'设置验证批次',self.validation_batches_dialog)
        self.validation_threshold=tk.StringVar(value='20')
        ttk.Label(bar,text='CV上限%').pack(side='left',padx=(12,4))
        ttk.Entry(bar,textvariable=self.validation_threshold,width=6).pack(side='left')
        self.button(bar,'计算CV',self.validation_preview,style='Primary.TButton')
        self.validation_status=tk.StringVar(value='所有管留档；全部数值指标自动计算，仅样本管参与CV。双击修改数据。')
        ttk.Label(t,textvariable=self.validation_status,style='Muted.TLabel').pack(anchor='w',padx=16,pady=4)
        self.validation_tree=self.tree(t,['ID','P/F','验证材料','分析员','Run','批次用途','Tube','类型','管组','复管编号'],[65,65,180,130,90,130,140,90,150,80])
        self.validation_tree.bind('<Double-1>',lambda e:self.guard(lambda:self.edit_record_dialog(self.selected_id(self.validation_tree))))
        self.validation_tree.bind('<ButtonRelease-1>',lambda e:self.guard(lambda:self.validation_pf_click(e)))

    def validation_pf_click(self,event):
        tree=self.validation_tree;row=tree.identify_row(event.y)
        if row and tree.identify_region(event.x,event.y)=='cell' and tree.identify_column(event.x)=='#2':
            tree.selection_set(row);self.pf_dialog(tree)

    def validation_select_samples(self):
        self.require_panel()
        ids=[str(x['rec']['id']) for x in self.store.validation_items(self.panel_id) if x['role']=='样本管']
        self.validation_tree.selection_set(ids)

    def validation_set_run(self,role):
        panel=self.require_panel();ids=set(self.validation_tree.selection())
        if not ids:raise ValueError('请选择对应Run中的记录')
        cfg=self.store.validation_config(panel)
        for x in self.store.validation_items(panel):
            if str(x['rec']['id']) in ids:
                run=x['payload']['meta']['Run number']
                if not run:raise ValueError('所选记录缺少Run，请核对导入来源')
                cfg.setdefault('runs',{})[run]=role
        self.store.save_validation_config(panel,cfg);self.refresh()

    def validation_edit_group(self):
        panel=self.require_panel();ids=set(self.validation_tree.selection())
        rows=[x for x in self.store.validation_items(panel) if str(x['rec']['id']) in ids]
        if not rows:raise ValueError('请选择验证记录，可多选批量设置共同材料和管组')
        win=tk.Toplevel(self);win.title('验证材料与管组');win.geometry('520x360');win.transient(self);win.grab_set()
        variables={}
        for label,key in [('验证材料（同一材料用同一名称）','material'),('管组（如rpt、saturation）','tube_group'),('类型','role')]:
            self.label(win,label);same={x[key] for x in rows};v=tk.StringVar(value=next(iter(same)) if len(same)==1 else '');variables[key]=v
            widget=ttk.Combobox(win,textvariable=v,values=['样本管','Control'],state='readonly') if key=='role' else ttk.Entry(win,textvariable=v)
            widget.pack(fill='x',padx=12)
        rep=None
        if len(rows)==1:
            self.label(win,'复管编号');rep=tk.StringVar(value=rows[0]['replicate'] or '');ttk.Entry(win,textvariable=rep).pack(fill='x',padx=12)
        def save():
            cfg=self.store.validation_config(panel)
            for x in rows:
                item={key:txt(v.get()) or x[key] for key,v in variables.items()}
                if not item['material'] or not item['tube_group']:raise ValueError('材料和管组不能为空')
                item['replicate']=x['replicate']
                if rep is not None:
                    if rep.get().strip() and not re.fullmatch(r'[1-9]\d*',rep.get().strip()):raise ValueError('复管编号请输入正整数或留空')
                    item['replicate']=int(rep.get()) if rep.get().strip() else None
                cfg.setdefault('records',{})[x['rec']['logical']]=item
            self.store.save_validation_config(panel,cfg);win.destroy();self.refresh()
        self.button(win,'保存分组',save,style='Primary.TButton',side='bottom',pady=12)

    def validation_select_metrics(self):
        panel=self.require_panel();cfg=self.store.validation_config(panel)
        fields={}
        for x in self.store.validation_items(panel):
            for f in self.store.editable_record_fields(x['rec']['id']):
                if f['area']=='raw' and f['kind']!='text':fields[f['key']]=f['name']
        if not fields:raise ValueError('没有可用的原始数值指标，请先导入验证数据')
        labels=list(fields.values());win=tk.Toplevel(self);win.title('选择验证CV指标');win.geometry('650x520');win.transient(self);win.grab_set()
        box=tk.Listbox(win,selectmode='extended',exportselection=False);box.pack(fill='both',expand=True,padx=12,pady=12)
        for i,label in enumerate(labels):
            box.insert('end',label)
            if label in cfg.get('metrics',[]):box.selection_set(i)
        bar=ttk.Frame(win);bar.pack(fill='x',padx=8,pady=8)
        def save():
            cfg['metrics']=[labels[i] for i in box.curselection()]
            cfg['threshold']=numeric(self.validation_threshold.get(),'CV上限')
            self.store.save_validation_config(panel,cfg);win.destroy()
            self.validation_status.set(f'已选{len(cfg["metrics"])}个指标；所有管留档，仅样本管计算CV。')
        self.button(bar,'全选',lambda:box.selection_set(0,'end'));self.button(bar,'保存',save,style='Primary.TButton')

    def validation_batches_dialog(self):
        panel=self.require_panel();cfg=self.store.validation_config(panel)
        runs=sorted({x['payload']['meta']['Run number'] for x in self.store.validation_items(panel) if x['payload']['meta']['Run number']},key=natural_sort_key)
        if not runs:raise ValueError('请先导入带Run编号的验证数据')
        win=tk.Toplevel(self);win.title('设置验证批次');win.geometry('600x600');win.transient(self);win.grab_set()
        self.label(win,'多选属于同一次验证的Run，指定一个Baseline。材料和样本管组仍分别计算。')
        chosen=tk.StringVar(value='（新建批次）');names=['（新建批次）']+[b['name'] for b in cfg.get('batches',[])]
        box=ttk.Combobox(win,textvariable=chosen,values=names,state='readonly');box.pack(fill='x',padx=16,pady=8)
        name=tk.StringVar(value='验证批次'+str(len(names)));ttk.Entry(win,textvariable=name).pack(fill='x',padx=16,pady=8)
        listing=tk.Listbox(win,selectmode='extended',exportselection=False);listing.pack(fill='both',expand=True,padx=16,pady=8)
        assignments={r:b['name'] for b in cfg.get('batches',[]) for r in b['runs']}
        for run in runs:listing.insert('end',run+('  ['+assignments[run]+']' if run in assignments else ''))
        baseline=tk.StringVar();self.label(win,'Baseline Run')
        basebox=ttk.Combobox(win,textvariable=baseline,state='readonly');basebox.pack(fill='x',padx=16,pady=8)
        def selection(_=None):
            selected=[runs[i] for i in listing.curselection()];basebox['values']=selected
            if baseline.get() not in selected:baseline.set('')
        listing.bind('<<ListboxSelect>>',selection)
        def load(_=None):
            listing.selection_clear(0,'end');baseline.set('')
            b=next((b for b in cfg.get('batches',[]) if b['name']==chosen.get()),None)
            name.set(b['name'] if b else '验证批次'+str(len(names)))
            if b:
                for i,r in enumerate(runs):
                    if r in b['runs']:listing.selection_set(i)
                selection();baseline.set(b['baseline'])
        box.bind('<<ComboboxSelected>>',load)
        def save():
            selected=[runs[i] for i in listing.curselection()]
            candidate=json.loads(dumps(cfg));candidate['batches']=[b for b in candidate.get('batches',[]) if b['name']!=chosen.get()]
            candidate['batches'].append({'name':name.get().strip(),'runs':selected,'baseline':baseline.get()})
            self.store.save_validation_config(panel,candidate);win.destroy();self.refresh()
        def delete():
            if chosen.get()=='（新建批次）':return
            candidate=json.loads(dumps(cfg));candidate['batches']=[b for b in candidate.get('batches',[]) if b['name']!=chosen.get()]
            self.store.save_validation_config(panel,candidate);win.destroy();self.refresh()
        bar=ttk.Frame(win);bar.pack(fill='x',padx=12,pady=12)
        self.button(bar,'保存批次',save,style='Primary.TButton');self.button(bar,'删除本批次',delete,style='Danger.TButton')

    def validation_inputs(self):
        panel=self.require_panel();cfg=self.store.validation_config(panel)
        items=self.store.validation_items(panel)
        ids=[x['rec']['id'] for x in items if x['role']=='样本管']
        threshold=numeric(self.validation_threshold.get(),'CV上限')
        if threshold<0:raise ValueError('CV上限不能小于0')
        cfg['threshold']=threshold;self.store.save_validation_config(panel,cfg)
        return panel,ids,validation_metrics(items),threshold

    def validation_preview(self):
        args=self.validation_inputs()
        def done(result):
            items,rows=result
            win=tk.Toplevel(self);win.title('验证CV（Control不参与）');win.geometry('1250x600');win.transient(self)
            self.label(win,'批内：同Run三复管；批间：当前Run均值与Baseline均值两个数的CV。全部数值指标自动计算；超限标红，不拦截导出。')
            tree=self.tree(win,VALIDATION_CV_HEADERS,[110,160,140,140,140,140,180,70,85,85,240,120,200])
            shown=[]
            for i,row in enumerate(rows):
                values=list(row);values[8]='' if row[8] is None else f'{row[8]:.2f}'
                shown.append(dict(iid=str(i),values=values,tags=('error',) if row[10]!='达标' else ()))
            PagedTable(self,tree).set_rows(shown,{},None)
        self.run_job('计算验证CV',lambda store:store.validation_report(*args),done,refresh=False)

    def validation_export_ui(self):
        args=self.validation_inputs()
        path=filedialog.asksaveasfilename(title='验证汇总＋所有管留档',initialfile=self.project_box.get()+'_'+self.panel_box.get()+'_验证.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')],parent=self)
        if path:
            self.run_job('导出验证报告',lambda store:store.export_validation(*args,path),
                         lambda n:self.status.set(f'全部验证管已留档，CV结果{n}项：'+path),refresh=False)

    def enable_sorting(self,tree,default=0):
        names=[tree.heading(c,'text') for c in tree['columns']]
        tree.sort_names=names;tree.sort_state=(default,False)
        for i,name in enumerate(names):tree.heading(str(i),command=lambda col=i:self.sort_table(tree,col))

    def sort_table(self,tree,column=None,descending=None):
        old=getattr(tree,'sort_state',(0,False))
        if column is None:column,descending=old
        elif descending is None:descending=not old[1] if old[0]==column else False
        tree.sort_state=(column,descending)
        if hasattr(tree,'page_view'):
            view=tree.page_view;view.page=0;view.sort_rows();view.render()
            for i,name in enumerate(tree.sort_names):tree.heading(str(i),text=name+(' ↓' if descending else ' ↑') if i==column else name)
            return
        items=list(tree.get_children());filled=[];empty=[]
        for iid in items:
            value=tree.set(iid,str(column))
            (filled if txt(value) else empty).append(iid)
        filled.sort(key=lambda iid:natural_sort_key(tree.set(iid,str(column))),reverse=descending)
        for index,iid in enumerate(filled+empty):tree.move(iid,'',index)
        for i,name in enumerate(tree.sort_names):tree.heading(str(i),text=name+(' ↓' if descending else ' ↑') if i==column else name)
        tree.schedule_redraw()

    def build(self):
        shell=ttk.Frame(self,style='Workspace.TFrame'); shell.pack(fill='both',expand=True)
        sidebar=ttk.Frame(shell,style='Sidebar.TFrame',width=244); sidebar.pack(side='left',fill='y'); sidebar.pack_propagate(False)
        ttk.Label(sidebar,text='流式数据汇总工具',style='Brand.TLabel').pack(anchor='w',padx=24,pady=(20,20))
        self.button(sidebar,'＋  导入文件夹',self.scan_folder,style='Compose.TButton',side='top',anchor='w',padx=8,pady=(0,14))
        nav=ttk.Frame(sidebar,style='Sidebar.TFrame'); nav.pack(fill='x',padx=(0,16))
        utility=ttk.Frame(sidebar,style='Sidebar.TFrame'); utility.pack(side='bottom',fill='x',padx=10,pady=18)
        for text,command in [('备份数据库',self.backup_db),('切换数据库',self.switch_db),('恢复已删除项',self.restore_dialog)]:
            ctk.CTkButton(utility,text=text,corner_radius=8,height=30,bg_color='#F5F5F5',fg_color='transparent',hover_color='#E8EAED',text_color='#5F6368',anchor='w',font=(self.ui_font,13),command=lambda fn=command:self.guard(fn)).pack(fill='x',pady=2)
        ttk.Label(utility,text='本机运行  ·  V'+VERSION,style='Sidebar.TLabel').pack(anchor='w',padx=12,pady=(14,0))
        ttk.Label(utility,text='作者 Guibin Huang',style='Sidebar.TLabel').pack(anchor='w',padx=12,pady=(4,0))
        content=ttk.Frame(shell,style='Workspace.TFrame'); content.pack(side='left',fill='both',expand=True)
        head=ttk.Frame(content,style='Workspace.TFrame',padding=(12,10)); head.pack(fill='x')
        selector=ctk.CTkFrame(head,fg_color='#E3E3E3',bg_color='#F5F5F5',corner_radius=12)
        selector.pack(side='left',fill='x',expand=True,padx=(0,14))
        bar=ttk.Frame(selector,style='Toolbar.TFrame',padding=(6,2));bar.pack(fill='x',padx=16,pady=5)
        ttk.Label(bar,text='项目',style='Toolbar.TLabel').pack(side='left',padx=(6,4))
        self.project_box=ttk.Combobox(bar,state='readonly',width=20); self.project_box.pack(side='left',padx=5,pady=5)
        self.project_box.bind('<<ComboboxSelected>>',lambda e:self.guard(self.choose_project))
        self.action_menu(bar,'项目操作',[('新建项目',self.new_project),('删除当前项目',self.delete_current_project),('恢复已删除项',self.restore_dialog)])
        ttk.Label(bar,text='Panel',style='Toolbar.TLabel').pack(side='left',padx=(14,4))
        self.panel_box=ttk.Combobox(bar,state='readonly',width=9); self.panel_box.pack(side='left',padx=5,pady=5)
        self.panel_box.bind('<<ComboboxSelected>>',lambda e:self.guard(self.choose_panel))
        self.action_menu(bar,'Panel操作',[('新增Panel',self.new_panel),('删除当前Panel',self.delete_current_panel),('恢复已删除项',self.restore_dialog)])
        self.toolbar_bar=bar;self.toolbar_selectors=list(bar.winfo_children());self.header_actions={}
        for widget in self.toolbar_selectors:widget.pack_forget()
        bar.toolbar_host=True;bar.pack_propagate(False)
        bar.bind('<Configure>',self.schedule_header_layout,add='+')
        footer=ttk.Frame(content,style='Workspace.TFrame',padding=(14,6)); footer.pack(side='bottom',fill='x')
        self.status=tk.StringVar(value='本机运行 · 数据保存在本地数据库')
        ttk.Label(footer,textvariable=self.status,style='Footer.TLabel').pack(anchor='w')
        self.path_label=ttk.Label(footer,text='数据库：'+str(self.store.path),style='Footer.TLabel'); self.path_label.pack(anchor='w',pady=(2,0))
        panel_card=ctk.CTkFrame(content,fg_color='#FFFFFF',bg_color='#F5F5F5',corner_radius=12)
        panel_card.pack(fill='both',expand=True,padx=(10,18),pady=(0,8))
        self.notebook=ttk.Notebook(panel_card,style='Workspace.TNotebook'); self.notebook.pack(fill='both',expand=True,padx=12,pady=10)
        pages=[
            ('数据与导出','数据工作台','导入更新、检查结果，再生成可追溯的汇总报告。'),
            ('验证样品','验证样品','所有管留档，仅样本管计算CV。'),
            ('汇报项配置','汇报项配置','选择本Panel指标，或配置跨Panel父群ABS回算。'),
            ('Beads设置','Beads设置','按Run设置Beads总数和血量。'),
            ('受试者库','受试者库','导入、核对和修正样品资料。'),
            ('重分析审核','重分析审核','比较不同分析版本，保留选择依据。'),
            ('导入字段配置','导入字段配置','对应不同导出模板的基础信息列。'),
            ('一致性核查','一致性核查','按BAS ID比较两份Excel的指定Sheet。')]
        self.tabs={};self.nav_buttons={};self.page_titles={}
        for name,label,subtitle in pages:
            frame=ttk.Frame(self.notebook);self.notebook.add(frame,text=label);self.tabs[name]=frame
            self.page_titles[name]=(label,subtitle)
        self.nav_frames=dict(self.tabs)
        for name,label,_ in pages:
            button=ctk.CTkButton(nav,text='         '+label,corner_radius=8,height=34,bg_color='#F5F5F5',fg_color='transparent',hover_color='#E8EAED',text_color='#202124',anchor='w',font=(self.ui_font,14),command=lambda key=name:self.notebook.select(self.nav_frames[key]))
            button.pack(fill='x');self.nav_buttons[name]=button
            self.sidebar_symbol(button,name)
        self.notebook.bind('<<NotebookTabChanged>>',self.section_changed)
        self.section_changed()
        t=self.tabs['数据与导出']
        cards=ttk.Frame(t); cards.pack(fill='x',padx=12,pady=(0,8))
        self.stat_labels={}
        for i,(key,title,detail,style) in enumerate([
            ('current','当前有效','当前 Panel 的结果','Stat.TLabel'),
            ('pending','待审核','需要确认的分析版本','StatWarn.TLabel'),
            ('issues','数据异常','点击筛选资料冲突或Beads待确认的记录','StatError.TLabel'),
            ('imports','导入异常','点击查看项目范围的问题文件','StatError.TLabel'),
            ('summary','Summary','当前 Panel 的导出状态','StatText.TLabel')]):
            cards.columnconfigure(i,weight=1,uniform='status')
            card=ttk.Frame(cards,padding=(10,8)); card.grid(row=0,column=i,sticky='nsew')
            ttk.Label(card,text=title,style='Muted.TLabel').pack(anchor='w')
            value=ctk.CTkLabel(card,text='—',height=28,font=(self.ui_font,21),text_color={'pending':'#9B6B00','issues':'#C83C35'}.get(key,'#1D1D1F'),anchor='w'); value.pack(fill='x',pady=(2,0)); self.stat_labels[key]=value
            if key in ('issues','imports'):
                callback=self.filter_data_issues if key=='issues' else self.show_import_issues
                value.configure(cursor='hand2')
                value.bind('<Button-1>',lambda e,fn=callback:self.guard(fn))
        ttk.Separator(t,orient='horizontal',style='Mac.TSeparator').pack(fill='x',padx=12,pady=(0,8))
        self.add_header_actions('数据与导出',[
            ('补充文件',self.select_files,'TButton'),
            ('查看记录／替代',lambda:self.review_selected(self.data_tree),'TButton')])
        lower=ttk.Frame(t); lower.pack(side='bottom',fill='x',padx=8,pady=(0,8))
        row=ttk.Frame(lower); row.pack(fill='x')
        self.button(row,'检查清单',self.check_project_ui)
        self.button(row,'预览结果',self.preview_report)
        self.button(row,'复管 CV',self.preview_cv)
        self.button(row,'导出 Summary',self.export_dialog,style='Primary.TButton')
        self.exclude_fail=tk.BooleanVar(value=False)
        ttk.Checkbutton(row,text='Summary 排除F（保留历史）',variable=self.exclude_fail,command=self.update_summary_status).pack(side='right',padx=8)
        self.summary_detail=ttk.Label(lower,text='',style='Muted.TLabel')
        self.summary_detail.pack(anchor='w',padx=8,pady=(4,0))
        caption=ttk.Frame(t); caption.pack(fill='x',padx=14,pady=(3,0))
        ttk.Label(caption,text='检测结果',style='Section.TLabel').pack(side='left')
        self.count_label=ttk.Label(caption,text='双击单元格修改／查看检测数据',style='Muted.TLabel'); self.count_label.pack(side='right')
        self.data_tree=self.tree(t,['ID','P/F ▾','BAS ID','Subject ID','Cycle（给药）','Visit','Timepoint','Tube','Run','Date','Time','Beads','血量µL'],[65,75,145,100,110,70,150,150,80,110,95,95,85])
        PagedTable(self,self.data_tree,run_column=8)
        self.enable_sorting(self.data_tree,2)
        self.data_tree.red_edit=lambda row,col:self.guard(lambda:self.edit_record_dialog(int(row),{3:'Subject ID',4:'Cycle',5:'Visit',6:'Timepoint'}.get(col))) if col!=11 else self.guard(lambda:self.open_beads_record(int(row)))
        self.data_tree.bind('<ButtonRelease-1>',lambda e:self.guard(lambda:self.open_pf_dropdown(e)))
        self.data_tree.bind('<MouseWheel>',self.close_pf_dropdown,add='+')
        self.data_tree.bind('<Configure>',self.close_pf_dropdown,add='+')
        self.data_tree.bind('<Double-1>',lambda e:self.guard(lambda:self.edit_data_cell(e)))
        self.empty_hint=ttk.Label(self.data_tree,text='这里将显示当前有效结果\n新建项目，或导入已有项目的数据文件',style='Muted.TLabel',justify='center',anchor='center',padding=(20,16))
        self.empty_hint.place(relx=.5,rely=.5,anchor='center')
        self.build_validation_tab()
        self.build_import_tab()
        self.check_page=CheckPage(self,self.tabs['一致性核查'])
        t = self.tabs['Beads设置']
        self.label(t, '每个Run保存独立参数。范围示例：1-10,12-20；Run11可单独填52900。总数请按实际批号输入。')
        form = ttk.Frame(t); form.pack(fill='x', padx=16, pady=(12,16))
        self.run_var = tk.StringVar(); self.bead_var = tk.StringVar(); self.vol_var = tk.StringVar(value='50'); self.lot_var = tk.StringVar()
        for i, (name, var, width) in enumerate([('Run范围', self.run_var, 13),('每管Beads总数',self.bead_var,13),('血量µL',self.vol_var,7),('批号（可空）',self.lot_var,12)]):
            ttk.Label(form, text=name).grid(row=0,column=i,padx=6,pady=(0,5),sticky='w')
            ttk.Entry(form, textvariable=var, width=width).grid(row=1,column=i,padx=6,pady=3,sticky='w')
        bar = ttk.Frame(form); bar.grid(row=1,column=4,padx=(16,0),sticky='w')
        self.button(bar, '预览并应用', self.apply_beads,style='Primary.TButton')
        self.button(bar, '填入选中Run', self.fill_beads)
        self.bead_tree = self.tree(t, ['Run','每管Beads总数','血量µL','批号','识别状态'], [100,160,100,160,380])
        self.bead_tree.red_edit=lambda row,col:(self.bead_tree.selection_set(row),self.guard(self.fill_beads))
        self.label(t, '修改会为受影响的当前/待审核结果建立新参数版本；旧版本保留。不会修改原始Excel。')
        self.build_config_tab()
        t = self.tabs['重分析审核']
        self.label(t, '同一样品、受试者、Panel和管名的不同结果进入审核；Cycle/Visit/Timepoint缺失时留空，有值时核对；不按文件名或时间自动选最新。')
        self.add_header_actions('重分析审核',[
            ('比较并处理选中候选',lambda:self.review_selected(self.pending_tree),'TButton'),
            ('全选待审核',lambda:self.pending_tree.selection_set(self.pending_tree.get_children()),'TButton'),
            ('批量采用选中项',lambda:self.batch_review_dialog('adopt'),'Primary.TButton'),
            ('批量不采用选中项',lambda:self.batch_review_dialog('reject'),'TButton')])
        self.pending_tree = self.tree(t, ['ID','BAS ID','Subject ID','Cycle（给药）','Visit','Timepoint','Tube','Run','Date','Time'], [70,150,110,110,80,100,180,90,120,100])
        self.pending_tree.bind('<Double-1>', lambda e: self.guard(lambda: self.review_selected(self.pending_tree)))
        self.label(t, '同一Run中的rpt1、rpt2为复管，分别保留并计算CV，不允许相互替代；同一复管的重分析仍需审核。')
        t = self.tabs['受试者库']
        self.add_header_actions('受试者库',[
            ('新增样品',lambda:self.subject_dialog(new=True),'TButton'),
            ('从Excel导入',self.import_subject_dialog,'TButton'),
            ('导出',self.export_subjects_ui,'Primary.TButton'),
            ('待修正清单',self.show_subject_import_issues,'TButton'),
            ('不采用选中行',self.dismiss_subject_issue,'TButton')])
        self.subject_tree = self.tree(t, [SUBJECT_SPLIT_LABELS.get(h,h) for h in SUBJECT_SPLIT_VIEW]+['状态','问题说明'], [145,155,110,130,110,145,110,110,140,140,100,420])
        PagedTable(self,self.subject_tree)
        self.enable_sorting(self.subject_tree,0)
        self.subject_tree.red_edit=lambda row,col:(self.subject_tree.selection_set(row),self.guard(self.subject_dialog))
        self.subject_tree.bind('<Double-1>',lambda _:self.guard(self.subject_dialog))
        self.schedule_header_layout()

    def require_panel(self):
        if self.panel_id is None:
            raise ValueError('请先新建或选择项目和Panel')
        return self.panel_id

    def mark_config_dirty(self, dirty=True):
        self.config_dirty=bool(dirty)
        if hasattr(self,'config_save_status'):
            self.config_save_status.set('待保存 · 当前Panel的汇报配置' if dirty else '已保存 · 当前Panel的汇报配置')

    def save_config_shortcut(self, event=None):
        self.guard(self.save_config_ui)
        return 'break'

    def discard_config(self):
        if self.config_dirty:
            choice=messagebox.askyesnocancel('汇报配置尚未保存','是：保存当前Panel配置后继续。\n否：放弃本次修改。\n取消：留在当前页面，继续编辑。',parent=self)
            if choice is None: return False
            if choice:
                self.save_config_ui()
                if self.config_dirty: return False
        return True

    def refresh_projects(self, select=None):
        self.project_rows = self.store.projects()
        self.project_box['values'] = [r['name'] for r in self.project_rows]
        if self.project_rows:
            index = next((i for i,r in enumerate(self.project_rows) if r['id'] == select), 0)
            self.project_box.current(index); self.choose_project(force=True)
        else:
            self.pid = self.panel_id = None; self.project_box.set(''); self.panel_box.set('')
            self.panel_box['values'] = []; self.clear_panel_selection()

    def choose_project(self, force=False):
        if not force and not self.discard_config():
            self.project_box.set(next((r['name'] for r in self.project_rows if r['id'] == self.pid), '')); return
        i = self.project_box.current()
        if i < 0:
            return
        self.pid = self.project_rows[i]['id']
        self.panel_rows = self.store.panels(self.pid)
        self.panel_box['values'] = [r['name'] for r in self.panel_rows]
        self.panel_id = None
        if self.panel_rows:
            self.panel_box.current(0); self.choose_panel(force=True)
        else:self.clear_panel_selection()

    def choose_panel(self, force=False):
        if not force and not self.discard_config():
            self.panel_box.set(next((r['name'] for r in self.panel_rows if r['id'] == self.panel_id), '')); return
        i = self.panel_box.current()
        if i < 0:
            return
        self.panel_id = self.panel_rows[i]['id']; self.load_config(); self.refresh()

    def new_project(self):
        if not self.discard_config(): return
        win=tk.Toplevel(self); win.title('新建项目'); win.geometry('480x230'); win.transient(self); win.grab_set()
        self.label(win,'项目名称（建议包含六位项目号）')
        name=tk.StringVar(); ttk.Entry(win,textvariable=name,width=40).pack(padx=12,pady=5)
        self.label(win,'初始汇报模板（现有项目不受影响）')
        template=ttk.Combobox(win,values=['通用项目：首次自行选择汇报项','770计数模板：现有35项与11项ABS'],state='readonly',width=42)
        template.pack(padx=12,pady=5); template.current(0)
        def commit():
            ident=self.store.add_project(name.get(),template.current()==1)
            win.destroy(); self.refresh_projects(ident); self.open_section('汇报项配置')
        bar=ttk.Frame(win); bar.pack(fill='x'); self.button(bar,'创建项目和Panel1',commit)

    def new_panel(self):
        if not self.pid: raise ValueError('请先选择项目')
        if not self.discard_config(): return
        ident=self.store.next_panel(self.pid)
        self.reload_panels(ident); self.open_section('汇报项配置')

    def clear_panel_selection(self):
        self.panel_id=None;self.panel_box.set('');self.cfg=empty_config()
        self.load_import_ui();self.render_config()
        self.bead_field_var.set(self.cfg['beads_field']);self.load_tube_names(self.cfg)
        self.cv_enabled_var.set(self.cfg.get('replicate_cv',True));self.cv_threshold_var.set('')
        self.parent_var.set('本Panel Beads Events');self.parent_confirmed.set(False)
        self.source_var.set('');self.name_var.set('')
        self.run_var.set('');self.bead_var.set('');self.vol_var.set('50');self.lot_var.set('')
        self.mark_config_dirty(False)
        self.config_save_status.set('尚未选择Panel')
        self.source_box['values']=[];self.bead_field_box['values']=[]
        self.refresh()

    def delete_current_panel(self):
        ident=self.require_panel()
        if not self.discard_config():return
        row=self.store.panel(ident);dependents=self.store.panel_dependents(ident)
        if dependents:raise ValueError('此Panel仍被引用，请先更改来源：\n'+'\n'.join(dependents))
        count=len(self.store.records(ident))
        if not messagebox.askyesno('删除当前Panel',f'删除 {row["name"]}？共有{count}条检测记录。\n会先备份，数据保留在可恢复删除项中。',parent=self):return
        self.store.delete_panel(ident);self.reload_panels()
        self.status.set('Panel已删除；如需撤回，请使用“恢复已删除项”。')

    def delete_current_project(self):
        if not self.pid:raise ValueError('请先选择项目')
        if not self.discard_config():return
        row=self.store.db.execute('SELECT * FROM projects WHERE id=?',(self.pid,)).fetchone()
        if not messagebox.askyesno('删除当前项目',f'删除项目 {row["name"]} 及其Panel？\n会先备份，数据保留，可从“恢复已删除项”恢复。',parent=self):return
        self.store.delete_project(self.pid);self.refresh_projects()
        self.status.set('项目已删除；如需撤回，请使用“恢复已删除项”。')

    def restore_dialog(self):
        if not self.discard_config():return
        items=self.store.deleted_items()
        if not items:messagebox.showinfo('恢复删除项','没有可恢复的项目或Panel。',parent=self);return
        win=tk.Toplevel(self);win.title('恢复已删除项目/Panel');win.geometry('760x520');win.transient(self);win.grab_set()
        self.label(win,'删除项的数据和配置仍在；如果出现同名，可在下方填写新的恢复名称。')
        table=self.tree(win,['类型','名称','所属项目'],[110,310,270],height=9)
        for i,item in enumerate(items):table.insert('','end',iid=str(i),values=['项目' if item['kind']=='project' else 'Panel',item['name'],item['project']])
        name=tk.StringVar();entry=ttk.Entry(win,textvariable=name,width=55);entry.pack(fill='x',padx=12,pady=8)
        def select(_=None):
            if table.selection():name.set(items[int(table.selection()[0])]['name'])
        table.bind('<<TreeviewSelect>>',select)
        def restore():
            if len(table.selection())!=1:raise ValueError('请选择一个删除项')
            item=items[int(table.selection()[0])];pid=self.store.restore_item(item['kind'],item['id'],name.get())
            win.destroy();self.refresh_projects(pid)
            if item['kind']=='panel':self.reload_panels(item['id'])
            self.status.set('已恢复删除项。')
        bar=ttk.Frame(win);bar.pack(fill='x',padx=8,pady=8)
        self.button(bar,'恢复选中项',restore,style='Primary.TButton');self.button(bar,'取消',win.destroy)

    def reload_panels(self, select=None):
        self.panel_rows=self.store.panels(self.pid)
        self.panel_box['values']=[r['name'] for r in self.panel_rows]
        if self.panel_rows:
            self.panel_box.current(next((i for i,r in enumerate(self.panel_rows) if r['id']==select),0))
            self.choose_panel(force=True)
        else:self.clear_panel_selection()

    def open_section(self, name):
        self.notebook.select(self.tabs[name])

    def refresh(self):
        self.close_pf_dropdown()
        for tree in [self.pending_tree,self.bead_tree,self.validation_tree]:
            tree.set_red_cells({})
            tree.delete(*tree.get_children())
        if not self.panel_id:
            for tree in (self.data_tree,self.subject_tree):tree.page_view.set_rows([],{},None)
            self.summary_detail.configure(text="")
            for label in self.stat_labels.values(): label.configure(text='—')
            self.count_label.configure(text='')
            self.empty_hint.place(relx=.5,rely=.5,anchor='center'); return
        records = self.store.records(self.panel_id,["current","pending"])
        data_rows=[];subject_rows=[]
        conflicts=self.store.metadata_conflicts(self.pid);red_cells={};subject_conflicts={}
        reviews=self.store.beads_reviews(self.panel_id)
        for record in records:
            fields=conflicts.get(record['id'],{})
            bas=json.loads(record['payload'])['meta'].get('BAS ID')
            if fields:subject_conflicts.setdefault(bas,set()).update(fields)
        counts = {'current': 0, 'pending': 0, 'superseded': 0, 'rejected': 0}
        for r in records:
            payload = json.loads(r['payload']); m = dict(payload['meta']); ident = str(r['id'])
            if m.get('BAS ID'):counts[r['state']] += 1
            if payload.get('sample_kind') == 'validation': m['BAS ID'] = '验证：'+sample_display(payload)
            if r['state'] == 'current' and payload['meta'].get('BAS ID'):
                for field in conflicts.get(r['id'],{}):
                    if field in ('Subject ID','Cycle','Visit','Timepoint'):red_cells[(ident,{'Subject ID':3,'Cycle':4,'Visit':5,'Timepoint':6}[field])]='待修正，暂不纳入有效汇报'
                if m['Run number'] and run_number(m['Run number']) in reviews:red_cells[(ident,11)]='样品名Beads待确认'
                data_rows.append(dict(iid=ident, tags=('error',) if r['pf']=='F' else (), values=[r['id'],(r['pf'] or '—')+' ▾'] + [m[h] for h in ['BAS ID','Subject ID','Cycle','Visit','Timepoint','Tube','Run number','Date','Time']] + [r['beads'] or '',r['volume'] or '']))
            if r['state'] == 'pending':
                self.pending_tree.insert('', 'end', iid=ident, tags=('warning',), values=[r['id']] + [m[h] for h in ['BAS ID','Subject ID','Cycle','Visit','Timepoint','Tube','Run number','Date','Time']])
        self.data_tree.page_view.set_rows(data_rows,red_cells,(str(self.store.path),self.panel_id))
        parameters={r['run']:r for r in self.store.parameters(self.panel_id)};bead_red={}
        for run in sorted(set(parameters)|set(reviews)):
            r=parameters.get(run);review=reviews.get(run)
            count=r['count'] if r else (review['counts'][0] if len(review['counts'])==1 else '')
            status=('冲突：'+', '.join(map(str,review['counts'])) if review['conflict'] else '样品名提取，待确认总数/血量/批号') if review else ''
            self.bead_tree.insert('','end',iid=str(run),values=['Run'+str(run),count,r['volume'] if r else '',r['lot'] if r else '',status])
            if review:bead_red[(str(run),1)]=status;bead_red[(str(run),4)]=status
        self.bead_tree.set_red_cells(bead_red)
        settings=self.store.validation_config(self.panel_id);self.validation_threshold.set(str(settings.get('threshold',20)))
        for item in sorted(self.store.validation_items(self.panel_id),key=lambda x:(natural_sort_key(x['payload']['meta']['Run number']),natural_sort_key(x['material']),x['replicate'] or 0)):
            r=item['rec'];m=item['payload']['meta']
            self.validation_tree.insert('','end',iid=str(r['id']),values=[r['id'],r['pf'],item['material'],m['Analyst'],m['Run number'],item['run_role'],m['Tube'],item['role'],item['tube_group'],item['replicate'] or ''])
        self.validation_status.set(f"全部验证管：{len(self.validation_tree.get_children())}；验证批次：{len(settings.get('batches',[]))}。全部数值指标自动计算；未分批只算批内CV。")
        subject_red={}
        for s in self.store.subject_rows(self.pid):
            split=split_subject(s)
            fields=subject_conflicts.get(s['样品编号'],set())
            for field in fields:
                for col in {'Subject ID':[2],'Cycle':[3],'Visit':[4],'Timepoint':[5,6,7]}.get(field,[]):subject_red[(s['样品编号'],col)]='与检测数据不一致'
            subject_rows.append(dict(iid=s['样品编号'], values=[split.get(h,'') for h in SUBJECT_SPLIT_VIEW]+['待核对' if fields else '有效','与检测数据不一致：'+', '.join(sorted(fields)) if fields else '']))
        for issue in self.store.subject_import_issues(self.pid):
            data=split_subject(issue['row'])
            subject_rows.append(dict(iid=f'issue:{issue["audit_id"]}:{issue["index"]}',values=[txt(data.get(h)) for h in SUBJECT_SPLIT_VIEW]+['待修正',f'{Path(issue["file"]).name} / {issue["sheet"]} / 行{issue["line"]}：'+issue['detail']],tags=('error',)))
        self.subject_tree.page_view.set_rows(subject_rows,subject_red,(str(self.store.path),self.pid))
        self.count_label.configure(text=f'{self.panel_box.get()} · {counts["current"]} 条正式记录 · 红色资料待修正后方可汇报 · 双击修改')
        self.stat_labels['current'].configure(text=str(counts['current']))
        self.stat_labels['pending'].configure(text=str(counts['pending']))
        self.stat_labels['issues'].configure(text=str(len({iid for iid,col in red_cells})))
        self.stat_labels['imports'].configure(text=str(len(self.store.unresolved_imports(self.pid))))
        if counts['current']: self.empty_hint.place_forget()
        else:
            self.empty_hint.configure(text='当前没有有效结果\n'+('请先处理待审核的分析版本' if counts['pending'] else '点击“导入文件夹”开始导入'))
            self.empty_hint.place(relx=.5,rely=.5,anchor='center')
        self.update_summary_status()
        available_fields = self.store.available_fields(self.panel_id)
        self.source_box['values'] = available_fields
        self.bead_field_box['values'] = available_fields
        self.path_label.configure(text='数据库：' + str(self.store.path))


    def filter_data_issues(self):
        self.require_panel();self.open_section('数据与导出')
        self.data_tree.page_view.show_issues()
        self.status.set('仅显示资料冲突或Beads待确认的记录；点击“显示全部”恢复。筛选不改变Summary导出范围。')

    def show_import_issues(self):
        self.require_panel();rows=self.store.unresolved_imports(self.pid)
        win=tk.Toplevel(self);win.title('项目导入异常');win.geometry('1100x600');win.transient(self)
        self.label(win,f'共 {len(rows)} 个问题文件；这些文件问题不等同于工作台中的数据异常。')
        tree=self.tree(win,['状态','文件','原因'],[90,450,480])
        for r in rows:tree.insert('','end',values=[r['status'],r['file'],r['detail']])
        self.button(win,'查看最近导入清单',self.last_import_dialog)

    def update_summary_status(self):
        if not self.panel_id:return
        summary=self.store.summary_status(self.panel_id,self.exclude_fail.get())
        self.stat_labels['summary'].configure(text='已导出' if summary['ready'] else '待导出')
        if not summary.get('stamp'):
            self.summary_detail.configure(text='尚未导出。搜索和分页仅影响查看，Summary导出完整范围。');return
        stamp=datetime.fromisoformat(summary['stamp']).astimezone().strftime('%Y-%m-%d %H:%M')
        scope='排除F' if summary['exclude_fail'] else '包含F'
        state='与当前数据及导出条件一致' if summary['ready'] else '数据或导出条件已变化，请重新导出'
        self.summary_detail.configure(text=f'上次导出：{stamp} · {scope} · {summary["rows"]} 条 · {state}')

    def run_job(self, title, job, done, refresh=True):
        """All Tk operations stay on the main thread; worker opens its own SQLite connection."""
        dialog = tk.Toplevel(self); dialog.title(title); dialog.geometry('440x140'); dialog.transient(self); dialog.grab_set()
        ttk.Label(dialog,text=title,style='Section.TLabel').pack(pady=(20,4))
        detail=ttk.Label(dialog,text='正在处理，请稍候…',style='Muted.TLabel');detail.pack(pady=(0,12))
        progress = ttk.Progressbar(dialog, mode='indeterminate'); progress.pack(fill='x',padx=25); progress.start()
        dialog.protocol('WM_DELETE_WINDOW', lambda: None)
        q = queue.Queue(); path = self.store.path
        def work():
            store = None
            try:
                store = Store(path)
                store.progress=lambda message:q.put(('progress',message))
                result=job(store)
                store.close();store=None
                q.put((True, result))
            except Exception as exc:
                q.put((False, str(exc)))
            finally:
                if store: store.close()
        threading.Thread(target=work, daemon=True).start()
        def poll():
            latest=None
            try:
                while True:
                    ok, result = q.get_nowait()
                    if ok!='progress':break
                    latest=result
            except queue.Empty:
                if latest is not None:detail.configure(text=latest)
                self.after(100, poll); return
            progress.stop(); dialog.grab_release(); dialog.destroy()
            if ok:
                self.guard(lambda: done(result))
            else:
                messagebox.showerror('未完成操作', result, parent=self)
            if refresh:self.refresh()
        self.after(100, poll)

    def select_files(self):
        panel = self.require_panel()
        paths = filedialog.askopenfilenames(title='选择当前Panel原始Excel（可多选）',filetypes=[('Excel','*.xlsx')],parent=self)
        if paths: self.preview_files(paths)

    def scan_folder(self):
        self.require_panel()
        if not self.discard_config(): return
        folder=filedialog.askdirectory(title='选择一个项目的上层文件夹（含所有Run和Panel）',parent=self)
        if not folder: return
        pid=self.pid
        self.run_job('扫描并识别Panel',lambda s:s.project_routes(pid,scan_project_folder(folder,[r[0] for r in s.db.execute('SELECT path FROM exclusions')])),self.route_scanned_files)

    def route_scanned_files(self, routes):
        if not routes:
            messagebox.showinfo('扫描结果','文件夹内没有Excel文件。',parent=self); return
        if any(route_needs_attention(r) for r in routes):
            self.preview_project_files(routes)
        else:
            self.begin_project_import(routes)

    def begin_project_import(self, routes):
        pid=self.pid; selected=self.panel_id
        def work(store):
            stats=store.import_project(pid,routes)
            stats['checks']=store.check_project(pid)
            return stats
        def done(stats):
            self.reload_panels(selected); self.show_project_import_result(stats)
        self.run_job('自动分Panel导入并检查',work,done)

    def preview_project_files(self, routes):
        if not routes:
            messagebox.showinfo('扫描结果','文件夹内没有Excel文件。',parent=self); return
        pid=self.pid
        win=tk.Toplevel(self); win.title('处理识别异常 · 正常文件已自动分配'); win.geometry('1180x720'); win.transient(self); win.grab_set()
        win.columnconfigure(0,weight=1)
        win.rowconfigure(1,weight=1)
        header=ttk.Frame(win); header.grid(row=0,column=0,sticky='ew')
        content=ttk.Frame(win); content.grid(row=1,column=0,sticky='nsew')
        footer=ttk.Frame(win); footer.grid(row=2,column=0,sticky='ew',pady=(0,10))
        def commit():
            if not any(not r.get('skip') and (r.get('panel_id') or r.get('create_name')) for r in routes): raise ValueError('没有已分配的可导入文件')
            win.destroy()
            self.begin_project_import(routes)
        actions=ttk.Frame(header); actions.pack(fill='x',padx=8,pady=(6,0))
        continue_button=self.button(actions,'继续导入',commit,style='Primary.TButton')
        self.button(actions,'取消',win.destroy)
        self.label(header,f'项目：{self.project_box.get()}。识别只依据文件名；子文件夹会一起扫描。未分配文件不导入，其余继续。')
        self.label(header,'可多选后分配Panel或跳过。旧名称Panel可手动选定，导入后记住编号对应关系；原名称和历史保留。')
        notice=self.label(header,'')
        only_problems=tk.BooleanVar(value=True)
        filterbar=ttk.Frame(header); filterbar.pack(fill='x',padx=12)
        tree=self.tree(content,['状态','目标Panel','文件Run','文件名','识别说明'],[100,160,85,330,420],height=12)
        empty_hint=ttk.Label(tree,text='',style='Muted.TLabel',justify='center',anchor='center')
        names={p['id']:p['name'] for p in self.store.panels(pid)}
        choices={f'{p["name"]}（已有）':('id',p['id']) for p in self.store.panels(pid)}
        numbers={r['number'] for r in routes if r['number']}
        numbers.add(max([self.store.panel_number(p['id']) or 0 for p in self.store.panels(pid)],default=0)+1)
        for n in sorted(numbers):
            if not any(self.store.panel_number(p['id'])==n for p in self.store.panels(pid)):
                choices[f'新建 Panel{n}']=('new',f'Panel{n}')
        def render():
            selected=tree.selection(); tree.delete(*tree.get_children())
            active=0; waiting=0; skipped=0; runs={}
            expected={self.store.panel_number(p['id']) for p in self.store.panels(pid)} | {r['number'] for r in routes}
            expected.discard(None)
            for i,r in enumerate(routes):
                target=names.get(r.get('panel_id')) or r.get('create_name','')
                state='跳过' if r.get('skip') else ('待分配' if not target else '准备导入')
                active+=state=='准备导入'; waiting+=state=='待分配'; skipped+=state=='跳过'
                if not only_problems.get() or (not r.get('skip') and route_needs_attention(r)):
                    tree.insert('','end',iid=str(i),values=[state,target or '待分配',r['run'],Path(r['file']).name,r['detail']])
                if r['run'] and state=='准备导入':
                    n=self.store.panel_number(r['panel_id']) if r.get('panel_id') else numbered_panel(r.get('create_name'))
                    runs.setdefault(r['run'],set()).add(n)
            tree.selection_set([i for i in selected if tree.exists(i)])
            if tree.get_children():
                empty_hint.place_forget()
            else:
                empty_hint.configure(text='异常文件已处理完。\n点击上方“继续导入”处理其余文件；已跳过文件不会导入。' if active else '当前筛选下没有文件。取消上方勾选可查看全部文件。')
                empty_hint.place(relx=.5,rely=.5,anchor='center')
            continue_button.configure(text=f'继续导入（{active} 个文件）',state='normal' if active else 'disabled')
            missing=[f'{run}缺'+','.join(f'Panel{n}' for n in sorted(expected-nums)) for run,nums in runs.items() if expected-nums]
            notice.configure(text=f'准备导入 {active}；待分配 {waiting}；跳过 {skipped}。'+('  文件名检查：'+'；'.join(missing[:5]) if missing else ''))
        ttk.Checkbutton(filterbar,text='只显示尚需处理的文件（取消可看全部，包括已跳过）',variable=only_problems,command=render).pack(side='left')
        render()
        detail=tk.StringVar(value='选中一行查看完整路径。')
        self.label(footer,'文件路径：'); ttk.Entry(footer,textvariable=detail,state='readonly').pack(fill='x',padx=12)
        tree.bind('<<TreeviewSelect>>',lambda _:detail.set(routes[int(tree.selection()[0])]['file'] if tree.selection() else ''))
        bar=ttk.Frame(footer); bar.pack(fill='x',padx=8,pady=8)
        combo=ttk.Combobox(bar,values=list(choices),state='readonly',width=30); combo.pack(side='left',padx=6)
        if choices: combo.current(0)
        def assign():
            if not tree.selection(): raise ValueError('先选中文件，可以Ctrl多选')
            kind,value=choices[combo.get()]
            for i in tree.selection():
                r=routes[int(i)]
                if r.get('protected_skip'): continue
                r.update(panel_id=value if kind=='id' else None,create_name=value if kind=='new' else '',skip=False,attention=False)
                r['detail']='手动分配；导入时核对表内Panel'
            render()
        def skip():
            for i in tree.selection(): routes[int(i)].update(skip=True,detail='本次手动跳过')
            render()
        # Protect outputs, unsupported formats and references from manual reassignment.
        for r in routes: r['protected_skip']=r.get('skip',False)
        self.button(bar,'分配给选中目标',assign); self.button(bar,'跳过选中文件',skip)
        bar=ttk.Frame(footer); bar.pack(fill='x',padx=8)
        self.button(bar,'开始导入所有已分配文件',commit,style='Primary.TButton'); self.button(bar,'取消',win.destroy)

    def show_project_import_result(self, stats):
        win=tk.Toplevel(self); win.title('项目导入结果'); win.geometry('1180x690'); win.transient(self); win.grab_set()
        files=stats['files']
        actions=ttk.Frame(win); actions.pack(fill='x',padx=8,pady=6)
        def mark_validation():
            chosen=[files[int(i)] for i in tree.selection()]
            if not chosen: raise ValueError('请先选中样品名无法识别的失败文件，可Ctrl多选')
            updates={}; targets=[]
            for f in chosen:
                if f['status']!='失败': raise ValueError('只可选择失败文件')
                if not any(word in f.get('detail','') for word in ('样品名格式不支持','无法识别完整样品编号')):
                    raise ValueError('所选文件不是样品名识别问题；请核对其实际报错，不能用验证Run绕过访视或时间点冲突')
                panel=f.get('panel_id')
                if not panel: raise ValueError('请先为文件分配正确Panel')
                runs=runs_in_name(Path(f['file']).stem)
                if len(runs)!=1:
                    raise ValueError('文件名没有唯一Run编号；请在源文件名中补充Run数字后重新导入，避免错标其他Run')
                n=runs[0]
                updates.setdefault(panel,set(self.store.config(panel).get('import_options',{}).get('validation_runs',[]))).add(n)
                targets.append((f,n))
            names=sorted({f'{f["panel_name"]} / Run{n}' for f,n in targets})
            if not messagebox.askyesno('确认验证实验', '\n'.join(names)+'\n\n将这些Panel中的整Run设为验证实验，支持任意样品名，单独汇总，不纳入正式Summary。\n仅重试选中的失败文件；既有记录不会改写。是否继续？',parent=win): return
            self.store.configure_validation_runs(updates)
            if self.panel_id in updates:
                self.cfg['import_options']=self.store.config(self.panel_id).get('import_options',{})
            routes=[]
            for f,n in targets:
                routes.append({'file':f['file'],'number':self.store.panel_number(f['panel_id']),'run':'Run'+str(n),
                               'panel_id':f['panel_id'],'create_name':'','skip':False,'attention':False,'detail':'人工确认验证Run'})
            win.destroy(); self.begin_project_import(routes)
        if stats['errors']: self.button(actions,'标记为验证Run并重试',mark_validation,style='Primary.TButton')
        self.label(win,f'文件成功 {sum(f["status"]=="成功" for f in files)}，失败 {len(stats["errors"])}，跳过 {sum(f["status"]=="跳过" for f in files)}。新增有效 {stats["current"]}，待审核 {stats["pending"]}，重复跳过 {stats["duplicate"]}。')
        self.label(win,'下一步运行项目检查，补齐汇报配置／Beads并处理待审核；跨Panel ABS请在汇报项配置中选择父群来源。')
        tree=self.tree(win,['状态','Panel','文件','有效','待审核','重复','详情'],[65,120,320,60,70,60,440],height=10)
        only_problems=tk.BooleanVar(value=True)
        def render():
            tree.delete(*tree.get_children())
            for i,f in enumerate(files):
                problem=f['status']=='失败' or f['pending']>0 or (f['status']=='跳过' and f['detail'] not in ('已登记的汇总文件或受试者资料','程序导出的汇总文件','已登记为受试者资料或程序导出文件'))
                if not only_problems.get() or problem:
                    tree.insert('','end',iid=str(i),values=[f['status'],f['panel_name'],Path(f['file']).name,f['current'],f['pending'],f['duplicate'],f['detail']])
        render()
        text=tk.Text(win,height=8,wrap='word'); text.pack(fill='both',expand=True,padx=12,pady=6)
        def show(value): text.configure(state='normal'); text.delete('1.0','end'); text.insert('end',value); text.configure(state='disabled')
        warning='\n'.join(stats['warnings']) or '无额外检查提示'
        checks=stats.get('checks',{})
        problems=[r['panel_name']+'：'+'；'.join(r['issues'][:6]) for r in checks.get('panels',[]) if r['issues']]
        show('\n'.join(problems) if problems else '正常文件已完成自动分配。表格仅显示异常／待审核；完整清单可展开。')
        tree.bind('<<TreeviewSelect>>',lambda _:show(files[int(tree.selection()[0])]['file']+'\n'+files[int(tree.selection()[0])]['detail']) if tree.selection() else None)
        bar=ttk.Frame(win); bar.pack(fill='x',padx=8)
        ttk.Checkbutton(bar,text='只看异常／待审核',variable=only_problems,command=render).pack(side='left',padx=6)
        self.button(bar,'全部检查提示',lambda:show(warning))
        def save():
            path=filedialog.asksaveasfilename(parent=win,initialfile='项目导入结果.txt',defaultextension='.txt',filetypes=[('文本','*.txt')])
            if path:
                if Path(path).suffix.lower()!='.txt': raise ValueError('请使用.txt后缀')
                Path(path).write_text('\n'.join(f'{f["status"]} | {f["panel_name"]} | {f["file"]}\n{f["detail"]}' for f in files)+'\n'+warning,encoding='utf-8-sig')
        self.button(bar,'保存清单',save)
        def retry():
            paths={f['file']:f for f in stats['errors']}; routes=[]
            for candidate in scan_project_folder_from_paths(paths):
                f=paths[candidate['file']]; candidate.update(panel_id=f.get('panel_id'),create_name='')
                routes.append(candidate)
            win.destroy(); self.preview_project_files(routes)
        if stats['errors']: self.button(bar,'重新分配／重试失败文件',retry)
        def check(): win.destroy(); self.check_project_ui()
        self.button(bar,'项目检查',check); self.button(bar,'关闭',win.destroy)

    def check_project_ui(self):
        self.require_panel()
        if self.config_dirty: raise ValueError('请先保存当前汇报项配置')
        pid=self.pid; exclude=self.exclude_fail.get()
        self.run_job('检查各Panel汇报条件',lambda s:s.check_project(pid,exclude),self.show_project_check)

    def show_project_check(self, report):
        win=tk.Toplevel(self); win.title('项目检查 · Summary前'); win.geometry('1100x680'); win.transient(self); win.grab_set()
        self.label(win,'检查结果用于核对，不阻止Summary导出：无法计算的指标留空，原因写入“汇总问题”。验证样品单独导出原始指标，不含ABS。')
        tree=self.tree(win,['Panel','当前有效','待审核','计算通过','排除F','Summary状态','问题数'],[150,100,100,100,100,210,90],height=6)
        rows={str(r['panel_id']):r for r in report['panels']}
        for ident,r in rows.items(): tree.insert('','end',iid=ident,values=[r['panel_name'],r['current'],r['pending'],r['ready'],r['excluded'],'已导出且未变更' if r['summary_ready'] else '尚未导出／已变更',len(r['issues'])])
        text=tk.Text(win,wrap='word',height=14); text.pack(fill='both',expand=True,padx=12,pady=8)
        messages=[]
        for r in report['panels']:
            messages.append(r['panel_name']+'：'+('；'.join(r['issues']) if r['issues'] else '汇报检查通过'))
        for f in report['files']: messages.append(f'文件{f["status"]}：{f["file"]}\n{f["detail"]}')
        text.insert('end','\n\n'.join(messages)); text.configure(state='disabled')
        def go(section):
            if not tree.selection(): raise ValueError('请先选择一个Panel')
            if not self.discard_config(): return
            ident=int(tree.selection()[0]); win.destroy(); self.reload_panels(ident); self.open_section(section)
        bar=ttk.Frame(win); bar.pack(fill='x',padx=8)
        for label,section in [('汇报项设置','汇报项配置'),('设置Beads','Beads设置'),('字段映射','导入字段配置'),('处理待审核','重分析审核'),('返回Summary','数据与导出')]:
            self.button(bar,label,lambda sec=section:go(sec))
        self.button(bar,'关闭',win.destroy)

    def preview_files(self, paths):
        if not paths:
            messagebox.showinfo('扫描结果','没有找到可识别的数据表。支持Keyword表头；自定义表头请先保存“导入字段配置”。',parent=self); return
        win = tk.Toplevel(self); win.title('选择要导入当前Panel的文件'); win.geometry('960x570'); win.transient(self); win.grab_set()
        self.label(win, f'当前：{self.project_box.get()} / {self.panel_box.get()}。请仅选择属于这个Panel的文件；Ctrl可多选。')
        tree = self.tree(win,['文件名','完整路径'],[320,590])
        for i,p in enumerate(paths): tree.insert('','end',iid=str(i),values=[Path(p).name,str(p)])
        tree.selection_set(tree.get_children())
        bar = ttk.Frame(win); bar.pack(fill='x',padx=8)
        self.button(bar,'全选',lambda: tree.selection_set(tree.get_children()))
        self.button(bar,'清空选择',lambda: tree.selection_remove(tree.selection()))
        def commit():
            chosen = [paths[int(i)] for i in tree.selection()]
            if not chosen: raise ValueError('请至少选择一个文件')
            panel = self.panel_id; win.destroy()
            self.run_job('逐文件导入原始数据',lambda s:s.import_batch(panel,chosen),lambda stats:self.show_import_result(panel,stats))
        self.button(bar,'导入选中文件',commit)
        self.button(bar,'取消',win.destroy)

    def last_import_dialog(self):
        panel=self.require_panel()
        for row in self.store.db.execute("SELECT action,detail FROM audit WHERE project_id=? AND action IN ('批量导入结果','项目自动分Panel导入') ORDER BY id DESC",(self.pid,)):
            stats=json.loads(row['detail'])
            if row['action']=='项目自动分Panel导入':
                self.show_project_import_result(stats); return
            if stats.get('panel')==panel:
                self.show_import_result(panel,stats); return
        raise ValueError('当前项目还没有导入清单')

    def show_import_result(self, panel, stats):
        win = tk.Toplevel(self); win.title('导入结果清单'); win.geometry('1080x680'); win.transient(self); win.grab_set()
        files = stats.get('files',[])
        summary = (f'文件：成功 {sum(f["status"]=="成功" for f in files)} / 失败 {len(stats.get("errors",[]))} / '
                   f'跳过 {sum(f["status"]=="跳过" for f in files)}\n'
                   f'结果：新增有效 {stats["current"]} / 待审核 {stats["pending"]} / 重复跳过 {stats["duplicate"]}')
        self.label(win,summary)
        self.label(win,'失败文件整份未导入；其他文件已保存。选中一行查看完整原因。修好源文件后可重试，已导入内容会去重。')
        tree = self.tree(win,['状态','文件名','新增有效','待审核','重复跳过','说明'],[65,280,80,80,90,420],height=9)
        for i,f in enumerate(files):
            tree.insert('','end',iid=str(i),values=[f['status'],Path(f['file']).name,f['current'],f['pending'],f['duplicate'],f['detail']])
        detail = tk.Text(win,height=9,wrap='word'); detail.pack(fill='both',expand=True,padx=12,pady=8)
        def display(value):
            detail.configure(state='normal'); detail.delete('1.0','end'); detail.insert('1.0',value); detail.configure(state='disabled')
        def select(_):
            if tree.selection():
                f = files[int(tree.selection()[0])]; display(f['file']+'\n\n'+f['detail'])
        tree.bind('<<TreeviewSelect>>',select)
        warning_text = '\n'.join(stats.get('warnings',[])) or '没有额外提示'
        display(warning_text)
        bar = ttk.Frame(win); bar.pack(fill='x',padx=8)
        self.button(bar,'查看全部检查提示',lambda:display(warning_text))
        def save():
            path = filedialog.asksaveasfilename(parent=win,title='保存导入结果清单',initialfile='导入结果清单.txt',defaultextension='.txt',filetypes=[('文本','*.txt')])
            if path:
                if Path(path).suffix.lower() != '.txt': raise ValueError('导入清单请保存为.txt，避免覆盖原始Excel或数据库')
                lines = [summary,''] + [f'{f["status"]} | {f["file"]}\n{f["detail"]}\n有效 {f["current"]}，待审核 {f["pending"]}，重复 {f["duplicate"]}\n' for f in files]
                Path(path).write_text('\n'.join(lines)+'\n检查提示：\n'+warning_text,encoding='utf-8-sig')
        self.button(bar,'保存清单为TXT',save)
        def retry():
            paths = [f['file'] for f in stats.get('errors',[])]; win.destroy()
            self.run_job('重试失败文件',lambda s:s.import_batch(panel,paths),lambda result:self.show_import_result(panel,result))
        if stats.get('errors'): self.button(bar,'重试失败文件',retry)
        self.button(bar,'关闭',win.destroy)

    def apply_beads(self):
        panel = self.require_panel(); runs = expand_runs(self.run_var.get())
        count = numeric(self.bead_var.get(),'每管Beads总数','events'); volume = numeric(self.vol_var.get(),'血量µL')
        if count <= 0 or volume <= 0: raise ValueError('Beads和血量必须大于0')
        affected = [r for r in self.store.records(panel,['current','pending']) if txt(json.loads(r['payload'])['meta']['Run number']) and run_number(json.loads(r['payload'])['meta']['Run number']) in runs and (r['beads'],r['volume'],r['lot']) != (count,volume,txt(self.lot_var.get()))]
        reason = simpledialog.askstring('参数修改原因',f'将设置 {len(runs)} 个Run：{count} beads / {volume} µL。\n影响 {len(affected)} 条当前/待审核结果，会建立新版本。\n填写初始设置或修改原因：',parent=self)
        if not reason: return
        if not messagebox.askyesno('确认应用',f'Run：{self.run_var.get()}\n每管Beads：{count}\n血量：{volume} µL\n旧版本保留，确认应用？',parent=self): return
        count_versions = self.store.set_parameters(panel,runs,count,volume,self.lot_var.get(),reason)
        self.refresh(); self.status.set(f'Beads已保存；建立{count_versions}个参数新版本。')

    def fill_beads(self):
        selected = self.bead_tree.selection()
        if not selected: raise ValueError('请先选择一个或多个Run')
        self.run_var.set(','.join(selected))
        values = self.bead_tree.item(selected[0],'values')
        self.bead_var.set(values[1]); self.vol_var.set(values[2]); self.lot_var.set(values[3])


    def build_import_tab(self):
        t = self.tabs['导入字段配置']
        self.add_header_actions('导入字段配置',[
            ('读取Excel表头',self.inspect_import_headers,'TButton'),
            ('恢复自动识别',self.reset_import_mapping,'TButton'),
            ('保存导入配置',self.save_import_ui,'Primary.TButton'),
            ('验证Run设置',self.edit_validation_runs,'TButton')])
        f = ttk.Frame(t); f.pack(fill='x',padx=18,pady=12)
        self.import_vars = {}; self.import_boxes = {}
        for i,field in enumerate(INPUT_FIELDS):
            row,col = i % 6,(i // 6)*2
            ttk.Label(f,text=field).grid(row=row,column=col,sticky='w',padx=8,pady=6)
            var = tk.StringVar(value='（自动识别）'); self.import_vars[field] = var
            box = ttk.Combobox(f,textvariable=var,width=36,state='readonly',values=['（自动识别）','（留空）'])
            box.grid(row=row,column=col+1,padx=8,pady=6,sticky='ew'); self.import_boxes[field] = box
        self.import_panel_var = tk.StringVar(); self.import_tube_var = tk.StringVar()
        ttk.Label(f,text='来源Panel校验').grid(row=6,column=0,sticky='w',padx=8,pady=6)
        ttk.Entry(f,textvariable=self.import_panel_var,width=22).grid(row=6,column=1,sticky='w',padx=8,pady=6)
        ttk.Label(f,text='缺失Tube时管名').grid(row=6,column=2,sticky='w',padx=8,pady=6)
        ttk.Entry(f,textvariable=self.import_tube_var,width=28).grid(row=6,column=3,sticky='w',padx=8,pady=6)

    def reset_import_mapping(self):
        for var in self.import_vars.values(): var.set('（自动识别）')

    def load_import_ui(self):
        options = self.cfg.get('import_options',{})
        for field,var in self.import_vars.items():
            source = options.get('mapping',{}).get(field)
            value = '（自动识别）' if source is None else ('（留空）' if source=='' else source)
            self.import_boxes[field]['values'] = ['（自动识别）','（留空）']+([source] if source else [])
            var.set(value)
        self.import_panel_var.set(options.get('source_panel','')); self.import_tube_var.set(options.get('tube_default',''))

    def save_import_ui(self):
        panel = self.require_panel(); mapping = {}
        for field,var in self.import_vars.items():
            value = var.get()
            if value != '（自动识别）': mapping[field] = '' if value=='（留空）' else value
        options = {**self.cfg.get('import_options',{}),'mapping':mapping,'source_panel':txt(self.import_panel_var.get()),'tube_default':txt(self.import_tube_var.get())}
        self.store.save_import_options(panel,options)
        self.cfg['import_options'] = options
        self.status.set('当前Panel导入配置已保存。下一步可从Excel生成汇报项。')

    def edit_validation_runs(self):
        panel=self.require_panel(); old=self.store.config(panel).get('import_options',{}).get('validation_runs',[])
        value=simpledialog.askstring('验证Run设置 · '+self.store.panel(panel)['name'],
            '仅适用于当前Panel。填写Run编号或范围，如1-3,8。\n这些Run全部按验证实验导入，不进入正式Summary。\n已有记录不自动重分类。清空可移除尚无验证记录的设置。',
            initialvalue=','.join(map(str,old)),parent=self)
        if value is None: return
        runs=expand_runs(value) if value.strip() else []
        self.store.configure_validation_runs({panel:runs})
        self.cfg['import_options']=self.store.config(panel).get('import_options',{})
        self.status.set('验证Run设置已保存；请重新导入失败文件。')

    def inspect_import_headers(self):
        self.require_panel()
        path = filedialog.askopenfilename(title='选择一个当前Panel原始Excel',filetypes=[('Excel','*.xlsx')],parent=self)
        if not path: return
        wb = load_workbook(path,read_only=True,data_only=False); candidates = []
        try:
            for ws in wb:
                for row,values in enumerate(ws.iter_rows(max_row=30,values_only=True),1):
                    if sum(v is not None and v!='' for v in values) >= 2:
                        candidates.append((ws.title,row,unique_headers(values)))
        finally: wb.close()
        if not candidates: raise ValueError('前30行没有可选择的表头')
        preferred = 0
        for i,(_,_,labels) in enumerate(candidates):
            try: fields = resolve_metadata(labels)
            except ValueError: continue
            if fields['BAS ID'] or fields['Specimen']: preferred = i; break
        win = tk.Toplevel(self); win.title('选择表头行'); win.geometry('950x260'); win.transient(self); win.grab_set()
        self.label(win,'确认Excel中的表头行。应用后，在主窗口逐项选择元数据来源并保存。')
        names = [f'{s} / 第{r}行：'+ ' | '.join(labels)[:140] for s,r,labels in candidates]
        box = ttk.Combobox(win,values=names,state='readonly',width=115); box.pack(padx=12,pady=10); box.current(preferred)
        def use():
            labels = [v for v in candidates[box.current()][2] if v]
            for field,var in self.import_vars.items():
                self.import_boxes[field]['values'] = ['（自动识别）','（留空）']+labels
                matches = [v for v in labels if header_token(v) in set(map(norm,META_ALIASES[field]))]
                var.set(matches[0] if len(matches)==1 else ('（留空）' if field=='Analyst' and not matches else '（自动识别）'))
            win.destroy(); self.status.set('已读取表头；核对映射后点击“保存当前Panel导入配置”。')
        self.button(win,'使用这行表头',use)

    def generate_config_from_excel(self):
        self.require_panel()
        path = filedialog.askopenfilename(title='从当前Panel原始Excel生成汇报项',filetypes=[('Excel','*.xlsx')],parent=self)
        if not path: return
        cfg,labels,fields = config_from_excel(path,self.cfg.get('import_options',{}))
        if not messagebox.askyesno('载入汇报项',f'识别到{len(cfg["columns"])}个指标，将替换当前编辑列表。\n只读取原始指标，不自动增加ABS或做FMO扣除。继续？',parent=self): return
        self.cfg = cfg; self.bead_field_var.set(cfg['beads_field']); self.load_tube_names(cfg)
        self.cv_enabled_var.set(True); self.cv_threshold_var.set(''); self.render_config(); self.mark_config_dirty()
        self.status.set('汇报项已生成；重名列以Excel列号区分，请核对名称后保存配置。')

    def build_config_tab(self):
        t = self.tabs['汇报项配置']
        self.add_header_actions('汇报项配置',[
            ('保存配置',self.save_config_ui,'Primary.TButton'),
            ('导出配置',self.export_config,'TButton'),
            ('载入配置',self.import_config,'TButton'),
            ('从Excel读取汇报项',self.generate_config_from_excel,'TButton')])
        self.config_save_status=tk.StringVar(value='尚未选择Panel')
        ttk.Label(t,textvariable=self.config_save_status,style='Muted.TLabel').pack(anchor='e',padx=18,pady=(4,12))
        pages=ttk.Notebook(t);pages.pack(fill='both',expand=True,padx=16,pady=(0,12))
        metrics=ttk.Frame(pages,padding=8);settings=ttk.Frame(pages,padding=12)
        pages.add(metrics,text='汇报指标');pages.add(settings,text='ABS / CV 设置')
        f = ttk.LabelFrame(settings,text='本Panel ABS参数',padding=12); f.pack(fill='x',pady=(0,16))
        ttk.Label(f,text='本Panel Beads Events列').grid(row=0,column=0,padx=5,sticky='w')
        self.bead_field_var = tk.StringVar(); self.bead_field_box = ttk.Combobox(f,textvariable=self.bead_field_var,width=26); self.bead_field_box.grid(row=0,column=1,padx=8)
        ttk.Label(f,text='参与ABS回算的管名').grid(row=0,column=2,padx=5)
        self.regex_var = tk.StringVar(); ttk.Entry(f,textvariable=self.regex_var,width=30).grid(row=0,column=3,padx=8)
        self.cv_enabled_var = tk.BooleanVar(value=True); self.cv_threshold_var = tk.StringVar()
        self.cv_digits_var=tk.StringVar(value='2')
        self.cv_low_boundary_var=tk.StringVar();self.cv_low_threshold_var=tk.StringVar()
        cv=ttk.LabelFrame(settings,text='复管统计与显示',padding=12);cv.pack(fill='x')
        low=ttk.Frame(cv);low.pack(fill='x')
        ttk.Checkbutton(low,text='rpt复管CV',variable=self.cv_enabled_var).pack(side='left',padx=(0,18))
        ttk.Label(low,text='常规CV ≤').pack(side='left',padx=(0,5))
        ttk.Entry(low,textvariable=self.cv_threshold_var,width=6).pack(side='left',padx=4)
        ttk.Label(low,text='%').pack(side='left',padx=(0,22))
        ttk.Label(low,text='平均 %parent <').pack(side='left',padx=5)
        ttk.Entry(low,textvariable=self.cv_low_boundary_var,width=6).pack(side='left',padx=4)
        ttk.Label(low,text='% 时，CV ≤').pack(side='left',padx=5)
        ttk.Entry(low,textvariable=self.cv_low_threshold_var,width=6).pack(side='left',padx=4)
        ttk.Label(low,text='%').pack(side='left',padx=5)
        precision_row=ttk.Frame(cv);precision_row.pack(fill='x',pady=(16,0))
        ttk.Label(precision_row,text='Summary CV小数位（独立于指标）').pack(side='left')
        ttk.Spinbox(precision_row,from_=0,to=10,textvariable=self.cv_digits_var,width=5).pack(side='left',padx=8)
        self.cv_digits_var.trace_add('write',lambda *_:self.mark_config_dirty())
        self.cv_low_boundary_var.trace_add('write',lambda *_:self.mark_config_dirty())
        self.cv_low_threshold_var.trace_add('write',lambda *_:self.mark_config_dirty())
        self.bead_field_var.trace_add('write',lambda *_:self.mark_config_dirty())
        self.regex_var.trace_add('write',lambda *_:self.mark_config_dirty())
        self.cv_enabled_var.trace_add('write',lambda *_:self.mark_config_dirty())
        self.cv_threshold_var.trace_add('write',lambda *_:self.mark_config_dirty())
        t=metrics
        f = ttk.LabelFrame(t,text='添加汇报指标',padding=10); f.pack(fill='x',pady=(0,10))
        f.columnconfigure(0,weight=1);f.columnconfigure(1,weight=1)
        self.source_var = tk.StringVar(); self.name_var = tk.StringVar(); self.mode_var = tk.StringVar(value='原始值'); self.kind_var = tk.StringVar(value='数值')
        for i,name in enumerate(['来源表头（可输入）','输出列名','方式','来源类型']): ttk.Label(f,text=name).grid(row=0,column=i,padx=6,pady=(0,6),sticky='w')
        self.source_box = ttk.Combobox(f,textvariable=self.source_var,width=20); self.source_box.grid(row=1,column=0,padx=6,sticky='ew')
        self.source_box.bind('<<ComboboxSelected>>',lambda e:self.source_selected())
        ttk.Entry(f,textvariable=self.name_var,width=20).grid(row=1,column=1,padx=6,sticky='ew')
        ttk.Combobox(f,textvariable=self.mode_var,values=['原始值','ABS回算'],state='readonly',width=12).grid(row=1,column=2,padx=6)
        ttk.Combobox(f,textvariable=self.kind_var,values=list(KIND_LABEL.values()),state='readonly',width=17).grid(row=1,column=3,padx=6)
        self.precision_mode_var=tk.StringVar(value='默认');self.precision_digits_var=tk.StringVar(value='3')
        precision=ttk.Frame(f);precision.grid(row=2,column=0,columnspan=4,sticky='w',pady=(12,0))
        ttk.Label(precision,text='该指标精度').pack(side='left',padx=6)
        ttk.Combobox(precision,textvariable=self.precision_mode_var,values=list(PRECISION_LABELS.values()),state='readonly',width=12).pack(side='left',padx=6)
        ttk.Spinbox(precision,from_=0,to=10,textvariable=self.precision_digits_var,width=5).pack(side='left',padx=6)
        ttk.Label(precision,text='位 · CV精度在“ABS / CV 设置”中单独设置').pack(side='left',padx=6)
        f=ttk.LabelFrame(t,text='当前指标的ABS来源',padding=8);self.abs_form=f
        self.parent_var=tk.StringVar(value='本Panel Beads Events');self.parent_rep_var=tk.StringVar(value='唯一一管（多管需指定）');self.parent_confirmed=tk.BooleanVar(value=False)
        ttk.Label(f,text='ABS回算来源').grid(row=0,column=0,padx=6,sticky='w')
        self.parent_box=ttk.Combobox(f,textvariable=self.parent_var,state='readonly',width=34);self.parent_box.grid(row=0,column=1,padx=6)
        ttk.Label(f,text='父群取管规则').grid(row=0,column=2,padx=6)
        ttk.Combobox(f,textvariable=self.parent_rep_var,values=['唯一一管（多管需指定）','同Run rpt复管均值'],state='readonly',width=25).grid(row=0,column=3,padx=6)
        self.parent_check=ttk.Checkbutton(f,text='已核对：百分比的分母与所选父群ABS一致',variable=self.parent_confirmed)
        self.parent_check.grid(row=1,column=0,columnspan=4,sticky='w',padx=6,pady=6)
        def reset_parent_confirmation(*_): self.parent_confirmed.set(False)
        self.source_var.trace_add('write',reset_parent_confirmation);self.parent_var.trace_add('write',reset_parent_confirmation);self.kind_var.trace_add('write',reset_parent_confirmation)
        self.parent_box.bind('<<ComboboxSelected>>',lambda _:self.parent_selected())
        bar = ttk.Frame(t); bar.pack(fill='x',padx=8);self.column_actions=bar
        self.button(bar,'＋ 添加汇报项',self.add_column)
        self.button(bar,'编辑选中项',self.edit_column)
        self.button(bar,'指定父群来源管',self.parent_choices_dialog)
        self.button(bar,'删除',self.remove_columns,style='Danger.TButton')
        self.button(bar,'↑',lambda:self.move_column(-1))
        self.button(bar,'↓',lambda:self.move_column(1))
        self.button(bar,'清空列表',self.clear_columns,style='Danger.TButton')
        self.config_tree = self.tree(t,['顺序','输出名称','来源表头','方式','来源类型','精度','ABS回算来源'],[50,170,190,100,100,100,170],height=12)
        self.config_tree.bind('<Double-1>',lambda e:self.guard(self.edit_column))
        self.mode_var.trace_add('write',lambda *_:self.show_abs_form())
        self.parent_var.trace_add('write',lambda *_:self.show_abs_form())
        self.show_abs_form()

    def show_abs_form(self):
        if self.mode_var.get()=='ABS回算':
            self.abs_form.pack(fill='x',padx=12,pady=6,before=self.column_actions)
            if self.parent_var.get()=='本Panel Beads Events': self.parent_check.grid_remove()
            else: self.parent_check.grid()
        else: self.abs_form.pack_forget()

    def load_tube_names(self,cfg):
        self.cv_digits_var.set(str(cfg.get('cv_digits',2)))
        display=tube_names_display(cfg)
        self.regex_var.set(display);self._tube_names_loaded=display
        if hasattr(self,'cv_low_boundary_var'):
            self.cv_low_boundary_var.set('' if cfg.get('cv_low_pct_boundary') is None else txt(cfg['cv_low_pct_boundary']))
            self.cv_low_threshold_var.set('' if cfg.get('cv_low_pct_threshold') is None else txt(cfg['cv_low_pct_threshold']))

    def load_config(self):
        self.cfg = self.store.config(self.panel_id)
        self.load_import_ui()
        self.bead_field_var.set(self.cfg['beads_field']); self.load_tube_names(self.cfg)
        self.cv_enabled_var.set(self.cfg.get('replicate_cv',True)); self.cv_threshold_var.set('' if self.cfg.get('cv_threshold') is None else str(self.cfg['cv_threshold']))
        self.render_config(); self.mark_config_dirty(False)
        self.parent_var.set('本Panel Beads Events');self.parent_confirmed.set(False)

    def refresh_parent_options(self):
        self.parent_options={o['label']:o for o in self.store.parent_abs_options(self.panel_id)} if self.panel_id else {}
        self.parent_box['values']=['本Panel Beads Events']+list(self.parent_options)

    def render_config(self):
        self.refresh_parent_options()
        self.config_tree.delete(*self.config_tree.get_children())
        for i,c in enumerate(self.cfg['columns']):
            self.config_tree.insert('','end',iid=str(i),values=[i+1,c['name'],c['source'],mode_label(c),KIND_LABEL[c['kind']],precision_label(c),abs_source_label(c)])

    def parent_selected(self):
        self.mode_var.set('ABS回算')
        if self.parent_var.get()!='本Panel Beads Events' and self.kind_var.get() not in (KIND_LABEL['pct100'],KIND_LABEL['pct1']):
            self.kind_var.set(KIND_LABEL['pct100'])

    def source_selected(self):
        self.name_var.set(self.source_var.get()); self.kind_var.set(KIND_LABEL[infer_kind(self.source_var.get())])

    def column_from_form(self):
        source = txt(self.source_var.get()); name = txt(self.name_var.get())
        if not source or not name: raise ValueError('请输入来源表头和输出名称')
        mode = 'abs' if self.mode_var.get() == 'ABS回算' else 'raw'
        kind = next(k for k,v in KIND_LABEL.items() if v == self.kind_var.get())
        result={'source':source,'name':name,'mode':mode,'kind':'events' if mode=='abs' else kind}
        if mode=='abs' and self.parent_var.get()!='本Panel Beads Events':
            chosen=self.parent_options.get(self.parent_var.get())
            if not chosen: raise ValueError('父群来源项不存在，请先保存计数Panel的ABS配置后重新选择')
            result.update(mode='parent_abs',kind=kind,parent_panel=chosen['panel'],parent_panel_name=chosen['panel_name'],parent_source=chosen['source'],
                parent_replicates='mean' if self.parent_rep_var.get()=='同Run rpt复管均值' else 'unique',parent_confirmed=self.parent_confirmed.get())
        result.update(precision_options(self.precision_mode_var.get(),self.precision_digits_var.get()))
        validate_config({**self.cfg,'columns':[result]})
        return result

    def add_column(self):
        self.require_panel(); c = self.column_from_form()
        if norm(c['name']) in list(map(norm,META)) + [norm(x['name']) for x in self.cfg['columns']]: raise ValueError('输出名称重复')
        self.cfg['columns'].append(c); self.mark_config_dirty(); self.render_config()

    def edit_column(self):
        selected=self.config_tree.selection()
        if len(selected)!=1: return
        i=int(selected[0]);c=self.cfg['columns'][i];self.refresh_parent_options()
        win=tk.Toplevel(self);win.title('编辑汇报项');win.transient(self);win.grab_set()
        source=tk.StringVar(value=c['source']);name=tk.StringVar(value=c['name']);mode=tk.StringVar(value='原始值' if c['mode']=='raw' else 'ABS回算')
        kind=tk.StringVar(value=c['kind']);parent=tk.StringVar(value=abs_source_label(c) or '本Panel Beads Events')
        reps=tk.StringVar(value='同Run rpt复管均值' if c.get('parent_replicates')=='mean' else '唯一一管（多管需指定）');confirmed=tk.BooleanVar(value=c.get('parent_confirmed',False))
        for j,(label,var) in enumerate([('来源表头（跨Panel选子群%）',source),('输出名称',name)]):
            ttk.Label(win,text=label).grid(row=j,column=0,padx=12,pady=8,sticky='w')
            ttk.Entry(win,textvariable=var,width=46).grid(row=j,column=1,padx=12,pady=8)
        for j,(label,var,values) in enumerate([('方式',mode,['原始值','ABS回算']),('来源类型',kind,list(KIND_LABEL)),
            ('ABS回算来源',parent,['本Panel Beads Events']+list(self.parent_options)),('父群取管规则',reps,['唯一一管（多管需指定）','同Run rpt复管均值'])],2):
            ttk.Label(win,text=label).grid(row=j,column=0,padx=12,pady=8,sticky='w')
            ttk.Combobox(win,textvariable=var,values=values,state='readonly',width=44).grid(row=j,column=1,padx=12,pady=8)
        ttk.Checkbutton(win,text='已核对百分比的分母与所选父群ABS一致',variable=confirmed).grid(row=6,column=0,columnspan=2,padx=12,pady=8,sticky='w')
        pm=tk.StringVar(value=PRECISION_LABELS[c.get('precision_mode','default')]);pd=tk.StringVar(value=str(c.get('precision_digits',3)))
        ttk.Label(win,text='Summary精度').grid(row=7,column=0,padx=12,pady=8,sticky='w')
        pf=ttk.Frame(win);pf.grid(row=7,column=1,sticky='w',padx=12)
        ttk.Combobox(pf,textvariable=pm,values=list(PRECISION_LABELS.values()),state='readonly',width=12).pack(side='left',padx=6)
        ttk.Spinbox(pf,from_=0,to=10,textvariable=pd,width=5).pack(side='left',padx=6)
        ttk.Label(pf,text='位').pack(side='left')
        for var in (source,parent,kind): var.trace_add('write',lambda *_:confirmed.set(False))
        def save():
            candidate={'source':txt(source.get()),'name':txt(name.get()),'mode':'raw' if mode.get()=='原始值' else 'abs','kind':kind.get()}
            if candidate['mode']=='abs':
                if parent.get()=='本Panel Beads Events': candidate['kind']='events'
                else:
                    chosen=self.parent_options.get(parent.get())
                    if not chosen: raise ValueError('父群来源不存在，请重新选择')
                    candidate.update(mode='parent_abs',parent_panel=chosen['panel'],parent_panel_name=chosen['panel_name'],parent_source=chosen['source'],
                        parent_replicates='mean' if reps.get()=='同Run rpt复管均值' else 'unique',parent_confirmed=confirmed.get())
                    if c.get('parent_panel')==chosen['panel'] and c.get('parent_source')==chosen['source']:candidate['parent_choices']=c.get('parent_choices',{})
            candidate.update(precision_options(pm.get(),pd.get()))
            cfg=json.loads(dumps(self.cfg));cfg['columns'][i]=candidate;validate_config(cfg)
            self.cfg=cfg;self.mark_config_dirty();self.render_config();win.destroy()
        ttk.Button(win,text='保存修改到列表',command=lambda:self.guard(save)).grid(row=8,column=1,pady=15)

    def parent_choices_dialog(self):
        panel=self.require_panel();selected=self.config_tree.selection()
        if len(selected)!=1: raise ValueError('请先选中一个跨Panel ABS汇报项')
        i=int(selected[0]);column=self.cfg['columns'][i]
        if column['mode']!='parent_abs': raise ValueError('该汇报项未配置跨Panel父群ABS来源')
        self.store.parent_source_config(panel,column)
        cfg=self.ui_config();groups={}
        for rec in self.store.records(panel,['current']):
            p=json.loads(rec['payload']);m=p['meta']
            if p.get('sample_kind')=='validation' or not re.fullmatch(cfg['sample_regex'],txt(m['Tube'])):continue
            key=m['BAS ID']+'|'+txt(m['Run number']);groups[key]=(m['BAS ID'],m['Run number'])
        win=tk.Toplevel(self);win.title('指定父群来源管 · '+column['name']);win.geometry('850x600');win.transient(self);win.grab_set()
        self.label(win,'按BAS匹配；同一样品的子群复管共用此来源。默认优先同Run，多来源不自动取第一条。')
        table=self.tree(win,['BAS ID','子群Run','父群来源'],[180,120,420],height=10)
        keys={};choices=tk.StringVar();box=ttk.Combobox(win,textvariable=choices,state='readonly',width=75);box.pack(fill='x',padx=12,pady=8)
        mapping={}
        def render():
            table.delete(*table.get_children());keys.clear()
            for n,(key,(bas,run)) in enumerate(sorted(groups.items())):
                keys[str(n)]=key;ids=column.get('parent_choices',{}).get(key)
                table.insert('','end',iid=str(n),values=[bas,run,'指定记录 '+', '.join(map(str,ids)) if ids else '自动匹配（'+column.get('parent_replicates','unique')+'）'])
        def select(_=None):
            if not table.selection():return
            key=keys[table.selection()[0]];bas,_=groups[key];mapping.clear();mapping['自动匹配']=None
            for rec in self.store.parent_candidates(panel,column,bas):
                m=json.loads(rec['payload'])['meta'];mapping[f'ID {rec["id"]} · {m["Run number"]} · {m["Tube"]} · P/F={rec["pf"] or "未判定"}']=[rec['id']]
            box['values']=list(mapping)
            current=column.get('parent_choices',{}).get(key);choices.set(next((label for label,ids in mapping.items() if ids==current),'自动匹配' if current is None else '原选择已失效，请重选'))
        table.bind('<<TreeviewSelect>>',select)
        def save():
            if not table.selection() or choices.get() not in mapping:raise ValueError('请选择样品和有效来源管')
            key=keys[table.selection()[0]];ids=mapping[choices.get()];column.setdefault('parent_choices',{})
            if ids is None:column['parent_choices'].pop(key,None)
            else:column['parent_choices'][key]=ids
            self.mark_config_dirty();self.render_config();render()
        bar=ttk.Frame(win);bar.pack(fill='x',padx=8,pady=8)
        self.button(bar,'保存当前样品取管到配置列表',save);self.button(bar,'关闭',win.destroy)
        render()

    def remove_columns(self):
        for i in sorted(map(int,self.config_tree.selection()),reverse=True): self.cfg['columns'].pop(i)
        self.mark_config_dirty();self.render_config()

    def move_column(self,delta):
        selected=self.config_tree.selection()
        if len(selected)!=1: raise ValueError('请选择一个汇报项')
        i=int(selected[0]);j=i+delta
        if 0<=j<len(self.cfg['columns']):
            self.cfg['columns'][i],self.cfg['columns'][j]=self.cfg['columns'][j],self.cfg['columns'][i]
            self.mark_config_dirty();self.render_config();self.config_tree.selection_set(str(j))

    def clear_columns(self):
        if messagebox.askyesno('清空汇报项','只清空当前编辑列表，保存后生效。继续？',parent=self):
            self.cfg['columns']=[];self.mark_config_dirty();self.render_config()

    def reset_columns(self):
        if messagebox.askyesno('载入模板','用原项目计数模板替换当前编辑列表？',parent=self):
            options=self.cfg.get('import_options',{});self.cfg=default_config();self.cfg['import_options']=options;self.bead_field_var.set(self.cfg['beads_field']);self.load_tube_names(self.cfg);self.mark_config_dirty();self.render_config()
            self.cv_enabled_var.set(True);self.cv_threshold_var.set('')

    def ui_config(self):
        cfg=json.loads(dumps(self.cfg));cfg['beads_field']=txt(self.bead_field_var.get())
        names=self.regex_var.get()
        # Preserve an unrecognized old custom rule until the user replaces it.
        old_names=tube_names_display(self.cfg)
        if 'abs_tube_names' in cfg or names!=getattr(self,'_tube_names_loaded','') or old_names:
            cfg['abs_tube_names']=names;cfg['sample_regex']=tube_names_regex(names)
        cfg['cv_digits']=int(self.cv_digits_var.get())
        if not 0<=cfg['cv_digits']<=10:raise ValueError('CV小数位支持0–10位')
        cfg['replicate_cv']=self.cv_enabled_var.get()
        cfg['cv_threshold']=numeric(self.cv_threshold_var.get(),'CV阈值') if txt(self.cv_threshold_var.get()) else None
        cfg['cv_low_pct_boundary']=numeric(self.cv_low_boundary_var.get(),'低表达%parent界限','pct100') if txt(self.cv_low_boundary_var.get()) else None
        cfg['cv_low_pct_threshold']=numeric(self.cv_low_threshold_var.get(),'低表达CV限值') if txt(self.cv_low_threshold_var.get()) else None
        validate_config(cfg);return cfg

    def save_config_ui(self):
        panel=self.require_panel();cfg=self.ui_config();self.store.save_config(panel,cfg);self.cfg=self.store.config(panel);self.mark_config_dirty(False)
        self.status.set('汇报配置已保存。原始数据保留。')

    def export_config(self):
        self.require_panel();cfg=self.ui_config()
        p=filedialog.asksaveasfilename(title='导出配置',defaultextension='.json',filetypes=[('JSON','*.json')],parent=self)
        if p: Path(p).write_text(json.dumps(cfg,ensure_ascii=False,indent=2),encoding='utf-8')

    def import_config(self):
        self.require_panel()
        p=filedialog.askopenfilename(title='载入配置',filetypes=[('JSON','*.json')],parent=self)
        if not p:return
        cfg=strip_retired_config(json.loads(Path(p).read_text(encoding='utf-8-sig')));validate_config(cfg)
        if not self.discard_config():return
        self.cfg=cfg;self.bead_field_var.set(cfg['beads_field']);self.load_tube_names(cfg);self.mark_config_dirty();self.render_config()
        self.load_import_ui()
        self.cv_enabled_var.set(cfg.get('replicate_cv',True));self.cv_threshold_var.set('' if cfg.get('cv_threshold') is None else str(cfg['cv_threshold']))

    def selected_id(self,tree):
        selected=tree.selection()
        if len(selected)!=1:raise ValueError('请先选中一条记录')
        return int(selected[0])

    def open_beads_record(self,ident):
        rec=self.store.record(ident);m=json.loads(rec['payload'])['meta']
        self.run_var.set(m['Run number']);self.bead_var.set('' if rec['beads'] is None else str(rec['beads']));self.vol_var.set('' if rec['volume'] is None else str(rec['volume']));self.lot_var.set(rec['lot'] or '')
        self.open_section('Beads设置')

    def edit_data_cell(self, event):
        tree=self.data_tree; row=tree.identify_row(event.y)
        if not row or tree.identify_region(event.x,event.y)!='cell': return
        tree.selection_set(row); ident=int(row)
        column=tree.identify_column(event.x); index=int(column[1:])-1
        keys=['ID','P/F','BAS ID','Subject ID','Cycle','Visit','Timepoint','Tube','Run number','Date','Time','Beads','血量µL']
        key=keys[index] if 0<=index<len(keys) else ''
        if key=='P/F': self.open_pf_dropdown(event); return
        if key in ('Beads','血量µL'):
            rec=self.store.record(ident); payload=json.loads(rec['payload'])
            run=payload['meta'].get('Run number')
            if not run: raise ValueError('记录缺少Run，无法定位Beads参数')
            self.run_var.set(run); self.bead_var.set('' if rec['beads'] is None else str(rec['beads']))
            self.vol_var.set('' if rec['volume'] is None else str(rec['volume'])); self.lot_var.set(rec['lot'] or '')
            self.open_section('Beads设置'); return
        self.edit_record_dialog(ident, key if key in ('Subject ID','Cycle','Visit','Timepoint','Date','Time') else None)

    def close_pf_dropdown(self, _event=None):
        editor=getattr(self,'pf_editor',None)
        if editor is not None:
            self.pf_editor=None
            if editor.winfo_exists(): editor.destroy()

    def open_pf_dropdown(self,event):
        tree=self.data_tree; row=tree.identify_row(event.y)
        if not row or tree.identify_region(event.x,event.y)!='cell' or tree.identify_column(event.x)!='#2':
            self.close_pf_dropdown(); return
        self.close_pf_dropdown();tree.selection_set(row)
        bbox=tree.bbox(row,'1')
        if not bbox: return
        rec=self.store.record(int(row))
        if rec['state']!='current': return
        editor=ttk.Combobox(tree,state='readonly',values=['未判定','P','F'],width=5)
        self.pf_editor=editor;editor.set(rec['pf'] or '未判定');editor.place(x=bbox[0],y=bbox[1],width=bbox[2],height=bbox[3]);editor.focus_set()
        def commit(_event=None):
            value=editor.get();pf='' if value=='未判定' else value
            try:
                current=self.store.record(int(row))
                if current['state']!='current': raise ValueError('记录状态已变化，请刷新后再修改')
                if current['pf']!=pf:self.store.set_pf(int(row),pf,'数据工作台P/F下拉选择')
            finally: self.close_pf_dropdown()
            self.refresh()
            if self.data_tree.exists(row):self.data_tree.selection_set(row);self.data_tree.see(row)
        editor.bind('<<ComboboxSelected>>',lambda e:self.guard(lambda:commit(e)))
        editor.bind('<Escape>',lambda e:self.close_pf_dropdown())

    def export_subjects_ui(self):
        if self.pid is None: raise ValueError('请先选择项目')
        path=filedialog.asksaveasfilename(title='导出受试者库',initialfile=self.project_box.get()+'_受试者库.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')],parent=self)
        if not path:return
        count=self.store.export_subjects(self.pid,path)
        self.status.set(f'受试者库已导出：{count}条 · '+path)

    def edit_record_dialog(self, ident, initial_field=None):
        rec=self.store.record(ident)
        if rec['state']!='current': raise ValueError('只可修改当前有效记录；待审核记录请先处理审核')
        win=tk.Toplevel(self); win.title('修改记录'); win.geometry('960x630'); win.transient(self); win.grab_set()
        active={'id':ident}; entries=[]
        footer=ttk.Frame(win); footer.pack(side='bottom',fill='x',padx=8,pady=8)
        title=self.label(win,'')
        self.label(win,'双击字段修改。保存后创建新版本，旧值保留，P/F清空待重新判定；随后重新导出Summary。')
        self.label(win,'红色资料可逐条修正；保存后重新核对。BAS ID、Run和管名请核对原文件。')
        table=self.tree(win,['类别','字段／来源指标','当前值','类型'],[95,380,240,130],height=10)
        def render():
            current=self.store.record(active['id']); payload=json.loads(current['payload'])
            title.configure(text=f'记录{active["id"]} · {sample_display(payload)} · {payload["meta"]["Run number"]} · {payload["meta"]["Tube"]}')
            entries[:]=self.store.editable_record_fields(active['id'])
            table.delete(*table.get_children())
            for i,f in enumerate(entries): table.insert('','end',iid=str(i),values=['信息' if f['area']=='meta' else '检测值',f['name'],txt(f['value']),KIND_LABEL[f['kind']]])
        def edit():
            if len(table.selection())!=1: raise ValueError('请先选择一个字段')
            field=entries[int(table.selection()[0])]; old_id=active['id']
            dialog=tk.Toplevel(win); dialog.title('修改 '+field['name']); dialog.geometry('680x500'); dialog.transient(win); dialog.grab_set()
            self.label(dialog,'字段：'+field['name']+'    当前值：'+txt(field['value']))
            issues=self.store.metadata_conflicts(self.pid).get(active['id'],{}).get(field['key'],[])
            if issues:self.label(dialog,'\n'.join(issues),wraplength=600)
            self.label(dialog,'新值'); value=tk.StringVar(value=txt(field['value'])); entry=ttk.Entry(dialog,textvariable=value); entry.pack(fill='x',padx=12); entry.focus_set(); entry.selection_range(0,'end')
            self.label(dialog,'修改原因（必填）'); reason=tk.StringVar(); ttk.Entry(dialog,textvariable=reason).pack(fill='x',padx=12)
            self.label(dialog,'保存后旧版本保留；此条记录需要重新判定P/F。')
            def close(): dialog.destroy(); win.grab_set()
            dialog.protocol('WM_DELETE_WINDOW',close)
            def save():
                new_id=self.store.correct_record(old_id,field['area'],field['key'],value.get(),reason.get())
                active['id']=new_id; close(); render(); self.refresh()
                if new_id!=old_id: self.status.set(f'已保存修改：记录{old_id} → {new_id}。请重新判定P/F并导出Summary。')
            bar=ttk.Frame(dialog); bar.pack(side='bottom',fill='x',padx=8,pady=10)
            self.button(bar,'保存修改',save,style='Primary.TButton'); self.button(bar,'取消',close)
        self.button(footer,'修改选中字段',edit,style='Primary.TButton'); self.button(footer,'关闭',win.destroy)
        table.bind('<Double-1>',lambda _:self.guard(edit)); render()
        if initial_field:
            match=next((i for i,f in enumerate(entries) if f['area']=='meta' and f['key']==initial_field),None)
            if match is not None:
                table.selection_set(str(match)); table.see(str(match)); edit()

    def batch_review_dialog(self, action):
        panel=self.require_panel();ids=list(map(int,self.pending_tree.selection()))
        plan=self.store.batch_review_plan(panel,ids,action)
        word='采用' if action=='adopt' else '不采用'
        win=tk.Toplevel(self);win.title('批量'+word+'重分析结果');win.geometry('1100x700');win.transient(self);win.grab_set()
        footer=ttk.Frame(win);footer.pack(side='bottom',fill='x',padx=12,pady=10)
        reason=tk.StringVar();ttk.Label(footer,text='统一处理原因（必填）').pack(anchor='w')
        ttk.Entry(footer,textvariable=reason).pack(fill='x',pady=5)
        old_fail=tk.BooleanVar(value=False)
        if action=='adopt':ttk.Checkbutton(footer,text='同时把被替代的旧结果标为F（仅确认失败时勾选）',variable=old_fail).pack(anchor='w',pady=4)
        bar=ttk.Frame(footer);bar.pack(fill='x')
        eligible=[r for r in plan if r['eligible']]
        self.label(win,f'选中{len(plan)}条，可{word}{len(eligible)}条，需另行处理{len(plan)-len(eligible)}条。旧记录和来源保留。')
        self.label(win,'批量采用仅按同一样品同管匹配；同组选择多个候选时跳过，不自动挑最新。rpt1/rpt2分别保留。')
        details=tk.Text(win,height=5,wrap='word');details.pack(side='bottom',fill='x',padx=12,pady=6)
        table=self.tree(win,['状态','候选ID','样品编号','Tube','Run','旧ID','说明','来源文件'],[70,80,145,135,90,70,350,300],height=10)
        for i,row in enumerate(plan):table.insert('','end',iid=str(i),values=['可处理' if row['eligible'] else '跳过',row['id'],row['bas'],row['tube'],row['run'],row['old_id'] or '',row['detail'],'；'.join(Path(f).name for f in row['files'])],tags=() if row['eligible'] else ('warning',))
        def show(_=None):
            if not table.selection():return
            row=plan[int(table.selection()[0])]
            lines=[f'{c["field"]}：{txt(c["old"])} → {txt(c["new"])}' for c in row['changes']]
            details.configure(state='normal');details.delete('1.0','end');details.insert('end','\n'.join(lines) or row['detail']);details.configure(state='disabled')
        table.bind('<<TreeviewSelect>>',show)
        def commit():
            why=txt(reason.get())
            if not why:raise ValueError('请填写统一处理原因')
            if not eligible:raise ValueError('没有可处理的记录，请根据提示重新选择')
            expected={r['id']:r['old_id'] for r in plan};fail=old_fail.get()
            if not messagebox.askyesno('确认批量'+word,f'将{word}{len(eligible)}条候选，跳过{len(plan)-len(eligible)}条。是否继续？',parent=win):return
            win.destroy()
            def done(result):
                messagebox.showinfo('批量审核完成',f'已处理{result["processed"]}条，跳过{result["skipped"]}条。\n跳过项仍保留在待审核中。',parent=self)
                self.status.set(f'批量{word}：处理{result["processed"]}条，跳过{result["skipped"]}条。')
            self.run_job('批量'+word+'重分析',lambda s:s.batch_review(panel,ids,action,why,fail,expected),done)
        self.button(bar,'确认批量'+word,commit,style='Primary.TButton');self.button(bar,'取消',win.destroy)

    def review_selected(self,tree):
        ident=self.selected_id(tree);rec=self.store.record(ident);payload=json.loads(rec['payload'])
        win=tk.Toplevel(self);win.title(f'记录{ident} · {STATE_LABEL[rec["state"]]}');win.geometry('1100x720');win.transient(self);win.grab_set()
        self.label(win,f'{sample_display(payload)} / {payload["meta"]["Tube"]} / {payload["meta"]["Run number"]}   P/F：{rec["pf"] or "未判定"}')
        if payload.get('manual_edits'): self.label(win,'人工修改记录：'+rec['reason']+'；完整旧值及修改原因见历史与操作日志。')
        candidates=[r for r in self.store.records(rec['panel_id'],['current']) if r['id']!=ident and sample_identity(json.loads(r['payload']))==sample_identity(payload) and not distinct_replicates(payload,json.loads(r['payload'])) and (norm(json.loads(r['payload'])['meta']['Tube'])=='isotype')==(norm(payload['meta']['Tube'])=='isotype')]
        targets={'无指定旧记录（同管名自动匹配）':None}
        default=list(targets)[0]
        for r in candidates:
            m=json.loads(r['payload'])['meta'];label=f'ID {r["id"]} | {m["Tube"]} | {m["Run number"]} | {m["Date"]} {m["Time"]}'
            targets[label]=r['id']
            if r['logical']==rec['logical']:default=label
        top=ttk.Frame(win);top.pack(fill='x',padx=12)
        ttk.Label(top,text='要替代的旧记录').pack(side='left')
        target_var=tk.StringVar(value=default);box=ttk.Combobox(top,textvariable=target_var,values=list(targets),state='readonly',width=90);box.pack(side='left',padx=10)
        details=self.tree(win,['字段','选中的新/候选值','旧值'],[320,360,360])
        details.tag_configure('diff',background='#fff1cf')
        def fill(*_):
            details.delete(*details.get_children());target=targets[target_var.get()]
            old=self.store.record(target) if target else None;op=json.loads(old['payload']) if old else None
            left={**{'[信息] '+k:v for k,v in payload['meta'].items()},**payload['raw'],'[每管Beads]':rec['beads'],'[血量µL]':rec['volume'],'[P/F]':rec['pf']}
            right={**{'[信息] '+k:v for k,v in op['meta'].items()},**op['raw'],'[每管Beads]':old['beads'],'[血量µL]':old['volume'],'[P/F]':old['pf']} if op else {}
            for k in dict.fromkeys(list(left)+list(right)):
                a,b=left.get(k),right.get(k)
                details.insert('','end',values=[payload['labels'].get(k,k),txt(a),txt(b)],tags=('diff',) if old and a!=b else ())
        box.bind('<<ComboboxSelected>>',fill);fill()
        sources=self.store.db.execute('SELECT file,sheet,line FROM sources WHERE record_id=?',(ident,)).fetchall()
        source_text='\n'.join(f'{r[0]} / {r[1]} / 行{r[2]}' for r in sources)
        src=tk.Text(win,height=3,wrap='word');src.pack(fill='x',padx=12);src.insert('1.0',source_text);src.configure(state='disabled')
        bottom=ttk.Frame(win);bottom.pack(fill='x',padx=12,pady=10)
        ttk.Label(bottom,text='处理原因').grid(row=0,column=0,sticky='w')
        reason=tk.StringVar();ttk.Entry(bottom,textvariable=reason,width=72).grid(row=0,column=1,padx=8,sticky='w')
        old_fail=tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom,text='同时把旧结果P/F标为F（仅在你确认其失败时勾选）',variable=old_fail).grid(row=1,column=1,sticky='w',pady=8)
        bar=ttk.Frame(win);bar.pack(fill='x',padx=8)
        def approve():
            why=txt(reason.get())
            if not why:raise ValueError('请填写处理原因')
            if not messagebox.askyesno('确认采用','采用此记录作为有效结果；指定旧结果将保留为已被替代。继续？',parent=win):return
            self.store.activate(ident,why,targets[target_var.get()],old_fail.get());win.destroy();self.refresh()
        def reject():
            self.store.reject(ident,reason.get());win.destroy();self.refresh()
        if rec['state'] in ('current','pending'):self.button(bar,'采用此结果／替代指定旧记录',approve)
        if rec['state']=='pending':self.button(bar,'不采用此候选',reject)
        self.button(bar,'关闭',win.destroy)

    def pf_dialog(self,tree):
        ident=self.selected_id(tree)
        pf=simpledialog.askstring('设置P/F','输入P或F；输入空字符串清除判定：',initialvalue=self.store.record(ident)['pf'],parent=self)
        if pf is None:return
        pf=pf.strip().upper()
        if pf not in ('','P','F'):raise ValueError('只能输入P、F或空白')
        reason=simpledialog.askstring('判定原因','填写判定或修改原因：',parent=self)
        if not reason:return
        self.store.set_pf(ident,pf,reason);self.refresh()

    def selected_subject_issue(self):
        selected=self.subject_tree.selection()
        if len(selected)!=1 or not selected[0].startswith('issue:'): return None
        _,audit_id,index=selected[0].split(':')
        return next((r for r in self.store.subject_import_issues(self.pid) if r['audit_id']==int(audit_id) and r['index']==int(index)),None)

    def subject_dialog(self,new=False):
        self.require_panel(); selected=() if new else self.subject_tree.selection(); issue=None if new else self.selected_subject_issue()
        existing=(issue.get('item') or issue['row']) if issue else next((r for r in self.store.subject_rows(self.pid) if selected and r['样品编号']==selected[0]),{})
        if issue and issue.get('header_error'): raise ValueError('此问题是表头缺失或重名；请先修改Excel表头并重新导入清单')
        win=tk.Toplevel(self);win.title('修正受试者问题行' if issue else '受试者信息');win.transient(self);win.grab_set()
        variables={};existing=split_subject(existing)
        if issue:
            ttk.Label(win,text=issue['detail'],foreground='#A2373E',wraplength=640).grid(row=0,column=0,columnspan=2,padx=12,pady=8)
        offset=1 if issue else 0
        for i,h in enumerate(SUBJECT_SPLIT_VIEW):
            color='#A2373E' if issue and (h in issue['errors'] or h in ('采血阶段','给药前后','间隔时长') and 'Time Point' in issue['errors']) else '#1C1C1E'
            ttk.Label(win,text=SUBJECT_SPLIT_LABELS.get(h,h),foreground=color).grid(row=i+offset,column=0,padx=16,pady=5,sticky='w')
            variables[h]=tk.StringVar(value=txt(existing.get(h)))
            if h=='给药前后':
                entry=ttk.Combobox(win,textvariable=variables[h],values=['','Pre','Post'],width=46)
            else: entry=ttk.Entry(win,textvariable=variables[h],width=48)
            entry.grid(row=i+offset,column=1,padx=16,pady=5)
        sync_visit=tk.BooleanVar(value=False)
        ttk.Checkbutton(win,text='同时修正该样品各Panel的访视信息',variable=sync_visit).grid(row=len(SUBJECT_SPLIT_VIEW)+offset,column=0,columnspan=2,padx=16,pady=8,sticky='w')
        def save():
            row=join_subject({h:v.get() for h,v in variables.items()})
            reason=''
            if sync_visit.get():
                reason=simpledialog.askstring('同步修正原因','只修正访视/采血资料；不改样品编号或受试者编号。请输入原因：',parent=win)
                if not txt(reason): return
            if issue: self.store.resolve_subject_issue(self.pid,issue['audit_id'],issue['index'],row,reason,_sync_visit=sync_visit.get())
            else: self.store.put_subjects(self.pid,[row],_sync_visit=sync_visit.get(),reason=reason)
            win.destroy(); self.refresh()
        ctk.CTkButton(win,text='修正并保存' if issue else '保存',fg_color='#2B74CF',hover_color='#2468BB',corner_radius=9,height=36,font=(self.ui_font,14),command=lambda:self.guard(save)).grid(row=len(SUBJECT_SPLIT_VIEW)+offset+1,column=1,sticky='e',padx=16,pady=16)

    def dismiss_subject_issue(self):
        self.require_panel(); issue=self.selected_subject_issue()
        if not issue: raise ValueError('请先选中一条红色待修正行；有效资料不能在此删除')
        reason=simpledialog.askstring('不采用问题行','填写不采用原因；该行仍保留在操作日志中，不作为正式受试者资料：',parent=self)
        if not txt(reason): return
        self.store.resolve_subject_issue(self.pid,issue['audit_id'],issue['index'],None,reason);self.refresh()

    def show_subject_import_issues(self, result=None):
        self.require_panel(); self.refresh(); self.open_section('受试者库')
        issues=self.store.subject_import_issues(self.pid)
        if result is not None:
            messagebox.showinfo('受试者清单已处理',f'正常处理 {result["accepted"]} 行，更新 {result["updated"]} 行，待修正 {result["failed"]} 行。\n红色行已保留在受试者库，可双击修正；不会覆盖有效资料。',parent=self)
        elif not issues: messagebox.showinfo('受试者清单','没有待修正行。',parent=self)
        if issues:
            iid=f'issue:{issues[0]["audit_id"]}:{issues[0]["index"]}'
            self.subject_tree.selection_set(iid);self.subject_tree.see(iid)

    def import_subject_dialog(self):
        self.require_panel();pid=self.pid
        p=filedialog.askopenfilename(title='选择受试者Excel',filetypes=[('Excel','*.xlsx')],parent=self)
        if p:self.run_job('逐行导入受试者清单',lambda s:s.import_subjects(pid,p),self.show_subject_import_issues)

    def preview_report(self):
        panel=self.require_panel()
        if self.config_dirty:
            if not messagebox.askyesno('保存配置','先保存当前汇报项配置再预览？',parent=self):return
            self.save_config_ui()
        exclude=self.exclude_fail.get()
        def work(store):
            cfg=store.config(panel)
            prepared=store.prepare_report(panel,exclude,_check_pending=False,_allow_partial=True)
            rows=[]
            for rec,p,values,calc,_ in prepared:
                display=[]
                for c,v in zip(cfg['columns'],values):
                    if isinstance(v,tuple):v=calc[v[1]][6]
                    display.append(formatted_report_value(v,c))
                rows.append(dict(iid=str(rec['id']),values=['验证样品' if p.get('sample_kind')=='validation' else 'Summary',p.get('sample_name','')]+[rec['pf'] if h=='P/F' else p['meta'][h] for h in META]+display))
            return cfg,rows
        def done(result):
            cfg,rows=result
            win=tk.Toplevel(self);win.title('报告预览');win.geometry('1200x640');win.transient(self)
            self.label(win,f'符合条件共{len(rows)}条，分页查看。ABS单位为cells/µL；按已保存的汇报精度显示，计算使用未舍入数据。')
            self.label(win,'搜索仅影响预览；缺失项留空，原因见导出报告的待处理问题。待审核候选不自动采用。')
            table=self.tree(win,['输出工作表','验证样品名称']+META+[c['name'] for c in cfg['columns']])
            PagedTable(self,table,run_column=2+META.index('Run number')).set_rows(rows,{},None)
        self.run_job('准备报告预览',work,done,refresh=False)

    def preview_cv(self):
        panel=self.require_panel()
        if self.config_dirty:
            if not messagebox.askyesno('保存配置','先保存当前配置再计算CV？',parent=self):return
            self.save_config_ui()
        panel_name=self.panel_box.get()
        def work(store):
            return replicate_cv_rows(store.prepare_report(panel,False,_check_pending=False,_allow_partial=True),store.config(panel),panel_name,_allow_partial=True)
        def done(rows):
            win=tk.Toplevel(self);win.title('同Run复管CV');win.geometry('1200x650');win.transient(self)
            self.label(win,'CV = 样本SD(n−1) / 均值 × 100%。包括当前F复管；阈值判断不自动改变P/F。')
            if not rows:self.label(win,'没有rpt复管数据，或未启用CV计算。')
            table=self.tree(win,CV_HEADERS)
            shown=[dict(iid=str(i),values=['' if v is None else (f'{v:.3f}' if isinstance(v,float) else v) for v in row]) for i,row in enumerate(rows)]
            PagedTable(self,table).set_rows(shown,{},None)
        self.run_job('计算复管CV',work,done,refresh=False)

    def export_dialog(self):
        panel=self.require_panel()
        if self.config_dirty:
            if not messagebox.askyesno('保存配置','汇报项有未保存修改。先保存，再导出？',parent=self):return
            self.save_config_ui()
        p=filedialog.asksaveasfilename(title='导出当前Panel',initialfile=self.panel_box.get()+'_Summary.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')],parent=self)
        if p:
            exclude=self.exclude_fail.get()
            self.run_job('生成Excel报告',lambda s:s.export(panel,p,exclude),lambda n:messagebox.showinfo('导出完成',f'已导出{n}条当前有效结果。\n{p}\n\n验证样品单独导出，不含ABS。留空项见“汇总问题”，未纳入文件见“批次文件提示”。\n请用Excel打开重算ABS公式；修改Excel不会回写数据库。',parent=self))

    def audit_dialog(self):
        self.require_panel();win=tk.Toplevel(self);win.title('项目操作日志');win.geometry('1000x600')
        text=tk.Text(win,wrap='word');text.pack(fill='both',expand=True)
        for r in self.store.db.execute('SELECT * FROM audit WHERE project_id=? ORDER BY id DESC',(self.pid,)):
            text.insert('end',f'{r["id"]}  {r["stamp"]}  {r["action"]}\n{r["detail"]}\n\n')
        text.configure(state='disabled')

    def backup_db(self):
        p=filedialog.asksaveasfilename(title='备份完整数据库',initialfile='express7_backup.sqlite3',defaultextension='.sqlite3',filetypes=[('SQLite数据库','*.sqlite3')],parent=self)
        if p:self.store.backup(p);messagebox.showinfo('备份完成',p,parent=self)

    def switch_db(self):
        if not self.discard_config():return
        path=self.choose_database_path()
        if path:
            new=Store(path);self.store.close();self.store=new;self.mark_config_dirty(False);self.refresh_projects()
            self.path_label.configure(text='数据库：'+str(self.store.path))

    def close_app(self):
        if self.discard_config():self.store.close();self.destroy()


if __name__ == '__main__':
    try:
        App().mainloop()
    except StartupCancelled:
        pass
    except Exception as error:
        traceback.print_exc()
        try:
            root=tk.Tk();root.withdraw();messagebox.showerror('启动失败',str(error));root.destroy()
        except Exception:
            pass
