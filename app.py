import io, os, re, json, shutil, subprocess, tempfile
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="CX AI Reporting Agent", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

DEFAULTS = {"aht":90, "within1":98, "within2":98, "abandon":2, "exclude":["Counter Staff"]}

# ---------- Helpers ----------
def clean_name(x):
    return re.sub(r"\s+", " ", str(x).strip()) if pd.notna(x) else ""

def parse_hms(x):
    if pd.isna(x): return 0.0
    s=str(x).strip()
    try:
        p=s.split(":")
        if len(p)==3: return int(p[0])*3600+int(p[1])*60+float(p[2])
        if len(p)==2: return int(p[0])*60+float(p[1])
        return float(s)
    except: return 0.0

def hms(sec):
    sec=max(0,int(round(float(sec or 0))))
    return f"{sec//3600:02d}:{(sec%3600)//60:02d}:{sec%60:02d}"

def read_table(upload):
    """Read CSV/XLSX/XLS exports.

    Some Intellicon/legacy BIFF .xls exports contain harmless BIFF records that
    newer xlrd versions flag as workbook corruption (e.g. "seen[2] == 4").
    We first ask xlrd to tolerate those records, then fall back to LibreOffice
    conversion when available.
    """
    b=upload.getvalue(); n=upload.name.lower()
    if n.endswith('.csv'):
        return pd.read_csv(io.BytesIO(b), low_memory=False)
    if n.endswith('.xlsx'):
        return pd.read_excel(io.BytesIO(b), engine='openpyxl')
    if n.endswith('.xls'):
        # Primary path: xlrd supports legacy .xls and can ignore benign
        # workbook-corruption markers produced by some call-center exports.
        try:
            return pd.read_excel(
                io.BytesIO(b),
                engine='xlrd',
                engine_kwargs={'ignore_workbook_corruption': True}
            )
        except Exception as first_error:
            # Local/server fallback: LibreOffice can repair/re-save many BIFF
            # workbooks that xlrd refuses. This keeps the web app compatible
            # with the exact .xls export shown in the user's screenshot.
            soffice = shutil.which('libreoffice') or shutil.which('soffice')
            if soffice:
                with tempfile.TemporaryDirectory() as td:
                    src=os.path.join(td, os.path.basename(upload.name))
                    with open(src, 'wb') as f: f.write(b)
                    # Convert directly to CSV. Some legacy BIFF exports become
                    # malformed XLSX files when LibreOffice writes NaN numeric
                    # cells; CSV avoids that second parser failure entirely.
                    cmd=[soffice,'--headless','--convert-to','csv','--outdir',td,src]
                    proc=subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60)
                    converted=os.path.join(td, os.path.splitext(os.path.basename(upload.name))[0]+'.csv')
                    if proc.returncode == 0 and os.path.exists(converted):
                        try:
                            df=pd.read_csv(converted, low_memory=False)
                            return df.replace({np.nan: None})
                        except Exception:
                            pass
            raise ValueError(
                'The Agent Summary .xls file uses a legacy BIFF format that could not be read. '
                'Please re-save it as .xlsx, or install/update xlrd. Original error: ' + str(first_error)
            )
    raise ValueError('Supported formats: CSV, XLS, XLSX')

def norm_key(x):
    """Normalize an agent name for matching across systems that spell it
    differently (dots vs spaces, extra whitespace, case)."""
    s = str(x or '').lower().replace('.', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def read_tickets_table(upload):
    """Read a tickets/CRM export and return per-agent Info/Complaint/SR/Total counts.

    Supports two formats seen in the wild:
      1) Pre-aggregated 'Report Agents Activity - Care Connect' export:
         UserName / TotalTicket / InfoCount / ComplaintCount / ServiceCount
      2) Raw ticket-level 'CRM Report - Care Connect' export (one row per
         ticket): 'Created By' (agent) + 'Category' (Info / Service Request /
         Complaints), which we aggregate ourselves.
    """
    b = upload.getvalue(); n = upload.name.lower()
    if n.endswith('.csv'):
        raw = pd.read_csv(io.BytesIO(b), header=None, low_memory=False)
    elif n.endswith('.xlsx'):
        raw = pd.read_excel(io.BytesIO(b), engine='openpyxl', header=None)
    elif n.endswith('.xls'):
        try:
            raw = pd.read_excel(io.BytesIO(b), engine='xlrd', header=None,
                                 engine_kwargs={'ignore_workbook_corruption': True})
        except Exception:
            raw = pd.read_excel(io.BytesIO(b), engine='xlrd', header=None)
    else:
        raise ValueError('Supported formats: CSV, XLS, XLSX')

    header_row = None; mode = None
    for i in range(min(10, len(raw))):
        vals = raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if 'username' in vals:
            header_row = i; mode = 'aggregated'; break
        if 'created by' in vals and 'category' in vals:
            header_row = i; mode = 'raw'; break
    if header_row is None:
        raise ValueError("Tickets report must contain either a 'UserName' column "
                          "(pre-summarized export) or 'Created By' + 'Category' "
                          "columns (raw ticket-log export).")

    headers = raw.iloc[header_row].astype(str).str.strip().tolist()
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how='all')

    colmap = {c.lower(): c for c in df.columns}
    def col(*names):
        for nm in names:
            if nm.lower() in colmap: return colmap[nm.lower()]
        return None

    if mode == 'aggregated':
        c_user = col('UserName')
        c_total = col('TotalTicket')
        c_info = col('InfoCount')
        c_complaint = col('ComplaintCount')
        c_service = col('ServiceCount')
        missing = [lbl for lbl, c in [('UserName', c_user), ('TotalTicket', c_total),
                   ('InfoCount', c_info), ('ComplaintCount', c_complaint),
                   ('ServiceCount', c_service)] if c is None]
        if missing:
            raise ValueError('Tickets report missing columns: ' + ', '.join(missing))
        out = pd.DataFrame({
            'agent_name': df[c_user].map(clean_name),
            'info': pd.to_numeric(df[c_info], errors='coerce').fillna(0),
            'complaints': pd.to_numeric(df[c_complaint], errors='coerce').fillna(0),
            'sr': pd.to_numeric(df[c_service], errors='coerce').fillna(0),
            'tickets_total': pd.to_numeric(df[c_total], errors='coerce').fillna(0),
        })
    else:  # mode == 'raw'
        c_user = col('Created By')
        c_cat = col('Category')
        df = df[df[c_user].notna()].copy()
        df['agent_name'] = df[c_user].map(clean_name)
        cat = df[c_cat].fillna('').astype(str).str.strip().str.lower()
        df['_info'] = cat.eq('info').astype(int)
        df['_complaint'] = cat.eq('complaints').astype(int)
        df['_sr'] = cat.eq('service request').astype(int)
        out = df.groupby('agent_name', as_index=False).agg(
            info=('_info', 'sum'), complaints=('_complaint', 'sum'), sr=('_sr', 'sum'))
        out['tickets_total'] = out['info'] + out['complaints'] + out['sr']

    out = out[out.agent_name.str.len() > 0]
    out['key'] = out.agent_name.map(norm_key)
    out = out.groupby('key', as_index=False).agg(
        agent_name=('agent_name', 'first'), info=('info', 'sum'),
        complaints=('complaints', 'sum'), sr=('sr', 'sum'),
        tickets_total=('tickets_total', 'sum'))
    return out

def shift_from_login(x, m_start=6, e_start=15, n_start=2):
    """Classify an agent's shift from their login hour, using three
    configurable start-hours (0-23). Handles wraparound past midnight
    automatically regardless of which shift wraps."""
    if pd.isna(x): return 'Unknown'
    h=x.hour + x.minute/60.0
    bounds=sorted([('Morning',m_start),('Evening',e_start),('Night',n_start)],key=lambda t:t[1])
    for i in range(3):
        name,start=bounds[i]
        nxt=bounds[(i+1)%3][1]
        if nxt>start:
            if start<=h<nxt: return name
        else:  # this shift wraps past midnight
            if h>=start or h<nxt: return name
    return 'Unknown'

def duration_bucket(s):
    return pd.cut(s,bins=[-0.01,60,120,180,np.inf],labels=['0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min'],include_lowest=True)

def wait_bucket(s):
    return pd.cut(s,bins=[-0.01,15,30,60,120,np.inf],labels=['0 - 15 SEC','16 SEC - 30 SEC','31 SEC - 1 MIN','1.1 MIN - 2 MIN','ABOVE 2 MIN'],include_lowest=True)

DEFAULT_ALIAS_GROUPS=[
    {'faisal mehmood','faisal mahmood'},
    {'m hassan','muhammad hasaan'},
]

def build_alias_map(groups):
    amap={}
    for g in groups:
        canon=sorted(g)[0]
        for v in g: amap[v]=canon
    return amap

def apply_aliases(key, alias_map):
    return alias_map.get(key, key)

def compute_shift_summary(m):
    if len(m):
        shift=m.groupby('shift',dropna=False).agg(agents=('agent_name','count'),calls=('calls','sum'),tickets_total=('tickets_total','sum'),avg_aht_sec=('aht_sec','mean')).reset_index()
        shift['avg_aht']=shift.avg_aht_sec.map(hms)
        shift['diff_qty']=shift.tickets_total-shift.calls
        order=['Morning','Evening','Night','Unknown']; shift['sort']=shift['shift'].map({x:i for i,x in enumerate(order)}); shift=shift.sort_values('sort').drop(columns='sort')
    else: shift=pd.DataFrame(columns=['shift','agents','calls','tickets_total','avg_aht_sec','avg_aht','diff_qty'])
    return shift

def build_report(queue, summary, cfg, tickets=None):
    q=queue.copy(); a=summary.copy()
    if 'disposition' not in q.columns: raise ValueError("Queue report must contain 'disposition'.")
    for c in ['agent_name','event']:
        if c not in q.columns: q[c]=''
    q['agent_name']=q['agent_name'].map(clean_name)
    q['event_clean']=q['event'].fillna('').astype(str).str.strip().str.lower()
    q['disp_clean']=q['disposition'].fillna('').astype(str).str.strip().str.upper()
    q['call_duration_sec']=pd.to_numeric(q.get('call_duration',0),errors='coerce').fillna(0)
    q['queue_holdtime_sec']=pd.to_numeric(q.get('queue_holdtime',0),errors='coerce').fillna(0)
    q['call_started_dt']=pd.to_datetime(q.get('call_started'),errors='coerce')

    answered=q[q.disp_clean.eq('ANSWERED')].copy()
    ivr=q[q.event_clean.eq('ivr hangup')].copy()
    abandoned=q[(q.event_clean.eq('abandoned')) & (~q.event_clean.eq('ivr hangup'))].copy()

    if len(a):
        a=a.copy()
        required=['agentName','loginTime','breakTime','channelName']
        missing=[x for x in required if x not in a.columns]
        if missing: raise ValueError('Agent Summary missing: '+', '.join(missing))
        a['agentName']=a.agentName.map(clean_name); a['channelName']=a.channelName.fillna('').astype(str).str.lower()
        a=a[a.channelName.eq('call')].copy()
        a['login_dt']=pd.to_datetime(a.get('loginStart'),errors='coerce')
        a['login_sec']=a.loginTime.map(parse_hms); a['break_sec']=a.breakTime.map(parse_hms)
        a['active_sec']=(a.login_sec-a.break_sec).clip(lower=0)
        a['login']=a.login_sec.map(hms); a['break']=a.break_sec.map(hms); a['active']=a.active_sec.map(hms)
        a['shift']=a.login_dt.map(lambda x: shift_from_login(x, cfg.get('morning_start',6), cfg.get('evening_start',15), cfg.get('night_start',2)))
        a=a.drop_duplicates('agentName')
    else: a=pd.DataFrame(columns=['agentName','shift','login','break','active'])

    calls=answered.groupby('agent_name').size().rename('calls')
    aht=answered.groupby('agent_name').call_duration_sec.mean().rename('aht_sec')
    buckets=answered.assign(bucket=duration_bucket(answered.call_duration_sec)).groupby(['agent_name','bucket'],observed=False).size().unstack(fill_value=0)
    for col in ['0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min']:
        if col not in buckets.columns: buckets[col]=0
    buckets=buckets[['0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min']]
    buckets.index=buckets.index.astype(str)

    m=a.rename(columns={'agentName':'agent_name'}).copy()
    m['agent_name']=m['agent_name'].astype(str)
    m=m.merge(calls.rename_axis('agent_name').reset_index(),on='agent_name',how='left').merge(aht.rename_axis('agent_name').reset_index(),on='agent_name',how='left').merge(buckets.reset_index().rename(columns={'index':'agent_name'}),on='agent_name',how='left')
    for c in ['calls','aht_sec','0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min']:
        m[c]=m[c].fillna(0)
    m['calls']=m.calls.astype(int)
    m['aht']=m.aht_sec.map(hms)

    cutting=answered[answered.call_duration_sec<=0].groupby('agent_name').size().rename('call_cutting')
    m=m.merge(cutting.rename_axis('agent_name').reset_index(),on='agent_name',how='left')
    m['call_cutting']=m['call_cutting'].fillna(0).astype(int)

    tickets_unmatched=pd.DataFrame(columns=['agent_name','info','complaints','sr','tickets_total'])
    if tickets is not None and len(tickets):
        t=tickets.copy()
        alias_map=build_alias_map(DEFAULT_ALIAS_GROUPS+(cfg.get('alias_groups') or []))
        t['key']=t['key'].map(lambda k: apply_aliases(k, alias_map))
        m['key']=m.agent_name.map(norm_key).map(lambda k: apply_aliases(k, alias_map))
        matched_keys=set(t.key) & set(m.key)
        tickets_unmatched=t[~t.key.isin(m.key)][['agent_name','info','complaints','sr','tickets_total']]
        m=m.merge(t[['key','info','complaints','sr','tickets_total']],on='key',how='left')
        m['has_tickets']=m.key.isin(matched_keys)
        m=m.drop(columns='key')
    else:
        m['info']=np.nan; m['complaints']=np.nan; m['sr']=np.nan; m['tickets_total']=np.nan
        m['has_tickets']=False

    for c in ['info','complaints','sr','tickets_total']:
        m[c]=pd.to_numeric(m[c],errors='coerce')
    m['diff_qty']=np.where(m.has_tickets, m.tickets_total.fillna(0)-m.calls, np.nan)
    def _status(row):
        if not row['has_tickets']: return 'No Ticket Data'
        if row['diff_qty']>0: return 'More Added'
        if row['diff_qty']<0: return 'Not All Generated'
        return 'All Added'
    m['status']=m.apply(_status,axis=1) if len(m) else pd.Series(dtype=str)

    excluded={clean_name(x).lower() for x in cfg['exclude']}
    eligible=answered[~answered.agent_name.str.lower().isin(excluded)]
    within1=int((eligible.queue_holdtime_sec<=60).sum()); within2=int((eligible.call_duration_sec<=120).sum())
    denom=len(eligible); within1_pct=within1/denom*100 if denom else 0; within2_pct=within2/denom*100 if denom else 0
    offered=len(answered)+len(abandoned); abandon_pct=len(abandoned)/offered*100 if offered else 0
    overall=answered.call_duration_sec.mean() if len(answered) else 0
    waits=wait_bucket(q.queue_holdtime_sec)
    wait_counts=pd.Series(waits).value_counts(sort=False).reindex(['0 - 15 SEC','16 SEC - 30 SEC','31 SEC - 1 MIN','1.1 MIN - 2 MIN','ABOVE 2 MIN'],fill_value=0)

    shift=compute_shift_summary(m)

    date=q.call_started_dt.dropna().dt.date.min() if q.call_started_dt.notna().any() else datetime.now().date()
    return dict(date=str(date), q=q, answered=answered, abandoned=abandoned, ivr=ivr, agents=m, shift=shift, waits=wait_counts.to_dict(), answered_n=len(answered), abandoned_n=len(abandoned), ivr_n=len(ivr), offered=offered, abandon_pct=abandon_pct, overall=overall, within1=within1, within1_pct=within1_pct, within2=within2, within2_pct=within2_pct, denom=denom, tickets_unmatched=tickets_unmatched)

def deterministic_ai(r,cfg):
    a=r['agents']; lines=[]
    lines.append(f"### Executive Summary\n{r['answered_n']:,} answered calls were analyzed on {r['date']}. Overall AHT was **{hms(r['overall'])}**. Non-IVR abandonment was **{r['abandoned_n']:,} ({r['abandon_pct']:.1f}%)**, while **{r['ivr_n']:,} IVR hangups** were kept separate.")
    lines.append(f"### KPI Findings\n- Within 1 minute: **{r['within1']:,} ({r['within1_pct']:.1f}%)** vs target {cfg['within1']}%.\n- Within 2 minutes: **{r['within2']:,} ({r['within2_pct']:.1f}%)** vs target {cfg['within2']}%.\n- AHT target: **{cfg['aht']} sec**.")
    if len(a):
        top=a.sort_values('calls',ascending=False).iloc[0]; high=a.sort_values('aht_sec',ascending=False).head(3)
        lines.append(f"### Agent Highlights\n- Highest call volume: **{top.agent_name}** ({int(top.calls):,} calls).\n- Highest AHT: " + ', '.join(f"{x.agent_name} ({x.aht})" for _,x in high.iterrows()) + '.')
        over=a[a.aht_sec>cfg['aht']]
        lines.append(f"### Recommended Actions\n- Review **{len(over)} agent(s)** above the AHT target.\n- Focus coaching on calls exceeding 2 minutes and periods with elevated queue waiting time.\n- Review abandonment separately from IVR hangups so service-level decisions use the correct denominator.")
    return '\n\n'.join(lines)

def ai_summary(r,cfg):
    key=os.getenv('OPENAI_API_KEY')
    if not key or OpenAI is None: return deterministic_ai(r,cfg)
    payload={"date":r['date'],"answered":r['answered_n'],"abandoned":r['abandoned_n'],"ivr_hangups":r['ivr_n'],"abandon_pct":round(r['abandon_pct'],2),"overall_aht":hms(r['overall']),"within1_pct":round(r['within1_pct'],2),"within2_pct":round(r['within2_pct'],2),"agents":r['agents'][['agent_name','shift','calls','aht']].to_dict('records')}
    prompt=f"You are a CX call-center operations manager. Analyze this daily report. Do not invent causes. Write Executive Summary, Key Findings, Agents to Review, Recommended Actions. Targets: AHT <= {cfg['aht']} sec, within1 >= {cfg['within1']}%, within2 >= {cfg['within2']}%, abandonment <= {cfg['abandon']}%. DATA: {json.dumps(payload)}"
    try:
        client=OpenAI(api_key=key)
        resp=client.responses.create(model=os.getenv('OPENAI_MODEL','gpt-4o-mini'),input=prompt)
        return resp.output_text
    except Exception:
        return deterministic_ai(r,cfg)

def _base_wb(title):
    wb=Workbook(); ws=wb.active; ws.title=title; ws.sheet_view.showGridLines=False
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    return wb, ws

def _cell(ws, r, c, value=None, fill=None, font=None, align='center', border=None, numfmt=None):
    x=ws.cell(r,c,value)
    if fill: x.fill=PatternFill('solid',fgColor=fill)
    if font: x.font=font
    x.alignment=Alignment(horizontal=align,vertical='center',wrap_text=True)
    if border: x.border=border
    if numfmt: x.number_format=numfmt
    return x

def _merge(ws,r,c1,c2,value,fill,font,size=None,border=None):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    f=font
    if size is not None: f=Font(name=f.name or 'Calibri',size=size,bold=f.bold,color=f.color)
    _cell(ws,r,c1,value,fill=fill,font=f,border=border)
    if border:
        for c in range(c1,c2+1): ws.cell(r,c).border=border

def _style_headers(ws,r,headers,fill,border):
    for c,h in enumerate(headers,1): _cell(ws,r,c,h,fill=fill,font=Font(bold=True,color='FFFFFF'),border=border)

def _apply_widths(ws,widths):
    for c,w in widths.items(): ws.column_dimensions[get_column_letter(c)].width=w
    ws.sheet_view.zoomScale=85

def _save_wb(wb):
    out=io.BytesIO(); wb.save(out); return out.getvalue()

def export_dashboard_excel(r, meta=None):
    meta=meta or {}; wb,ws=_base_wb('CX Daily Dashboard')
    navy='244F78'; blue='9DC3E6'; white='FFFFFF'; green='C6EFCE'; border=Border(*(Side(style='thin',color='000000'),)*4)
    _merge(ws,1,1,13,'CX DAILY AGENT PERFORMANCE DASHBOARD',blue,Font(bold=True),size=15,border=border); ws.row_dimensions[1].height=26
    labels=['Date','Total Agents','Present','Leave','Rest','Absent','Overall Avg AHT']; vals=[r['date'],meta.get('total_agents',len(r['agents'])),meta.get('present',len(r['agents'])),meta.get('leave',0),meta.get('rest',0),meta.get('absent',0),hms(r['overall'])]
    spans=[(1,2),(3,4),(5,6),(7,7),(8,8),(9,9),(10,13)]
    for (c1,c2),lab,val in zip(spans,labels,vals):
        _merge(ws,2,c1,c2,lab,navy,Font(bold=True,color=white),border=border)
        _merge(ws,3,c1,c2,val,'FFFFFF',Font(bold=True,size=12),border=border)
    agents=r['agents'].copy(); order=['Morning','Evening','Night','Unknown']
    present_shifts=[s for s in order if not agents[agents['shift'].fillna('Unknown').eq(s)].empty]
    first_shift=present_shifts[0] if present_shifts else None
    agent_header='Agent\n'+first_shift if first_shift else 'Agent'
    headers=['Sr#',agent_header,'Login Time','Break Time','Active Time','AHT','Calls\nAttended','Complaints','Info','SR','Tickets\nTotal','Difference','Status']
    _style_headers(ws,5,headers,navy,border); ws.row_dimensions[5].height=34
    row=6; sr=1
    for shift in order:
        g=agents[agents['shift'].fillna('Unknown').eq(shift)]
        if g.empty: continue
        if shift!=first_shift:
            _merge(ws,row,1,13,shift,'FFFFFF',Font(bold=True,size=11),border=border); row+=1
        for _,x in g.iterrows():
            has_t=bool(x.get('has_tickets',False))
            comp=int(x['complaints']) if has_t else ''
            info=int(x['info']) if has_t else ''
            svc=int(x['sr']) if has_t else ''
            tix=int(x['tickets_total']) if has_t else ''
            diff=int(x['diff_qty']) if has_t else ''
            status=x['status'] if has_t else 'No Ticket Data'
            vals=[sr,x.agent_name,x.get('login','00:00:00'),x.get('break','00:00:00'),x.get('active','00:00:00'),x.aht,int(x.calls),comp,info,svc,tix,diff,status]
            for c,v in enumerate(vals,1): _cell(ws,row,c,v,border=border)
            ws.cell(row,2).alignment=Alignment(horizontal='left',vertical='center')
            if status=='Not All Generated':
                ws.cell(row,13).font=Font(bold=True,color='FF0000')
            elif status in ('All Added','More Added'):
                ws.cell(row,13).fill=PatternFill('solid',fgColor=green); ws.cell(row,13).font=Font(color='006100')
            else:
                ws.cell(row,13).font=Font(italic=True,color='808080')
            row+=1; sr+=1
    row+=1
    _merge(ws,row,1,6,'Shift Summary',navy,Font(bold=True,color=white),border=border); row+=1
    _style_headers(ws,row,['Shift','Agents','Calls','Tickets','Difference','Avg AHT'],navy,border); row+=1
    for _,x in r['shift'].iterrows():
        vals=[x['shift'],int(x['agents']),int(x['calls']),int(x['tickets_total']),int(x['diff_qty']),x['avg_aht']]
        for c,v in enumerate(vals,1): _cell(ws,row,c,v,border=border)
        row+=1
    s_agents=int(r['shift']['agents'].sum()) if len(r['shift']) else len(agents)
    s_calls=int(r['shift']['calls'].sum()) if len(r['shift']) else total_calls
    s_tix=int(r['shift']['tickets_total'].sum()) if len(r['shift']) else 0
    vals=['Total',s_agents,s_calls,s_tix,s_tix-s_calls,hms(r['overall'])]
    for c,v in enumerate(vals,1): _cell(ws,row,c,v,fill=navy,font=Font(bold=True,color=white),border=border)
    _apply_widths(ws,{1:10,2:24,3:13,4:13,5:13,6:11,7:13,8:12,9:10,10:10,11:14,12:13,13:18}); ws.freeze_panes='A5'; ws.print_title_rows='1:5'
    return _save_wb(wb)

def export_call_handling_excel(r, meta=None):
    meta=meta or {}; wb,ws=_base_wb('Call Handling Report')
    peach='FCE4D6'; blue='DDEBF7'; green='E2F0D9'; red='FF0000'; redfill='FFC7CE'; navy='244F78'; white='FFFFFF'; border=Border(*(Side(style='thin',color='000000'),)*4)
    _merge(ws,1,1,11,'Agents Call Handling Time Analysis',peach,Font(bold=True),size=14,border=border)
    _cell(ws,1,12,'Date',fill=peach,font=Font(bold=True),border=border); _cell(ws,1,13,r['date'],fill=peach,font=Font(bold=True),border=border); _cell(ws,1,14,'',fill=peach,border=border)
    headers=['Sr#','Agent Name','UAN','0 - 1 min','%','1.1- 2 Min','%','2.1- 3 Min','%','Above 3 min','%','Call Cutting','Total','Receiving Age%']
    _style_headers(ws,2,headers,blue,border); ws.row_dimensions[2].height=32
    row=3; agents=r['agents'].copy(); order=['Morning','Evening','Night','Unknown']; sr=1
    for shift in order:
        g=agents[agents['shift'].fillna('Unknown').eq(shift)]
        if g.empty: continue
        _merge(ws,row,1,14,shift,green,Font(bold=True),border=border); row+=1
        for _,x in g.iterrows():
            total=int(x.calls); b=[int(x[c]) for c in ['0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min']]
            cutting=int(x.get('call_cutting',0))
            vals=[sr,x.agent_name,'UAN',b[0],b[0]/total if total else 0,b[1],b[1]/total if total else 0,b[2],b[2]/total if total else 0,b[3],b[3]/total if total else 0,cutting,total,1]
            for c,v in enumerate(vals,1): _cell(ws,row,c,v,border=border)
            for c in [5,7,9,11,14]: ws.cell(row,c).number_format='0%'
            ws.cell(row,2).alignment=Alignment(horizontal='left',vertical='center'); row+=1; sr+=1
        total=int(g.calls.sum()); bs=[int(g[c].sum()) for c in ['0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min']]
        cut_shift=int(g.get('call_cutting',pd.Series(dtype=int)).sum())
        vals=[f'{shift} Total','UAN','',bs[0],bs[0]/total if total else 0,bs[1],bs[1]/total if total else 0,bs[2],bs[2]/total if total else 0,bs[3],bs[3]/total if total else 0,cut_shift,total,1]
        for c,v in enumerate(vals,1): _cell(ws,row,c,v,fill=green,font=Font(bold=True),border=border)
        for c in [5,7,9,11,14]: ws.cell(row,c).number_format='0%'
        row+=1
    total=int(agents.calls.sum()) if len(agents) else 0; bs=[int(agents[c].sum()) for c in ['0 - 1 min','1.1 - 2 min','2.1 - 3 min','Above 3 min']]
    cut_grand=int(agents.get('call_cutting',pd.Series(dtype=int)).sum()) if len(agents) else 0
    vals=['Grand Total','UAN','',bs[0],bs[0]/total if total else 0,bs[1],bs[1]/total if total else 0,bs[2],bs[2]/total if total else 0,bs[3],bs[3]/total if total else 0,cut_grand,total,1]
    for c,v in enumerate(vals,1): _cell(ws,row,c,v,fill=blue,font=Font(bold=True),border=border)
    for c in [5,7,9,11,14]: ws.cell(row,c).number_format='0%'
    row+=2
    _merge(ws,row,1,14,'Note: All agents should try to handle the query within 2 minutes.',red,Font(bold=True,color=white),border=border); row+=2
    target_pct=meta.get('within2_target',98)/100; target_calls=round(r['answered_n']*target_pct); achieved=r['within2']; achieved_pct=r['within2_pct']/100
    _merge(ws,row,1,4,"KPI's",'FFFFFF',Font(bold=True,size=12),border=border); row+=1
    for c,h in enumerate(['','Calls','%age'],1): _cell(ws,row,c,h,font=Font(bold=True),border=border)
    row+=1
    data=[('Target',target_calls,target_pct),('Achieved',achieved,achieved_pct),('Variance',achieved-target_calls,achieved_pct-target_pct)]
    for label,calls,pct in data:
        fill=redfill if label=='Variance' and calls<0 else ('C6EFCE' if label=='Variance' else 'FFFFFF')
        _cell(ws,row,1,label,fill=fill,border=border); _cell(ws,row,2,calls,fill=fill,border=border); _cell(ws,row,3,pct,fill=fill,border=border,numfmt='0%'); row+=1
    _apply_widths(ws,{1:7,2:23,3:11,4:12,5:8,6:12,7:8,8:12,9:8,10:13,11:8,12:13,13:10,14:14}); ws.freeze_panes='A3'; ws.print_title_rows='1:2'
    return _save_wb(wb)

def export_ivr_excel(r, meta=None):
    meta=meta or {}; wb,ws=_base_wb('IVR Report')
    yellow='FFF2CC'; blue='DDEBF7'; navy='244F78'; green='C6EFCE'; redfill='FFC7CE'; white='FFFFFF'; border=Border(*(Side(style='thin',color='000000'),)*4)
    _merge(ws,1,1,7,f"Analysis of Call Waiting (IVR) Time ({r['date']})",yellow,Font(bold=True),size=13,border=border)
    headers=['Time - Slot','Answered','%age','Abandoned','%age','Grand Total','%age']; _style_headers(ws,2,headers,blue,border)
    row=3
    # IVR waiting-time analysis is based ONLY on calls actually answered by an agent.
    # All abandoned calls and IVR-hangup calls are excluded from the denominator and buckets.
    q=r['q'].copy(); q['wait_bucket']=wait_bucket(q.queue_holdtime_sec)
    ivr_answered=q[(q.disp_clean=='ANSWERED') & (~q.event_clean.eq('abandoned')) & (~q.event_clean.eq('ivr hangup'))].copy()
    total=len(ivr_answered)
    for slot in ['0 - 15 SEC','16 SEC - 30 SEC','31 SEC - 1 MIN','1.1 MIN - 2 MIN','ABOVE 2 MIN']:
        a=int((ivr_answered.wait_bucket.astype(str)==slot).sum()); ab=0; gt=a
        vals=[slot,a,a/total if total else 0,ab,0,gt,gt/total if total else 0]
        for c,v in enumerate(vals,1): _cell(ws,row,c,v,border=border)
        for c in [3,5,7]: ws.cell(row,c).number_format='0%'
        row+=1
    vals=['Grand Total',r['answered_n'],r['answered_n']/total if total else 0,r['abandoned_n'],r['abandoned_n']/total if total else 0,r['answered_n']+r['abandoned_n'],(r['answered_n']+r['abandoned_n'])/total if total else 0]
    for c,v in enumerate(vals,1): _cell(ws,row,c,v,fill=blue,font=Font(bold=True),border=border)
    for c in [3,5,7]: ws.cell(row,c).number_format='0%'
    row+=2
    target1=meta.get('within1_target',98)/100; actual1=r['within1_pct']/100; target1_calls=round(r['denom']*target1)
    targetA=meta.get('abandon_target',2)/100; actualA=r['abandon_pct']/100; targetA_calls=round((r['answered_n']+r['abandoned_n'])*targetA)
    for c0,title,target,actual,tgt_calls,act_calls in [(1,'Call Answered (Within 1 Min)',target1,actual1,target1_calls,r['within1']),(5,'Abandoned',targetA,actualA,targetA_calls,r['abandoned_n'])]:
        _merge(ws,row,c0,c0+2,title,'FFFFFF',Font(bold=True),border=border); rr=row+1
        for c,h in enumerate(['','%age','Calls'],c0): _cell(ws,rr,c,h,font=Font(bold=True),border=border)
        rr+=1
        for label,pct,calls in [('Target',target,tgt_calls),('Achieved',actual,act_calls),('Variance',actual-target,act_calls-tgt_calls)]:
            fill='FFFFFF'
            if label=='Variance': fill=green if ((title.startswith('Call') and pct>=0) or (title=='Abandoned' and pct<=0)) else redfill
            _cell(ws,rr,c0,label,fill=fill,border=border); _cell(ws,rr,c0+1,pct,fill=fill,border=border,numfmt='0%'); _cell(ws,rr,c0+2,calls,fill=fill,border=border); rr+=1
    _apply_widths(ws,{1:25,2:12,3:10,4:12,5:10,6:14,7:10}); ws.freeze_panes='A3'; ws.print_title_rows='1:2'
    return _save_wb(wb)

def export_full_zip(r, meta=None):
    import zipfile
    meta=meta or {}
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"CX_Daily_Dashboard_{r['date']}.xlsx", export_dashboard_excel(r,meta))
        z.writestr(f"Agents_Call_Handling_{r['date']}.xlsx", export_call_handling_excel(r,meta))
        z.writestr(f"IVR_Call_Waiting_{r['date']}.xlsx", export_ivr_excel(r,meta))
    return out.getvalue()

def export_excel(r, ai, meta=None):
    return export_dashboard_excel(r,meta)

# ---------- Styling ----------
st.markdown('''<style>
.main .block-container{padding-top:1.2rem;padding-bottom:2rem;max-width:1450px}
.hero{padding:22px 26px;border-radius:16px;background:linear-gradient(135deg,#163b63,#285f8f);color:white;margin-bottom:18px;box-shadow:0 8px 24px rgba(0,0,0,.12)}
.hero h1{margin:0;font-size:32px}.hero p{margin:6px 0 0;opacity:.9}
[data-testid="stMetric"]{background:#fff;border:1px solid #e6eaf0;border-radius:12px;padding:12px;box-shadow:0 2px 10px rgba(0,0,0,.04)}
</style>''',unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>🤖 CX AI Call Center Reporting Agent</h1><p>Upload your daily Queue + Agent Summary files and generate the CX dashboard, handling-time analysis, IVR report, KPIs and AI management insights.</p></div>',unsafe_allow_html=True)

with st.sidebar:
    st.header('⚙️ Reporting Rules')
    cfg={'aht':st.number_input('AHT target (seconds)',30,600,DEFAULTS['aht']), 'within1':st.number_input('Within 1 minute target (%)',50,100,DEFAULTS['within1']), 'within2':st.number_input('Within 2 minutes target (%)',50,100,DEFAULTS['within2']), 'abandon':st.number_input('Abandon target (%)',0.0,20.0,float(DEFAULTS['abandon'])), 'exclude':[x.strip() for x in st.text_input('Exclude from KPI denominator','Counter Staff').split(',') if x.strip()]}; meta={'total_agents':st.number_input('Total agents',1,200,17), 'present':st.number_input('Present',0,200,12), 'leave':st.number_input('Leave',0,200,2), 'rest':st.number_input('Rest',0,200,2), 'absent':st.number_input('Absent',0,200,1), 'within1_target':98, 'within2_target':98, 'abandon_target':2}
    st.divider(); st.caption('IVR hangups are always kept separate from actual abandoned calls.')
    st.divider()
    st.caption('Automatic shift suggestion only — you can manually assign every agent below. Evening duty starts at 3:00 PM by default; manual assignment always overrides login-time grouping.')
    sc1,sc2,sc3=st.columns(3)
    cfg['morning_start']=sc1.number_input('Morning starts',0,23,6)
    cfg['evening_start']=sc2.number_input('Evening starts',0,23,15)
    cfg['night_start']=sc3.number_input('Night starts',0,23,2)
    st.divider()
    st.caption('Name aliases: same agent spelled differently across files. Built-in: Faisal Mehmood/Mahmood, M Hassan/Muhammad Hasaan.')
    alias_text=st.text_area('Extra aliases (one pair per line, "Name A = Name B")','',help='Example:\nJohn Smith = Jon Smith')
    extra_groups=[]
    for line in alias_text.splitlines():
        if '=' in line:
            a,b=line.split('=',1); a,b=norm_key(a),norm_key(b)
            if a and b: extra_groups.append({a,b})
    cfg['alias_groups']=extra_groups

u1,u2,u3=st.columns(3)
with u1: qfile=st.file_uploader('📥 Queue Report',type=['csv','xls','xlsx'],help='Export from Intellicon/call-center system.')
with u2: afile=st.file_uploader('📥 Agent Summary Report',type=['csv','xls','xlsx'],help='Agent login, break and channel summary.')
with u3: tfile=st.file_uploader('📥 Tickets Report (Care Connect)',type=['csv','xls','xlsx'],help='Optional. "Report Agents Activity - Care Connect" export with Complaints/Info/SR ticket counts. Without it, those columns show as No Ticket Data.')

if qfile and afile:
    try:
        tdf=read_tickets_table(tfile) if tfile else None
        r=build_report(read_table(qfile),read_table(afile),cfg,tickets=tdf)
        if tfile and len(r['tickets_unmatched']):
            names=', '.join(r['tickets_unmatched'].agent_name.tolist())
            st.warning(f"⚠️ {len(r['tickets_unmatched'])} agent(s) in the Tickets Report could not be matched to the Agent Summary (name spelling mismatch?): {names}")
        st.success(f"Report generated successfully for {r['date']}. Assign each agent's actual duty shift below; the selected shifts will be used for all summaries and Excel exports.")

        # Manual shift assignment is the source of truth for reporting.
        # The login-time shift is only used as the initial/default suggestion.
        st.subheader('🕒 Shift Assignment')
        st.caption(
            'Select the actual duty shift for each agent. Manual assignments override the '
            'automatic login-time grouping and are used in the Daily Dashboard, Call Handling '
            'report, IVR report and Shift Summary. Evening duty starts at **3:00 PM** by default.'
        )
        shift_key = 'shift_assignment_editor'
        base = r['agents'][['agent_name','shift']].copy()
        base.columns = ['Agent', 'Shift']
        edited = st.data_editor(
            base,
            use_container_width=True,
            hide_index=True,
            key=shift_key,
            column_config={
                'Agent': st.column_config.TextColumn('Agent', disabled=True),
                'Shift': st.column_config.SelectboxColumn(
                    'Duty Shift',
                    options=['Morning','Evening','Night','Unknown'],
                    required=True
                )
            },
            disabled=['Agent']
        )
        overrides = dict(zip(edited['Agent'].astype(str), edited['Shift'].astype(str)))
        r['agents']['shift'] = r['agents']['agent_name'].map(overrides).fillna(r['agents']['shift'])
        r['shift'] = compute_shift_summary(r['agents'])

        # Clear visual confirmation of the configured shift rules.
        sm = r['shift'].copy()
        if len(sm):
            sm['calls'] = sm['calls'].astype(int)
            st.markdown(
                '<div style="padding:10px 14px;border-radius:10px;background:#eef6ff;'
                'border:1px solid #c9def5;margin:8px 0 14px 0;">'
                '<b>Current shift calculation:</b> Morning → Evening → Night according to '
                'your manual assignments. Evening agents are counted in the Evening summary '
                'regardless of their login time.'
                '</div>',
                unsafe_allow_html=True
            )

        cols=st.columns(6)
        cols[0].metric('Answered',f"{r['answered_n']:,}")
        cols[1].metric('Non-IVR Abandoned',f"{r['abandoned_n']:,}",f"{r['abandon_pct']:.1f}%")
        cols[2].metric('IVR Hangups',f"{r['ivr_n']:,}")
        cols[3].metric('Overall AHT',hms(r['overall']))
        cols[4].metric('Within 1 Min',f"{r['within1_pct']:.1f}%",f"Target {cfg['within1']}%")
        cols[5].metric('Within 2 Min',f"{r['within2_pct']:.1f}%",f"Target {cfg['within2']}%")

        t1,t2,t3,t4,t5=st.tabs(['📊 Daily Dashboard','⏱️ Call Handling','☎️ IVR / Queue','🤖 AI Insights','💬 Ask the Agent'])
        with t1:
            d=r['agents'][['agent_name','shift','login','break','active','aht','calls','complaints','info','sr','tickets_total','diff_qty','status']].copy()
            d.columns=['Agent','Shift','Login','Break','Active','AHT','Calls','Complaints','Info','SR','Tickets Total','Difference','Status']  # diff_qty renamed for display only
            st.dataframe(d,use_container_width=True,hide_index=True)
            st.subheader('Shift Summary'); st.dataframe(r['shift'][['shift','agents','calls','tickets_total','diff_qty','avg_aht']],use_container_width=True,hide_index=True)
            c1,c2=st.columns(2)
            with c1:
                chart=r['agents'].groupby('shift',dropna=False).calls.sum().reset_index(); st.plotly_chart(px.pie(chart,names='shift',values='calls',title='Calls by Shift',hole=.45),use_container_width=True)
            with c2:
                st.plotly_chart(px.bar(r['agents'].sort_values('calls',ascending=False).head(10),x='agent_name',y='calls',title='Top 10 Agents by Calls'),use_container_width=True)
        with t2:
            bh=r['answered'].copy(); bh['bucket']=duration_bucket(bh.call_duration_sec); bc=bh.groupby('bucket',observed=False).size().reset_index(name='Calls')
            c1,c2=st.columns(2)
            with c1: st.plotly_chart(px.bar(bc,x='bucket',y='Calls',title='Calls by Handling-Time Slot'),use_container_width=True)
            with c2: st.plotly_chart(px.bar(r['agents'].sort_values('aht_sec',ascending=False),x='agent_name',y='aht_sec',title='Agent AHT (seconds)'),use_container_width=True)

            st.subheader('Calls Answered: Within 1 Min / Within 2 Min / Above 2 Min (per agent)')
            wa=r['agents'].copy()
            wa['Within 1 Min']=wa['0 - 1 min']
            wa['Within 2 Min']=wa['1.1 - 2 min']
            wa['Above 2 Min']=wa['2.1 - 3 min']+wa['Above 3 min']
            wt=wa[['agent_name','shift','calls','Within 1 Min','Within 2 Min','Above 2 Min']].rename(columns={'agent_name':'Agent','shift':'Shift','calls':'Total Calls'})
            st.dataframe(wt,use_container_width=True,hide_index=True)
            tot=wt[['Total Calls','Within 1 Min','Within 2 Min','Above 2 Min']].sum()
            k1,k2,k3,k4=st.columns(4)
            k1.metric('Total Calls',f"{int(tot['Total Calls']):,}")
            k2.metric('Within 1 Min',f"{int(tot['Within 1 Min']):,}",f"{tot['Within 1 Min']/tot['Total Calls']*100:.1f}%" if tot['Total Calls'] else None)
            k3.metric('Within 2 Min',f"{int(tot['Within 2 Min']):,}",f"{tot['Within 2 Min']/tot['Total Calls']*100:.1f}%" if tot['Total Calls'] else None)
            k4.metric('Above 2 Min',f"{int(tot['Above 2 Min']):,}",f"{tot['Above 2 Min']/tot['Total Calls']*100:.1f}%" if tot['Total Calls'] else None)
            st.plotly_chart(px.bar(wt,x='Agent',y=['Within 1 Min','Within 2 Min','Above 2 Min'],title='Calls Handled by Time Band per Agent',barmode='stack'),use_container_width=True)

            over=r['agents'][r['agents'].aht_sec>cfg['aht']][['agent_name','aht','calls']]
            st.subheader(f"Agents above {cfg['aht']} sec AHT target")
            st.dataframe(over.rename(columns={'agent_name':'Agent','aht':'AHT','calls':'Calls'}),use_container_width=True,hide_index=True)
        with t3:
            # IVR waiting analysis counts answered calls only.
            ivr_answered=r['q'][(r['q'].disp_clean=='ANSWERED') & (~r['q'].event_clean.eq('abandoned')) & (~r['q'].event_clean.eq('ivr hangup'))].copy()
            ivr_answered['wait_bucket']=wait_bucket(ivr_answered.queue_holdtime_sec)
            slots=['0 - 15 SEC','16 SEC - 30 SEC','31 SEC - 1 MIN','1.1 MIN - 2 MIN','ABOVE 2 MIN']
            w=ivr_answered['wait_bucket'].value_counts().reindex(slots,fill_value=0).rename_axis('Time Slot').reset_index(name='Calls')
            st.plotly_chart(px.bar(w,x='Time Slot',y='Calls',title='Analysis of Call Waiting (IVR) Time — Answered Calls Only'),use_container_width=True)
            k1,k2,k3=st.columns(3)
            k1.metric('Answered Calls',f"{len(ivr_answered):,}")
            k2.metric('Within 1 Min',f"{r['within1']:,} / {r['denom']:,}")
            k3.metric('Excluded',f"{r['abandoned_n']+r['ivr_n']:,}",'Abandoned + IVR Hangup')
            st.info(f"IVR waiting-time call numbers exclude all abandoned calls ({r['abandoned_n']:,}) and IVR-hung-up calls ({r['ivr_n']:,}). Only answered calls are included.")
        with t4:
            ai=ai_summary(r,cfg); st.markdown(ai)
            st.download_button('⬇️ Download Complete Report — 3 Excel Files',export_full_zip(r,meta),f"CX_Daily_Reports_{r['date']}.zip",'application/zip')
            st.caption('Each report is also available separately in the exact Excel layout of your supplied templates.')
            d1,d2,d3=st.columns(3)
            with d1: st.download_button('📊 Daily Dashboard Excel',export_dashboard_excel(r,meta),f"CX_Daily_Dashboard_{r['date']}.xlsx",'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='dl_dashboard')
            with d2: st.download_button('⏱️ Call Handling Excel',export_call_handling_excel(r,meta),f"Agents_Call_Handling_{r['date']}.xlsx",'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='dl_handling')
            with d3: st.download_button('☎️ IVR Report Excel',export_ivr_excel(r,meta),f"IVR_Call_Waiting_{r['date']}.xlsx",'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='dl_ivr')
        with t5:
            st.markdown('Ask questions such as **Who had the highest AHT?**, **Which shift handled the most calls?**, or **Why is the 2-minute KPI below target?**')
            q=st.text_input('Your question',key='agent_question')
            if q:
                # Useful local agent for common report questions without requiring an API key.
                ql=q.lower(); a=r['agents']
                if 'highest aht' in ql or 'max aht' in ql:
                    x=a.sort_values('aht_sec',ascending=False).iloc[0]; st.success(f"{x.agent_name} has the highest AHT at {x.aht}.")
                elif 'most calls' in ql or 'highest calls' in ql or 'top agent' in ql:
                    x=a.sort_values('calls',ascending=False).iloc[0]; st.success(f"{x.agent_name} handled the most calls: {int(x.calls):,}.")
                elif 'shift' in ql and ('most' in ql or 'highest' in ql):
                    x=r['shift'].sort_values('calls',ascending=False).iloc[0]; st.success(f"{x['shift']} shift handled the most calls: {int(x['calls']):,}.")
                elif 'abandon' in ql:
                    st.info(f"Non-IVR abandonment is {r['abandoned_n']:,} calls ({r['abandon_pct']:.1f}%). IVR hangups ({r['ivr_n']:,}) are excluded.")
                elif '2 min' in ql or 'two minute' in ql:
                    st.info(f"{r['within2']:,} of {r['denom']:,} eligible calls were handled within 2 minutes ({r['within2_pct']:.1f}%). Target is {cfg['within2']}%.")
                else:
                    st.markdown(ai_summary(r,cfg))
    except Exception as e:
        st.error(f'Could not generate report: {e}')
else:
    st.info('Upload both files above to generate the web report.')
    st.markdown('### What you will get')
    st.write('Daily dashboard • agent rankings • shift summary • AHT analysis • 0–1 / 1.1–2 / 2.1–3 / >3 minute buckets • IVR waiting analysis • non-IVR abandonment • KPI variance • AI management summary • Excel export • report Q&A.')
