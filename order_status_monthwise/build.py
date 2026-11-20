"""Build the month-wise Order Status report from the ERP workbook.

Input : OrderStatusRep(ort)-CBE_RODTEP-_CBE_EXP_<dd.mm.yyyy>.xlsx  (sheet "Order Status-By shipment Date")
Output: Order_Status_Monthwise_<dd-mm-yyyy>.xlsx  (one sheet per month: past months with orders,
        running month, +1, +2; orders beyond running+2 are left out and counted).
Format: copied from the template workbook's "Nov-26" sheet (widths, borders, fonts, row heights).
"""
import datetime as dt
import re
from copy import copy

import openpyxl
from openpyxl.formula.translate import Translator  # noqa: F401  (kept for parity with earlier builds)
from openpyxl.styles import Alignment, Border
from openpyxl.utils import get_column_letter as L

from .parse import blocks
from .styles import roles

TEMPLATE_SHEET = "Nov-26"


def build(SRC_PATH, REPORT, OUT_PATH, TPL_PATH):
    run=(REPORT.year,REPORT.month)
    def addm(ym,n): m=ym[0]*12+ym[1]-1+n; return (m//12,m%12+1)
    nxt2=addm(run,2)
    MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    lab=lambda ym:f"{MON[ym[1]-1]}-{str(ym[0])[2:]}"; full=lambda ym:f"{MON[ym[1]-1]} {ym[0]}"

    wb=openpyxl.load_workbook(TPL_PATH); tp=wb[TEMPLATE_SHEET]; old=list(wb.sheetnames)
    R=roles(tp)
    src=openpyxl.load_workbook(SRC_PATH)['Order Status-By shipment Date']
    MAXC=src.max_column
    hdr={c:str(src.cell(6,c).value) for c in range(1,MAXC+1)}
    PL=[c for c in hdr if hdr[c].startswith('Planned in')]; assert hdr[PL[-1]+1]=='Planned Qty Total' and hdr[PL[-1]+2]=='Un Planned' and MAXC==PL[-1]+2
    NPL=len(PL); print('planned weeks',NPL,[hdr[c] for c in PL])
    SC=[1,2,3,4,5,6,10,11,12,13,14,15,16]+list(range(22,MAXC+1)); ND=len(SC)   # Buyer Contact No shown (dest col 5); container/ship/vsl/posted/location dropped
    def TC(dc):   # destination column -> template (user's Nov-26) column
        if dc<=4: return dc
        if dc==5: return 4            # Buyer Contact No takes NAV Doc No's style/width
        if dc<=13: return dc-1        # Req date .. CAN PACK
        return 13                     # planned weeks, Planned Qty Total, Un Planned: same format as first planned column
    D=lambda sc: SC.index(sc)+1
    COLDIM={}
    for k,d in tp.column_dimensions.items():
        for c in range(d.min,d.max+1):
            COLDIM[L(c)]=(d.width,d.hidden)
    bl=blocks(src)
    for b in bl:
        b['date']=dt.datetime.strptime(next(str(src.cell(r,6).value) for r in b['rows'] if str(src.cell(r,6).value).strip()),'%d-%m-%Y').date()
        b['ym']=(b['date'].year,b['date'].month); b['last']=b['sub'] if b['sub'] else b['rows'][-1]
    by={}
    for b in bl: by.setdefault(b['ym'],[]).append(b)
    months=sorted(set(k for k in by if k<=nxt2)|{run,addm(run,1),nxt2})
    left_out=[b for b in bl if b['ym']>nxt2]
    merges=list(src.merged_cells.ranges)
    ref_sub=next(bb for bb in bl if bb['sub'])['sub']
    SUMD=list(range(8,ND+1))           # all numeric columns, incl. CAN PACK (13)
    TAB=lambda k:'A6A6A6' if k<run else ('00B050' if k==run else '4F81BD')
    COLMAP={L(sc):L(i+1) for i,sc in enumerate(SC)}      # source column letter -> destination column letter
    def remap(f,drow):
        """Re-point every cell reference in a source formula to its destination column (and shifted row)."""
        def sub(m):
            col,row=m.group(1),int(m.group(2))
            if col not in COLMAP: raise ValueError(f'formula {f} references column {col} which is not in the report')
            return f"{COLMAP[col]}{row+drow}"
        return re.sub(r'(?<![A-Za-z])([A-Z]{1,3})(\d+)(?![\d(])',sub,f)
    def sty(ws,dr,tr):
        for dc in range(1,ND+1): ws.cell(dr,dc)._style=copy(tp.cell(tr,TC(dc))._style)
        c=ws.cell(dr,13); al=c.alignment; c.alignment=Alignment(horizontal=al.horizontal,vertical=al.vertical,wrap_text=True)   # CAN PACK wraps like every other column
    stats={}
    for k in months:
        ws=wb.create_sheet('new_'+lab(k)); ws.sheet_properties.tabColor=TAB(k)
        for dc in range(1,ND+1):
            w,h=COLDIM[L(TC(dc))]; ws.column_dimensions[L(dc)].width=w; ws.column_dimensions[L(dc)].hidden=h
        period=full(k)
        # row 1 title, row 2 header
        sty(ws,1,1); ws.row_dimensions[1].height=tp.row_dimensions[1].height
        ws['C1']=f"Order Status-Shipment Date Wise On {REPORT:%d-%m-%Y}  |  Buyer Requested Shipment: {period}  "
        ws.merge_cells('C1:G1')
        ws.merge_cells('I1:L1'); ws['I1']='No. of Sales Order'
        sty(ws,2,2); ws.row_dimensions[2].height=tp.row_dimensions[2].height
        for dc in range(1,ND+1): ws.cell(2,dc).value=src.cell(6,SC[dc-1]).value
        def put(sr,dr):
            for dc,sc in enumerate(SC,1):
                s=src.cell(sr,sc)
                if type(s).__name__=='MergedCell' or s.value is None: continue
                v=s.value
                if isinstance(v,str) and v.startswith('='): v=remap(v,dr-sr)
                ws.cell(dr,dc).value=v
        row=3; subs=[]; nl=0
        for b in sorted(by.get(k,[]),key=lambda b:b['date']):
            n=len(b['rows']); first=row
            for i,sr in enumerate(b['rows']):
                tr=R['single'] if n==1 else (R['first'] if i==0 else (R['last'] if i==n-1 else R['mid']))
                sty(ws,row,tr); put(sr,row); ws.row_dimensions[row].height=20.25 if i==0 else 21.0
                row+=1; nl+=1
            sty(ws,row,R['single_sub'] if n==1 else R['sub3'])
            if b['sub']: put(b['sub'],row)
            else:
                for dc in SUMD: ws.cell(row,dc).value=f"=SUM({L(dc)}{first}:{L(dc)}{row-1})"
                ws.cell(row,7).value=' '
            ws.cell(row,13).value=f"=SUM(M{first}:M{row-1})"   # CAN PACK subtotal (not in ERP file)
            subs.append(row); row+=1
            o=first-b['rows'][0]
            for m in merges:
                if b['rows'][0]<=m.min_row and m.max_row<=b['last'] and m.min_col==m.max_col and m.min_col in SC:
                    dc=D(m.min_col); ws.merge_cells(start_row=m.min_row+o,start_column=dc,end_row=m.max_row+o,end_column=dc)
            if n>1 and not any(b['rows'][0]<=m.min_row and m.max_row<=b['last'] for m in merges):
                # block left unmerged by the ERP export: merge Buyer Name / NAV Doc No down the product rows and centre (as in user's Nov-26)
                for rr in range(first+1,first+n):
                    ws.cell(rr,6)._style=copy(ws.cell(first,6)._style)
                    for dc in (1,2): ws.cell(rr,dc).border=copy(ws.cell(first,dc).border)
                for dc in (1,2,3,4,5):
                    bo=ws.cell(first+n,dc).border; ws.cell(first+n,dc).border=Border(left=copy(bo.left),right=copy(bo.right),top=copy(ws.cell(first,dc).border.top),bottom=copy(bo.bottom))
                ws.cell(first+n,6)._style=copy(ws.cell(first,6)._style)
                for dc in (3,4,5):
                    for rr in range(first+1,first+n): ws.cell(rr,dc).value=None
                    ws.merge_cells(start_row=first,start_column=dc,end_row=first+n-1,end_column=dc)
                    al=copy(ws.cell(first,dc).alignment); ws.cell(first,dc).alignment=Alignment(horizontal='center',vertical=al.vertical,wrap_text=al.wrap_text)
            sty(ws,row,R['single_spacer'] if n==1 else R['spacer']); ws.cell(row,12).value=0; row+=1
        sty(ws,row,R['total']); ws.cell(row,3).value=f"TOTAL - {period}"
        for dc in SUMD: ws.cell(row,dc).value=("="+"+".join(f"{L(dc)}{s}" for s in subs)) if subs else 0
        ws['M1']=f'=SUMPRODUCT((LEN(TRIM(D3:D{row-1}))>0)/COUNTIF(D3:D{row-1},D3:D{row-1}&""))'; ws['M1'].number_format='0'
        ws.freeze_panes='A3'; ws.auto_filter.ref=f'C2:{L(ND)}2'
        stats[k]=dict(orders=len(by.get(k,[])),lines=nl,gt=row)
    for n in old: del wb[n]
    for k in months: wb['new_'+lab(k)].title=lab(k)
    wb.active=wb.sheetnames.index(lab(run))
    for ws in wb: ws.sheet_view.tabSelected=(ws.title==lab(run))
    wb.save(OUT_PATH)
    print('saved',{lab(k):(v['orders'],v['lines']) for k,v in stats.items()}); print('left out (after',full(nxt2)+')',len(left_out),'orders')
    return {lab(k):v for k,v in stats.items()}

