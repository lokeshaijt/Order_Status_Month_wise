"""Row-role detection on a month sheet of the template."""
import openpyxl  # noqa: F401
def roles(ws,gc=7,fc=6):
    """Row roles on a month sheet. gc = column of Pending Ord Qty, fc = column of Prod Name."""
    out={}; n=ws.max_row; r=3
    while r<=n:
        c=str(ws.cell(r,3).value or '')
        if c.startswith('TOTAL'): out.setdefault('total',r); r+=1; continue
        g=ws.cell(r,gc).value
        if isinstance(g,str) and g.startswith('=SUM') and str(ws.cell(r,fc).value or '')==' ':
            out.setdefault('sub',r); r+=1; continue
        r+=1
    r=3
    while r<=n:
        if isinstance(ws.cell(r,gc).value,(int,float)) and (r==3 or ws.cell(r-1,gc).value in (None,0)):
            e=r
            while isinstance(ws.cell(e,gc).value,(int,float)): e+=1
            if e-r>=3 and 'first' not in out:
                out['first']=r; out['mid']=r+1; out['last']=e-1; out['sub3']=e; out['spacer']=e+1
            if e-r==1 and 'single' not in out: out['single']=r; out['single_sub']=e; out['single_spacer']=e+1
            r=e+2; continue
        r+=1
    return out
def sig(c):
    b=c.border; f=c.fill; fo=c.font; a=c.alignment
    return (''.join((s[0] if s else '-') for s in (b.left.style,b.right.style,b.top.style,b.bottom.style)),
            f.fgColor.rgb if f.fill_type else None, fo.sz, bool(fo.b), fo.color.rgb if fo.color and fo.color.type=='rgb' else (fo.color.theme if fo.color else None), a.horizontal, a.vertical, bool(a.wrap_text), c.number_format)
