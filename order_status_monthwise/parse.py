"""Parse order blocks in the ERP sheet "Order Status-By shipment Date"."""


def blocks(ws):
    """Return order blocks: {'ord', 'rows' (product rows), 'sub' (subtotal row or None)}.

    Product row  = column I (Pending Ord Qty) non-empty.
    Subtotal row = column I is ' ' and column K starts with =SUM.
    The last block of the ERP export has no subtotal row / merges.
    """
    out = []
    cur = None
    for r in range(7, ws.max_row + 1):
        i = ws.cell(r, 9).value
        a = ws.cell(r, 1).value
        isprod = i is not None and str(i).strip() != ""
        issub = (i is not None and str(i).strip() == ""
                 and isinstance(ws.cell(r, 11).value, str) and str(ws.cell(r, 11).value).startswith("=SUM"))
        if isprod:
            if cur is None or (a not in (None, "") and str(a).strip() != "" and str(a) != cur["ord"]):
                cur = {"ord": str(a).strip() if a not in (None, "") else None, "rows": [], "sub": None}
                out.append(cur)
            cur["rows"].append(r)
        elif issub and cur:
            cur["sub"] = r
            cur = None
    return out
