import re
from smartsheet.models import Cell, Row

def norm(v):
    if v is None:
        return ""
    s = str(v).strip().lower()
    return re.sub(r"\.0+$", "", s)

def disp_or_val(cell):
    # prefer display_value when Smartsheet provides a formatted value
    dv = getattr(cell, "display_value", None)
    return dv if dv not in (None, "") else cell.value

def title_to_index(sheet):
    # authoritative positions from Smartsheet (not Python enumerate order)
    return {c.title: c.index for c in sheet.columns}

def index_to_id(sheet):
    # authoritative positions from Smartsheet (not Python enumerate order)
    return {c.index: c.id for c in sheet.columns}

def id_to_index(sheet):
    # authoritative positions from Smartsheet (not Python enumerate order)
    return {c.id: c.index for c in sheet.columns}

def id_to_title(sheet):
    return {c.id: c.title for c in sheet.columns}

def title_to_id(sheet):
    return {c.title: c.id for c in sheet.columns}

def guard_row(row, *idxs):
    # ensure row has enough cells for all requested positions
    return max(idxs) < len(row.cells)

def new_cell(column_id, value=None, strict=False, formula=None):
    new_cell = Cell()
    new_cell.column_id = column_id
    if formula is not None:
        new_cell.formula = formula
    else:
        new_cell.value = value
    if strict:
        new_cell.strict = True
    return new_cell

def new_row(cells=None, id=None, parent_id=None, to_top=False):
    new_row = Row()
    if cells:
        new_row.cells = cells
    if id:
        new_row.id = id
    if parent_id:
        new_row.parent_id = parent_id
    if to_top:
        new_row.to_top = to_top
    return new_row