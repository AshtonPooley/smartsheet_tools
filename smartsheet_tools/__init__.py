from datetime import datetime
import re
from smartsheet.models import Cell, Row

# Cache for column types to minimize API calls when correcting date formats
_COLUMN_TYPE_CACHE = {}

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

def datetime_to_isoformat(dt):
    if dt is None:
        return None
    return dt.replace(microsecond=0).isoformat() + 'Z'

def standard_time_to_isoformat(st):
    if st is None:
        return None
    return datetime_to_isoformat(datetime.strptime(st, "%m/%d/%Y"))

def get_cached_column_type(column_id, sheet_obj):
    if sheet_obj.id not in _COLUMN_TYPE_CACHE:
        _COLUMN_TYPE_CACHE[sheet_obj.id] = {}
        
    if column_id not in _COLUMN_TYPE_CACHE[sheet_obj.id]:
        _COLUMN_TYPE_CACHE[sheet_obj.id][column_id] = str(sheet_obj.get_column(column_id).type)
    
    return _COLUMN_TYPE_CACHE[sheet_obj.id][column_id]

def get_col_names_of_date_cols(sheet_obj):
    return [c.title for c in sheet_obj.columns if get_cached_column_type(c.id, sheet_obj) in ("DATE", "DATETIME")]

def brute_force_date_string(s):
    # attempt to parse a date string in common formats to ISO 8601
    s = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime_to_isoformat(datetime.strptime(s, fmt))
        except ValueError:
            continue
    return None

def is_date_col(column_id, sheet_obj):
    column_type = get_cached_column_type(column_id, sheet_obj)
    return column_type in ("DATE", "DATETIME")

def correct_date_format(isoformat_datetime, column_id, sheet_obj):
    if isinstance(isoformat_datetime, datetime):
        isoformat_datetime = datetime_to_isoformat(isoformat_datetime)

    column_type = get_cached_column_type(column_id, sheet_obj)
    if column_type == "DATE":
        return isoformat_datetime.split("T",1)[0]
    elif column_type == "DATETIME":
        return isoformat_datetime
    return None

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