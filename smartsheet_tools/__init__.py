from datetime import datetime
import re
from smartsheet.models import Cell, Row, Folder, Sheet
from smartsheet.models import Column

# Cache for column types to minimize API calls when correcting date formats
_COLUMN_TYPE_CACHE = {}
_TITLE_TO_ID_CACHE = {}
_ID_TO_INDEX_CACHE = {}

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
    if sheet.id not in _TITLE_TO_ID_CACHE:
        _TITLE_TO_ID_CACHE[sheet.id] = {c.title: c.index for c in sheet.columns}
    return _TITLE_TO_ID_CACHE[sheet.id]

def index_to_id(sheet):
    # authoritative positions from Smartsheet (not Python enumerate order)
    if sheet.id not in _ID_TO_INDEX_CACHE:
        _ID_TO_INDEX_CACHE[sheet.id] = {c.index: c.id for c in sheet.columns}
    return _ID_TO_INDEX_CACHE[sheet.id]

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

def get_cached_column_type(column_id, sheet_obj, prefill=False):
    if sheet_obj.id not in _COLUMN_TYPE_CACHE:
        _COLUMN_TYPE_CACHE[sheet_obj.id] = {}
        
    if column_id not in _COLUMN_TYPE_CACHE[sheet_obj.id]:
        if not prefill:
            
            # Value is not in there and no prefill, so look it up
            for col in sheet_obj.columns:
                _COLUMN_TYPE_CACHE[sheet_obj.id][column_id] = col.type

        else:
            _COLUMN_TYPE_CACHE[sheet_obj.id][column_id] = prefill
    
    return _COLUMN_TYPE_CACHE[sheet_obj.id][column_id]

def get_col_names_of_date_cols(sheet_obj):
    return [c.title for c in sheet_obj.columns if get_cached_column_type(c.id, sheet_obj, prefill=c.type) in ("DATE", "DATETIME")]

def get_col_names_of_bool_cols(sheet_obj):
    return [c.title for c in sheet_obj.columns if get_cached_column_type(c.id, sheet_obj, prefill=c.type) == "CHECKBOX"]

def brute_force_date_string(s, nonetype_if_fail=False):
    # attempt to parse a date string in common formats to ISO 8601
    if isinstance(s, datetime):
        return datetime_to_isoformat(s)
    
    if not isinstance(s, str):
        return None if nonetype_if_fail else s
    
    s = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime_to_isoformat(datetime.strptime(s, fmt))
        except ValueError:
            continue
    return None if nonetype_if_fail else s


def is_date_col(column_id, sheet_obj):
    column_type = get_cached_column_type(column_id, sheet_obj)
    return column_type in ("DATE", "DATETIME")

def correct_date_format(value, column_id, sheet_obj, nonetype_if_fail=False):
    if isinstance(value, datetime):
        value = datetime_to_isoformat(value)

    column_type = get_cached_column_type(column_id, sheet_obj)
    if column_type == "DATE":
        return value.split("T",1)[0]
    elif column_type == "DATETIME":
        return value
    return None if nonetype_if_fail else value

def new_cell(column_id=None, value=None, strict=False, formula=None):
    new_cell = Cell()
    if column_id is not None:
        new_cell.column_id = column_id
    if formula is not None:
        new_cell.formula = formula
    else:
        new_cell.value = value
    new_cell.strict = strict
    return new_cell

def new_row(cells=None, id=None, parent_id=None, to_top=False, locked=False):
    new_row = Row()
    if cells:
        new_row.cells = cells
    if id:
        new_row.id = id
    if parent_id:
        new_row.parent_id = parent_id
    if to_top:
        new_row.to_top = to_top
    if locked:
        new_row.locked = locked
    return new_row

def walk_folder_for_sheets(smartsheet_client, folder_id):
    for item in smartsheet_client.Folders.get_folder_children(folder_id).data:
        if isinstance(item, Folder):
            yield from walk_folder_for_sheets(smartsheet_client, item.id)
        elif isinstance(item, Sheet):
            yield item

def walk_workspace_for_sheets(smartsheet_client, workspace_id):
    for item in smartsheet_client.Workspaces.get_workspace_children(workspace_id).data:
        if isinstance(item, Folder):
            yield from walk_folder_for_sheets(smartsheet_client, item.id)
        elif isinstance(item, Sheet):
            yield item
            
def walk_folder_for_folders(smartsheet_client, folder_id):
    for item in smartsheet_client.Folders.get_folder_children(folder_id).data:
        if isinstance(item, Folder):
            yield item
            yield from walk_folder_for_folders(smartsheet_client, item.id)
            
def walk_workspace_for_folders(smartsheet_client, workspace_id):
    for item in smartsheet_client.Workspaces.get_workspace_children(workspace_id).data:
        if isinstance(item, Folder):
            yield item
            yield from walk_folder_for_folders(smartsheet_client, item.id)
            
def walk_sheet_names_from_workspace(smartsheet_client, workspace_id):
    for sheet in walk_workspace_for_sheets(smartsheet_client, workspace_id):
        yield sheet.name
        
def new_column(column_type, title, index=None, id=None, options=None, symbol=None, primary=False, hidden=False, locked=False):
    new_column = Column()
    
    new_column.type = column_type
    new_column.title = title
    if index is not None:
        new_column.index = index
    if id is not None:
        new_column.id = id
    if options is not None and column_type in ("PICKLIST", "MULTI_PICKLIST"):
        new_column.options = options
    if symbol is not None and column_type == "CHECKBOX":
        new_column.symbol = symbol
    if primary:
        new_column.primary = True
    if hidden:
        new_column.hidden = True
    if locked:
        new_column.locked = True
    return new_column