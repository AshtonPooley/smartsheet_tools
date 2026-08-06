from datetime import datetime
import time
import re
import warnings
from smartsheet.models import Cell, Row, Folder, Sheet, Error
from smartsheet.models import Column

# Cache for column types to minimize API calls when correcting date formats
_COLUMN_TYPE_CACHE = {}
_TITLE_TO_ID_CACHE = {}
_ID_TO_INDEX_CACHE = {}
# Per-sheet set of column ids that carry a Column Formula. Populated for free
# from the already-fetched sheet (see get_formula_col_ids / safe_grab_sheet_by_name)
# so the row wrappers can strip un-writable formula-column cells before a write.
_FORMULA_COL_IDS_CACHE = {}
_MAX_SHEET_LIMIT_ERROR_CODES = {5732}
_MAX_SHEET_LIMIT_STRINGS = (
    "reached the cell limit",
    "limit for cells in the sheet",
    "cell limit of 500,000 cells",
    "reached the row limit",
    "row limit of 20,000 rows",
)

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

def get_formula_col_ids(sheet_obj):
    """Return the set of column ids on this sheet that carry a Column Formula.

    Smartsheet rejects any written value for such a column with errorCode 1302,
    which aborts the whole add/update batch (and is non-retryable). Computed once
    per sheet id from the already-fetched sheet object -- no extra API call -- and
    cached, so callers (and the safe row wrappers) can cheaply strip those cells
    before writing. Returns an empty set for sheets with no formula columns.
    """
    sheet_id = sheet_obj.id
    if sheet_id not in _FORMULA_COL_IDS_CACHE:
        _FORMULA_COL_IDS_CACHE[sheet_id] = {
            c.id for c in sheet_obj.columns if getattr(c, "formula", None)
        }
    return _FORMULA_COL_IDS_CACHE[sheet_id]

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

def _fetch_children(fetch, obj_id, what, max_tries=4, delay_seconds=10):
    # get_folder_children / get_workspace_children return an Error object on
    # failure (or raise) instead of a Result with .data -- accessing .data on
    # that Error blows up with an opaque "'Error' object has no attribute
    # 'data'" AttributeError. Folder/workspace listing is idempotent, and these
    # calls occasionally hit a transient 500/503 or a 404 "Not Found" that
    # self-heals (Smartsheet replication lag under load), so retry a few times
    # before giving up and surface the real API error if we do.
    last_error = None
    for attempt in range(1, max_tries + 1):
        result = None
        try:
            result = fetch(obj_id)
        except Exception as exc:
            last_error = exc
        else:
            if isinstance(result, Error):
                last_error = result
            else:
                data = getattr(result, "data", None)
                if data is not None:
                    return data
                last_error = RuntimeError(f"unexpected response {result!r}")

        if attempt < max_tries:
            warnings.warn(f"failed to list {what} (attempt {attempt}/{max_tries}); retrying in {delay_seconds}s: {last_error}")
            time.sleep(delay_seconds)

    raise RuntimeError(
        f"failed to list {what} after {max_tries} tries: {last_error}"
    ) from (last_error if isinstance(last_error, Exception) else None)

def walk_folder_for_sheets(smartsheet_client, folder_id):
    for item in _fetch_children(smartsheet_client.Folders.get_folder_children, folder_id, f"folder {folder_id} children"):
        if isinstance(item, Folder):
            yield from walk_folder_for_sheets(smartsheet_client, item.id)
        elif isinstance(item, Sheet):
            yield item

def walk_workspace_for_sheets(smartsheet_client, workspace_id):
    for item in _fetch_children(smartsheet_client.Workspaces.get_workspace_children, workspace_id, f"workspace {workspace_id} children"):
        if isinstance(item, Folder):
            yield from walk_folder_for_sheets(smartsheet_client, item.id)
        elif isinstance(item, Sheet):
            yield item

def walk_folder_for_folders(smartsheet_client, folder_id):
    for item in _fetch_children(smartsheet_client.Folders.get_folder_children, folder_id, f"folder {folder_id} children"):
        if isinstance(item, Folder):
            yield item
            yield from walk_folder_for_folders(smartsheet_client, item.id)

def walk_workspace_for_folders(smartsheet_client, workspace_id):
    for item in _fetch_children(smartsheet_client.Workspaces.get_workspace_children, workspace_id, f"workspace {workspace_id} children"):
        if isinstance(item, Folder):
            yield item
            yield from walk_folder_for_folders(smartsheet_client, item.id)
            
def walk_sheet_names_from_workspace(smartsheet_client, workspace_id):
    for sheet in walk_workspace_for_sheets(smartsheet_client, workspace_id):
        yield sheet.name

def _collect_error_text_and_codes(error_obj):
    text_chunks = []
    codes = set()
    seen = set()

    def _walk(value):
        if value is None:
            return

        obj_id = id(value)
        if obj_id in seen:
            return
        seen.add(obj_id)

        text_chunks.append(str(value))

        message = getattr(value, "message", None)
        if message not in (None, ""):
            text_chunks.append(str(message))

        for attr_name in ("code", "error_code"):
            code = getattr(value, attr_name, None)
            if code is None:
                continue
            try:
                codes.add(int(code))
            except (TypeError, ValueError):
                pass

        _walk(getattr(value, "result", None))
        _walk(getattr(value, "error", None))

    _walk(error_obj)
    return " ".join(text_chunks).lower(), codes


def _is_max_sheet_limit_error(error_obj):
    error_text, error_codes = _collect_error_text_and_codes(error_obj)
    if _MAX_SHEET_LIMIT_ERROR_CODES.intersection(error_codes):
        return True
    return any(msg in error_text for msg in _MAX_SHEET_LIMIT_STRINGS)


def _find_sheet_id_by_name(smartsheet_client, sheet_name):
    all_sheets = smartsheet_client.Sheets.list_sheets(include_all=True)
    for sheet_obj in all_sheets.data:
        if sheet_obj.name == sheet_name:
            return sheet_obj.id
    return None


def _delete_last_row_of_sheet(smartsheet_client, sheet_name):
    sheet_id = _find_sheet_id_by_name(smartsheet_client, sheet_name)
    if sheet_id is None:
        return False

    first_page = smartsheet_client.Sheets.get_sheet(sheet_id, page_size=1, page=1)
    total_row_count = getattr(first_page, "total_row_count", None)
    if not total_row_count:
        rows = getattr(first_page, "rows", []) or []
        total_row_count = len(rows)

    if total_row_count <= 0:
        return False

    last_row_page = smartsheet_client.Sheets.get_sheet(
        sheet_id,
        row_numbers=[total_row_count],
        page_size=1,
    )
    rows = getattr(last_row_page, "rows", []) or []
    if not rows:
        return False

    last_row_id = rows[-1].id
    if last_row_id is None:
        return False

    smartsheet_client.Sheets.delete_rows(sheet_id, [last_row_id])
    return True


def safe_grab_sheet_by_name(
    smartsheet_client,
    name,
    max_tries=5,
    delay_seconds=15,
    raise_error=True,
    fix_max_limit_sheets=False,
):
    last_error = None
    for attempt in range(1, max_tries + 1):
        try:
            result = smartsheet_client.Sheets.get_sheet_by_name(name)
        except Exception as exc:
            last_error = exc
            result = None
            if (
                fix_max_limit_sheets
                and _is_max_sheet_limit_error(exc)
            ):
                warnings.warn(f"row/cell limit reached for {name}; deleting last row and retrying")
                try:
                    if _delete_last_row_of_sheet(smartsheet_client, name):
                        result = smartsheet_client.Sheets.get_sheet_by_name(name)
                except Exception as retry_exc:
                    last_error = retry_exc
                    result = None

        if isinstance(result, Sheet):
            get_formula_col_ids(result)  # cache formula-column ids for the row wrappers
            return result

        if isinstance(result, Error):
            last_error = result
            if (
                fix_max_limit_sheets
                and _is_max_sheet_limit_error(result)
            ):
                warnings.warn(f"row/cell limit reached for {name}; deleting last row and retrying")
                try:
                    if _delete_last_row_of_sheet(smartsheet_client, name):
                        result = smartsheet_client.Sheets.get_sheet_by_name(name)
                        if isinstance(result, Sheet):
                            get_formula_col_ids(result)  # cache formula-column ids for the row wrappers
                            return result
                        if isinstance(result, Error):
                            last_error = result
                except Exception as retry_exc:
                    last_error = retry_exc

        if attempt < max_tries:
            warnings.warn(f"failed to grab sheet {name} trying again in {delay_seconds}s")
            time.sleep(delay_seconds)

    if raise_error:
        raise RuntimeError(f"failed to grab sheet {name} after {max_tries} tries") from (last_error if isinstance(last_error, Exception) else None)



def sheet_exists(smartsheet_client, sheet_name, max_tries=3, delay_seconds=15, fix_max_limit_sheets=False):
    last_error = None
    for attempt in range(1, max_tries + 1):
        try:
            result = smartsheet_client.Sheets.get_sheet_by_name(sheet_name)
        except Exception as exc:
            last_error = exc
            result = None
            if (
                fix_max_limit_sheets
                and _is_max_sheet_limit_error(exc)
            ):
                warnings.warn(f"Sheet exists but row/cell limit reached for {sheet_name}; using safe_grab_sheet_by_name to recover")
                recovered = safe_grab_sheet_by_name(
                    smartsheet_client,
                    sheet_name,
                    max_tries=max_tries,
                    delay_seconds=delay_seconds,
                    raise_error=False,
                    fix_max_limit_sheets=True,
                )
                if isinstance(recovered, Sheet):
                    return recovered
                result = recovered
                
        if isinstance(result, Sheet):
            return result

        if result is False:
            return None

        if isinstance(result, Error):
            last_error = result
            if (
                fix_max_limit_sheets
                and _is_max_sheet_limit_error(result)
            ):
                warnings.warn(f"row/cell limit reached for {sheet_name}; using safe_grab_sheet_by_name to recover")
                recovered = safe_grab_sheet_by_name(
                    smartsheet_client,
                    sheet_name,
                    max_tries=max_tries,
                    delay_seconds=delay_seconds,
                    raise_error=False,
                    fix_max_limit_sheets=True,
                )
                if isinstance(recovered, Sheet):
                    return recovered
                if isinstance(recovered, Error):
                    last_error = recovered

        if attempt < max_tries:
            warnings.warn("error grabbing sheet; trying again")
            time.sleep(delay_seconds)

    return None
        
def safe_get_column_by_title(
    smartsheet_client,
    sheet_id,
    column_title,
    max_tries=3,
    delay_seconds=15,
    raise_error=True,
):
    last_error = None
    for attempt in range(1, max_tries + 1):
        try:
            result = smartsheet_client.Sheets.get_column_by_title(sheet_id, column_title)
        except Exception as exc:
            last_error = exc
            result = None

        if isinstance(result, Column):
            return result

        if isinstance(result, Error):
            last_error = result

        if attempt < max_tries:
            warnings.warn(f"failed to grab column {column_title!r} from sheet {sheet_id} trying again in {delay_seconds}s")
            time.sleep(delay_seconds)

    if raise_error:
        raise RuntimeError(
            f"failed to grab column {column_title!r} from sheet {sheet_id} after {max_tries} tries"
        ) from (last_error if isinstance(last_error, Exception) else None)
    return None


# API error codes where retrying cannot help -- the request itself is the
# problem (bad data, missing value, full sheet, not found). Surface these
# immediately instead of burning retries on them.
_NON_RETRYABLE_ERROR_CODES = {
    1006,   # not found
    1012,   # required object attribute(s) missing (e.g. cell.value)
    1057,   # column type does not support the supplied value
    1302,   # cannot edit a cell in a Column-Formula column
    5536,   # CELL_VALUE_FAILS_VALIDATION (e.g. PICKLIST)
    5636,   # reached the cell limit (legacy phrasing)
    5732,   # reached the 500,000 cell limit
}


# Message fragments that mark a non-retryable error when the code can't be
# extracted (e.g. the SDK raised an exception rather than returning an Error).
_NON_RETRYABLE_ERROR_STRINGS = (
    "did not conform to the strict requirements",  # 5536 PICKLIST / type validation
    "required object attribute",                    # 1012 missing cell.value
    "cannot edit cells with column formula",        # 1302 formula-column write
)


def _is_non_retryable_error(error_obj):
    text, codes = _collect_error_text_and_codes(error_obj)
    if _NON_RETRYABLE_ERROR_CODES.intersection(codes):
        return True
    if any(s in text for s in _MAX_SHEET_LIMIT_STRINGS):  # cell/row limit phrases
        return True
    return any(s in text for s in _NON_RETRYABLE_ERROR_STRINGS)


def _safe_row_op(op, sheet_id, rows, op_name, max_tries=4, delay_seconds=10, raise_error=True):
    """Run a Smartsheet row mutation (add/update/delete_rows) with retries
    through transient failures.

    Two transient failure modes are handled:
      - a raised exception: a 503 with an empty body crashes the SDK's
        response .json() with JSONDecodeError before it can build an Error;
        network blips raise too.
      - a returned Error object: 500s (errorCode 4000/4004) come back as
        Error results rather than raising.
    Both are retried with a delay. Known data errors (cell limit, validation,
    missing value, not found) are non-retryable -- they are raised (or returned
    when raise_error=False) immediately since retrying cannot help.

    On success returns the SDK Result object, so callers can use it exactly
    like the raw smartsheet.Sheets.<op> return value.
    """
    last_error = None
    for attempt in range(1, max_tries + 1):
        result = None
        try:
            result = op(sheet_id, rows)
        except Exception as exc:
            last_error = exc
            # The SDK may *raise* the API error (rather than return an Error
            # object). Short-circuit non-retryable conditions here too instead
            # of burning the full retry budget on something that can't succeed.
            if _is_non_retryable_error(exc):
                if raise_error:
                    raise RuntimeError(f"{op_name} on sheet {sheet_id} failed (non-retryable): {exc}") from exc
                return exc

        if isinstance(result, Error):
            if _is_non_retryable_error(result):
                if raise_error:
                    raise RuntimeError(f"{op_name} on sheet {sheet_id} failed (non-retryable): {result}")
                return result
            last_error = result
        elif result is not None:
            return result  # success Result object

        if attempt < max_tries:
            warnings.warn(f"{op_name} on sheet {sheet_id} failed (attempt {attempt}/{max_tries}); retrying in {delay_seconds}s: {last_error}")
            time.sleep(delay_seconds)

    if raise_error:
        # Keep the original error text in the message -- downstream callers
        # (e.g. QSOytd's cell-limit detection) string-match on it.
        raise RuntimeError(
            f"{op_name} on sheet {sheet_id} failed after {max_tries} tries: {last_error}"
        ) from (last_error if isinstance(last_error, Exception) else None)
    return last_error


def _strip_formula_cells(sheet_id, rows):
    """Drop cells that target a Column-Formula column from an add/update payload.

    Smartsheet rejects any written value for a formula column (errorCode 1302),
    which fails the whole batch. Uses the per-sheet formula-column cache that is
    populated for free when the sheet is fetched (safe_grab_sheet_by_name /
    get_formula_col_ids), so this is a single dict lookup and an immediate no-op
    for sheets with no formula columns (the common case). A row left with no
    cells is dropped (Smartsheet rejects a cell-less row); those cells were
    un-writable anyway, so no real data is lost.
    """
    formula_ids = _FORMULA_COL_IDS_CACHE.get(sheet_id)
    if not formula_ids:
        return rows
    cleaned = []
    dropped = 0
    for row in rows:
        cells = getattr(row, "cells", None)
        if not cells:
            cleaned.append(row)
            continue
        kept = [c for c in cells if getattr(c, "column_id", None) not in formula_ids]
        if len(kept) == len(cells):
            cleaned.append(row)
        elif kept:
            row.cells = kept
            cleaned.append(row)
        else:
            dropped += 1
    if dropped:
        warnings.warn(f"skipped {dropped} row(s) with only formula-column cells on sheet {sheet_id}")
    return cleaned


def safe_add_rows(smartsheet_client, sheet_id, rows, max_tries=4, delay_seconds=10, raise_error=True):
    rows = _strip_formula_cells(sheet_id, rows)
    if not rows:
        return None
    return _safe_row_op(smartsheet_client.Sheets.add_rows, sheet_id, rows, "add_rows",
                        max_tries=max_tries, delay_seconds=delay_seconds, raise_error=raise_error)


def safe_update_rows(smartsheet_client, sheet_id, rows, max_tries=4, delay_seconds=10, raise_error=True):
    rows = _strip_formula_cells(sheet_id, rows)
    if not rows:
        return None
    return _safe_row_op(smartsheet_client.Sheets.update_rows, sheet_id, rows, "update_rows",
                        max_tries=max_tries, delay_seconds=delay_seconds, raise_error=raise_error)


def safe_delete_rows(smartsheet_client, sheet_id, row_ids, max_tries=4, delay_seconds=10, raise_error=True):
    return _safe_row_op(smartsheet_client.Sheets.delete_rows, sheet_id, row_ids, "delete_rows",
                        max_tries=max_tries, delay_seconds=delay_seconds, raise_error=raise_error)


def safe_sort_sheet(smartsheet_client, sheet_id, sort_specifier, max_tries=4, delay_seconds=10, raise_error=True):
    """Sort a sheet with retries through transient failures.

    Same transient failure modes as _safe_row_op: a 503 with an empty body
    crashes the SDK's response .json() with JSONDecodeError before it can build
    an Error object; transient 500s come back as Error results rather than
    raising. Both are retried with a delay. On success returns the SDK Result
    (the sorted Sheet), so callers that assign the return value keep working.
    """
    last_error = None
    for attempt in range(1, max_tries + 1):
        result = None
        try:
            result = smartsheet_client.Sheets.sort_sheet(sheet_id, sort_specifier)
        except Exception as exc:
            last_error = exc
            if _is_non_retryable_error(exc):
                if raise_error:
                    raise RuntimeError(f"sort_sheet on sheet {sheet_id} failed (non-retryable): {exc}") from exc
                return exc

        if isinstance(result, Error):
            if _is_non_retryable_error(result):
                if raise_error:
                    raise RuntimeError(f"sort_sheet on sheet {sheet_id} failed (non-retryable): {result}")
                return result
            last_error = result
        elif result is not None:
            return result  # success -- the sorted Sheet

        if attempt < max_tries:
            warnings.warn(f"sort_sheet on sheet {sheet_id} failed (attempt {attempt}/{max_tries}); retrying in {delay_seconds}s: {last_error}")
            time.sleep(delay_seconds)

    if raise_error:
        raise RuntimeError(
            f"sort_sheet on sheet {sheet_id} failed after {max_tries} tries: {last_error}"
        ) from (last_error if isinstance(last_error, Exception) else None)
    return last_error


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
