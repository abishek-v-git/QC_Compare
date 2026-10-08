"""Build Mercury-style part numbers from a dev export and check which exist in a Mercury export."""
import io
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_SCAN_ROWS = 20
THREE_DIGIT_TAIL = re.compile(r'^(.+-)\d(\d{2})$')


class ExportError(ValueError):
    pass


def _clean(value):
    return '' if value is None else str(value).strip()


def mercury_base(item_number):
    """ASY-71940-131 -> ASY-71940-31: a three-digit last segment drops its first digit.
    Anything else (ASY-02505-21, IC-00688) is kept as-is."""
    m = THREE_DIGIT_TAIL.match(item_number)
    if m and item_number.count('-') >= 2:
        return m.group(1) + m.group(2)
    return item_number


def combine(item_number, revision):
    base = mercury_base(item_number)
    return f'{base}-{revision}' if revision else base


def read_dev_export(file):
    """Return [(item_number, revision)] from the dev export, locating the header row by name."""
    try:
        ws = load_workbook(file, read_only=True, data_only=True).worksheets[0]
    except Exception as exc:
        raise ExportError(f'Could not read the dev export: {exc}') from exc

    rows = ws.iter_rows(values_only=True)
    item_col = rev_col = None
    for _, row in zip(range(HEADER_SCAN_ROWS), rows):
        headers = [_clean(c).lower() for c in row]
        if 'item number' in headers and 'revision' in headers:
            item_col, rev_col = headers.index('item number'), headers.index('revision')
            break
    if item_col is None:
        raise ExportError('Dev export has no header row with "Item Number" and "Revision" columns.')

    items = []
    for row in rows:
        item = _clean(row[item_col]) if item_col < len(row) else ''
        if item:
            rev = _clean(row[rev_col]) if rev_col < len(row) else ''
            items.append((item, rev))
    return items


def read_mercury_export(file):
    """Every non-empty cell in the first sheet, normalised for matching."""
    try:
        ws = load_workbook(file, read_only=True, data_only=True).worksheets[0]
    except Exception as exc:
        raise ExportError(f'Could not read the Mercury export: {exc}') from exc
    return {_clean(c).upper() for row in ws.iter_rows(values_only=True) for c in row if _clean(c)}


def build_rows(dev_items, mercury_parts):
    rows = []
    for item, rev in dev_items:
        combined = combine(item, rev)
        rows.append({'item_number': item, 'revision': rev, 'combined': combined,
                     'matched': combined.upper() in mercury_parts})
    return rows


def rows_to_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Mercury check'
    ws.append(['Item Number', 'Revision', 'Combined', 'Matched with Mercury'])
    for r in rows:
        ws.append([r['item_number'], r['revision'], r['combined'], 'Yes' if r['matched'] else 'No'])

    header_fill = PatternFill('solid', fgColor='5850EC')
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    yes_fill, no_fill = PatternFill('solid', fgColor='E9F7EE'), PatternFill('solid', fgColor='FDEDED')
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].fill = yes_fill if row[0].value == 'Yes' else no_fill
        row[0].alignment = Alignment(horizontal='center')
    for i, width in enumerate([22, 12, 26, 22], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
