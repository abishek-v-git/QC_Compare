import base64

from django.shortcuts import render

from .forms import UploadForm
from .services.mercury import (AVAILABLE, NO_NEED, NOT_AVAILABLE, ExportError, build_rows, change_number,
                               output_filename, read_dev_export, read_mercury_export, rows_to_xlsx)


def upload(request):
    """Runs the check and renders the result straight away; nothing is stored."""
    form = UploadForm(request.POST or None, request.FILES or None)
    error = None
    if request.method == 'POST' and form.is_valid():
        dev, mercury = form.cleaned_data['dev_export'], form.cleaned_data['mercury_export']
        try:
            dev_items = read_dev_export(dev)
            rows = build_rows(dev_items, read_mercury_export(mercury))
        except ExportError as exc:
            error = str(exc)
        else:
            if not rows:
                error = 'No ASY or PCA item numbers were found in the dev export.'
            else:
                return _result(request, dev.name, mercury.name, rows, skipped=len(dev_items) - len(rows))
    return render(request, 'checker/upload.html', {'form': form, 'error': error})


def _result(request, dev_filename, mercury_filename, rows, skipped):
    counts = {s: sum(1 for r in rows if r['status'] == s) for s in (AVAILABLE, NOT_AVAILABLE, NO_NEED)}
    to_check = counts[AVAILABLE] + counts[NOT_AVAILABLE]  # Outback parts don't count towards the rate
    # The download is embedded in the page as a data URI, since there is no stored run to fetch it from later.
    change = change_number(dev_filename)
    xlsx = rows_to_xlsx(rows, change)
    return render(request, 'checker/result.html', {
        'change': change,
        'output_filename': output_filename(change),
        'dev_filename': dev_filename,
        'mercury_filename': mercury_filename,
        'rows': rows,
        'total': len(rows),
        'skipped': skipped,
        'available_count': counts[AVAILABLE],
        'not_available_count': counts[NOT_AVAILABLE],
        'no_need_count': counts[NO_NEED],
        'match_pct': round(100 * counts[AVAILABLE] / to_check) if to_check else 100,
        'xlsx_b64': base64.b64encode(xlsx).decode(),
        'AVAILABLE': AVAILABLE, 'NOT_AVAILABLE': NOT_AVAILABLE, 'NO_NEED': NO_NEED,
    })
