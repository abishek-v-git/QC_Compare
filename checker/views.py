import base64

from django.shortcuts import render

from .forms import UploadForm
from .services.mercury import ExportError, build_rows, read_dev_export, read_mercury_export, rows_to_xlsx


def upload(request):
    """Runs the check and renders the result straight away; nothing is stored."""
    form = UploadForm(request.POST or None, request.FILES or None)
    error = None
    if request.method == 'POST' and form.is_valid():
        dev, mercury = form.cleaned_data['dev_export'], form.cleaned_data['mercury_export']
        try:
            rows = build_rows(read_dev_export(dev), read_mercury_export(mercury))
        except ExportError as exc:
            error = str(exc)
        else:
            if not rows:
                error = 'No item numbers were found in the dev export.'
            else:
                return _result(request, dev.name, mercury.name, rows)
    return render(request, 'checker/upload.html', {'form': form, 'error': error})


def _result(request, dev_filename, mercury_filename, rows):
    total = len(rows)
    matched = sum(1 for r in rows if r['matched'])
    # The download is embedded in the page as a data URI, since there is no stored run to fetch it from later.
    xlsx = rows_to_xlsx([r for r in rows if r['matched']])
    return render(request, 'checker/result.html', {
        'dev_filename': dev_filename,
        'mercury_filename': mercury_filename,
        'rows': rows,
        'total': total,
        'matched_count': matched,
        'unmatched_count': total - matched,
        'match_pct': round(100 * matched / total),
        'xlsx_b64': base64.b64encode(xlsx).decode(),
    })
