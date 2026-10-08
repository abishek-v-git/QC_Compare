from django import forms

EXCEL_EXTENSIONS = ('.xlsx', '.xlsm')


def _validate_excel(f):
    if not f.name.lower().endswith(EXCEL_EXTENSIONS):
        raise forms.ValidationError('Upload an Excel file (.xlsx).')
    return f


class UploadForm(forms.Form):
    dev_export = forms.FileField()
    mercury_export = forms.FileField()

    def clean_dev_export(self):
        return _validate_excel(self.cleaned_data['dev_export'])

    def clean_mercury_export(self):
        return _validate_excel(self.cleaned_data['mercury_export'])
