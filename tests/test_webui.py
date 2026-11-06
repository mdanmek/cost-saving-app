'''The web page: logic in costsaving/webui.py, the page script streamlit_app.py and the Pages build.

Streamlit itself is not installed in CI (it comes with stlite in the browser), so the page script is
run against a small stand-in module that records what would be drawn. That checks the flow of the
script (upload → check → summary → download, blocked on Error); it does not check how it looks.
'''
from pathlib import Path
import io
import json
import re
import runpy
import sys
import types
import zipfile

import pytest
from openpyxl import load_workbook

from costsaving import webui

ROOT = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------- webui.run_uploads
def run_sample(with_errors, erp=None):
    ic, po, master = webui.sample_inputs(with_errors)
    return webui.run_uploads(ic, po, master, erp)


def test_clean_sample_has_no_error_and_four_readable_files():
    r = run_sample(False)
    assert r.problem is None and r.errors == 0 and r.can_download
    assert sorted(n.rsplit('_for_', 1)[1] for n in r.files) == ['Collect.xlsx', 'Contract.xlsx', 'PD.xlsx', 'Summary.xlsx']
    for data in r.files.values():
        assert load_workbook(io.BytesIO(data)).sheetnames


def test_errors_block_download():
    r = run_sample(True)
    assert r.errors == 3 and not r.can_download
    first = webui.sorted_check(r.result['IC_Check'])
    assert list(first['Level'][:3]) == ['Error'] * 3


def test_erp_system_id_typed_by_user_removes_warning():
    without = run_sample(False)
    with_id = run_sample(False, erp='TEST-ERP')
    msg = 'ERP System ID ว่าง'
    assert without.result['IC_Check']['Message'].str.contains(msg).any()
    assert not with_id.result['IC_Check']['Message'].str.contains(msg).any()


def test_master_tables_may_sit_in_the_ic_file():
    from costsaving import sample
    r = webui.run_uploads(sample.ic_workbook_bytes(with_masters=True), [sample.po_csv_bytes()], None)
    assert r.problem is None and r.files


@pytest.mark.parametrize('ic, expected', [
    (None, 'ยังไม่ได้อัปโหลดไฟล์ IC'),
    (b'not an excel file', 'อ่านไฟล์ไม่สำเร็จ'),
])
def test_bad_input_returns_a_message_not_an_exception(ic, expected):
    r = webui.run_uploads(ic, [], None)
    assert expected in r.problem and not r.can_download


def test_missing_master_table_is_reported():
    from costsaving import sample
    r = webui.run_uploads(sample.ic_workbook_bytes(with_masters=False), [sample.po_csv_bytes()], None)
    assert 'ไม่พบตาราง' in r.problem


def test_zip_holds_the_four_files():
    r = run_sample(False)
    names = zipfile.ZipFile(io.BytesIO(webui.zip_files(r.files))).namelist()
    assert sorted(names) == sorted(r.files)


def test_summary_display_formats_percent_and_numbers():
    s = webui.summary_for_display(run_sample(False).result['Summary'])
    value = dict(zip(s['Measure'], s['Value']))
    assert value['% Hospital Saving to Managed Spend'].endswith('%')
    assert re.fullmatch(r'-?[\d,]+\.\d\d', value['Hospital Saving (Annualize)'])


# ---------------------------------------------------------------- nothing leaves the browser
def test_app_code_has_no_network_calls():
    '''The promise "ไม่ส่งข้อมูลออก" also holds in the code: no HTTP / socket / JS fetch anywhere.'''
    banned = re.compile(r'^\s*(import|from)\s+(requests|urllib|http|socket|httpx|aiohttp|pyodide\.http|js)\b', re.M)
    for f in [ROOT / 'streamlit_app.py', *sorted((ROOT / 'costsaving').glob('*.py'))]:
        assert not banned.search(f.read_text(encoding='utf-8')), f.name


# ---------------------------------------------------------------- page script with a stand-in Streamlit
class StopScript(Exception):
    pass


class FakeStreamlit(types.ModuleType):
    def __init__(self, choice, button=True):
        super().__init__('streamlit')
        self.choice, self.pressed = choice, button
        self.session_state = {}
        self.downloads, self.errors, self.frames = [], [], 0
        self.sidebar = self

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def __getattr__(self, name):  # title, caption, header, subheader, info, warning, success, metric, ...
        return lambda *a, **k: None

    def columns(self, spec):
        return [self] * (spec if isinstance(spec, int) else len(spec))

    def text_input(self, *a, **k):
        return ''

    def radio(self, label, options, **k):
        return options[self.choice]

    def file_uploader(self, *a, **k):
        return None

    def button(self, *a, disabled=False, **k):
        return self.pressed and not disabled

    def spinner(self, *a, **k):
        return self

    def expander(self, *a, **k):
        return self

    def stop(self):
        raise StopScript

    def error(self, msg, *a, **k):
        self.errors.append(msg)

    def dataframe(self, *a, **k):
        self.frames += 1

    def download_button(self, label, data=None, disabled=False, **k):
        self.downloads.append((label, disabled, len(data or b'')))


def run_page(fake, monkeypatch):
    monkeypatch.setitem(sys.modules, 'streamlit', fake)
    monkeypatch.syspath_prepend(str(ROOT))
    try:
        runpy.run_path(str(ROOT / 'streamlit_app.py'), run_name='__main__')
    except StopScript:
        pass
    return fake


def test_page_clean_sample_enables_five_downloads(monkeypatch):
    fake = run_page(FakeStreamlit(choice=1), monkeypatch)
    assert len(fake.downloads) == 5 and not any(d for _, d, _ in fake.downloads)
    assert all(size > 0 for _, _, size in fake.downloads)
    assert fake.frames >= 3 and not fake.errors


def test_page_sample_with_errors_disables_downloads(monkeypatch):
    fake = run_page(FakeStreamlit(choice=2), monkeypatch)
    assert len(fake.downloads) == 5 and all(d for _, d, _ in fake.downloads)
    assert any('Error 3' in e for e in fake.errors)


def test_page_without_upload_waits(monkeypatch):
    fake = run_page(FakeStreamlit(choice=0), monkeypatch)
    assert fake.downloads == [] and 'run' not in fake.session_state


# ---------------------------------------------------------------- GitHub Pages build
def test_site_build_lists_only_app_files(tmp_path):
    sys.path.insert(0, str(ROOT / 'site'))
    try:
        import build
    finally:
        sys.path.pop(0)
    manifest = build.build(tmp_path / '_site')
    files = manifest['files']
    assert manifest['entrypoint'] == 'streamlit_app.py' and 'streamlit_app.py' in files
    assert 'costsaving/webui.py' in files and 'costsaving/config/dim_method.csv' in files
    assert not any(f.startswith('tests/') or f.endswith(('.xlsx', '.xlsm')) for f in files)
    for f in files + ['index.html', 'files.json']:
        assert (tmp_path / '_site' / f).is_file(), f
    assert json.loads((tmp_path / '_site' / 'files.json').read_text())['files'] == files


def test_index_html_pins_one_stlite_version():
    html = (ROOT / 'site' / 'index.html').read_text(encoding='utf-8')
    versions = set(re.findall(r'@stlite/browser@([\d.]+)/', html))
    assert len(versions) == 1, versions
    assert 'Content-Security-Policy' in html and "form-action 'none'" in html
