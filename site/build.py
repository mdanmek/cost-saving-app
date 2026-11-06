'''Build the GitHub Pages site into _site/.

    _site/
      index.html          loads stlite from the CDN, then mounts the files listed in files.json
      files.json          {entrypoint, requirements, files: [...]}
      streamlit_app.py
      costsaving/*.py, costsaving/config/*.csv

Only the app code and the public config are copied: no tests, no data. Run: python site/build.py
then serve _site/ (python -m http.server -d _site) to try it locally.
'''
from pathlib import Path
import json
import shutil
import sys

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / '_site'
ENTRYPOINT = 'streamlit_app.py'
# packages of the Pyodide distribution; streamlit itself comes with stlite
REQUIREMENTS = ['pandas', 'openpyxl']


def app_files():
    files = [ROOT / ENTRYPOINT]
    files += sorted((ROOT / 'costsaving').glob('*.py'))
    files += sorted((ROOT / 'costsaving' / 'config').glob('*.csv'))
    return files


def build(out=OUT):
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    paths = []
    for f in app_files():
        rel = f.relative_to(ROOT).as_posix()
        (out / rel).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(f, out / rel)
        paths.append(rel)
    shutil.copy2(ROOT / 'site' / 'index.html', out / 'index.html')
    (out / '.nojekyll').write_text('')
    manifest = {'entrypoint': ENTRYPOINT, 'requirements': REQUIREMENTS, 'files': paths}
    (out / 'files.json').write_text(json.dumps(manifest, indent=2), encoding='utf-8')
    return manifest


if __name__ == '__main__':
    m = build(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT)
    print(f"built {len(m['files'])} files into _site/")
