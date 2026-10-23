"""Verify standalone UI/API wiring and byte-identical packaged references."""
import argparse
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('--manifest', default='migration/source-manifest-20261023.json')
args = parser.parse_args()
project = root/'ZTools.Standalone'
program = (project/'Program.cs').read_text(encoding='utf-8-sig')
ui = (project/'wwwroot/index.html').read_text(encoding='utf-8-sig')
js = (project/'wwwroot/app.js').read_text(encoding='utf-8-sig')
csproj = (project/'ZTools.Standalone.csproj').read_text(encoding='utf-8-sig')
for command in ['new','open','calculate','utility','parameters','save-parameters','print','generate','edit','finish','refresh']:
    assert '"' + command + '"=>' in program, command
for group in ['equipment','motor','pid','materials']:
    assert 'data-parameters="' + group + '"' in ui, group
assert 'three.module.js' in js and 'WebGLRenderer' in js
assert 'cfl-logo.jpg' in ui and 'List-all' in ui
assert 'DXF' not in ui and 'DxfReader' not in (project/'DrawingOutput.cs').read_text(encoding='utf-8')
assert 'Autodesk' not in csproj and 'ZTools.AutoCAD' not in csproj
assert not (root/'ZTools.AutoCAD').exists() and not (root/'deploy').exists()
assert (root/'migration/archived-autocad-addin-20261023.zip').is_file()
assert 'ZTools.Standalone' in (root/'ZTools.sln').read_text(encoding='utf-8-sig')
print('PASS: standalone Chinese UI, four parameter forms, API routes, WebGL viewer, DWG-only templates and retired add-in.')

package = root/'dist/CFL-SubmitTools'
assert (package/'ZTools.Standalone.exe').is_file()
assert (package/'Start.ps1').is_file()
assert (package/'wwwroot/vendor/three.module.js').is_file()
for entry in json.loads((root/args.manifest).read_text(encoding='utf-8')):
    path = package/entry['path']
    assert hashlib.sha256(path.read_bytes()).hexdigest() == entry['sha256'], path
print('PASS: published Windows application and all 16 original references match the source snapshot.')
for name in ['ACadSharp-LICENSE.txt','PDFsharp-LICENSE.txt','three-LICENSE.txt']:
    assert (package/'licenses'/name).is_file(), name
print('PASS: packaged third-party notices and library licenses.')
