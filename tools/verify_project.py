"""Verify standalone wiring and the absence of external business templates."""
import argparse
import json
import subprocess
from pathlib import Path

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('--manifest', default='migration/source-manifest-20261023-embedded.json')
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
assert 'EmbeddedResource' in csproj and 'ZTools.Templates.' in csproj
assert not (package/'Referrings').exists()
external = [p for p in package.rglob('*') if p.is_file() and p.suffix.lower() in {'.dwg','.xls','.dvb','.exb'}]
assert not external, external
inspection = subprocess.run([str(package/'ZTools.Standalone.exe'),'--verify-templates'], capture_output=True, text=True, encoding='utf-8-sig', check=True,
                            creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
inventory = json.loads(inspection.stdout)
assert len(inventory) == 4
source_entries = json.loads((root/args.manifest).read_text(encoding='utf-8'))
for asset in inventory:
    relative = 'Referrings/Ztools/' + ('' if asset['name'] == 'ZTools.xls' else 'TEMPLATE/') + asset['name']
    expected = next(e for e in source_entries if e['path'] == relative)
    assert asset['sha256'] == expected['sha256'] and asset['bytes'] == expected['bytes'], relative
print('PASS: published application embeds the three DWGs and Excel template; no external business templates are packaged.')
for name in ['ACadSharp-LICENSE.txt','PDFsharp-LICENSE.txt','three-LICENSE.txt']:
    assert (package/'licenses'/name).is_file(), name
assert (package/'ZTools.Calculation.dll').is_file()
assert (package/'licenses/mobile/NPOI-LICENSE').is_file()
print('PASS: packaged third-party notices and library licenses.')
