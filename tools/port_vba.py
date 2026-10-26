"""Targeted source-to-source port for the supplied VBA, producing reviewable C#.

Not a general VBA converter. Unsupported syntax fails rather than disappearing.
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'ZTools.Standalone/Core/Generated'
OUT.mkdir(parents=True, exist_ok=True)
TOKEN = re.compile(r'"(?:""|[^"])*"|\d+(?:\.\d*)?#?|[A-Za-z_][\w]*|<>|<=|>=|:=|[^\s]')
OPS = {'or': 1, 'and': 2, '=': 3, '<>': 3, '<': 3, '>': 3, '<=': 3, '>=': 3,
       '&': 4, '+': 5, '-': 5, '*': 6, '/': 6, '^': 7}
FN = {'len', 'lenb', 'mid', 'left', 'right', 'instr', 'val', 'asc', 'replace',
      'ucase', 'trim', 'cstr', 'cint', 'isnumeric', 'typename', 'format'}
MEMBERS = {n.lower(): n for n in '''Application ActiveDocument Path Name Documents Open Activate Visible
Range Offset Value Value2 Cells Worksheets Sheets Rows Count End Row ClearContents Select Selection
Interior ColorIndex ScreenUpdating VLookup HLookup CountA Union WorksheetFunction ModelSpace AddMText Text
Height Width Update ScaleEntity InsertBlock Layers Item LayerOn ActiveTextStyle SetFont Layouts Block
ObjectName GetAttributes TagString TextString Move ScaleEntity SelectionSets Add Select On AddItems
Plot PlotToFile ConfigName StyleSheet CanonicalMediaName PaperUnits PlotRotation StandardScale
UseStandardScale CenterPlot PlotWithLineweights PlotWithPlotStyles NumberOfCopies QuietErrorMode
PurgeAll Regen SetVariable SetWindowToPlot PlotType SaveAs Close Wblock ZoomExtents Explode Delete
SummaryInfo NumCustomInfo GetCustomByIndex SetCustomByKey AddCustomInfo GetCustomByKey Copy Paste
PasteSpecial Cut SendCommand FullName ActiveLayout ActiveWorkbook Save HasAttributes TextStyle
Caption TextString TrueColor GetBoundingBox EffectiveName RefreshPlotDeviceInfo GetCanonicalMediaNames
'''.split()}
CONST = {'true':'true', 'false':'false', 'nothing':'null', 'xlup':'-4162', 'xlnone':'-4142',
         'xlall':'-4104', 'xlpasteall':'-4104', 'acallviewports':'1', 'acmillimeters':'1',
         'ac90degrees':'1', 'ac0degrees':'0', 'acscaletofit':'0', 'acwindow':'4',
         'acselectionsetall':'5', 'vbcrlf':'"\\r\\n"', 'vbmodeless':'0'}

class Node:
    def __init__(self, kind, value, args=None):
        self.kind, self.value, self.args = kind, value, args or []

class Port:
    def __init__(self):
        self.variables = {}
        self.arrays = set()
        self.lines = []
        self.selects = []
        self.sequence = 0

    def expr(self, text):
        self.tokens = TOKEN.findall(text.strip())
        self.pos = 0
        node = self.parse(0)
        if self.pos != len(self.tokens):
            raise ValueError('Unparsed expression: ' + text + ' @ ' + str(self.tokens[self.pos:]))
        return node

    def parse(self, minimum):
        tok = self.tokens[self.pos]; self.pos += 1
        low = tok.lower()
        if tok == '(':
            left = self.parse(0)
            assert self.tokens[self.pos] == ')'
            self.pos += 1
        elif low in ('not', '-', '+'):
            left = Node('unary', low, [self.parse(3 if low == 'not' else 8)])
        elif tok.startswith('"'):
            left = Node('literal', json.dumps(tok[1:-1].replace('""', '"'), ensure_ascii=False))
        elif tok[0].isdigit():
            left = Node('literal', tok.rstrip('#'))
        else:
            left = Node('id', low)
        while self.pos < len(self.tokens):
            tok = self.tokens[self.pos]
            if tok == '.':
                self.pos += 1
                left = Node('member', self.tokens[self.pos].lower(), [left]); self.pos += 1
            elif tok == '(':
                self.pos += 1
                args = []
                if self.tokens[self.pos] != ')':
                    while True:
                        args.append(self.parse(0))
                        if self.tokens[self.pos] != ',': break
                        self.pos += 1
                assert self.tokens[self.pos] == ')'
                self.pos += 1
                left = Node('call', left, args)
            elif tok.lower() in OPS and OPS[tok.lower()] >= minimum:
                self.pos += 1
                left = Node('op', tok.lower(), [left, self.parse(OPS[tok.lower()] + 1)])
            else: break
        return left

    def render(self, node, mode='value'):
        kind, v, a = node.kind, node.value, node.args
        if kind == 'literal': return v
        if kind == 'id':
            if v in CONST: return CONST[v]
            if v in ('application','thisworkbook','selection','rows','thisdrawing'):
                return {'application':'ExcelApp', 'thisworkbook':'Book', 'selection':'ExcelApp.Selection',
                        'rows':'ExcelApp.Rows','thisdrawing':'Drawing'}[v]
            if v in FN: return 'Vb.' + v.capitalize()
            if v == 'changefld': return 'ChangeField'
            if v == 'updateprogress': return 'ReportProgress'
            if v == 'finddatabaseoffset': return 'FindDatabaseOffset'
            if v == 'pumpdatabasesheet': return 'PumpDatabaseSheet'
            if v == 'insertpidplan': return 'InsertPidPlan'
            if v == 'calculatesectionnames': return 'CalculateSectionNames'
            if v == 'insertsection': return 'InsertSection'
            self.variables.setdefault(v, 'null')
            return '@' + v
        if kind == 'op':
            names = {'&':'Cat', '+':'Add', '-':'Sub', '*':'Mul', '/':'Div', '^':'Pow',
                     '=':'Eq', '<>':'Ne', '<':'Lt', '>':'Gt', '<=':'Le', '>=':'Ge',
                     'or':'Or', 'and':'And'}
            return 'Vb.' + names[v] + '(' + ', '.join(self.render(x) for x in a) + ')'
        if kind == 'unary':
            return {'not':'Vb.Not(', '-':'Vb.Neg(', '+':'Vb.Num('}[v] + self.render(a[0]) + ')'
        if kind == 'member':
            if a[0].kind == 'id' and a[0].value == 'vba':
                return 'Vb.' + v.capitalize()
            if a[0].kind == 'id' and a[0].value == 'application' and v in ('vlookup','hlookup','counta'):
                return 'ExcelApp.WorksheetFunction.' + MEMBERS[v]
            suffix = MEMBERS.get(v, v)
            base = self.render(a[0], 'object') + '.' + suffix
            if v in ('getattributes', 'explode', 'numcustominfo'): return base + '()'
            if v == 'value': return self.render(a[0], 'object') + '.Value2'
            return base
        if kind == 'call':
            if v.kind == 'id' and v.value in ('sheets','worksheets'):
                return 'Book.Worksheets[' + ', '.join(self.render(x) for x in a) + ']'
            isrange = v.kind == 'member' and v.value in ('range','offset','cells')
            if v.kind == 'id' and v.value in self.arrays:
                return self.render(v) + '[' + ', '.join('Vb.Int(' + self.render(x) + ')' for x in a) + ']'
            if v.kind == 'member' and v.value in ('range','cells','worksheets','sheets','offset','end'):
                base = self.render(v.args[0], 'object') + '.' + MEMBERS[v.value]
                code = base + '[' + ', '.join(self.render(x) for x in a) + ']'
            elif v.kind == 'member' and v.value == 'block':
                code = self.render(v, 'object') + '.Item(' + ', '.join(self.render(x) for x in a) + ')'
            else:
                args = []
                for i,x in enumerate(a):
                    islookup = v.kind == 'member' and v.value in ('vlookup','hlookup','counta')
                    args.append(self.render(x, 'object' if islookup and (i == 1 or v.value == 'counta') else 'value'))
                code = self.render(v, 'object') + '(' + ', '.join(args) + ')'
            return code + ('.Value2' if isrange and mode == 'value' else '')
        raise ValueError(kind)

    def e(self, s, mode='value'):
        return self.render(self.expr(s), mode)

    def emit(self, s): self.lines.append(s)

    def statement(self, s):
        low = s.lower().strip()
        if not low: return
        if low.startswith('dim '):
            for entry in split_args(s[4:]):
                m = re.fullmatch(r'(\w+)\s*(\(\s*0\s+To\s+(\d+)\s*\))?(?:\s+As\s+(\w+))?', entry.strip(), re.I)
                if not m: raise ValueError(s)
                name = m[1].lower(); typ = (m[4] or 'variant').lower()
                self.variables[name] = '""' if typ == 'string' else ('0' if typ in ('integer','long','double','single') else ('false' if typ == 'boolean' else 'null'))
                if m[2]:
                    self.arrays.add(name); self.variables[name] = 'new double[' + str(int(m[3])+1) + ']'
            return
        if low.startswith('if '):
            m = re.match(r'If (.*?) Then\s*(.*)', s, re.I)
            if not m: raise ValueError(s)
            self.emit('if (Vb.Bool(' + self.e(m[1]) + ')) {')
            if m[2]:
                self.statement(m[2]); self.emit('}')
            return
        if low.startswith('elseif '):
            m = re.match(r'ElseIf (.*?) Then\s*(.*)', s, re.I)
            self.emit('} else if (Vb.Bool(' + self.e(m[1]) + ')) {')
            if m[2]: self.statement(m[2])
            return
        if low == 'else': self.emit('} else {'); return
        if low == 'end if': self.emit('}'); return
        if low.startswith('for each '):
            m = re.match(r'For Each (\w+) In (.*)', s, re.I)
            self.sequence += 1
            name = self.e(m[1]); expr = self.e(m[2], 'object')
            self.emit('foreach (dynamic each' + str(self.sequence) + ' in ' + expr + ') {')
            self.emit(name + ' = each' + str(self.sequence) + ';'); return
        if low.startswith('for '):
            m = re.match(r'For (\w+) = (.*?) To (.*?)(?: Step (.*))?$', s, re.I)
            name, start, end, step = self.e(m[1]), self.e(m[2]), self.e(m[3]), self.e(m[4] or '1')
            self.emit('for (' + name + ' = ' + start + '; Vb.Le(' + name + ', ' + end + '); ' + name + ' = Vb.Add(' + name + ', ' + step + ')) {'); return
        if low.startswith('next'): self.emit('}'); return
        if low == 'exit for': self.emit('break;'); return
        if low == 'exit sub': self.emit('throw new InvalidOperationException("计算流程提前终止，请检查输入参数及数据库匹配情况。");'); return
        if low.startswith('select case '):
            self.sequence += 1; name = 'case' + str(self.sequence)
            self.emit('dynamic ' + name + ' = ' + self.e(s[12:]) + ';')
            self.selects.append([name,False]); return
        if low.startswith('case '):
            name, seen = self.selects[-1]; val = s[5:].rstrip(':')
            if val.lower().startswith('is '): cond = self.e(name + ' ' + val[3:])
            elif re.search(' To ', val, re.I):
                v1,v2 = re.split(' To ', val, flags=re.I)
                cond = 'Vb.And(Vb.Ge(' + name + ', ' + self.e(v1) + '), Vb.Le(' + name + ', ' + self.e(v2) + '))'
            elif val.lower() == 'else': cond = None
            else: cond = 'Vb.Eq(' + name + ', ' + self.e(val) + ')'
            # Case selector is a C# temporary, not a VBA local.
            if name.lower() in self.variables: self.variables.pop(name.lower())
            if cond: cond = cond.replace('@'+name.lower(), name)
            self.emit(('} else ' if seen else '') + ('if (Vb.Bool(' + cond + ')) {' if cond else '{'))
            self.selects[-1][1] = True; return
        if low == 'end select':
            self.selects.pop(); self.emit('}'); return
        if low.startswith('redim preserve '):
            m = re.match(r'ReDim Preserve (\w+)\(0 To (\d+)\)', s, re.I)
            self.emit(self.e(m[1]) + ' = Vb.Resize(' + self.e(m[1]) + ', ' + str(int(m[2])+1) + ');'); return
        if low.startswith('msgbox '): self.emit('Notify(' + self.e(s[7:].strip()) + ');'); return
        if low.startswith('call '): s = s[5:]; low = s.lower()
        if low.startswith('set '): s = s[4:]
        # Find a statement assignment, outside strings and parentheses.
        depth = 0; quoted = False; assignment = -1
        for i,c in enumerate(s):
            if c == '"': quoted = not quoted
            if quoted: continue
            if c == '(': depth += 1
            if c == ')': depth -= 1
            if c == '=' and depth == 0: assignment = i; break
        if assignment >= 0:
            rhs = self.e(s[assignment+1:])
            if s[:assignment].strip().lower().endswith('.textstring'):
                rhs = 'Vb.Text(' + rhs + ')'
            self.emit(self.e(s[:assignment], 'value') + ' = ' + rhs + ';'); return
        # VBA allows method calls without parentheses.
        tokens = TOKEN.findall(s); depth = 0; cut = None
        for match in re.finditer(r'"(?:""|[^"])*"|\S+', s):
            pass
        quoted = False
        for i,c in enumerate(s):
            if c == '"': quoted = not quoted
            if quoted: continue
            if c == '(': depth += 1
            if c == ')': depth -= 1
            if c.isspace() and depth == 0: cut = i; break
        if cut is not None:
            target, args = s[:cut], s[cut:].strip()
            self.emit(self.e(target, 'object') + '(' + ', '.join(self.e(x) for x in split_args(args)) + ');')
        else:
            code = self.e(s, 'object')
            self.emit(code + ('' if code.endswith(')') else '()') + ';')

    def convert(self, source):
        for number,line in enumerate(source.splitlines(),1):
            line = strip_comment(line).strip()
            if not line: continue
            self.emit('// Normalized rule line ' + str(number))
            try: self.statement(line)
            except Exception as ex: raise RuntimeError(str(number) + ': ' + line + '\n' + str(ex)) from ex
        declarations = ['dynamic @' + n + ' = ' + val + ';' for n,val in sorted(self.variables.items())]
        # Format braces without changing source strings.
        result = []; indent = 3
        for line in declarations + self.lines:
            if line.startswith('}'): indent -= 1
            result.append('    '*indent + line)
            if line.endswith('{'): indent += 1
        return '\n'.join(result)

def split_args(s):
    tokens = TOKEN.findall(s); parts=[]; current=[]; depth=0
    for tok in tokens:
        if tok == ',' and depth == 0:
            parts.append(' '.join(current)); current=[]; continue
        if tok == '(': depth += 1
        if tok == ')': depth -= 1
        current.append(tok)
    if current: parts.append(' '.join(current))
    return parts

def strip_comment(s):
    quoted = False
    for i,c in enumerate(s):
        if c == '"': quoted = not quoted
        if c == "'" and not quoted: return s[:i]
    return s

def generate():
    source = (ROOT/'migration/legacy/ZTools/Module1.bas').read_text(encoding='utf-8')
    source = source[source.index('    Dim Pump_Type'):source.index("'used for progress")]
    source = source[:source.rfind('End Sub')]
    # Keep the pump family ZE while selecting ZE7 as its database. Mutating
    # Pump_Type to ZE7 mixes the database name with the pump family.
    start = source.index("      'calculate available rows in ZE,ZF,ZA data source")
    end = source.index("'calculate progress course---13", start)
    ga = source[start:end]
    version_if = 'If Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then'
    counter_start = ga.index(version_if)
    counter_end = ga.index('End If', counter_start) + len('End If')
    old = ga[counter_start:counter_end]
    replacement = 'gaDatabaseSheet = PumpDatabaseSheet(Pump_Type, k)\nCounter = Sheets(gaDatabaseSheet).Range("A" & Rows.Count).End(xlUp).Row - 31'
    ga = ga[:counter_start] + replacement + '\n' * (old.count('\n') - 1) + ga[counter_end:]
    ga = re.sub(re.escape(version_if) + r'\s+origpumptype = Pump_Type\s+Pump_Type = "ZE7"\s*End If',
                lambda m: '\n' * m[0].count('\n'), ga)
    ga = ga.replace('Sheets(Pump_Type)', 'Sheets(gaDatabaseSheet)')
    # ZE7 has f in T; its f* column U is empty. Keep the legacy 7th-edition
    # f selection explicit when Pump_Type remains ZE.
    ga = ga.replace('Sheets("ZE7").Range("$a$31:$bb$2000"), 21, False)', 'Sheets(gaDatabaseSheet).Range("$a$31:$bb$2000"), 20, False)')
    ga = ga.replace('If origpumptype <> "" Then Pump_Type = origpumptype', '')
    source = source[:start] + ga + source[end:]
    # Accept both the Chinese UI value and the original workbook's 7th spelling,
    # including harmless whitespace, for weight/nozzle lookups as well.
    source = source.replace(version_if, 'If PumpDatabaseSheet(Pump_Type, k) = "ZE7" Then')
    # Host handles validation, state restoration, and visible progress.
    source = source.replace('ThisWorkbook.Save', '')
    source = re.sub(r'^.*(?:UserForm2|Unload |UpdateProgress|\.Select$|TypeName\(Selection\)|ScreenUpdating|MsgBox).*$','',source,flags=re.M)
    # The original blank-cell highlight pass flags intentionally unused cells. The host
    # validates required output columns instead, with row/cell diagnostics.
    source = source[:source.index('    CounterCheck = 0')]
    # The old code carries ZE7 state between rows; reset per pump.
    source = source.replace('For k = 0 To i - 1', 'For k = 0 To i - 1\norigpumptype = ""')
    source = source.replace('VBA.Replace(Sheets("input").Range("AB" & 21 + k), VBA.Left(str2, 2), Sheets("input").Range("AB" & 21 + k))', 'VBA.Replace(str2, VBA.Left(str2, 2), Sheets("input").Range("AB" & 21 + k))')
    source = source.replace('For X = 0 To Counter - 1', 'Dim databaseOffset As Long\ndatabaseOffset = FindDatabaseOffset(gaDatabaseSheet, str4, k)\nFor X = databaseOffset To databaseOffset')
    # The legacy SEC branches have an erroneous nested If in the LK6 case.
    # Resolve both names once in the host, and address Sheet1 by name throughout.
    start = source.rfind('\n', 0, source.index('If InStr(Sheets(1).Range("C" & 40 + k)'))
    end = source.index("'calculate progress course---14")
    source = source[:start] + '\nCalculateSectionNames k\n' + source[end:]
    source = re.sub(r'Sheets\(1\)', 'Sheets("Sheet1")', source, flags=re.I)
    body = Port().convert(source)
    # Fail with the selected database/model/baseplate/input row before Excel's
    # first VLookup can replace the useful missing-record diagnosis with 1004.
    first_lookup = 'Book.Worksheets["Sheet1"].Range[Vb.Cat("F", Vb.Add(40, @k))].Value2 = ExcelApp.WorksheetFunction.VLookup'
    assert body.count(first_lookup) == 1
    body = body.replace(first_lookup,
        'report("正在查询第 " + (Vb.Int(@k) + 1) + " 台泵：" + Vb.Text(@str4) + "（" + Vb.Text(@pump_type) + "）……");\n'
        '                @databaseoffset = FindDatabaseOffset(@pump_type, @str4, @k);\n                ' + first_lookup)
    write('WorkbookCalculations', body, 'public void TransferData()')
    for kind in ('GA','PID','SEC'):
        source = (ROOT/f'migration/legacy/{kind}/Module1.bas').read_text(encoding='utf-8')
        source = source[source.index('Dim ptMin'):source.index('\nEnd Sub')]
        # Opening, ownership, output naming and plotting are handled by the host.
        first = source.index('On Error Resume Next')
        end = source.index("'对程序中变量赋值") if "'对程序中变量赋值" in source else source.index('i = Val(')
        mappings = {'GA': 'Set xlsheet1 = Book.Worksheets("input")\n',
                    'PID': 'Set xlsheet1 = Book.Worksheets("PID")\nSet xlsheet2 = Book.Worksheets("input")\n',
                    'SEC': 'Set xlsheet1 = Book.Worksheets("Sec-List")\nSet xlsheet2 = Book.Worksheets("input")\nSet xlsheet3 = Book.Worksheets("CZ-BearingSize")\n'}
        source = source[:first] + '\nSet xlbook = Book\nSet xlsheet = Book.Worksheets("Sheet1")\n' + mappings[kind] + source[end:]
        # Extract only drawing construction, ending before PDF configuration.
        plot = source.index('Set objLayout =')
        source = source[:source.rfind('Set objdoc =', 0, plot)]
        source = re.sub(r'^.*(?:ThisDrawing\.Application\.Visible|ThisDrawing\.ActiveTextStyle\.SetFont).*$','',source,flags=re.M)
        # Source incorrectly asks ObjectName (a string) for a Name property.
        source = re.sub(r'tempString\.Name', 'ThisDrawing.Layouts.Item("Model").Block(Index).Name', source, flags=re.I)
        # Actual templates split title data across hetongxinxi/title/etc., and name
        # the BOM block List. Apply the supplied tag maps to every block reference.
        source = re.sub(r'If ThisDrawing\.Layouts\.Item\("Model"\)\.Block\(Index\)\.Name = "(?:AAA|BBB)" Then', 'If True Then', source, flags=re.I)
        # DATE is a formatted Excel date; Value2 would expose the OLE serial number.
        source = re.sub(r'(attrib\.TextString\s*=\s*xlsheet\.range\(\("C" & 6\)\))', r'\1.Text', source, flags=re.I)
        if kind == 'GA':
            # Only the main pump block varies by edition. Keep the ZE family
            # and k/m=8 so all auxiliary block selections match the ZE8 path.
            source = source.replace('If Ver = 7 Then pump = "ZE7"', '')
            source = source.replace('InsertBlock p, pump, 1, 1, 1, 0', 'InsertBlock p, PumpDatabaseSheet(pump, i - 40), 1, 1, 1, 0')
        if kind == 'PID':
            source = re.sub(r'ThisDrawing\.Application\.ActiveDocument\.ModelSpace\.InsertBlock ([pu]), (p[1-6]), 1, 1, 1, 0', r'InsertPidPlan \1, \2', source)
        if kind == 'SEC':
            # Insert the exact Sheet1 CA result, once, including the CZ case.
            start = source.index('If pump = "CZ" Then',source.index('"PUREOIL"'))
            end = source.index('Count = ThisDrawing.Layouts.Item("Model").Block.Count',start)
            source = source[:start] + 'InsertSection p, i\n' + source[end:]
        # Bare project-context Book must refer to the host, not a local.
        body = Port().convert(source).replace('dynamic @book = null;','').replace('@book','Book')
        write(kind + 'Drawing', body, 'public void Build' + kind + '()')
    print('Generated C# for workbook calculations and GA/PID/SEC drawing construction.')

def write(name, body, signature):
    (OUT/(name+'.cs')).write_text('// Ported from migration/legacy; see tools/port_vba.py and migration/README.md.\n'
        'using System;\nnamespace ZTools.Engine {\n    internal sealed partial class SubmissionEngine {\n        ' + signature + ' {\n' + body + '\n        }\n    }\n}\n',encoding='utf-8')

if __name__ == '__main__': generate()
