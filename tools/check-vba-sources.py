"""Static checks for this repair, not a VBA compiler. Run native tests in Visio."""
from pathlib import Path
import re, subprocess, collections
ROOT=Path(__file__).resolve().parents[1]
def read(p):
 b=p.read_bytes()
 try:return b.decode('utf8').replace('\r\n','\n')
 except UnicodeDecodeError:return b.decode('cp1251').replace('\r\n','\n')
files=[p for p in (ROOT/'_codes').rglob('*') if p.suffix in ('.bas','.cls','.frm')]
issues=[];projects=collections.defaultdict(list)
for p in files:projects[p.parent].append(p)
for project,modules in projects.items():
 helper=project/'m_GFSSafety.bas'
 for p in modules:
  s=read(p)
  if re.search(r'\bGFS(?:ResolveShape|BeginCellEvent|EndCellEvent|CurrentTime|IsActiveDrawingShape)\s*\(?',s) and not helper.exists():issues.append(f'{p}: missing safety helper')
  if p.name=='ThisDocument.bas':
   for m in re.finditer(r'(?im)^Private Sub \w+_CellChangedCore\([^\n]*\)(.*?)(?=^End Sub)',s,re.S):
    if re.search(r'(?im)^\s*\w+Import\s*\(ShpInd\)',m[1]):issues.append(f'{p}: event still passes ID without source context')
  # Every new optional-context routine must resolve through its source shape.
  for m in re.finditer(r'(?im)^[^\n]*(?:Sub|Function) \w+\([^\n]*sourceShape As Visio.Shape[^\n]*\)(.*?)(?=^End (?:Sub|Function))',s,re.S):
   if 'Application.ActivePage.Shapes.ItemFromID(ShpIndex)' in m[1]:issues.append(f'{p}: optional source context bypassed')
  # Core/wrapper signatures remain paired, including object parameters passed ByRef.
  for core,args in re.findall(r'(?im)^(?:Public |Private )?Sub (\w+EventsCore)\(([^\n]*)\)',s):
   name=core[:-len('EventsCore')]
   wrapper=re.search(r'(?im)^(?:Public |Private )?Sub '+name+r'\(([^\n]*)\)',s)
   if wrapper is None or wrapper[1]!=args:issues.append(f'{p}: wrapper signature mismatch {name}')
  # Detect duplicated Sub/Function definitions newly introduced by edits.
  declarations=re.findall(r'(?im)^(?:Public |Private |Friend )?(?:Sub|Function) (\w+)\(',s)
  dups=[n for n,c in collections.Counter(x.lower() for x in declarations).items() if c>1]
  if dups:issues.append(f'{p}: duplicate procedures {dups}')
# Cross-project helper copies must be identical.
helpers=list((ROOT/'_codes').rglob('m_GFSSafety.bas'))
if len({p.read_bytes() for p in helpers})!=1:issues.append('Safety helpers diverged across projects')
# Catch changed modules with missing module headers / accidental null bytes.
changed=subprocess.check_output(['git','diff','HEAD','--name-only','-z'],cwd=ROOT).decode().split('\0')
for rel in filter(None,changed):
 p=ROOT/rel
 if p.suffix not in ('.bas','.cls','.frm') or not p.exists():continue
 s=read(p)
 if not re.search(r'(?m)^Attribute VB_Name = "\w+"',s):issues.append(f'{rel}: missing module name')
 if '\x00' in s:issues.append(f'{rel}: NUL byte in source')
if issues:
 print('\n'.join(issues));raise SystemExit(1)
print(f'PASS: static consistency checks across {len(files)} VBA source modules; {len(helpers)} identical helpers')
