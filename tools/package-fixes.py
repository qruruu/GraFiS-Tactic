"""Build a module patch package against VBA extracted from the reviewed stencils."""
from pathlib import Path
import argparse,hashlib,json,re,subprocess,zipfile
p=argparse.ArgumentParser();p.add_argument('--embedded-dir',type=Path,required=True);p.add_argument('--output',type=Path,required=True);args=p.parse_args()
root=Path(__file__).resolve().parents[1]
def decode(path):
 b=path.read_bytes()
 try:return b.decode('utf8'),'utf-8'
 except UnicodeDecodeError:return b.decode('cp1251'),'windows-1251'
def code(text):
 ls=text.splitlines();start=next((i for i,l in enumerate(ls) if l.startswith('Attribute VB_Name')),0)
 return '\n'.join(l.rstrip() for l in ls[start:] if not l.startswith('Attribute ')).strip()
def digest(text):return hashlib.sha256(code(text).encode('utf8')).hexdigest()
changed=set(filter(None,subprocess.check_output(['git','diff','HEAD','--name-only','-z'],cwd=root).decode().split('\0')))
changed.update(filter(None,subprocess.check_output(['git','ls-files','--others','--exclude-standard','-z'],cwd=root).decode().split('\0')))
projects={};source_only=[];module_paths=[]
for rel in sorted(changed):
 f=root/rel
 if not rel.startswith('_codes/') or f.suffix not in ('.bas','.cls','.frm'):continue
 name=f.parent.name;stencil=root/(name+'.vss');text,encoding=decode(f)
 module=re.search(r'(?m)^Attribute VB_Name = "([^"\n]+)"',text)[1]
 embedded=args.embedded_dir/name/('ThisDocument.cls' if module=='ThisDocument' else f.name)
 is_new=f.name=='m_GFSSafety.bas'
 if not stencil.exists() or (not embedded.exists() and not is_new):
  source_only.append(rel);continue
 original_hash=digest(embedded.read_text()) if embedded.exists() else None
 projects.setdefault(name,{'stencil':name+'.vss','modules':[]})['modules'].append({'name':module,'path':rel,'encoding':encoding,'is_new':is_new,'original_hash':original_hash,'new_hash':digest(text)})
 module_paths.append(rel)
manifest={'base_commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=root).decode().strip(),'projects':list(projects.values()),'source_only':source_only,'runtime_verified':False}
(root/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf8')
patch=subprocess.check_output(['git','diff','--binary','HEAD'],cwd=root)
(root/'changes.patch').write_bytes(patch)
include=module_paths+source_only+['manifest.json','README-FIXES.md','AUDIT.md','changes.patch','tools/apply-visio-fixes.ps1','tools/check-vba-sources.py','tests/GFS_ContextTests.bas','tests/GFS_YamlFailureTests.bas']
args.output.parent.mkdir(parents=True,exist_ok=True)
with zipfile.ZipFile(args.output,'w',zipfile.ZIP_DEFLATED) as z:
 for rel in include:z.write(root/rel,rel)
print(json.dumps({'archive':str(args.output),'projects':len(projects),'installable_modules':len(module_paths),'source_only':source_only},ensure_ascii=False))
