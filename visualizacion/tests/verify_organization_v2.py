"""Final read-only syntax and local Markdown link audit (no model execution)."""
from pathlib import Path
import ast
import json
import re
from urllib.parse import unquote, urlsplit
ROOT=Path(__file__).resolve().parents[2]
roots=['pharmacy/scripts','pharmacy/tests','visualizacion/scripts','visualizacion/tests']
python_files=[p for folder in roots for p in (ROOT/folder).glob('*.py')]+[ROOT/'visualizacion/server.py']
failures=[]
for p in python_files:
    try: ast.parse(p.read_text(encoding='utf-8-sig'),filename=str(p))
    except (SyntaxError,UnicodeError) as e: failures.append({'path':str(p.relative_to(ROOT)),'error':str(e)})
documents=[ROOT/'README.md',ROOT/'visualizacion/README.md',ROOT/'visualizacion/QA.md']
documents+=list((ROOT/'loop').glob('*.md'))+list((ROOT/'pharmacy/analysis').rglob('*.md'))
broken=[];checked=0;external=0
for p in documents:
    for target in re.findall(r'\[[^\]\n]+\]\(([^)\n]+)\)',p.read_text(encoding='utf-8-sig')):
        target=target.strip().split(' \"',1)[0].strip('<>')
        if urlsplit(target).scheme or target.startswith('#'):
            external+=1;continue
        dest=unquote(target.split('#',1)[0])
        checked+=1
        if not (p.parent/dest).exists():broken.append({'document':str(p.relative_to(ROOT)),'target':target})
path=ROOT/'visualizacion/tests/organization_v2_controls.json'
report=json.loads(path.read_text(encoding='utf-8'))
report['ast']={'files_checked':len(python_files),'failure_count':len(failures),'failures':failures,'method':'ast.parse; -B; no imports'}
report['markdown_links']={'documents_checked':len(documents),'links_checked':checked,'external_links_skipped':external,'broken_count':len(broken),'broken':broken}
report['passed']=not failures and not broken and report['iqvia_inputs']['all_match']
path.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:report[k] for k in ['ast','markdown_links','passed']},ensure_ascii=False))
assert report['passed']
