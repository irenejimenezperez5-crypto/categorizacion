"""Read-only verification after two actual arrancar_app.cmd executions."""
from pathlib import Path
from urllib.request import urlopen
import argparse
import hashlib
import json
ROOT=Path(__file__).resolve().parents[2]
parser=argparse.ArgumentParser()
parser.add_argument('--expected-pid',type=int,required=True,help='PID recorded after first actual launch')
args=parser.parse_args()
with urlopen('http://127.0.0.1:8765/api/health',timeout=10) as r: health=json.load(r)
with urlopen('http://127.0.0.1:8765/data/metadata.json',timeout=10) as r: metadata=json.load(r)
with urlopen('http://127.0.0.1:8765/',timeout=30) as r: html_sha=hashlib.sha256(r.read()).hexdigest()
local_sha=hashlib.sha256((ROOT/'visualizacion/app/dist/index.html').read_bytes()).hexdigest()
server=json.loads((ROOT/'visualizacion/logs/server.json').read_text())
snapshot=json.loads((ROOT/'visualizacion/app/src/data.json').read_text(encoding='utf-8'))
checks={'own_server_ready':health=={'app_id':'farmacias-potencial-2027','project_key':'c338604c1b558b7e','ready':True},'current_version':metadata.get('version')=='2027.2','current_book_hash':metadata['excel']['sha256']=='6991b58caf1d10bd05cd21b84e40b5f450b8e7333895b013846c33553ab812df','served_html_matches_local_build':html_sha==local_sha,'complete_commercial_app':snapshot['buildStatus']=='complete' and snapshot['commercialReport'],'second_launch_reused_server':server['pid']==args.expected_pid and server['port']==8765}
report={'date':'2026-10-07','launcher':'arrancar_app.cmd','executions':2,'start_method':'Actual Windows CMD; opens default browser; hidden pythonw server','first_launch_server':{'pid':args.expected_pid,'port':8765},'second_launch_server':server,'health':health,'version':metadata.get('version'),'total_stores':metadata['total_stores'],'total_crm':metadata['total_crm'],'national_market_eur':metadata['total_market_eur'],'workbook_sha256':metadata['excel']['sha256'],'html_sha256':html_sha,'checks':checks,'passed':all(checks.values()),'left_running':True}
(ROOT/'visualizacion/tests/launcher_v2_controls.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
assert report['passed'],report
