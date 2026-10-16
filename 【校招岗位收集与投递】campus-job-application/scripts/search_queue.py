"""Persistent dual-channel search coordination; no account/form operations."""
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlsplit
import hashlib
import json
import sqlite3
import time
import uuid

CHANNELS=('chrome','iab')
DEFAULT_COMPANIES_PER_CHANNEL=10
STAGES=('collection','verification','structuring','recommendation')
CHECKS={1:{'company_identity','cohort','plans_enumerated','catalog_pages_enumerated','counts_reconciled','source_saved'},
        2:{'plans','catalog','jd','eligibility','rules','referrals','preferences'},
        3:{'typed_fields','foreign_keys','counts','source_provenance','history_preserved'},
        4:{'hard_scope','three_tiers','qualification','quota_matrix','special_plan_decisions','selection_preserved'}}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class Queue:
    def __init__(self, root):
        self.root=Path(root).resolve();self.root.mkdir(parents=True,exist_ok=True)
        with self.db() as db:
            db.executescript('''
              CREATE TABLE IF NOT EXISTS batches(id TEXT PRIMARY KEY, manifest TEXT NOT NULL, stage INTEGER NOT NULL);
              CREATE TABLE IF NOT EXISTS tasks(id TEXT PRIMARY KEY,batch TEXT NOT NULL,company TEXT NOT NULL,
                source_key TEXT NOT NULL,position INTEGER NOT NULL,stage INTEGER NOT NULL,status TEXT NOT NULL,
                channel TEXT,token TEXT,expires REAL,evidence TEXT,evidence_sha TEXT,gap TEXT,
                UNIQUE(batch,company,stage));
              CREATE TABLE IF NOT EXISTS events(id INTEGER PRIMARY KEY AUTOINCREMENT,batch TEXT,at REAL,event TEXT,payload TEXT);
            ''')

    @contextmanager
    def db(self):
        db=sqlite3.connect(self.root/'queue.sqlite',timeout=30,isolation_level=None)
        db.row_factory=sqlite3.Row
        try:
            db.execute('BEGIN IMMEDIATE');yield db;db.commit()
        except BaseException:
            db.rollback();raise
        finally:db.close()

    def ref(self, relative):
        path=(self.root/relative).resolve()
        if not path.is_relative_to(self.root) or not path.is_file():
            raise ValueError('Evidence must be an existing file inside the private queue directory')
        return path

    def event(self,db,batch,event,payload):
        db.execute('INSERT INTO events(batch,at,event,payload) VALUES(?,?,?,?)',
                   (batch,time.time(),event,json.dumps(payload,ensure_ascii=False)))

    def create(self, manifest):
        required={'batch_id','scope','authorization_ref','workbook_sha256','companies'}
        if not required<=manifest.keys() or not manifest['companies'] or not manifest['scope']:
            raise ValueError('Missing frozen scope, company queue, workbook hash or authorization')
        self.ref(manifest['authorization_ref'])
        companies=manifest['companies'];ids=[c['company_key'] for c in companies]
        if any(not isinstance(c.get('name'),str) or not c['name'] or not c.get('source_key') for c in companies) or len(set(ids))!=len(ids):
            raise ValueError('Company identity/source missing or duplicate')
        batch=manifest['batch_id']
        with self.db() as db:
            old=db.execute('SELECT manifest FROM batches WHERE id=?',(batch,)).fetchone()
            encoded=json.dumps(manifest,ensure_ascii=False,sort_keys=True)
            if old:
                if old['manifest']!=encoded:raise ValueError('Batch already frozen with different scope')
                return {'batch_id':batch,'created':False,'companies':len(ids)}
            db.execute('INSERT INTO batches VALUES(?,?,1)',(batch,encoded))
            for i,c in enumerate(companies):
                for stage in range(1,5):
                    db.execute('INSERT INTO tasks(id,batch,company,source_key,position,stage,status) VALUES(?,?,?,?,?,?,?)',
                               (str(uuid.uuid4()),batch,c['company_key'],c['source_key'],i,stage,'pending'))
            self.event(db,batch,'create',{'companies':len(ids),'channels':CHANNELS})
        return {'batch_id':batch,'created':True,'companies':len(ids),'execution_batches':(len(ids)+9)//10}

    def dispatch(self,batch,channels=CHANNELS,lease_seconds=900,companies_per_channel=DEFAULT_COMPANIES_PER_CHANNEL):
        if not channels or len(set(channels))!=len(channels) or not set(channels)<=set(CHANNELS):
            raise ValueError('Channels must be distinct chrome/iab')
        if type(lease_seconds)!=int or not 30<=lease_seconds<=3600:raise ValueError('Invalid lease duration')
        if type(companies_per_channel)!=int or not 1<=companies_per_channel<=10:
            raise ValueError('Company concurrency per channel must be an integer from 1 to 10')
        assignments=[]
        with self.db() as db:
            b=db.execute('SELECT * FROM batches WHERE id=?',(batch,)).fetchone()
            if not b:raise ValueError('Unknown batch')
            db.execute("UPDATE tasks SET status='blocked',gap='Lease expired; inspect existing page/evidence before explicit retry' WHERE batch=? AND status='running' AND expires<?",(batch,time.time()))
            manifest=json.loads(b['manifest']);index={c['company_key']:c for c in manifest['companies']}
            for slot in range(companies_per_channel):
                for channel in channels:
                    if db.execute("SELECT COUNT(*) FROM tasks WHERE batch=? AND channel=? AND status='running'",(batch,channel)).fetchone()[0]>=companies_per_channel:continue
                    task=db.execute("""SELECT * FROM tasks t WHERE batch=? AND stage=? AND status='pending'
                      AND NOT EXISTS(SELECT 1 FROM tasks a WHERE a.batch=t.batch AND a.source_key=t.source_key AND a.status='running')
                      ORDER BY position LIMIT 1""",(batch,b['stage'])).fetchone()
                    if not task:continue
                    token=str(uuid.uuid4());expires=time.time()+lease_seconds
                    db.execute("UPDATE tasks SET status='running',channel=?,token=?,expires=? WHERE id=?",(channel,token,expires,task['id']))
                    assignments.append({'task_id':task['id'],'lease_token':token,'channel':channel,
                        'stage':task['stage'],'operation':STAGES[task['stage']-1],
                        'company':index[task['company']],'scope':manifest['scope'],
                        'execution_batch':task['position']//10+1,'expires':expires,
                        'browser_executor':'host_browser_mcp','application_actions':'forbidden'})
            self.event(db,batch,'dispatch',assignments)
        return {'batch_id':batch,'assignments':assignments,'concurrency':len(assignments),
                'company_capacity_per_channel':companies_per_channel,'capacity_unit':'company',
                'execution_required':True,'note':'Assignments do not prove browser execution or verified coverage'}

    def record(self,batch,task_id,lease_token,evidence_ref):
        path=self.ref(evidence_ref);e=json.loads(path.read_text(encoding='utf-8'))
        with self.db() as db:
            t=db.execute('SELECT * FROM tasks WHERE id=? AND batch=?',(task_id,batch)).fetchone()
            if not t or t['status']!='running' or t['token']!=lease_token or t['expires']<time.time():
                raise ValueError('Stale, duplicate or foreign lease')
            if any(e.get(k)!=v for k,v in {'task_id':task_id,'company_key':t['company'],'channel':t['channel'],'stage':t['stage']}.items()):
                raise ValueError('Evidence identity/channel/stage mismatch')
            if not e.get('observed_at') or not e.get('observations') or not e.get('checks') or not isinstance(e.get('gaps'),list):
                raise ValueError('Missing observations, source time, checks or explicit gaps')
            sources=e.get('source_urls')
            if not isinstance(sources,list) or not sources or any(urlsplit(u).scheme not in ('http','https') or not urlsplit(u).netloc for u in sources):
                raise ValueError('Actual source URLs required')
            if any(type(v)!=bool for v in e['checks'].values()):raise ValueError('Checks must be booleans')
            if not CHECKS[t['stage']]<=e['checks'].keys():raise ValueError('Missing required stage checks')
            artifacts=e.get('artifacts',[])
            if not artifacts:raise ValueError('Browser capture or structured source artifact required')
            for a in artifacts:
                if sha(self.ref(a['path']))!=a['sha256']:raise ValueError('Artifact missing or changed')
            if t['stage']==1 and all(e['checks'].values()) and not e['gaps']:
                inventory=e.get('inventory',{});plans=inventory.get('plans');roles=inventory.get('roles')
                if not isinstance(plans,list) or not isinstance(roles,list):raise ValueError('Complete plan/job name inventory required')
                for name,items in [('plans',plans),('roles',roles)]:
                    expected=inventory.get('expected_'+name)
                    if type(expected)!=int or expected!=len(items) or expected<0:raise ValueError('Inventory count mismatch')
                    ids=[x['id'] for x in items]
                    if len(set(ids))!=len(ids) or any(not x.get('title' if name=='roles' else 'name') for x in items):raise ValueError('Inventory identity/name missing or duplicate')
                pids={p['id'] for p in plans}
                if any(not r.get('plan_ids') or not set(r['plan_ids'])<=pids for r in roles):raise ValueError('Inventory role plan mismatch')
            if t['stage']>=3:
                import company_handoff as h
                c=json.loads(self.ref(e['company_payload_ref']).read_text(encoding='utf-8'))
                if c['company_key']!=t['company']:raise ValueError('Company payload mismatch')
                # A structured unknown stays unknown; research and user-confirmation
                # readiness are checked separately by the application handoff gate.
                errors=h.validate_company(c)+h.validate_research(c)
                if errors:raise ValueError('; '.join(errors))
                if not any(a['path']==e['company_payload_ref'] for a in artifacts):raise ValueError('Company payload must be hashed')
            status='checked' if all(e['checks'].values()) and not e['gaps'] else 'blocked'
            db.execute('UPDATE tasks SET status=?,evidence=?,evidence_sha=?,gap=? WHERE id=?',
                       (status,evidence_ref,sha(path),json.dumps(e['gaps'],ensure_ascii=False),task_id))
            self.event(db,batch,'record',{'task_id':task_id,'status':status,'evidence_ref':evidence_ref})
        return {'task_id':task_id,'status':status,'evidence_ref':evidence_ref,'fact_verification':'executor-responsibility'}

    def check(self,batch):
        with self.db() as db:
            b=db.execute('SELECT * FROM batches WHERE id=?',(batch,)).fetchone()
            if not b:raise ValueError('Unknown batch')
            tasks=[dict(t) for t in db.execute('SELECT * FROM tasks WHERE batch=? AND stage=? ORDER BY position',(batch,b['stage']))]
        issues=[];counts={s:sum(t['status']==s for t in tasks) for s in ('pending','running','blocked','checked')}
        for t in tasks:
            if t['evidence']:
                try:
                    path=self.ref(t['evidence'])
                    if sha(path)!=t['evidence_sha']:raise ValueError('changed')
                    evidence=json.loads(path.read_text(encoding='utf-8'))
                    for a in evidence['artifacts']:
                        if sha(self.ref(a['path']))!=a['sha256']:raise ValueError('artifact changed')
                except (OSError,ValueError,KeyError):issues.append({'task_id':t['id'],'reason':'Evidence or referenced artifact changed'})
        result={'batch_id':batch,'stage':b['stage'],'operation':STAGES[b['stage']-1],
                'companies':len(tasks),'counts':counts,'issues':issues,
                'passed':counts['checked']==len(tasks) and not issues,
                'audited':counts['pending']==0 and counts['running']==0 and not issues,
                'unresolved':[{'task_id':t['id'],'company_key':t['company'],'status':t['status'],'gap':t['gap']} for t in tasks if t['status']!='checked']}
        # Batch identity is data, never a path segment.
        report=self.root/'round-checks'/(hashlib.sha256(batch.encode()).hexdigest()[:16]+f'-{b["stage"]}.json')
        report.parent.mkdir(parents=True,exist_ok=True)
        report.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        result['unresolved_count']=len(result['unresolved']);result['unresolved']=result['unresolved'][:10]
        result['report_ref']=str(report.relative_to(self.root))
        return result

    def advance(self,batch):
        result=self.check(batch)
        if not result['audited']:raise ValueError('Round not fully checked; independent pending companies may continue')
        with self.db() as db:
            stage=db.execute('SELECT stage FROM batches WHERE id=?',(batch,)).fetchone()['stage']
            if stage!=result['stage']:raise ValueError('Round changed concurrently')
            if stage>=4:raise ValueError('Research stages finished; workbook write and user HTML review remain separate gates')
            db.execute('UPDATE batches SET stage=stage+1 WHERE id=?',(batch,))
            self.event(db,batch,'advance',{'from':stage,'to':stage+1,'unresolved_carried':result['counts']['blocked'],'check_report':result['report_ref']})
        following=self.check(batch)
        following['previous_round_blocked']=result['counts']['blocked']
        following['previous_round_passed']=result['passed']
        following['previous_round_report_ref']=result['report_ref']
        return following

    def retry(self,batch,task_id,reason):
        if not reason.strip():raise ValueError('Recovery evidence/reason required')
        with self.db() as db:
            t=db.execute('SELECT * FROM tasks WHERE id=? AND batch=?',(task_id,batch)).fetchone()
            if not t or t['status']!='blocked':raise ValueError('Only explicitly blocked task may be retried')
            self.event(db,batch,'retry',{'task_id':task_id,'reason':reason,'previous_evidence':t['evidence']})
            db.execute("UPDATE tasks SET status='pending',token=NULL,expires=NULL,evidence=NULL,evidence_sha=NULL,gap=NULL WHERE id=?",(task_id,))
        return {'task_id':task_id,'status':'pending'}
