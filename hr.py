"""Effective-dated personnel, job classification and internally calculated reports."""
import io, json, re, secrets
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from flask import request, jsonify, session, abort, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from docx import Document
from docx.shared import Pt
from source_data import source_status, education_at, is_person_row

GROUPS = {'management':'Quản lý','technology':'Công nghệ','auxiliary':'Phụ trợ','service':'Phục vụ','unknown':'Chưa xác định'}
EDUCATION = {'postgraduate':'Trên đại học','university':'Đại học','college':'Cao đẳng','intermediate':'Trung cấp','vocational':'CNKT','untrained':'Chưa qua đào tạo','unknown':'Chưa xác định'}
STATUSES = {'active':'Đang làm việc','terminated':'Đã nghỉ','review':'Cần rà soát'}
FIELDS = ('name','birth_date','hire_date','gender','education','job_id','status')
KINDS = {'transfer':'Điều động','title':'Thay đổi chức danh','termination':'Chấm dứt HĐLĐ'}

def today(): return datetime.now(timezone(timedelta(hours=7))).date()

def iso(value, required=True):
    if not value and not required: return ''
    if not isinstance(value,str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}',value): raise ValueError('Ngày phải có dạng yyyy-mm-dd.')
    return date.fromisoformat(value).isoformat()

def parse_excel_date(v):
    if isinstance(v,dict): v=v.get('date')
    if isinstance(v,datetime): return v.date().isoformat()
    if isinstance(v,date): return v.isoformat()
    for fmt in ('%Y-%m-%d','%Y-%m-%dT%H:%M:%S','%d/%m/%Y'):
        try: return datetime.strptime(str(v).strip(),fmt).date().isoformat()
        except ValueError: pass
    return ''

def age_group(birth, asof, scheme='standard'):
    if not birth: return None,'Chưa xác định'
    born=date.fromisoformat(birth);at=date.fromisoformat(asof)
    age=at.year-born.year-((at.month,at.day)<(born.month,born.day))
    if age<0:return None,'Ngày sinh sau ngày chốt'
    if scheme=='primary':
        return age, 'Dưới 30' if age<30 else '30–40' if age<=40 else '41–50' if age<=50 else '51–60' if age<=60 else 'Trên 60'
    return age,'Dưới 25' if age<25 else '25–35' if age<=35 else '36–45' if age<=45 else '46–55' if age<=55 else '56–60' if age<=60 else 'Trên 60'

def make_document(root, kind, fields):
    path=root/('template-'+kind+'.docx')
    if path.exists():
        doc=Document(path)
        def container(obj):
            for p in obj.paragraphs:
                # Replace placeholders across runs while preserving surrounding formatting.
                for k,v in fields.items():
                    token='{{'+k+'}}'
                    # Do not re-scan the replacement text as a template.
                    original=p.text; matches=list(re.finditer(re.escape(token),original))
                    for match in reversed(matches):
                        start,end=match.span();offset=0;hits=[]
                        for run in p.runs:
                            a=offset;b=a+len(run.text);offset=b
                            if b>start and a<end:hits.append((run,a))
                        if not hits:continue
                        first,a=hits[0];prefix=first.text[:start-a]
                        if len(hits)==1:first.text=prefix+str(v)+first.text[end-a:]
                        else:
                            suffix=hits[-1][0].text[end-hits[-1][1]:]
                            first.text=prefix+str(v)
                            for run,_ in hits[1:-1]:run.text=''
                            hits[-1][0].text=suffix
            for t in obj.tables:
                for row in t.rows:
                    for cell in row.cells:container(cell)
        container(doc)
        for sec in doc.sections:container(sec.header);container(sec.footer)
    else:
        doc=Document();doc.styles['Normal'].font.name='Times New Roman';doc.styles['Normal'].font.size=Pt(13)
        doc.add_paragraph('BẢN DỰ THẢO — CHƯA ÁP DỤNG MẪU ĐƠN VỊ',style='Subtitle')
        doc.add_heading('QUYẾT ĐỊNH '+KINDS[kind].upper(),0)
        for label,key in [('Số','so_quyet_dinh'),('Người lao động','ho_ten'),('Mã nhân sự','ma_nv'),('Chức danh cũ','chuc_danh_cu'),('Đơn vị cũ','don_vi_cu'),('Chức danh mới','chuc_danh_moi'),('Đơn vị mới','don_vi_moi'),('Ngày hiệu lực','ngay_hieu_luc'),('Căn cứ / nội dung','can_cu'),('Người ký','nguoi_ky')]:
            doc.add_paragraph(label+': '+fields.get(key,''))
        doc.add_paragraph('Kiểm tra nội dung, căn cứ và thẩm quyền trước khi ký. Xuất bản dự thảo chưa thay đổi hồ sơ; chỉ thao tác Áp dụng quyết định đã duyệt mới ghi nhận biến động.')
    out=io.BytesIO();doc.save(out);out.seek(0);return out

def register_hr(app, db, lock, auth, audit, book, patches, root):
    with db() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS hr_jobs(id TEXT PRIMARY KEY,title TEXT NOT NULL,unit TEXT NOT NULL,version INTEGER NOT NULL DEFAULT 0,UNIQUE(title,unit));
        CREATE TABLE IF NOT EXISTS hr_job_rules(id INTEGER PRIMARY KEY,job TEXT NOT NULL,effective TEXT NOT NULL,labor_group TEXT NOT NULL,technical INTEGER NOT NULL,grade INTEGER,UNIQUE(job,effective));
        CREATE TABLE IF NOT EXISTS hr_people(id TEXT PRIMARY KEY,code TEXT NOT NULL UNIQUE,version INTEGER NOT NULL DEFAULT 0,source_book TEXT,source_row INTEGER,UNIQUE(source_book,source_row));
        CREATE TABLE IF NOT EXISTS hr_events(id INTEGER PRIMARY KEY,person TEXT NOT NULL,effective TEXT NOT NULL,kind TEXT NOT NULL,snapshot TEXT NOT NULL,decision TEXT UNIQUE,actor TEXT NOT NULL,stamp TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS hr_events_person_date ON hr_events(person,effective,id);
        CREATE TABLE IF NOT EXISTS hr_source_evidence(person TEXT PRIMARY KEY,book TEXT NOT NULL,source_row INTEGER NOT NULL,source_revision INTEGER NOT NULL,end_column TEXT,end_code TEXT,status TEXT,observed_date TEXT,source_code TEXT);
        CREATE TABLE IF NOT EXISTS hr_decisions(id TEXT PRIMARY KEY,person TEXT NOT NULL,version INTEGER NOT NULL,payload TEXT NOT NULL,applied INTEGER NOT NULL DEFAULT 0,stamp TEXT NOT NULL,actor TEXT NOT NULL);
        ''')
    with db() as c:
        if 'source_code' not in {r['name'] for r in c.execute('PRAGMA table_info(hr_source_evidence)')}:
            c.execute('ALTER TABLE hr_source_evidence ADD COLUMN source_code TEXT')
    def fail(message): abort(400,message)
    def read_date(v,required=True):
        try:return iso(v,required)
        except (ValueError,TypeError):fail('Ngày không hợp lệ, cần yyyy-mm-dd.')
    def string(d,key,required=False):
        v=d.get(key,'')
        if not isinstance(v,str) or len(v)>4000 or (required and not v.strip()):fail('Thông tin '+key+' không hợp lệ.')
        return v.strip()
    def job(c,jid,asof):
        j=c.execute('SELECT * FROM hr_jobs WHERE id=?',(jid,)).fetchone()
        if not j:return dict(id='',title='Chưa xác định',unit='Chưa xác định',labor_group='unknown',technical=False,grade=None,version=0)
        rule=c.execute('SELECT * FROM hr_job_rules WHERE job=? AND effective<=? ORDER BY effective DESC,id DESC LIMIT 1',(jid,asof)).fetchone()
        return dict(j) | dict(labor_group=rule['labor_group'] if rule else 'unknown',technical=bool(rule['technical']) if rule else False,grade=rule['grade'] if rule else None)
    def person(c,pid):
        p=c.execute('SELECT * FROM hr_people WHERE id=?',(pid,)).fetchone()
        if not p:abort(404,'Không có hồ sơ nhân sự.')
        return p
    def latest(c,pid,asof=None):
        return c.execute('SELECT * FROM hr_events WHERE person=?'+(' AND effective<=?' if asof else '')+' ORDER BY effective DESC,id DESC LIMIT 1',(pid,asof) if asof else (pid,)).fetchone()
    def classified(c,p,event,asof,scheme='standard'):
        data=json.loads(event['snapshot']);j=job(c,data['job_id'],asof);age,band=age_group(data['birth_date'],asof,scheme)
        warnings=[]
        if not data['birth_date']:warnings.append('Thiếu ngày sinh')
        if not data['hire_date']:warnings.append('Thiếu ngày vào làm')
        if j['technical'] and j['grade'] is None:warnings.append('CNKT chưa xác định bậc thợ')
        if data['education']=='unknown':warnings.append('Chưa xác định trình độ')
        if j['labor_group']=='unknown':warnings.append('Vị trí chưa có nhóm lao động')
        if data['status']=='review':warnings.append('Chưa xác nhận trạng thái làm việc')
        if not data['gender']:warnings.append('Chưa xác nhận giới tính')
        evidence=c.execute('SELECT * FROM hr_source_evidence WHERE person=?',(p['id'],)).fetchone()
        if p['code'].startswith('LUU-TRU-') and evidence:
            warnings.append('Mã nhân sự nguồn trùng; cần rà soát trùng người' if evidence['source_code'] else 'Thiếu mã nhân sự nguồn; dùng mã lưu trữ nội bộ')
        if data['status']=='terminated' and event['kind'] in ('baseline','source_status'):warnings.append('Đã nghỉ theo mã cuối kỳ; ngày ghi nhận không phải ngày nghỉ thực tế')
        return dict(source_evidence=dict(evidence) if evidence else None,id=p['id'],code=p['code'],version=p['version'],**data,age=age,age_group=band,job=j,warnings=warnings,effective=event['effective'],event_kind=event['kind'])
    def validate_snapshot(c,d,effective):
        values={k:string(d,k,k in ('name','status','education')) for k in FIELDS}
        values['birth_date']=read_date(values['birth_date'],False);values['hire_date']=read_date(values['hire_date'],False)
        if values['birth_date'] and values['birth_date']>effective:fail('Ngày sinh phải trước ngày ghi nhận.')
        if values['hire_date'] and values['birth_date'] and values['hire_date']<values['birth_date']:fail('Ngày vào làm phải sau ngày sinh.')
        if values['hire_date'] and values['hire_date']>effective:fail('Ngày vào làm không được sau ngày ghi nhận hồ sơ.')
        if values['gender'] not in ('','male','female','other') or values['education'] not in EDUCATION or values['status'] not in STATUSES:fail('Danh mục không hợp lệ.')
        if values['job_id'] and not c.execute('SELECT 1 FROM hr_jobs WHERE id=?',(values['job_id'],)).fetchone():fail('Vị trí công việc không tồn tại.')
        return values
    def add_event(c,p,effective,kind,snapshot,decision=None):
        previous=latest(c,p['id'])
        if previous and effective<previous['effective']:abort(409,'Đã có biến động sau ngày này. Chọn ngày từ '+previous['effective']+' trở đi để bảo toàn lịch sử.')
        c.execute('INSERT INTO hr_events(person,effective,kind,snapshot,decision,actor,stamp) VALUES(?,?,?,?,?,?,?)',(p['id'],effective,kind,json.dumps(snapshot,ensure_ascii=False),decision,session['user'],datetime.now().isoformat()))
        c.execute('UPDATE hr_people SET version=version+1 WHERE id=?',(p['id'],))
        audit(c,'hr_'+kind,dict(person=p['id'],effective=effective,decision=decision))
    @app.get('/api/hr/options')
    @auth()
    def options():return jsonify(groups=GROUPS,education=EDUCATION,statuses=STATUSES,today=today().isoformat())
    @app.get('/api/hr/jobs')
    @auth()
    def jobs():
        asof=read_date(request.args.get('asof',today().isoformat()))
        with db() as c:
            out=[]
            for r in c.execute('SELECT * FROM hr_jobs ORDER BY unit,title').fetchall():
                j=job(c,r['id'],asof);j['rules']=[dict(x) for x in c.execute('SELECT * FROM hr_job_rules WHERE job=? ORDER BY effective,id',(r['id'],))];out.append(j)
        return jsonify(out)
    @app.post('/api/hr/jobs')
    @auth(('admin','editor'))
    def create_job():
        d=request.get_json() or {};title=string(d,'title',True);unit=string(d,'unit',True)
        with lock,db() as c:
            if c.execute('SELECT 1 FROM hr_jobs WHERE title=? AND unit=?',(title,unit)).fetchone():abort(409,'Vị trí ở đơn vị này đã tồn tại.')
            jid=secrets.token_hex(16);c.execute('INSERT INTO hr_jobs(id,title,unit) VALUES(?,?,?)',(jid,title,unit));audit(c,'hr_job_created',dict(job=jid,title=title,unit=unit))
        return jsonify(id=jid)
    @app.post('/api/hr/jobs/<jid>/rules')
    @auth(('admin','editor'))
    def rule(jid):
        d=request.get_json() or {};effective=read_date(d.get('effective'));group=d.get('labor_group');technical=d.get('technical');grade=d.get('grade')
        if group not in GROUPS or type(technical) is not bool or (grade is not None and (type(grade) is not int or not 1<=grade<=7)):fail('Nhóm lao động, CNKT hoặc bậc thợ không hợp lệ.')
        if not technical and grade is not None:fail('Chỉ gán bậc thợ cho vị trí công nhân kỹ thuật.')
        with lock,db() as c:
            j=c.execute('SELECT * FROM hr_jobs WHERE id=?',(jid,)).fetchone()
            if not j:abort(404)
            if d.get('version')!=j['version']:abort(409,'Danh mục vừa thay đổi. Tải lại trước khi lưu.')
            last=c.execute('SELECT effective FROM hr_job_rules WHERE job=? ORDER BY effective DESC LIMIT 1',(jid,)).fetchone()
            if last and effective<=last['effective']:abort(409,'Quy tắc mới phải có ngày sau '+last['effective']+'. Quy tắc cũ được giữ để tính lịch sử.')
            c.execute('INSERT INTO hr_job_rules(job,effective,labor_group,technical,grade) VALUES(?,?,?,?,?)',(jid,effective,group,int(technical),grade))
            c.execute('UPDATE hr_jobs SET version=version+1 WHERE id=?',(jid,));audit(c,'hr_rule_created',dict(job=jid,effective=effective,group=group,technical=technical,grade=grade))
        return jsonify(ok=True)
    def people_filter(rows):
        status=request.args.get('status','all')
        if status not in ('all','',*STATUSES):fail('Trạng thái lọc không hợp lệ.')
        q=request.args.get('q','').strip().casefold()
        return [p for p in rows if (status in ('','all') or p['status']==status)
                and (not request.args.get('unit') or p['job']['unit']==request.args['unit'])
                and (not request.args.get('group') or p['job']['labor_group']==request.args['group'])
                and (not request.args.get('education') or p['education']==request.args['education'])
                and (not request.args.get('age_group') or p['age_group']==request.args['age_group'])
                and (not q or q in (p['code']+' '+p['name']+' '+p['job']['unit']+' '+p['job']['title']+' '+((p['source_evidence'] or {}).get('source_code') or '')).casefold())]
    @app.get('/api/hr/people')
    @auth()
    def list_people():
        asof=read_date(request.args.get('asof',today().isoformat()));scheme=request.args.get('scheme','standard')
        with db() as c:
            out=[]
            for p in c.execute('SELECT * FROM hr_people ORDER BY code').fetchall():
                event=latest(c,p['id'],asof)
                if event:out.append(classified(c,p,event,asof,scheme))
        return jsonify(people_filter(out))
    @app.get('/api/hr/people/<pid>/history')
    @auth()
    def hr_history(pid):
        with db() as c:
            p=person(c,pid);return jsonify([dict(r) for r in c.execute('SELECT * FROM hr_events WHERE person=? ORDER BY effective,id',(pid,))])
    @app.post('/api/hr/people')
    @auth(('admin','editor'))
    def create_person():
        d=request.get_json() or {};code=string(d,'code',True);effective=read_date(d.get('effective'))
        with lock,db() as c:
            if c.execute('SELECT 1 FROM hr_people WHERE code=?',(code,)).fetchone():abort(409,'Mã nhân sự đã tồn tại.')
            data=validate_snapshot(c,d,effective)
            if data['status']=='active' and (not data['birth_date'] or not data['hire_date'] or not data['job_id']):fail('Hồ sơ đang làm cần ngày sinh, ngày vào làm và vị trí; nếu thiếu hãy chọn Cần rà soát.')
            pid=secrets.token_hex(16);c.execute('INSERT INTO hr_people(id,code) VALUES(?,?)',(pid,code));p=person(c,pid);add_event(c,p,effective,'hire',data)
        return jsonify(id=pid)
    @app.post('/api/hr/people/<pid>')
    @auth(('admin','editor'))
    def update_person(pid):
        d=request.get_json() or {};effective=read_date(d.get('effective'))
        with lock,db() as c:
            p=person(c,pid)
            if d.get('version')!=p['version']:abort(409,'Hồ sơ đã thay đổi, tải lại trước khi sửa.')
            data=validate_snapshot(c,d,effective);add_event(c,p,effective,'correction',data)
        return jsonify(ok=True)
    @app.post('/api/hr/people/batch')
    @auth(('admin','editor'))
    def batch_update_people():
        d=request.get_json() or {};effective=read_date(d.get('effective'));selected=d.get('selected');changes=d.get('changes')
        if not isinstance(selected,list) or not 1<=len(selected)<=1000 or not isinstance(changes,dict) or not changes or set(changes)-{'status','education','job_id'}:fail('Chọn hồ sơ và ít nhất một trường cần cập nhật.')
        if any(not isinstance(x,dict) or not isinstance(x.get('id'),str) for x in selected):fail('Danh sách hồ sơ không hợp lệ.')
        if len({x['id'] for x in selected})!=len(selected):fail('Danh sách có hồ sơ lặp.')
        with lock,db() as c:
            for item in selected:
                p=person(c,item['id'])
                if p['version']!=item.get('version'):abort(409,'Một hồ sơ đã thay đổi; chưa cập nhật hồ sơ nào. Tải lại và chọn lại.')
                event=latest(c,p['id']);snapshot=json.loads(event['snapshot']);snapshot.update(changes)
                data=validate_snapshot(c,snapshot,effective);add_event(c,p,effective,'correction',data)
            audit(c,'hr_batch',dict(count=len(selected),effective=effective,fields=list(changes)))
        return jsonify(updated=len(selected))
    def import_candidates(bid):
        _,(w,v)=book(bid)
        if 'DU LIEU CO SO' not in w.sheetnames:fail('Tệp cần trang DU LIEU CO SO.')
        s=w['DU LIEU CO SO'];saved=v[s.title];p=patches(bid,s.title);out=[]
        with db() as c:codes={x['code'] for x in c.execute('SELECT code FROM hr_people')};sources={(x['source_book'],x['source_row']) for x in c.execute('SELECT source_book,source_row FROM hr_people')}
        for r in range(4,s.max_row+1):
            def val(col):
                cell=s.cell(r,col);return p.get((s.title,cell.coordinate),saved.cell(r,col).value if cell.data_type=='f' else cell.value)
            def text(col):
                raw=val(col);return str(raw).strip() if raw is not None else ''
            name=text(6)
            if not name or not is_person_row(s,saved,p,r):continue
            code=text(5);birth=parse_excel_date(val(7));hire=parse_excel_date(val(13));edu=education_at(s,saved,p,r)
            status_info=source_status(s,saved,p,r)
            problem=[]
            storage_code=code or f'LUU-TRU-{bid[:8]}-{r}'
            if storage_code in codes:problem.append('Mã đã có trong hồ sơ')
            if not any(ch.isalpha() for ch in name):problem.append('Tên không hợp lệ')
            if (bid,r) in sources:problem.append('Dòng đã nhập')
            out.append(dict(row=r,code=code,storage_code=storage_code,name=name,birth_date=birth,hire_date=hire,title=text(2),unit=text(206) or text(14),education=edu,gender='',**status_info,problems=problem,eligible=not problem))
        counts=Counter(item['code'] for item in out if item['code'])
        for item in out:
            if item['code'] and counts[item['code']]>1:
                item['storage_code']=f'LUU-TRU-{bid[:8]}-{item["row"]}'
                item['status_reason']+=' Trùng mã nhân sự nguồn: giữ riêng từng dòng bằng mã lưu trữ nội bộ, cần rà soát trùng người.'
                if item['storage_code'] in codes and 'Mã đã có trong hồ sơ' not in item['problems']:item['problems'].append('Mã lưu trữ đã có trong hồ sơ')
                item['eligible']=not item['problems']
        return out
    def record_evidence(c,pid,bid,item,revision,effective):
        c.execute('INSERT OR REPLACE INTO hr_source_evidence(person,book,source_row,source_revision,end_column,end_code,status,observed_date,source_code) VALUES(?,?,?,?,?,?,?,?,?)',(pid,bid,item['row'],revision,item['end_column'],item['end_code'],item['status'],effective,item.get('source_code',item.get('code',''))))
    def status_sync_candidates(bid):
        items={item['row']:item for item in import_candidates(bid)};out=[]
        with db() as c:
            for p in c.execute('SELECT * FROM hr_people WHERE source_book=? ORDER BY code',(bid,)).fetchall():
                item=items.get(p['source_row'])
                if not item or item['storage_code']!=p['code'] or item['status']=='review':continue
                e=latest(c,p['id']);data=json.loads(e['snapshot'])
                if data['status']!=item['status']:
                    out.append(dict(id=p['id'],version=p['version'],code=p['code'],source_code=item['code'],name=data['name'],row=p['source_row'],current_status=data['status'],proposed_status=item['status'],end_column=item['end_column'],end_code=item['end_code'],status_reason=item['status_reason'],last_effective=e['effective']))
        return out
    @app.get('/api/hr/import/<bid>/sync-status')
    @auth(('admin','editor'))
    def preview_status_sync(bid):
        with lock:return jsonify(status_sync_candidates(bid))
    @app.post('/api/hr/import/<bid>/sync-status')
    @auth(('admin','editor'))
    def sync_source_status(bid):
        d=request.get_json() or {};effective=read_date(d.get('effective'));selected=d.get('selected')
        if not isinstance(selected,list) or not 1<=len(selected)<=1000 or any(not isinstance(x,dict) or not isinstance(x.get('id'),str) for x in selected):fail('Chọn danh sách hồ sơ đã xem trước.')
        if len({x['id'] for x in selected})!=len(selected):fail('Danh sách lặp hồ sơ.')
        with lock:
            b,_=book(bid)
            if d.get('revision')!=b['revision']:abort(409,'Tệp nguồn vừa thay đổi. Xem trước lại trạng thái.')
            candidates={x['id']:x for x in status_sync_candidates(bid)}
            with db() as c:
                for item in selected:
                    change=candidates.get(item['id'])
                    if not change or item.get('version')!=change['version']:abort(409,'Một hồ sơ đã thay đổi; chưa ghi nhận hồ sơ nào. Xem trước lại.')
                    p=person(c,item['id']);e=latest(c,p['id']);data=json.loads(e['snapshot']);data['status']=change['proposed_status']
                    add_event(c,p,effective,'source_status',data)
                    record_evidence(c,p['id'],bid,dict(row=change['row'],source_code=change['source_code'],end_column=change['end_column'],end_code=change['end_code'],status=change['proposed_status']),b['revision'],effective)
                audit(c,'hr_status_sync',dict(book=bid,count=len(selected),effective=effective))
        return jsonify(updated=len(selected))
    @app.get('/api/hr/import/<bid>')
    @auth(('admin','editor'))
    def preview_import(bid):
        with lock:return jsonify(import_candidates(bid))
    @app.post('/api/hr/import/<bid>')
    @auth(('admin','editor'))
    def import_people(bid):
        d=request.get_json() or {};effective=read_date(d.get('effective'))
        with lock:
            b,_=book(bid)
            if d.get('revision')!=b['revision']:abort(409,'Tệp nguồn vừa thay đổi; xem lại trước khi nhập.')
            candidates=import_candidates(bid);count=0;skipped=0
            with db() as c:
                for item in candidates:
                    if not item['eligible']:skipped+=1;continue
                    jid=''
                    if item['title'] and item['unit']:
                        j=c.execute('SELECT id FROM hr_jobs WHERE title=? AND unit=?',(item['title'],item['unit'])).fetchone()
                        if j:jid=j['id']
                        else:
                            jid=secrets.token_hex(16);c.execute('INSERT INTO hr_jobs(id,title,unit) VALUES(?,?,?)',(jid,item['title'],item['unit']))
                    data={k:item.get(k,'') for k in FIELDS};data['job_id']=jid
                    if data['birth_date'] and data['birth_date']>effective:data['birth_date']=''
                    if data['hire_date'] and data['hire_date']>effective:data['hire_date']=''
                    pid=secrets.token_hex(16);c.execute('INSERT INTO hr_people(id,code,source_book,source_row) VALUES(?,?,?,?)',(pid,item['storage_code'],bid,item['row']))
                    add_event(c,person(c,pid),effective,'baseline',data)
                    record_evidence(c,pid,bid,item,b['revision'],effective);count+=1
                audit(c,'hr_import',dict(book=bid,count=count,skipped=skipped,baseline=effective))
        return jsonify(imported=count,skipped=skipped)
    def decision_data(c,d):
        kind=d.get('kind')
        if kind not in KINDS:fail('Loại quyết định không hợp lệ.')
        effective=read_date(d.get('effective'));p=person(c,d.get('person'));e=latest(c,p['id'])
        if d.get('version')!=p['version']:abort(409,'Hồ sơ vừa thay đổi. Tải lại và lập lại quyết định.')
        if effective<e['effective']:abort(409,'Ngày quyết định trước biến động đã lưu gần nhất.')
        before=json.loads(e['snapshot'])
        if before['status']!='active':fail('Chỉ áp dụng quyết định cho hồ sơ đã xác nhận đang làm việc.')
        after=before.copy();old_job=job(c,before['job_id'],effective)
        if kind=='termination':after['status']='terminated'
        else:
            jid=string(d,'new_job',True)
            if not c.execute('SELECT 1 FROM hr_jobs WHERE id=?',(jid,)).fetchone():fail('Chọn vị trí mới hợp lệ.')
            after['job_id']=jid
        new_job=job(c,after['job_id'],effective)
        fields=dict(ho_ten=before['name'],ma_nv=p['code'],chuc_danh_cu=old_job['title'],don_vi_cu=old_job['unit'],chuc_danh_moi=new_job['title'] if kind!='termination' else '',don_vi_moi=new_job['unit'] if kind!='termination' else '',so_quyet_dinh=string(d,'number',True),ngay_hieu_luc=date.fromisoformat(effective).strftime('%d/%m/%Y'),can_cu=string(d,'reason',True),nguoi_ky=string(d,'signer'),loai_quyet_dinh=KINDS[kind])
        return p,dict(kind=kind,effective=effective,before=before,after=after,fields=fields)
    @app.post('/api/hr/decisions')
    @auth(('admin','editor'))
    def draft_decision():
        d=request.get_json() or {}
        with lock,db() as c:
            p,payload=decision_data(c,d);did=secrets.token_hex(16)
            c.execute('INSERT INTO hr_decisions(id,person,version,payload,stamp,actor) VALUES(?,?,?,?,?,?)',(did,p['id'],p['version'],json.dumps(payload,ensure_ascii=False),datetime.now().isoformat(),session['user']))
            audit(c,'hr_decision_draft',dict(decision=did,person=p['id']))
        return jsonify(id=did,effective=payload['effective'])
    @app.get('/api/hr/decisions')
    @auth()
    def decisions():
        with db() as c:
            out=[]
            for r in c.execute('SELECT * FROM hr_decisions ORDER BY stamp DESC LIMIT 200'):
                d=json.loads(r['payload']);out.append(dict(id=r['id'],person=r['person'],applied=bool(r['applied']),kind=d['kind'],effective=d['effective'],fields=d['fields']))
        return jsonify(out)
    @app.get('/api/hr/decisions/<did>/word')
    @auth()
    def decision_word(did):
        with lock,db() as c:
            r=c.execute('SELECT * FROM hr_decisions WHERE id=?',(did,)).fetchone()
            if not r:abort(404)
            d=json.loads(r['payload']);out=make_document(root,d['kind'],d['fields'])
        return send_file(out,as_attachment=True,download_name='Quyet_dinh_du_thao.docx',mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    @app.post('/api/hr/decisions/<did>/apply')
    @auth(('admin','editor'))
    def apply_decision(did):
        with lock,db() as c:
            r=c.execute('SELECT * FROM hr_decisions WHERE id=?',(did,)).fetchone()
            if not r:abort(404)
            if r['applied']:return jsonify(ok=True,already_applied=True)
            p=person(c,r['person'])
            if p['version']!=r['version']:abort(409,'Hồ sơ đã thay đổi từ khi lập dự thảo. Hãy lập lại quyết định để không ghi đè thông tin mới.')
            d=json.loads(r['payload']);add_event(c,p,d['effective'],d['kind'],d['after'],did)
            c.execute('UPDATE hr_decisions SET applied=1 WHERE id=?',(did,))
        return jsonify(ok=True,effective=d['effective'])
    def build_report(start,end,scheme):
        opening=(date.fromisoformat(start)-timedelta(days=1)).isoformat()
        with db() as c:
            people=c.execute('SELECT * FROM hr_people ORDER BY code').fetchall();rows=[];units=defaultdict(lambda:dict(opening=0,increase=0,decrease=0,closing=0));baseline_dates=[];moves=[]
            for p in people:
                first=c.execute('SELECT effective,kind FROM hr_events WHERE person=? ORDER BY effective,id LIMIT 1',(p['id'],)).fetchone()
                if first and first['kind']=='baseline':baseline_dates.append(first['effective'])
                before=latest(c,p['id'],opening);last=latest(c,p['id'],end)
                if before:
                    b=classified(c,p,before,opening,scheme)
                    if b['status']=='active':units[b['job']['unit']]['opening']+=1
                if last:
                    row=classified(c,p,last,end,scheme);rows.append(row)
                    if row['status']=='active':units[row['job']['unit']]['closing']+=1
                for e in c.execute('SELECT * FROM hr_events WHERE person=? AND effective>=? AND effective<=? ORDER BY effective,id',(p['id'],start,end)).fetchall():
                    prev=c.execute('SELECT * FROM hr_events WHERE person=? AND (effective<? OR (effective=? AND id<?)) ORDER BY effective DESC,id DESC LIMIT 1',(p['id'],e['effective'],e['effective'],e['id'])).fetchone()
                    old=classified(c,p,prev,e['effective'],scheme) if prev else None;new=classified(c,p,e,e['effective'],scheme)
                    old_unit=old['job']['unit'] if old and old['status']=='active' else None;new_unit=new['job']['unit'] if new['status']=='active' else None
                    if old_unit!=new_unit:
                        if old_unit is not None:units[old_unit]['decrease']+=1
                        if new_unit is not None:units[new_unit]['increase']+=1
                    moves.append(dict(code=p['code'],name=new['name'],effective=e['effective'],kind=e['kind'],old_unit=old_unit,new_unit=new_unit))
            active=[r for r in rows if r['status']=='active'];review=[r for r in rows if r['status']=='review']
            def counts(key):return dict(sorted(Counter(key(r) for r in active).items()))
            warnings=[]
            if baseline_dates and start<=max(baseline_dates):warnings.append('Kỳ báo cáo chạm hoặc trước ngày chốt nhập Excel. Không có đủ lịch sử trước ngày chốt; cần đối chiếu số đầu kỳ và tăng/giảm.')
            if review:warnings.append(str(len(review))+' hồ sơ cần rà soát chưa đưa vào tổng đang làm.')
            if any(r['warnings'] for r in active):warnings.append('Một số hồ sơ đang làm còn thiếu thông tin phân nhóm; xem cột Cần rà soát trong danh sách.')
            first_effective=c.execute('SELECT MIN(effective) FROM hr_events').fetchone()[0]
            source_count=c.execute('SELECT COUNT(*) FROM books').fetchone()[0]
            if not people:
                state='no_people';message='Chưa có hồ sơ quản lý để tổng hợp. '+('Đã có tệp Excel nguồn, nhưng chưa nhập thành hồ sơ nhân sự.' if source_count else 'Hãy nhập Excel nguồn rồi chuyển sang hồ sơ nhân sự, hoặc thêm hồ sơ mới.')
            elif not rows:
                state='before_records';message='Ngày cuối kỳ đang trước ngày hồ sơ bắt đầu có hiệu lực ('+str(first_effective)+'). Không có đủ dữ liệu để tính kỳ này.'
            elif not active and review:
                state='needs_review';message='Có '+str(len(review))+' hồ sơ chờ xác nhận tại ngày cuối kỳ; chưa hồ sơ nào được xác nhận đang làm. Báo cáo đang làm bằng 0, không phải số liệu nhân sự đã hoàn tất.'
            elif not active:
                state='no_active';message='Không có hồ sơ ở trạng thái đang làm tại ngày cuối kỳ đã chọn. Kiểm tra ngày hiệu lực và trạng thái trong hồ sơ nếu số 0 chưa đúng thực tế.'
            else:
                state='ready';message='Đã tổng hợp '+str(len(active))+' hồ sơ đang làm tại ngày '+end+'.'
            readiness=dict(state=state,message=message,total_people=len(people),effective_people=len(rows),review_count=len(review),first_effective=first_effective,source_books=source_count)
            return dict(start=start,end=end,scheme=scheme,readiness=readiness,total=len(active),female=sum(r['gender']=='female' for r in active),review_count=len(review),technical=sum(r['job']['technical'] for r in active),groups=counts(lambda r:GROUPS[r['job']['labor_group']]),education=counts(lambda r:EDUCATION[r['education']]),ages=counts(lambda r:r['age_group']),grades=dict(sorted(Counter(str(r['job']['grade'] or 'Chưa xác định') for r in active if r['job']['technical']).items())),units=[dict(unit=k,**v) for k,v in sorted(units.items())],people=rows,movements=moves,warnings=warnings)
    def report_args():
        start=read_date(request.args.get('start',today().replace(day=1).isoformat()));end=read_date(request.args.get('end',today().isoformat()));scheme=request.args.get('scheme','standard')
        if start>end or start<'1900-01-02' or scheme not in ('standard','primary'):fail('Kỳ báo cáo hoặc cách chia nhóm tuổi không hợp lệ.')
        return start,end,scheme
    @app.get('/api/hr/reports')
    @auth()
    def reports():
        with lock:return jsonify(build_report(*report_args()))
    @app.get('/api/hr/reports/export')
    @auth()
    def report_export():
        with lock:d=build_report(*report_args())
        w=Workbook();s=w.active;s.title='Tổng hợp nội bộ'
        s.append(['BÁO CÁO TỔNG HỢP NỘI BỘ — BẢN THỬ 0.3.2']);s.append(['Từ',d['start'],'Đến',d['end']]);s.append(['Đang làm',d['total'],'Nữ',d['female'],'CNKT',d['technical'],'Cần rà soát',d['review_count']])
        s.append(['Tình trạng dữ liệu',d['readiness']['message']])
        for warning in d['warnings']:s.append(['Lưu ý',warning])
        for key,title in [('groups','Nhóm lao động'),('education','Trình độ'),('ages','Nhóm tuổi'),('grades','Bậc CNKT')]:
            s.append([]);s.append([title,'Số người'])
            for name,count in d[key].items():s.append([name,count])
        s=w.create_sheet('Biến động đơn vị');s.append(['Đơn vị','Đầu kỳ','Tăng / chuyển vào / xác nhận','Giảm / chuyển ra','Cuối kỳ'])
        for u in d['units']:s.append([u[k] for k in ('unit','opening','increase','decrease','closing')])
        s=w.create_sheet('Danh sách chốt kỳ');s.append(['Mã / mã lưu trữ nội bộ','Họ tên','Ngày sinh','Tuổi','Nhóm tuổi','Trình độ','Chức danh','Đơn vị','Nhóm lao động','CNKT','Bậc','Trạng thái','Cần rà soát'])
        for p in d['people']:s.append([p['code'],p['name'],p['birth_date'],p['age'],p['age_group'],EDUCATION[p['education']],p['job']['title'],p['job']['unit'],GROUPS[p['job']['labor_group']],p['job']['technical'],p['job']['grade'],STATUSES[p['status']],'; '.join(p['warnings'])])
        s=w.create_sheet('Lịch sử trong kỳ');s.append(['Mã','Họ tên','Ngày hiệu lực','Loại biến động','Đơn vị trước','Đơn vị sau'])
        for m in d['movements']:s.append([m[k] for k in ('code','name','effective','kind','old_unit','new_unit')])
        for s in w:
            s.freeze_panes='B2';s.auto_filter.ref=s.dimensions
            for row in s:
                for cell in row:
                    # User-supplied values are always literal text in generated reports.
                    if cell.data_type=='f':cell.data_type='s'
            for cell in s[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='126C56')
            for col in s.columns:s.column_dimensions[col[0].column_letter].width=min(50,max(16,max(len(str(c.value or '')) for c in col)+2))
        out=io.BytesIO();w.save(out);out.seek(0)
        return send_file(out,as_attachment=True,download_name='Tong_hop_nhan_su_'+d['end']+'.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
