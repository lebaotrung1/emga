import unittest
import test_source_status as source

class BaselineDate(unittest.TestCase):
    setUp=source.SourceStatus.setUp
    tearDown=source.SourceStatus.tearDown
    upload=source.SourceStatus.upload
    post=source.SourceStatus.post
    def imported(self,rows=None):
        bid=self.upload(rows or [dict(code='ACTIVE',end='PX1CÔ'),dict(code='LEFT',end=None)])
        self.post('/api/hr/import/'+bid,dict(revision=0,effective='2026-10-28'))
        return bid
    def preview(self,bid,target='2026-09-30'):
        r=self.client.get('/api/hr/import/'+bid+'/baseline',query_string=dict(effective=target));self.assertEqual(r.status_code,200);return r.json
    def payload(self,p):return dict(effective=p['effective'],revision=p['revision'],reason='Sửa nhầm ngày nhập; đã đối chiếu ngày chốt tệp',confirmed=True,selected=[{k:r[k] for k in ('id','version','event')} for r in p['rows'] if r['eligible']])
    def test_correct_import_date_makes_prior_period_report_available(self):
        bid=self.imported();before=self.client.get('/api/hr/people?asof=2026-10-28').json
        report=self.client.get('/api/hr/reports?start=2026-01-01&end=2026-09-30').json
        self.assertEqual(report['total'],0);self.assertEqual(report['readiness']['state'],'before_records')
        p=self.preview(bid);self.assertEqual(p['eligible'],2);self.assertEqual(p['blocked'],0)
        r=self.post('/api/hr/import/'+bid+'/baseline',self.payload(p));self.assertEqual(r['updated'],2)
        report=self.client.get('/api/hr/reports?start=2026-01-01&end=2026-09-30').json
        self.assertEqual(report['total'],1);self.assertEqual(len(report['people']),2);self.assertTrue(report['warnings'])
        for person in before:
            history=self.client.get('/api/hr/people/'+person['id']+'/history').json
            self.assertEqual(len(history),1);self.assertEqual(history[0]['effective'],'2026-09-30')
            change=history[0]['baseline_corrections'][0]
            self.assertEqual(change['old_effective'],'2026-10-28');self.assertEqual(change['new_effective'],'2026-09-30')
        after=self.client.get('/api/hr/people?asof=2026-09-30').json
        self.assertEqual({(p['id'],p['code'],p['name'],p['status']) for p in before},{(p['id'],p['code'],p['name'],p['status']) for p in after})
        self.assertTrue(all(p['source_evidence']['observed_date']=='2026-09-30' for p in after))
        self.assertEqual(self.preview(bid)['unchanged'],2)
    def test_modified_person_is_blocked_and_batch_rolls_back(self):
        bid=self.imported();p=self.preview(bid);payload=self.payload(p)
        person=self.client.get('/api/hr/people?asof=2026-10-28').json[0];person.update(effective='2026-10-29',education='postgraduate')
        self.post('/api/hr/people/'+person['id'],person)
        # Update the untouched row first, then encounter the stale row; nothing may commit.
        payload['selected'].sort(key=lambda r:r['id']==person['id'])
        r=self.client.post('/api/hr/import/'+bid+'/baseline',json=payload,headers=self.h);self.assertEqual(r.status_code,409)
        self.assertEqual(self.client.get('/api/hr/people?asof=2026-09-30').json,[])
        current=self.preview(bid);self.assertEqual(current['eligible'],1);self.assertEqual(current['blocked'],1)
    def test_decision_draft_protects_initial_date(self):
        bid=self.imported([dict(code='ACTIVE',end='PX1CÔ')]);person=self.client.get('/api/hr/people?asof=2026-10-28').json[0]
        self.post('/api/hr/decisions',dict(person=person['id'],version=person['version'],kind='title',new_job=person['job_id'],effective='2026-10-30',number='TEST-QD',reason='Bản dự thảo để kiểm tra bảo vệ lịch sử'))
        p=self.preview(bid);self.assertEqual(p['eligible'],0);self.assertEqual(p['blocked'],1)
        self.assertIn('quyết định',' '.join(p['rows'][0]['blocked_reasons']))
    def test_confirmation_and_reason_required(self):
        bid=self.imported();p=self.preview(bid);d=self.payload(p);d['confirmed']=False
        self.assertEqual(self.client.post('/api/hr/import/'+bid+'/baseline',json=d,headers=self.h).status_code,400)
        d['confirmed']=True;d['reason']='';self.assertEqual(self.client.post('/api/hr/import/'+bid+'/baseline',json=d,headers=self.h).status_code,400)
        self.assertEqual(self.client.get('/api/hr/people?asof=2026-09-30').json,[])
    def test_dates_before_birth_or_hire_are_rejected(self):
        bid=self.imported();p=self.preview(bid,'2019-12-31');self.assertEqual(p['eligible'],0);self.assertEqual(p['blocked'],2)
        self.assertTrue(all(any('Ngày vào làm' in s for s in r['blocked_reasons']) for r in p['rows']))
    def test_original_snapshot_fields_and_audit_are_preserved(self):
        bid=self.imported();people=self.client.get('/api/hr/people?asof=2026-10-28').json
        snapshots={p['id']:self.client.get('/api/hr/people/'+p['id']+'/history').json[0]['snapshot'] for p in people}
        self.post('/api/hr/import/'+bid+'/baseline',self.payload(self.preview(bid)))
        for p in people:self.assertEqual(self.client.get('/api/hr/people/'+p['id']+'/history').json[0]['snapshot'],snapshots[p['id']])
        self.assertTrue(any(a['action']=='hr_baseline_date_corrected' for a in self.client.get('/api/audit').json))
    def test_source_revision_change_requires_new_preview(self):
        bid=self.imported();p=self.preview(bid)
        self.post('/api/book/'+bid+'/cell',dict(revision=0,sheet='DU LIEU CO SO',cell='F4',type='text',value='Tên điều chỉnh kiểm thử'))
        r=self.client.post('/api/hr/import/'+bid+'/baseline',json=self.payload(p),headers=self.h);self.assertEqual(r.status_code,409)
        self.assertEqual(self.client.get('/api/hr/people?asof=2026-09-30').json,[])

if __name__=='__main__':unittest.main()
