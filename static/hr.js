'use strict';
let hrOptions=null,hrJobs=[],hrPeople=[],hrCandidates=[],hrImportSource=null,hrReportQuery='';
const hrEdit=window.EMGA_ROLE!=='viewer';
let hrSelected=new Set(),hrSyncCandidates=[];
function hrFilteredPeople(){const q=$('#hr-search').value.toLocaleLowerCase('vi');const state=$('#hr-filter-status').value,unit=$('#hr-filter-unit').value,group=$('#hr-filter-group').value,edu=$('#hr-filter-education').value,age=$('#hr-filter-age').value;return hrPeople.filter(p=>(p.code+' '+p.name+' '+p.job.unit+' '+p.job.title+' '+(p.source_evidence?.source_code||'')).toLocaleLowerCase('vi').includes(q)&&(state==='all'||p.status===state)&&(!unit||p.job.unit===unit)&&(!group||p.job.labor_group===group)&&(!edu||p.education===edu)&&(!age||p.age_group===age));}
function hrSelectedCount(){if($('#hr-selected-count'))$('#hr-selected-count').textContent=hrSelected.size+' hồ sơ được chọn';}
function hrTable(target,headers,rows){const t=document.createElement('table'),head=t.createTHead().insertRow();for(const label of headers){const th=document.createElement('th');th.textContent=label;head.append(th);}const body=t.createTBody();for(const values of rows){const tr=body.insertRow();for(const value of values){const td=tr.insertCell();if(value instanceof Node)td.append(value);else td.textContent=value??'';}}$(target).replaceChildren(t);}
function hrButton(label,fn){const b=document.createElement('button');b.type='button';b.className='secondary';b.textContent=label;b.onclick=handle(fn);return b;}
function hrSelect(el,values,empty){const previous=el.value;el.replaceChildren(...(empty?[option('',empty)]:[]),...values.map(([v,t])=>option(v,t)));if([...el.options].some(o=>o.value===previous))el.value=previous;}
async function hrLoadJobs(){hrJobs=await(await api('/api/hr/jobs?asof='+$('#hr-asof').value)).json();const values=hrJobs.map(j=>[j.id,j.unit+' · '+j.title+' · '+hrOptions.groups[j.labor_group]]);for(const id of ['hr-person-job','hr-rule-job','hr-decision-job',...(hrEdit?['hr-batch-job']:[])])hrSelect($('#'+id),values,id==='hr-person-job'?'Chưa xác định':id==='hr-batch-job'?'Giữ nguyên':null);hrRenderJobs();hrRuleSelected();}
function hrRenderJobs(){hrTable('#hr-jobs-grid',['Đơn vị','Vị trí','Nhóm lao động','CNKT','Bậc',''],hrJobs.map(j=>[j.unit,j.title,hrOptions.groups[j.labor_group],j.technical?'Có':'Không',j.grade??'',hrButton('Xem quy tắc',()=>{$('#hr-rule-job').value=j.id;hrRuleSelected();})]));}
function hrRuleSelected(){const j=hrJobs.find(j=>j.id===$('#hr-rule-job').value);const f=$('#hr-rule-form');if(!j){$('#hr-rules-history').textContent='Chưa có vị trí; thêm vị trí trước.';return;}f.elements.labor_group.value=j.labor_group;f.elements.technical.value=String(j.technical);f.elements.grade.value=j.grade??'';$('#hr-rules-history').replaceChildren(...j.rules.map(r=>{const p=document.createElement('p');p.textContent=r.effective+' · '+hrOptions.groups[r.labor_group]+' · '+(r.technical?'CNKT':'Không CNKT')+(r.grade?' · Bậc '+r.grade:'');return p;}));}
async function hrLoadPeople(){hrSelected.clear();hrSelectedCount();hrPeople=await(await api('/api/hr/people?asof='+$('#hr-asof').value)).json();hrSelect($('#hr-filter-unit'),[...new Set(hrPeople.map(p=>p.job.unit))].sort().map(u=>[u,u]),'Tất cả');hrSelect($('#hr-filter-age'),[...new Set(hrPeople.map(p=>p.age_group))].sort().map(u=>[u,u]),'Tất cả');hrRenderPeople();hrSelect($('#hr-decision-person'),hrPeople.filter(p=>p.status==='active').map(p=>[p.id,p.code+' · '+p.name+' · '+p.job.unit]));}
function hrRenderPeople(){const list=hrFilteredPeople();$('#hr-count').textContent=list.length+' hồ sơ tại '+$('#hr-asof').value+' · '+hrPeople.filter(p=>p.status==='review').length+' cần rà soát';hrTable('#hr-people-grid',['Chọn','Mã','Họ tên','Tuổi / nhóm tuổi','Trình độ','Chức danh / Đơn vị','Nhóm lao động','Trạng thái / lưu ý',''],list.map(p=>{const actions=document.createElement('div');actions.append(hrButton(hrEdit?'Sửa hồ sơ':'Xem hồ sơ',()=>hrOpenPerson(p)),hrButton('Lịch sử',()=>hrHistory(p.id)));const check=document.createElement('input');check.type='checkbox';check.disabled=!hrEdit;check.checked=hrSelected.has(p.id);check.setAttribute('aria-label','Chọn '+p.code);check.onchange=()=>{if(check.checked)hrSelected.add(p.id);else hrSelected.delete(p.id);hrSelectedCount();};return[check,p.code+(p.code.startsWith('LUU-TRU-')?' (mã lưu trữ nội bộ)':''),p.name,(p.age??'?')+' · '+p.age_group,hrOptions.education[p.education],p.job.title+' / '+p.job.unit,hrOptions.groups[p.job.labor_group]+(p.job.technical?' · CNKT':'')+(p.job.grade?' bậc '+p.job.grade:''),hrOptions.statuses[p.status]+(p.source_evidence?' · '+p.source_evidence.end_column+': '+(p.source_evidence.end_code||'(trống)'):'')+(p.warnings.length?' · '+p.warnings.join('; '):''),actions];}));}
function hrOpenPerson(p=null){const f=$('#hr-person-form');f.reset();f.elements.id.value=p?.id||'';f.elements.version.value=p?.version||'';for(const key of ['code','name','birth_date','hire_date','gender','education','job_id','status'])f.elements[key].value=p?.[key]??({'education':'unknown','status':'review'}[key]||'');f.elements.effective.value=$('#hr-asof').value;f.elements.code.readOnly=!!p;$('#hr-person-heading').textContent=p?'Hồ sơ · '+p.name:'Thêm nhân sự';$('#hr-person-classification').textContent=p?'Nhóm hiện tại: '+p.age_group+' · '+hrOptions.groups[p.job.labor_group]+'; nhóm mới sẽ được tính theo vị trí và ngày hiệu lực.':'Chọn vị trí để tự gán nhóm. Ngày sinh sẽ tự chia nhóm tuổi theo kỳ báo cáo.';$('#hr-person-dialog').showModal();}
async function hrHistory(pid){const events=await(await api('/api/hr/people/'+pid+'/history')).json();$('#hr-history-panel').hidden=false;hrTable('#hr-history',['Ngày hiệu lực','Loại','Người ghi nhận','Nội dung hồ sơ'],events.map(e=>{const p=JSON.parse(e.snapshot);const j=hrJobs.find(j=>j.id===p.job_id);return[e.effective,e.kind,e.actor,p.name+' · '+hrOptions.education[p.education]+' · '+(j?j.unit+' / '+j.title:'Chưa xác định vị trí')+' · '+hrOptions.statuses[p.status]];}));}
$('#hr-person-close').onclick=()=>$('#hr-person-dialog').close();
$('#hr-person-form').onsubmit=handle(async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(e.target));if(d.id)d.version=Number(d.version);await json('/api/hr/people'+(d.id?'/'+d.id:''),d);$('#hr-person-dialog').close();await hrLoadPeople();notice('Đã lưu hồ sơ và lịch sử. Các nhóm và báo cáo tự tổng hợp cập nhật theo ngày hiệu lực.');});
if($('#hr-new'))$('#hr-new').onclick=()=>hrOpenPerson();$('#hr-search').oninput=hrRenderPeople;
$('#hr-refresh').onclick=handle(async()=>{await hrLoadJobs();await hrLoadPeople();});
$('#hr-job-form').onsubmit=handle(async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(e.target));const created=await(await json('/api/hr/jobs',d)).json();await hrLoadJobs();$('#hr-rule-job').value=created.id;hrRuleSelected();notice('Đã tạo vị trí. Khai báo quy tắc phân nhóm cho vị trí ở bên cạnh.');});
$('#hr-rule-job').onchange=hrRuleSelected;
$('#hr-rule-form').onsubmit=handle(async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(e.target));const j=hrJobs.find(j=>j.id===d.job);if(!j)throw Error('Chọn vị trí trước.');d.version=j.version;d.technical=d.technical==='true';d.grade=d.grade?Number(d.grade):null;await json('/api/hr/jobs/'+d.job+'/rules',d);await hrLoadJobs();await hrLoadPeople();notice('Đã lưu quy tắc. Người giữ vị trí này tự nhận nhóm theo ngày hiệu lực.');});
if($('#hr-import-preview'))$('#hr-import-preview').onclick=handle(async()=>{
    if(!bid)throw Error('Vào Excel nguồn và nhập/chọn tệp trước.');
    const bookInfo=await(await api('/api/book/'+bid)).json();hrImportSource={id:bid,revision:bookInfo.revision};
    hrCandidates=await(await api('/api/hr/import/'+bid)).json();hrSyncCandidates=await(await api('/api/hr/import/'+bid+'/sync-status')).json();
    $('#hr-import-box').hidden=false;
    const eligible=hrCandidates.filter(p=>p.eligible);
    $('#hr-import-confirm').disabled=!eligible.length;
    $('#hr-import-summary').textContent=bookInfo.name+': '+eligible.length+' dòng mới có thể nhập; '+hrCandidates.filter(p=>!p.eligible).length+' dòng đã có hoặc không hợp lệ. Trong các dòng mới: '+eligible.filter(p=>p.status==='active').length+' đang làm; '+eligible.filter(p=>p.status==='terminated').length+' đã nghỉ; '+eligible.filter(p=>p.status==='review').length+' cần rà soát theo mã cuối kỳ.';
    hrTable('#hr-import-grid',['Dòng','Mã nguồn','Mã lưu hồ sơ','Họ tên','Ngày sinh','Trình độ dự kiến','Chức danh','Đơn vị nguồn','Mã cuối kỳ','Trạng thái từ mã','Lưu ý'],hrCandidates.map(p=>[p.row,p.code,p.storage_code,p.name,p.birth_date,hrOptions.education[p.education],p.title,p.unit,p.end_column+': '+(p.end_code||'(trống)'),hrOptions.statuses[p.status],p.problems.join('; ')||p.status_reason]));
    $('#hr-sync-box').hidden=!hrSyncCandidates.length;
    $('#hr-sync-summary').textContent=hrSyncCandidates.length+' hồ sơ có trạng thái khác mã cuối kỳ trong tệp. Áp dụng sẽ đổi trạng thái các hồ sơ sau từ ngày ghi nhận đã chọn:';
    hrTable('#hr-sync-grid',['Mã','Họ tên','Trạng thái đang lưu','Trạng thái theo tệp','Mã cuối kỳ','Biến động gần nhất'],hrSyncCandidates.map(p=>[p.code,p.name,hrOptions.statuses[p.current_status],hrOptions.statuses[p.proposed_status],p.end_column+': '+(p.end_code||'(trống)'),p.last_effective]));
});
$('#hr-import-confirm').onclick=handle(async()=>{
    if(!hrImportSource)throw Error('Xem trước tệp cần nhập.');
    const r=await(await json('/api/hr/import/'+hrImportSource.id,{effective:$('#hr-baseline').value,revision:hrImportSource.revision})).json();
    $('#hr-import-box').hidden=true;await hrLoadJobs();await hrLoadPeople();notice('Đã nhập '+r.imported+' hồ sơ theo mã cuối kỳ; bỏ qua '+r.skipped+' dòng đã có hoặc không hợp lệ. Người đã nghỉ vẫn giữ trong danh sách Tất cả.');
});
$('#hr-sync-confirm').onclick=handle(async()=>{
    if(!hrImportSource||!hrSyncCandidates.length)throw Error('Xem trước danh sách cần cập nhật.');
    const effective=$('#hr-baseline').value;
    if(!confirm('Áp dụng trạng thái từ mã cuối kỳ cho '+hrSyncCandidates.length+' hồ sơ từ '+effective+'? Tên và lịch sử vẫn được giữ.'))return;
    const r=await(await json('/api/hr/import/'+hrImportSource.id+'/sync-status',{effective,revision:hrImportSource.revision,selected:hrSyncCandidates.map(p=>({id:p.id,version:p.version}))})).json();
    $('#hr-import-box').hidden=true;await hrLoadPeople();notice('Đã cập nhật trạng thái '+r.updated+' hồ sơ theo mã cuối kỳ, giữ nguyên tên và lịch sử trước '+effective+'.');
});
function hrReportParams(){return new URLSearchParams(new FormData($('#hr-report-form'))).toString();}
let hrReportPending=null;
function hrReportStatus(title,text,error=false){
    $('#hr-report-status-title').textContent=title;
    $('#hr-report-status-text').textContent=text;
    $('#hr-report-status-counts').textContent='';
    $('#hr-report-status-actions').replaceChildren();
    $('#hr-report-status').classList.toggle('warning',error);
}
function hrReportReadiness(d){
    const r=d.readiness;
    hrReportStatus(r.state==='ready'?'Đã tính xong báo cáo':'Báo cáo cần kiểm tra dữ liệu',r.message,r.state!=='ready');
    $('#hr-report-status-counts').textContent='Tổng hồ sơ đã nhập: '+r.total_people+' · Có hiệu lực đến cuối kỳ: '+r.effective_people+' · Đang làm: '+d.total+' · Chờ xác nhận: '+r.review_count;
    const actions=$('#hr-report-status-actions');
    if(r.state==='no_people'){
        actions.append(hrButton(r.source_books?'Mở bước nhập hồ sơ từ Excel':'Mở Excel nguồn',()=>{
            document.querySelector('[data-tab='+(r.source_books?'personnel':'work')+']').click();
            if(r.source_books&&hrEdit)$('#hr-import-preview').click();
        }));
    }else if(r.state==='needs_review'||r.state==='no_active'){
        actions.append(hrButton('Xem hồ sơ tại ngày cuối kỳ',async()=>{
            $('#hr-asof').value=d.end;document.querySelector('[data-tab=personnel]').click();
        }));
    }else if(r.state==='before_records'&&r.first_effective){
        actions.append(hrButton('Xem kỳ từ ngày có dữ liệu',async()=>{
            const form=$('#hr-report-form');form.elements.start.value=r.first_effective;
            if(form.elements.end.value<r.first_effective)form.elements.end.value=r.first_effective;
            await hrReport();
        }));
    }
}
async function hrReport(){
    if(hrReportPending)return hrReportPending;
    const form=$('#hr-report-form');
    if(!form.reportValidity())throw Error('Điền đủ ngày bắt đầu và kết thúc kỳ báo cáo.');
    const buttons=[...form.querySelectorAll('button')];
    buttons.forEach(b=>b.disabled=true);$('#hr-report-result').hidden=true;
    hrReportStatus('Đang tổng hợp…','Đang kiểm tra hồ sơ và tính số liệu của kỳ đã chọn.');
    hrReportPending=(async()=>{
        try{
            await hrReady;
            const query=hrReportParams();const d=await(await api('/api/hr/reports?'+query)).json();hrReportQuery=query;
            $('#hr-report-result').hidden=false;$('#hr-report-total').textContent=d.total;$('#hr-report-female').textContent=d.female;$('#hr-report-technical').textContent=d.technical;
            $('#hr-report-warnings').replaceChildren(...d.warnings.map(w=>{const e=document.createElement('div');e.className='warning';e.textContent=w;return e;}));
            for(const key of ['groups','education','ages','grades'])hrTable('#hr-report-'+key,['Phân nhóm','Số người'],Object.entries(d[key]));
            hrTable('#hr-report-units',['Đơn vị','Đầu kỳ','Tăng / vào','Giảm / ra','Cuối kỳ'],d.units.map(u=>[u.unit,u.opening,u.increase,u.decrease,u.closing]));
            hrReportReadiness(d);return d;
        }catch(e){hrReportStatus('Chưa tính được báo cáo',e.message,true);throw e;}
        finally{buttons.forEach(b=>b.disabled=false);hrReportPending=null;}
    })();
    return hrReportPending;
}
$('#hr-report-form').onsubmit=handle(async e=>{e.preventDefault();await hrReport();});
$('#hr-report-export').onclick=handle(async()=>{await hrReport();await download(await api('/api/hr/reports/export?'+hrReportQuery),'Tong_hop_nhan_su.xlsx');notice('Đã xuất báo cáo tự tổng hợp và danh sách đối chiếu.');});
$('#hr-decision-kind').onchange=()=>{$('#hr-new-job-label').hidden=$('#hr-decision-kind').value==='termination';};
$('#hr-decision-form').onsubmit=handle(async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(e.target));const p=hrPeople.find(p=>p.id===d.person);if(!p)throw Error('Chọn nhân sự đã xác nhận đang làm.');d.version=p.version;const r=await(await json('/api/hr/decisions',d)).json();await download(await api('/api/hr/decisions/'+r.id+'/word'),'Quyet_dinh_du_thao.docx');await hrDecisions();notice('Đã lưu và xuất dự thảo; hồ sơ chưa đổi. Sau khi được duyệt, dùng nút Áp dụng trong danh sách quyết định.');});
async function hrDecisions(){const list=await(await api('/api/hr/decisions')).json();hrTable('#hr-decisions-grid',['Số / Nhân sự','Loại','Ngày hiệu lực','Vị trí mới','Trạng thái','Thao tác'],list.map(d=>{const actions=document.createElement('div');actions.append(hrButton('Tải Word',async()=>download(await api('/api/hr/decisions/'+d.id+'/word'),'Quyet_dinh_du_thao.docx')));if(hrEdit&&!d.applied)actions.append(hrButton('Áp dụng quyết định đã duyệt',async()=>{if(!confirm('Ghi nhận quyết định '+d.fields.so_quyet_dinh+' đã được duyệt? Hồ sơ và phân nhóm sẽ thay đổi từ '+d.effective+'.'))return;await json('/api/hr/decisions/'+d.id+'/apply',{});await hrLoadPeople();await hrDecisions();notice('Đã ghi nhận quyết định. Hồ sơ và báo cáo tự cập nhật từ '+d.effective+'.');}));return[d.fields.so_quyet_dinh+' / '+d.fields.ho_ten,d.fields.loai_quyet_dinh,d.effective,d.fields.chuc_danh_moi+' / '+d.fields.don_vi_moi,d.applied?'Đã ghi nhận theo ngày hiệu lực':'Dự thảo — chưa áp dụng',actions];}));}
$('#hr-decisions-refresh').onclick=handle(hrDecisions);
window.hrActivate=async tab=>{await hrReady;if(['personnel','classification','hr-decisions'].includes(tab)){await hrLoadJobs();await hrLoadPeople();}if(tab==='hr-decisions')await hrDecisions();if(tab==='reports')await hrReport();};
const hrReady=(async()=>{hrOptions=await(await api('/api/hr/options')).json();$('#hr-asof').value=hrOptions.today;$('#hr-baseline').value=hrOptions.today;for(const [id,key] of [['hr-education','education'],['hr-status','statuses'],['hr-group','groups']])hrSelect($('#'+id),Object.entries(hrOptions[key]));hrSelect($('#hr-filter-group'),Object.entries(hrOptions.groups),'Tất cả');hrSelect($('#hr-filter-education'),Object.entries(hrOptions.education),'Tất cả');if(hrEdit){hrSelect($('#hr-batch-status'),Object.entries(hrOptions.statuses),'Giữ nguyên');hrSelect($('#hr-batch-education'),Object.entries(hrOptions.education),'Giữ nguyên');$('#hr-batch-form').elements.effective.value=hrOptions.today;}$('#hr-rule-form').elements.effective.value=hrOptions.today;$('#hr-decision-form').elements.effective.value=hrOptions.today;$('#hr-report-form').elements.start.value=hrOptions.today.slice(0,8)+'01';$('#hr-report-form').elements.end.value=hrOptions.today;await hrLoadJobs();await hrLoadPeople();})();hrReady.catch(e=>notice(e.message,true));

if(hrEdit){
$('#hr-select-filtered').onclick=()=>{hrSelected=new Set(hrFilteredPeople().map(p=>p.id));hrSelectedCount();hrRenderPeople();};
$('#hr-clear-selected').onclick=()=>{hrSelected.clear();hrSelectedCount();hrRenderPeople();};
$('#hr-batch-form').onsubmit=handle(async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(e.target));const changes={};for(const key of ['status','education','job_id'])if(d[key])changes[key]=d[key];if(!hrSelected.size||!Object.keys(changes).length)throw Error('Chọn hồ sơ và trường cần cập nhật trước.');if(!confirm('Ghi nhận '+hrSelected.size+' hồ sơ từ '+d.effective+'? Chỉ tiếp tục khi đã rà soát đúng thông tin.'))return;const r=await(await json('/api/hr/people/batch',{effective:d.effective,changes,selected:hrPeople.filter(p=>hrSelected.has(p.id)).map(p=>({id:p.id,version:p.version}))})).json();await hrLoadPeople();notice('Đã cập nhật '+r.updated+' hồ sơ và lịch sử phân nhóm.');});
}

for(const id of ['hr-filter-status','hr-filter-unit','hr-filter-group','hr-filter-education','hr-filter-age'])$('#'+id).onchange=()=>{hrSelected.clear();hrSelectedCount();hrRenderPeople();};
$('#hr-filter-reset').onclick=()=>{$('#hr-filter-status').value='all';for(const id of ['hr-filter-unit','hr-filter-group','hr-filter-education','hr-filter-age','hr-search'])$('#'+id).value='';hrSelected.clear();hrSelectedCount();hrRenderPeople();};
