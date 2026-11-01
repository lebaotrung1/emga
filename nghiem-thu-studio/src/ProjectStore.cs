using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.Serialization.Json;

namespace NghiemThuStudio
{
    public static class ProjectStore
    {
        public static T Copy<T>(T value)
        {
            var serializer=new DataContractJsonSerializer(typeof(T));
            using(var stream=new MemoryStream()) { serializer.WriteObject(stream,value); stream.Position=0; return (T)serializer.ReadObject(stream); }
        }
        public static Project Read(string path)
        {
            using(var stream=File.OpenRead(path))
            {
                var value=(Project)new DataContractJsonSerializer(typeof(Project)).ReadObject(stream);
                if(value==null || value.Version!=1 || value.Current==null || value.Closed==null) throw new InvalidDataException("Hồ sơ không đúng định dạng hoặc phiên bản không được hỗ trợ.");
                if(value.Drafts==null) value.Drafts=new List<Dossier>();
                CheckDossier(value.Current);
                foreach(var d in value.Closed.Concat(value.Drafts)) CheckDossier(d);
                return value;
            }
        }
        static void CheckDossier(Dossier d)
        {
            DateTime date; Guid id;
            if(d==null || d.RuleVersion!="1.0.0" || d.Month<1 || d.Month>12 || d.Year<2000 || d.Year>2100
                || !Guid.TryParseExact(d.Id,"N",out id) || !Catalog.Units.Any(x=>x[0]==d.Code) || d.Inputs==null || d.Works==null || d.Log==null || d.Diagnostics==null || d.References==null
                || !DateTime.TryParseExact(d.MeetingDay,"yyyy-MM-dd",System.Globalization.CultureInfo.InvariantCulture,System.Globalization.DateTimeStyles.None,out date))
                throw new InvalidDataException("Hồ sơ thiếu trường hoặc phiên bản công thức không được hỗ trợ.");
            if(d.Name!=Catalog.Units.First(x=>x[0]==d.Code)[1]) throw new InvalidDataException("Tên đơn vị không khớp mã phân xưởng.");
            var expected=Catalog.New(d.Code,d.Year,d.Month).Inputs.Select(x=>x.Key).OrderBy(x=>x).ToArray();
            if(!expected.SequenceEqual(d.Inputs.Select(x=>x.Key).OrderBy(x=>x))) throw new InvalidDataException("Danh sách đầu vào hồ sơ không hợp lệ.");
            Calc.Evaluate(d); // Reject malformed/overflowing records before replacing the open project.
            if(d.Approved && d.ApprovedResult==null) throw new InvalidDataException("Bản chốt thiếu ảnh chụp kết quả.");
        }
        public static void GuardOutput(Project project,string path,string extension)
        {
            if(!string.Equals(Path.GetExtension(path),extension,StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Đầu ra phải có đuôi "+extension+". Không ghi đè nguồn.");
            string full=Path.GetFullPath(path);
            foreach(var d in project.Closed.Concat(project.Drafts).Concat(new[]{project.Current}))
                if(!string.IsNullOrEmpty(d.SourcePath) && string.Equals(full,Path.GetFullPath(d.SourcePath),StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Không ghi đè Excel nguồn.");
        }
        public static void ArchiveDraft(Project project)
        {
            var d=project.Current;
            if(d==null || d.Approved || (string.IsNullOrEmpty(d.SourcePath) && !d.Inputs.Any(x=>x.Manual) && d.Works.Count==0 && string.IsNullOrWhiteSpace(d.AttendeesOffice) && string.IsNullOrWhiteSpace(d.Explanation))) return;
            var snapshot=Copy(d); project.Drafts.RemoveAll(x=>x.Id==d.Id); project.Drafts.Add(snapshot);
        }
        public static void Save(string path,Project project)
        {
            GuardOutput(project,path,".ntproject");
            ArchiveDraft(project);
            path=Path.GetFullPath(path); string directory=Path.GetDirectoryName(path); Directory.CreateDirectory(directory);
            string temp=path+".tmp-"+Guid.NewGuid().ToString("N");
            try
            {
                using(var stream=new FileStream(temp,FileMode.CreateNew,FileAccess.Write,FileShare.None))
                { new DataContractJsonSerializer(typeof(Project)).WriteObject(stream,project); stream.Flush(true); }
                if(File.Exists(path)) File.Replace(temp,path,path+".bak"); else File.Move(temp,path);
            }
            finally { if(File.Exists(temp)) File.Delete(temp); }
        }
        public static List<string> ApprovalProblems(Dossier d)
        {
            CheckDossier(d);
            var calc=Calc.Evaluate(d);
            var problems=calc.Issues.Where(s=>!s.StartsWith("Chênh nguồn") && !s.StartsWith("Khoản lệch nguồn")).ToList();
            if(string.IsNullOrWhiteSpace(d.AttendeesOffice) || string.IsNullOrWhiteSpace(d.AttendeesUnit)) problems.Add("Chưa nhập đủ thành phần TCHC và phân xưởng.");
            if(string.IsNullOrWhiteSpace(d.SignerOffice) || string.IsNullOrWhiteSpace(d.SignerUnit)) problems.Add("Chưa nhập đủ người ký hai bên.");
            if(string.IsNullOrWhiteSpace(d.QuantityUnit) || d.QuantityUnit.Contains("cần xác nhận")) problems.Add("Chưa xác nhận đơn vị sản lượng/cơ sở tính.");
            if(!d.SourceReviewed) problems.Add("Chưa xác nhận đã đối chiếu đầu vào và nguồn.");
            if(!Workbook.Unchanged(d)) problems.Add("Tệp nguồn đã thay đổi/không còn; nạp lại nguồn để đối chiếu trước khi chốt.");
            if((calc.Issues.Any(s=>s.StartsWith("Chênh nguồn") || s.StartsWith("Khoản lệch nguồn")) || d.Diagnostics.Any(s=>s.StartsWith("KỲ NGUỒN"))) && string.IsNullOrWhiteSpace(d.Explanation))
                problems.Add("Có chênh lệch/kỳ nguồn khác: cần ghi giải trình trước khi chốt.");
            return problems;
        }
        public static void Approve(Project project)
        {
            var d=project.Current; var problems=ApprovalProblems(d);
            if(problems.Count>0) throw new InvalidOperationException(string.Join("\n",problems));
            d.ApprovedAt=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss",System.Globalization.CultureInfo.InvariantCulture);
            d.ApprovedResult=new Dictionary<string,decimal?>(Calc.Evaluate(d).Values);
            d.ApprovedHash=d.Fingerprint(); d.Note("Chốt kỳ","Chốt "+d.Code+" "+d.Month+"/"+d.Year+"; mã "+d.ApprovedHash);
            project.Closed.Add(Copy(d)); project.Drafts.RemoveAll(x=>x.Id==d.Id);
        }
        public static Result Aggregate(Project project,int firstMonth,int lastMonth,out List<int> missing)
        {
            var current=project.Current; missing=new List<int>(); var rows=new List<Dictionary<string,decimal?>>();
            for(int month=firstMonth;month<=lastMonth;month++)
            {
                Dossier d=project.Closed.LastOrDefault(x=>x.Code==current.Code && x.Year==current.Year && x.Month==month && x.Approved);
                if(month==current.Month) d=current;
                if(d==null || !d.Approved) missing.Add(month);
                if(d!=null) rows.Add(d.Approved && d.ApprovedResult!=null?d.ApprovedResult:Calc.Evaluate(d).Values);
            }
            var result=new Result();
            foreach(string key in new[]{"productT","productM","extraProduct","gnv","regimeT","regimeM","adjustT","adjustM","direct","management","total","meal","mealDays","bddh","bddh1","bddh2","bddh3","bddhDays1","bddhDays2","bddhDays3","days","spdays","otherDays","bhDays","peopleMonths"})
                result.Values[key]=rows.Count==lastMonth-firstMonth+1 && rows.All(r=>r.ContainsKey(key)&&r[key].HasValue)?rows.Sum(r=>r[key].Value):(decimal?)null;
            result.Values["average"]=result.Get("peopleMonths")>0?result.Get("total")/result.Get("peopleMonths"):null;
            return result;
        }
    }
}
