using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.Text;

namespace NghiemThuStudio
{
    [DataContract] public sealed class Measure
    {
        [DataMember] public string Key;
        [DataMember] public string Label;
        [DataMember] public string Unit;
        [DataMember] public decimal? Value;
        [DataMember] public decimal? SourceValue;
        [DataMember] public string SourceCell;
        [DataMember] public string Formula;
        [DataMember] public string SourceState;
        [DataMember] public bool Manual;
    }
    [DataContract] public sealed class WorkLine
    {
        [DataMember] public string Kind="GNV";
        [DataMember] public string Description="";
        [DataMember] public string Unit="Công";
        [DataMember] public decimal? Quantity;
        [DataMember] public decimal? Price;
        [DataMember] public decimal Factor=1;
        [DataMember] public decimal Rounding=1;
        [DataMember] public decimal? SourceAmount;
        [DataMember] public string SourceCell;
        [DataMember] public string SourceFormula;
        public decimal? Amount { get { return Quantity.HasValue && Price.HasValue && Rounding>0 ? Calc.Round(Quantity.Value*Price.Value*Factor,Rounding) : (decimal?)null; } }
    }
    [DataContract] public sealed class Audit
    {
        [DataMember] public string Time;
        [DataMember] public string Action;
        [DataMember] public string Detail;
    }
    [DataContract] public sealed class Dossier
    {
        [DataMember] public string Id=Guid.NewGuid().ToString("N");
        [DataMember] public string Code;
        [DataMember] public string Name;
        [DataMember] public int Year=DateTime.Today.Year;
        [DataMember] public int Month=DateTime.Today.Month;
        [DataMember] public string MeetingDay=DateTime.Today.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture);
        public DateTime MeetingDate { get { return DateTime.ParseExact(MeetingDay,"yyyy-MM-dd",CultureInfo.InvariantCulture); } set { MeetingDay=value.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture); } }
        [DataMember] public string Location="Phòng họp";
        [DataMember] public string AttendeesOffice="";
        [DataMember] public string AttendeesUnit="";
        [DataMember] public string SignerOffice="";
        [DataMember] public string SignerUnit="";
        [DataMember] public string Recommendations="";
        [DataMember] public string Explanation="";
        [DataMember] public string SourcePath;
        [DataMember] public string SourceHash;
        [DataMember] public bool SourceReviewed;
        [DataMember] public string ApprovedHash;
        [DataMember] public string ApprovedAt;
        [DataMember] public Dictionary<string,decimal?> ApprovedResult;
        [DataMember] public string RuleVersion="1.0.0";
        [DataMember] public string QuantityUnit="tấn";
        [DataMember] public List<Measure> Inputs=new List<Measure>();
        [DataMember] public List<WorkLine> Works=new List<WorkLine>();
        [DataMember] public List<string> Diagnostics=new List<string>();
        [DataMember] public Dictionary<string,decimal?> References=new Dictionary<string,decimal?>();
        [DataMember] public List<Audit> Log=new List<Audit>();
        public Measure Input(string key) { return Inputs.First(m=>m.Key==key); }
        public decimal? V(string key) { return Input(key).Value; }
        public void Note(string action,string detail)
        { Log.Add(new Audit { Time=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss",CultureInfo.InvariantCulture),Action=action,Detail=detail }); }
        public void Change(string key,decimal? value)
        {
            var m=Input(key); if(m.Value==value) return;
            Note("Sửa đầu vào",m.Label+": "+Fmt.Number(m.Value)+" → "+Fmt.Number(value));
            m.Value=value; m.Manual=true; SourceReviewed=false; ApprovedHash=null;
        }
        static string Canon(decimal? value) { return value.HasValue?value.Value.ToString("G29",CultureInfo.InvariantCulture):"?"; }
        public string Fingerprint()
        {
            var s=new StringBuilder(); s.Append(RuleVersion).Append('|').Append(Code).Append('|').Append(Name).Append('|').Append(Year).Append('|').Append(Month).Append('|').Append(SourceHash).Append('|').Append(SourceReviewed);
            foreach(var m in Inputs) s.Append('|').Append(m.Key).Append('=').Append(m.Value.HasValue?m.Value.Value.ToString("G29",CultureInfo.InvariantCulture):"?");
            foreach(var w in Works) s.Append('|').Append(w.Kind).Append('|').Append(w.Description).Append('|').Append(w.Unit).Append('|').Append(Canon(w.Quantity)).Append('|').Append(Canon(w.Price)).Append('|').Append(Canon(w.Factor)).Append('|').Append(Canon(w.Rounding));
            s.Append('|').Append(MeetingDate.ToString("yyyyMMdd",CultureInfo.InvariantCulture)).Append('|').Append(Location).Append('|').Append(AttendeesOffice).Append('|').Append(AttendeesUnit).Append('|').Append(SignerOffice).Append('|').Append(SignerUnit).Append('|').Append(Recommendations).Append('|').Append(Explanation).Append('|').Append(QuantityUnit);
            using(var h=SHA256.Create()) return BitConverter.ToString(h.ComputeHash(Encoding.UTF8.GetBytes(s.ToString()))).Replace("-","");
        }
        public bool Approved { get { return !string.IsNullOrEmpty(ApprovedHash) && ApprovedHash==Fingerprint(); } }
    }
    [DataContract] public sealed class Project
    {
        [DataMember] public int Version=1;
        [DataMember] public Dossier Current;
        [DataMember] public List<Dossier> Closed=new List<Dossier>();
        [DataMember] public List<Dossier> Drafts=new List<Dossier>();
    }
    public static class Catalog
    {
        public static readonly string[][] Units={ new[]{"TK","Phân xưởng Thiêu kết","NT thiêu kết"},new[]{"LG","Phân xưởng Luyện gang","NT gang"},new[]{"LT","Phân xưởng Luyện thép","NT Thép"},new[]{"CDNL","Phân xưởng Cơ điện Năng lượng","NT CĐNL"},new[]{"KT","Phân xưởng Khai tuyển","NT. K. tuyển "} };
        static readonly string[][] Defs={
            new[]{"q1","Sản lượng loại 1","ĐVT sản lượng"},new[]{"q2","Sản lượng loại 2","ĐVT sản lượng"},new[]{"qh","Sản lượng ngày lễ","ĐVT sản lượng"},
            new[]{"p1","Đơn giá trực tiếp loại 1","đ/ĐVT"},new[]{"p2","Đơn giá trực tiếp loại 2","đ/ĐVT"},new[]{"m1","Đơn giá quản lý loại 1","đ/ĐVT"},new[]{"m2","Đơn giá quản lý loại 2","đ/ĐVT"},
            new[]{"rt1","Làm tròn tiền trực tiếp loại 1","đồng"},new[]{"rt2","Làm tròn tiền trực tiếp loại 2","đồng"},new[]{"rth","Làm tròn tiền trực tiếp ngày lễ","đồng"},new[]{"rm","Làm tròn tiền quản lý từng dòng","đồng"},
            new[]{"cdt","Tiền chế độ trực tiếp (tổng đã đối chiếu)","đồng"},new[]{"cdm","Tiền chế độ quản lý (tổng đã đối chiếu)","đồng"},
            new[]{"at","Điều chỉnh trực tiếp (+/−)","đồng"},new[]{"am","Điều chỉnh quản lý (+/−)","đồng"},
            new[]{"head","Lao động cuối kỳ","người"},new[]{"headm","Lao động quản lý","người"},
            new[]{"dt","Công sản phẩm trực tiếp","công"},new[]{"dm","Công sản phẩm quản lý","công"},new[]{"other","Công lễ/phép/học/chế độ (không BHXH)","công"},new[]{"bh","Công nghỉ hưởng BHXH","công"},
            new[]{"standard","Ngày công chuẩn tính bình quân","công/tháng"},
            new[]{"meal","Công ăn ca đã đối chiếu","công"},new[]{"mealprice","Đơn giá ăn ca","đ/công"},
            new[]{"b1","Công BDĐH mức 1","công"},new[]{"b2","Công BDĐH mức 2","công"},new[]{"b3","Công BDĐH mức 3","công"},
            new[]{"bp1","Đơn giá BDĐH mức 1","đ/công"},new[]{"bp2","Đơn giá BDĐH mức 2","đ/công"},new[]{"bp3","Đơn giá BDĐH mức 3","đ/công"}
        };
        public static Dossier New(string code,int year,int month)
        {
            var u=Units.First(x=>x[0]==code); var d=new Dossier { Code=code,Name=u[1],Year=year,Month=month,QuantityUnit=code=="CDNL"?"ĐVT cần xác nhận":"tấn" };
            foreach(var x in Defs) d.Inputs.Add(new Measure { Key=x[0],Label=x[1],Unit=x[2],SourceState="Chưa nhập" });
            foreach(string key in new[]{"at","am"}) { d.Input(key).Value=0; d.Input(key).SourceState="Không điều chỉnh"; }
            foreach(var pair in new[]{new[]{"standard","26"},new[]{"rt1","100"},new[]{"rt2","100"},new[]{"rth",code=="LG"?"1000":"100"},new[]{"rm","1"},new[]{"mealprice","30000"},new[]{"bp1","13000"},new[]{"bp2","20000"},new[]{"bp3","26000"}})
            { d.Input(pair[0]).Value=decimal.Parse(pair[1],CultureInfo.InvariantCulture); d.Input(pair[0]).SourceState="Mức khởi tạo — cần đối chiếu"; }
            return d;
        }
    }
    public sealed class Result
    {
        public Dictionary<string,decimal?> Values=new Dictionary<string,decimal?>();
        public List<string> Issues=new List<string>();
        public decimal? Get(string key) { decimal? value; return Values.TryGetValue(key,out value)?value:null; }
    }
    public static class Calc
    {
        public static decimal Round(decimal v,decimal unit=1) { if(unit<=0) throw new ArgumentException("Mức làm tròn phải > 0."); return Math.Round(v/unit,0,MidpointRounding.AwayFromZero)*unit; }
        static decimal? Add(params decimal?[] xs) { return xs.All(x=>x.HasValue)?xs.Sum(x=>x.Value):(decimal?)null; }
        static decimal? Mul(decimal? a,decimal? b,decimal unit=1) { return a.HasValue && b.HasValue?Round(a.Value*b.Value,unit):(decimal?)null; }
        static decimal? Line(Dossier d,string q,string p,string rounding,int multiplier=1)
        { var unit=d.V(rounding); return unit>0?Mul(d.V(q),d.V(p)*multiplier,unit.Value):null; }
        public static Result Evaluate(Dossier d)
        {
            var r=new Result(); var v=r.Values;
            foreach(var m in d.Inputs)
            {
                if(!m.Value.HasValue) r.Issues.Add("Chưa đủ đầu vào: "+m.Label);
                else if(m.Key!="at" && m.Key!="am" && m.Value<0) r.Issues.Add("Không nhận số âm: "+m.Label);
            }
            foreach(string key in new[]{"rt1","rt2","rth","rm"}) if(d.V(key)<=0) r.Issues.Add("Mức làm tròn phải > 0: "+d.Input(key).Label);
            if(d.V("standard")<=0) r.Issues.Add("Ngày công chuẩn phải lớn hơn 0.");
            foreach(string key in new[]{"head","headm"}) if(d.V(key).HasValue && d.V(key).Value%1!=0) r.Issues.Add("Số người phải là số nguyên: "+d.Input(key).Label);
            if(d.V("headm")>d.V("head")) r.Issues.Add("Lao động quản lý vượt tổng lao động.");
            v["tt1"]=Line(d,"q1","p1","rt1"); v["tt2"]=Line(d,"q2","p2","rt2");
            v["tth"]=Line(d,"qh","p1","rth",2);
            v["ql1"]=Line(d,"q1","m1","rm"); v["ql2"]=Line(d,"q2","m2","rm"); v["qlh"]=Line(d,"qh","m1","rm",2);
            v["productT"]=Add(v["tt1"],v["tt2"],v["tth"]); v["productM"]=Add(v["ql1"],v["ql2"],v["qlh"]);
            bool workMissing=false,extraMissing=false; decimal work=0,extra=0;
            foreach(var w in d.Works)
            {
                if(w.Kind!="GNV" && w.Kind!="SP phụ") { workMissing=true; r.Issues.Add("Loại khoản chưa hợp lệ: chọn GNV hoặc SP phụ."); }
                if(string.IsNullOrWhiteSpace(w.Description) || w.Description.StartsWith("#") || !w.Amount.HasValue || w.Rounding<=0)
                { if(w.Kind=="SP phụ") extraMissing=true; else workMissing=true; r.Issues.Add(w.Kind+" thiếu nội dung/khối lượng/giá hoặc mức làm tròn không hợp lệ."); }
                else if(w.Kind=="SP phụ") extra+=w.Amount.Value; else work+=w.Amount.Value;
                if(w.SourceAmount.HasValue && w.Amount.HasValue && w.SourceAmount!=w.Amount)
                    r.Issues.Add("Khoản lệch nguồn: "+w.Description+" (cần giải trình hệ số/giảm trừ).");
            }
            v["extraProduct"]=extraMissing?(decimal?)null:extra;
            v["productT"]=Add(v["productT"],v["extraProduct"]);
            v["gnv"]=workMissing?(decimal?)null:work;
            v["regimeT"]=d.V("cdt"); v["regimeM"]=d.V("cdm"); v["adjustT"]=d.V("at"); v["adjustM"]=d.V("am");
            v["direct"]=Add(v["productT"],v["gnv"],d.V("cdt"),d.V("at"));
            v["management"]=Add(v["productM"],d.V("cdm"),d.V("am"));
            v["total"]=Add(v["direct"],v["management"]);
            v["meal"]=Mul(d.V("meal"),d.V("mealprice"));
            for(int i=1;i<=3;i++) { v["bddh"+i]=Mul(d.V("b"+i),d.V("bp"+i)); v["bddhDays"+i]=d.V("b"+i); }
            v["bddh"]=Add(v["bddh1"],v["bddh2"],v["bddh3"]); v["mealDays"]=d.V("meal"); v["otherDays"]=d.V("other"); v["bhDays"]=d.V("bh");
            v["days"]=Add(d.V("dt"),d.V("dm"),d.V("other"),d.V("bh"));
            v["spdays"]=Add(d.V("dt"),d.V("dm"));
            v["peopleMonths"]=d.V("standard")>0?v["spdays"]/d.V("standard"):null;
            v["average"]=v["peopleMonths"]>0?v["total"]/v["peopleMonths"]:null;
            foreach(var kv in d.References)
                if(r.Get(kv.Key).HasValue && kv.Value.HasValue && Math.Abs(r.Get(kv.Key).Value-kv.Value.Value)>1)
                    r.Issues.Add("Chênh nguồn "+Labels(kv.Key)+": tính "+Fmt.Number(r.Get(kv.Key))+"đ; nguồn "+Fmt.Number(kv.Value)+"đ; chênh "+Fmt.Number(r.Get(kv.Key)-kv.Value)+"đ.");
            return r;
        }
        public static string Labels(string key)
        {
            var map=new Dictionary<string,string>{{"productT","Quỹ sản phẩm trực tiếp"},{"productM","Quỹ sản phẩm quản lý"},{"gnv","Công việc phát sinh"},{"extraProduct","Trong đó: sản phẩm phụ"},{"regimeT","Chế độ trực tiếp"},{"regimeM","Chế độ quản lý"},{"adjustT","Điều chỉnh trực tiếp"},{"adjustM","Điều chỉnh quản lý"},{"direct","Lương trực tiếp"},{"management","Lương quản lý"},{"total","Tổng quỹ tiền lương"},{"meal","Tiền ăn ca"},{"bddh","Tiền BDĐH"},{"days","Tổng ngày công"},{"spdays","Công sản phẩm"},{"peopleMonths","Lao động thực tế bình quân"},{"average","Thu nhập bình quân"},{"tt1","TT loại 1"},{"tt2","TT loại 2"},{"tth","TT ngày lễ"},{"ql1","QL loại 1"},{"ql2","QL loại 2"},{"qlh","QL ngày lễ"}};
            string text; return map.TryGetValue(key,out text)?text:key;
        }
    }
    public static class Fmt
    {
        public static readonly CultureInfo Vi=new CultureInfo("vi-VN");
        public static string Number(decimal? v) { return v.HasValue?v.Value.ToString("#,##0.####",Vi):"Chưa có dữ liệu"; }
        public static string Money(decimal? v) { return v.HasValue?Math.Round(v.Value/1000000m,1,MidpointRounding.AwayFromZero).ToString("#,##0.0",Vi):"Chưa đủ dữ liệu"; }
        public static bool Parse(string text,out decimal? value)
        {
            if(string.IsNullOrWhiteSpace(text)) { value=null; return true; }
            decimal d; bool ok=decimal.TryParse(text,NumberStyles.Number,Vi,out d); value=ok?(decimal?)d:null; return ok;
        }
    }
}
