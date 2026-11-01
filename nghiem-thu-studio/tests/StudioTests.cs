using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml;
using NghiemThuStudio;
class StudioTests
{
    static int count;
    static void Check(bool ok,string text) { count++; if(!ok) throw new Exception(text); }
    static void Reject(Action action,string text) { bool rejected=false; try { action(); } catch { rejected=true; } Check(rejected,text); }
    public static Dossier Complete(string code="LG",int month=9)
    {
        var d=Catalog.New(code,2026,month);
        foreach(var m in d.Inputs) { m.Value=0; m.Manual=true; m.SourceState="Dữ liệu giả lập"; }
        var values=new Dictionary<string,decimal>{{"q1",100},{"q2",10},{"qh",5},{"p1",100000},{"p2",80000},{"m1",10000},{"m2",8000},{"cdt",2000000},{"cdm",500000},{"at",-100000},{"head",20},{"headm",2},{"dt",260},{"dm",26},{"other",10},{"bh",4},{"standard",26},{"meal",100},{"mealprice",30000},{"b1",10},{"b2",20},{"bp1",13000},{"bp2",20000},{"bp3",26000},{"rt1",100},{"rt2",100},{"rth",code=="LG"?1000:100},{"rm",1}};
        foreach(var kv in values) d.Input(kv.Key).Value=kv.Value;
        d.QuantityUnit="tấn"; d.AttendeesOffice="Thành phần TCHC giả"; d.AttendeesUnit="Thành phần PX giả"; d.SignerOffice="Người ký TCHC giả"; d.SignerUnit="Người ký PX giả";
        d.SourceReviewed=true; d.MeetingDate=new DateTime(2026,10,1);
        d.Works.Add(new WorkLine { Description="GNV giả",Quantity=3.5m,Price=120000,Factor=1,Rounding=1,SourceAmount=420000 }); return d;
    }
    static void Engine()
    {
        var d=Complete(); string before=d.Fingerprint(); var r=Calc.Evaluate(d);
        Check(r.Get("productT")==11800000 && r.Get("productM")==1180000,"Three production lines, independent QL rates, holiday ×2");
        Check(r.Get("gnv")==420000 && r.Get("direct")==14120000 && r.Get("management")==1680000 && r.Get("total")==15800000,"Funds add once, include regime and signed adjustment");
        Check(r.Get("meal")==3000000 && r.Get("bddh")==530000,"Meals and BDDH");
        Check(r.Get("peopleMonths")==11 && r.Get("average")==15800000m/11,"Weighted denominator");
        d.Works.Add(new WorkLine { Kind="SP phụ",Description="SP phụ giả",Quantity=2,Price=1000,Rounding=1 });
        Check(Calc.Evaluate(d).Get("productT")==11802000 && Calc.Evaluate(d).Get("gnv")==420000,"Auxiliary production belongs to product fund, not GNV twice");
        d.Works.RemoveAt(1);
        Check(before==d.Fingerprint() && d.Works[0].SourceAmount==420000,"Calculation is pure and does not mutate source/results");
        d.Change("qh",0); r=Calc.Evaluate(d); Check(r.Get("tth")==0 && r.Get("qlh")==0 && !d.SourceReviewed,"Zero removes both holiday amounts, requires re-review");
        d.Change("q1",null); r=Calc.Evaluate(d); Check(!r.Get("productT").HasValue && !r.Get("total").HasValue,"Missing is not zero");
        d=Complete(); d.Input("p1").Value=0; r=Calc.Evaluate(d); Check(r.Get("tt1")==0 && r.Get("tth")==0,"Zero direct rate clears L1 and holiday");
        d=Complete(); d.Input("m1").Value=0; r=Calc.Evaluate(d); Check(r.Get("ql1")==0 && r.Get("qlh")==0 && r.Get("productT")==11800000,"Independent rates");
        d.Works[0].Quantity=0; Check(Calc.Evaluate(d).Get("gnv")==0,"GNV zero clears cached source amount");
        d.Works[0].Quantity=3.5m; d.Works[0].Factor=-1; Check(Calc.Evaluate(d).Get("gnv")==-420000,"Explicit signed factor");
        d.Works[0].Rounding=0; Check(!Calc.Evaluate(d).Get("gnv").HasValue,"Invalid rounding blocks instead of throwing/using zero");
        Check(Calc.Round(150,100)==200 && Calc.Round(-150,100)==-200,"Excel midpoint rounding");
        Check(Fmt.Money(123456789)=="123,5" && Fmt.Money(1000000)=="1,0" && Fmt.Money(0)=="0,0" && Fmt.Money(null).Contains("Chưa"),"Report million-dong precision");
        decimal? parsed; Check(Fmt.Parse("1.234,5678",out parsed) && parsed==1234.5678m && Fmt.Parse("",out parsed) && parsed==null,"Vietnamese numeric input and empty");
        Check(!Fmt.Parse("abc",out parsed),"Invalid number rejected");
        d=Complete(); var culture=Thread.CurrentThread.CurrentCulture;
        try { Thread.CurrentThread.CurrentCulture=new CultureInfo("en-US"); string hash=d.Fingerprint(); Thread.CurrentThread.CurrentCulture=new CultureInfo("vi-VN"); Check(hash==d.Fingerprint(),"Approval hash stable across machine number culture"); Thread.CurrentThread.CurrentCulture=new CultureInfo("ar-SA"); Check(hash==d.Fingerprint(),"Approval hash has invariant calendar"); }
        finally { Thread.CurrentThread.CurrentCulture=culture; }
    }
    static string ReadDoc(string path)
    { using(var z=ZipFile.OpenRead(path)) using(var s=z.GetEntry("word/document.xml").Open()) { var d=new XmlDocument(); d.Load(s); return d.InnerText; } }
    static void Workflow(string dir)
    {
        var p=new Project { Current=Complete() }; Check(ProjectStore.ApprovalProblems(p.Current).Count==0,"Complete reviewed record can close"); ProjectStore.Approve(p);
        Check(p.Current.Approved && p.Closed.Count==1,"Close creates snapshot"); decimal? frozen=p.Closed[0].ApprovedResult["total"];
        p.Current.Change("q1",0); Check(!p.Current.Approved && p.Closed[0].Approved && p.Closed[0].ApprovedResult["total"]==frozen,"Editing cannot mutate closed history");
        Reject(()=>ProjectStore.Approve(p),"Must re-review after edit");
        p.Current.SourceReviewed=true; ProjectStore.Approve(p); Check(p.Closed.Count==2,"Revision appends, not overwrites");
        string path=Path.Combine(dir,"project.ntproject"); ProjectStore.Save(path,p); var read=ProjectStore.Read(path);
        Check(read.Current.Approved && read.Closed.Count==2 && read.Current.MeetingDate==new DateTime(2026,10,1),"Project roundtrip including date and approval hash");
        read.Current.Change("p2",90000); ProjectStore.Save(path,read); Check(File.Exists(path+".bak") && ProjectStore.Read(path+".bak").Current.Approved,"Atomic save keeps prior valid backup");
        Reject(()=>ProjectStore.Save(Path.Combine(dir,"source.xlsx"),read),"Cannot save project over Excel");
        Reject(()=>Report.Export(read,Path.Combine(dir,"source.xlsx"),false),"Cannot export Word over Excel");
        p=new Project { Current=Complete("TK",1) }; ProjectStore.Approve(p); p.Current=Complete("TK",2); ProjectStore.Approve(p); p.Current=Complete("TK",3); ProjectStore.Approve(p);
        List<int> missing; var agg=ProjectStore.Aggregate(p,1,3,out missing);
        Check(missing.Count==0 && agg.Get("total")==47400000 && agg.Get("average")==15800000m/11,"Quarter sums closed periods and weighted average, no extra division by 3");
        p.Current=ProjectStore.Copy(p.Closed[1]); p.Current.Change("q1",0); p.Current.SourceReviewed=true; ProjectStore.Approve(p);
        agg=ProjectStore.Aggregate(p,1,3,out missing); Check(agg.Get("total")==p.Closed[0].ApprovedResult["total"]+p.Closed[2].ApprovedResult["total"]+p.Closed[3].ApprovedResult["total"],"Latest revision only, no double count");
        var drafts=new Project { Current=Complete("TK",1) }; ProjectStore.ArchiveDraft(drafts);
        drafts.Current=Complete("LG",1); ProjectStore.ArchiveDraft(drafts);
        Check(drafts.Drafts.Count==2 && drafts.Drafts.Select(x=>x.Code).Distinct().Count()==2,"Different unit drafts preserved before switching");
        ProjectStore.Approve(drafts); Check(drafts.Drafts.Count==1 && drafts.Closed.Count==1,"Closing removes only matching draft, retains other unit draft");
        var absent=new Project { Current=Complete("LG",9) }; ProjectStore.Approve(absent); agg=ProjectStore.Aggregate(absent,7,9,out missing);
        Check(missing.SequenceEqual(new[]{7,8}) && !agg.Get("total").HasValue,"Missing months never treated as zero");
        foreach(var unit in Catalog.Units)
        {
            var report=new Project { Current=Complete(unit[0]) }; string doc=Path.Combine(dir,unit[0]+".docx");
            Report.Export(report,doc,false); string text=ReadDoc(doc); Check(text.Contains("DỰ THẢO") && text.Contains("triệu đồng") && text.Contains("15,8") && text.Contains(unit[1]),"Draft report with units and correct factory: "+unit[0]);
            ProjectStore.Approve(report); Report.Export(report,doc,false); text=ReadDoc(doc);
            Check(!text.Contains("DỰ THẢO") && text.Contains("09/2026")==false && text.Contains("9/2026") && text.Contains("THÁNG 9 NĂM 2026") && File.Exists(doc+".bak"),"Closed monthly report, correct period and backup: "+unit[0]);
            Report.Export(report,doc,true); text=ReadDoc(doc); Check(text.Contains("DỰ THẢO") && text.Contains("Chưa đủ kỳ") && text.Contains("QUÝ III") && text.Contains("9 THÁNG ĐẦU NĂM 2026"),"Incomplete quarterly export cannot look official: "+unit[0]);
        }
    }
    static void Entry(ZipArchive zip,string name,string text) { using(var w=new StreamWriter(zip.CreateEntry(name).Open())) w.Write(text); }
    static void Source(string dir)
    {
        string path=Path.Combine(dir,"source.xlsx");
        using(var zip=ZipFile.Open(path,ZipArchiveMode.Create))
        {
            Entry(zip,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='NT gang' sheetId='1' r:id='rId1'/></sheets></workbook>");
            Entry(zip,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>");
            Entry(zip,"xl/worksheets/sheet1.xml","<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData><row r='2'><c r='I2' t='inlineStr'><is><t>Tháng 9 năm 2026</t></is></c></row><row r='15'><c r='D15'><v>20</v></c></row><row r='17'><c r='D17' t='e'><f>#REF!</f><v>#REF!</v></c></row><row r='12'><c r='D12'><f>1-1</f><v>0</v></c></row><row r='52'><c r='D52'><f>[1]Sheet1!A1</f></c></row></sheetData></worksheet>");
            Entry(zip,"xl/externalLinks/_rels/externalLink1.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='file:///C:/Missing/file.xlsx' TargetMode='External'/></Relationships>");
        }
        string before=Workbook.Hash(path); var d=Workbook.Import(path,"LG",2026,9);
        Check(d.V("head")==20 && d.V("headm")==null && d.V("dt")==0 && d.V("q1")==null,"Read cached zero, preserve Excel error/missing cache as unknown");
        Check(d.V("m1")==null && d.V("m2")==null,"Never invent management rates");
        Check(d.Diagnostics.Any(x=>x.Contains("file.xlsx")) && before==Workbook.Hash(path),"External file diagnostics, no source writes");
        var mismatch=Workbook.Import(path,"LG",2026,10); Check(mismatch.Diagnostics.Any(x=>x.StartsWith("KỲ NGUỒN")),"Source period mismatch detected");
        Reject(()=>Workbook.Import(path,"LT",2026,9),"Do not pick a similar sheet if canonical sheet missing");
        var p=new Project { Current=Complete() }; p.Current.SourcePath=path; p.Current.SourceHash=before; ProjectStore.Approve(p);
        using(var stream=new FileStream(path,FileMode.Append)) stream.WriteByte(0);
        Check(ProjectStore.ApprovalProblems(p.Current).Any(x=>x.Contains("thay đổi")),"Source modification invalidates readiness to approve");
        string output=Path.Combine(dir,"changed.docx"); Report.Export(p,output,false); Check(ReadDoc(output).Contains("DỰ THẢO"),"Changed source cannot export official report under old approval");
    }
    static void Real(string path)
    {
        string hash=Workbook.Hash(path); bool sawError=false;
        foreach(var unit in Catalog.Units)
        {
            var watch=System.Diagnostics.Stopwatch.StartNew();
            var d=Workbook.Import(path,unit[0],2026,9); var r=Calc.Evaluate(d);
            Console.WriteLine("Imported "+unit[0]+" in "+watch.ElapsedMilliseconds+" ms");
            Check(d.Inputs.Count>20 && d.SourceHash==hash && d.V("m1")==null && !r.Get("productM").HasValue,"Real import remains unresolved without approved QL rates: "+unit[0]);
            foreach(var m in d.Inputs.Where(x=>(x.SourceState??"").Contains("Lỗi Excel"))) { sawError=true; Check(!m.Value.HasValue,"Real Excel error not converted to zero"); }
            if(unit[0]=="KT") Check(d.Diagnostics.Any(x=>x.StartsWith("KỲ NGUỒN")),"Actual Khai tuyen stale period flagged");
        }
        Check(sawError && hash==Workbook.Hash(path),"Real workbook errors exposed, source unchanged byte-for-byte");
    }
    static void RealWord(string path)
    {
        string before=Workbook.Hash(path); var lines=TemplateInspector.Describe(path);
        Check(lines.Any(x=>x.Contains("9 tiêu đề")) && lines.Any(x=>x.Contains("tập nhiều kỳ")),"Actual Word is a nine-period dossier, not a single replacement template");
        Check(lines.Any(x=>x.Contains("năm 5 chữ số")),"Actual Word copy/paste date anomaly detected");
        Check(before==Workbook.Hash(path),"Original Word source remains unchanged");
    }
    static int Main(string[] args)
    {
        string dir=Path.Combine(Path.GetTempPath(),"nghiemthu-tests-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        try { Console.WriteLine("Checking engine..."); Engine(); Console.WriteLine("Checking project snapshots and reports..."); Workflow(dir); Console.WriteLine("Checking source diagnostics..."); Source(dir); Console.WriteLine("Checking selected real workbook..."); for(int i=0;i+1<args.Length;i+=2) { if(args[i]=="--real-workbook") Real(args[i+1]); else if(args[i]=="--real-word") RealWord(args[i+1]); else throw new Exception("Unknown test argument"); } Console.WriteLine("PASS: "+count+" Nghiệm Thu Studio checks."); return 0; }
        finally { Directory.Delete(dir,true); }
    }
}
