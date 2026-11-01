using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Xml;
using CiscoLDTLApp;
class ClassificationAndSaveTests
{
    static int checks;
    static void Check(bool ok,string text) { checks++; if(!ok) throw new Exception(text); }
    static void Entry(ZipArchive z,string path,string text) { using(var w=new StreamWriter(z.CreateEntry(path).Open())) w.Write(text); }
    static string Text(string a,string s) { return "<c r='"+a+"' t='inlineStr'><is><t>"+System.Security.SecurityElement.Escape(s)+"</t></is></c>"; }
    static string Num(string a,double n) { return "<c r='"+a+"'><v>"+n.ToString(System.Globalization.CultureInfo.InvariantCulture)+"</v></c>"; }
    static string Sheet(string body) { return "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData>"+body+"</sheetData></worksheet>"; }
    static void Fixture(string file,bool differentName=false)
    {
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)) {
            Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/></Types>");
            Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/></Relationships>");
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='thép' sheetId='1' r:id='rId1'/><sheet name='K. tuyển' sheetId='2' r:id='rId2'/><sheet name='Tổng hợp công' sheetId='3' r:id='rId3'/></sheets><calcPr/></workbook>");
            Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/><Relationship Id='rId2' Target='worksheets/sheet2.xml'/><Relationship Id='rId3' Target='worksheets/sheet3.xml'/></Relationships>");
            string header="<row r='4'>"+Text("C4","Bộ phận quản lý")+Num("D4",0)+Text("E4","TT")+"</row>";
            string employee="<row r='5'>"+Num("A5",1)+Text("C5","Nhân sự giả A")+Text("D5","171")+Text("E5","TT")+Num("G5",5200000)+Num("K5",26)+Num("M5",3)+Num("N5",3)+Num("P5",1)+Num("R5",2)+Num("T5",10)+Num("AB5",1000)+Num("AD5",100)+"</row>";
            Entry(z,"xl/worksheets/sheet1.xml",Sheet(header+employee));
            string copy="<row r='5'>"+Num("A5",1)+Text("C5",differentName?"Nhân sự giả B":"Nhân sự giả A")+Text("D5","171")+Text("E5","CN")+Num("G5",5200000)+Num("K5",0)+Num("M5",0)+Num("T5",10)+"</row>";
            Entry(z,"xl/worksheets/sheet2.xml",Sheet(header+copy));
            string master="<row r='6'>"+Text("C6","171")+Text("D6","Nhân sự giả A")+Num("K6",5200000)+Num("L6",26)+Num("M6",1)+Num("N6",2)+Num("O6",3)+Num("P6",4)+Num("Q6",5)+Num("W6",1000)+Num("AB6",100)+Num("AF6",10)+"</row>";
            if(differentName) master+="<row r='7'>"+Text("C7","171")+Text("D7","Nhân sự giả B")+Num("L7",13)+"</row>";
            Entry(z,"xl/worksheets/sheet3.xml",Sheet(master));
        }
    }
    static double Cell(string file,string sheet,string cell) { using(var z=ZipFile.OpenRead(file)) { var map=ExcelService.GetSheetMap(z);return ExcelService.GetCellDouble(ExcelService.LoadSheetXml(z,map[sheet]),ExcelService.LoadSharedStrings(z),cell,-999); } }
    static string Hash(string file) { using(var sha=SHA256.Create()) using(var s=File.OpenRead(file)) return Convert.ToBase64String(sha.ComputeHash(s)); }
    static EmployeeRecord Find(WorkshopData w,string code) { foreach(var e in w.ListNhanVien) if(e.SDB.TrimStart('0')==code.TrimStart('0')) return e; throw new Exception("Missing synthetic/test source identity "+code); }
    static void Roles()
    {
        foreach(string role in new[] {"QĐ","PQĐ","QĐ PX","Quản đốc","Phó Quản đốc","P.Quản đốc","KTV","KTV cơ khí","KTV CN","Kỹ thuật viên","TK","T.KÊ","Thống kê"}) Check(EmployeeClassification.IsManagementRole(role),"Eligible role "+role);
        foreach(string role in new[] {"TT","TRC","Tr.ca","P.ca","TTR","TTr","TP","Tổ trưởng","T. trưởng","Tổ phó","Trưởng ca","Phó trưởng ca","Phó ca","Phó ca - nghiền bi","CN","CNKT","ATV","","KTV kiêm Tổ trưởng"}) Check(!EmployeeClassification.IsManagementRole(role),"Direct role "+role);
        var w=new WorkshopData { LuongSPQL=100000 };
        var q=new EmployeeRecord {ChucDanh="QĐ",CongSP=26,HeSoCoDinh=2,XepLoai="A"};
        var direct=new EmployeeRecord {ChucDanh="TRC",LaQuanLy=true,QuanLyTinhMoi=true,CongSP=26,HeSoCoDinh=9,XepLoai="A",LuongSanPham=1234,BoPhan="Quản lý"};
        w.ListNhanVien.Add(q); w.ListNhanVien.Add(direct); ManagementPayroll.Recalculate(w);
        Check(q.LaQuanLy && q.LuongSanPham==100000,"Only valid title receives the management fund");
        Check(!direct.LaQuanLy && !direct.QuanLyTinhMoi && direct.LuongSanPham==1234,"Legacy flag/team cannot promote shift leader or reapply management pay");
        Check(w.LDQuanLy==1 && w.CongQL==26,"Counts and management attendance use eligible roster");
        q.BoPhan=direct.BoPhan="Cùng tổ"; var groups=w.GetToSummaries(); Check(groups.Count==2,"Same team splits management/direct summaries");
    }
    static void Synthetic(string dir)
    {
        string f=Path.Combine(dir,"duplicates.xlsx"); Fixture(f); var all=ExcelService.LoadWorkshopsFromExcel(f); var steel=Find(all[2],"171"); var stale=Find(all[4],"171");
        Check(!steel.LaQuanLy && steel.ChucDanh=="TT","Employee under management heading still classified by title");
        Check(steel.TongHopCongRow==6 && stale.TongHopCongRow==6,"Matching full identity in both workshop copies");
        steel.CoBangCong=true; steel.CongK3=4; steel.CoAnCa=true; steel.AnCaNhaAn=7;steel.AnCaTienMat=3;steel.AnCaDonVi=10;steel.AnCaNguon="DV";
        steel.CoBDDH=true; steel.BDDH_M1=3;steel.BDDH_M2=2;steel.BDDH_M3=1;steel.TienBDDH=105000;
        ExcelService.SaveWorkshopsToExcel(f,all);
        Check(Cell(f,"Tổng hợp công","L6")==26 && Cell(f,"Tổng hợp công","O6")==3,"Inactive later workshop cannot erase working days or leave");
        Check(Cell(f,"Tổng hợp công","R6")==15 && Cell(f,"Tổng hợp công","P6")==4 && Cell(f,"Tổng hợp công","Q6")==5,"Independent time columns P/Q preserved in total");
        Check(Cell(f,"Tổng hợp công","S6")==4,"Imported third shift saved to its dedicated column");
        all=ExcelService.LoadWorkshopsFromExcel(f);steel=Find(all[2],"171");stale=Find(all[4],"171");
        Check(steel.CongSP==26 && steel.CoBangCong && steel.CongK3==4,"Attendance metadata and K3 survive reload for direct employee");
        Check(steel.CoAnCa && steel.AnCaNhaAn==7 && steel.AnCaTienMat==3 && steel.AnCaNguon=="DV","Meal source and components survive reload");
        Check(steel.CoBDDH && steel.BDDH_M1==3 && steel.BDDH_M2==2 && steel.BDDH_M3==1,"Hazardous-work components survive reload");
        // Conflicting active workshop data must not partially overwrite the workbook.
        stale.CongSP=7; string before=Hash(f); bool rejected=false;
        try {ExcelService.SaveWorkshopsToExcel(f,all);} catch(ArgumentException) {rejected=true;}
        Check(rejected && Hash(f)==before,"Conflict leaves original workbook byte-for-byte intact");
        Check(Directory.GetFiles(dir,".emga-save-*.tmp").Length==0,"Failed staging file cleaned up");
        stale.CongSP=0; steel.CongSP=steel.CongLePhep=steel.CongHocHop=steel.CongVHTT=steel.CongBHXH=0;steel.LuongSanPham=0;
        ExcelService.SaveWorkshopsToExcel(f,all); Check(Cell(f,"Tổng hợp công","L6")==0 && Cell(f,"Tổng hợp công","W6")==0,"Intentional clear to zero is saved, not ignored");
        string names=Path.Combine(dir,"same-code-different-name.xlsx"); Fixture(names,true);var distinct=ExcelService.LoadWorkshopsFromExcel(names);
        Check(Find(distinct[2],"171").TongHopCongRow==6 && Find(distinct[4],"171").TongHopCongRow==7,"Same code, different name never maps to the first person");
        ExcelService.SaveWorkshopsToExcel(names,distinct); Check(Cell(names,"Tổng hợp công","L6")==26 && Cell(names,"Tổng hợp công","L7")==13,"Unchanged empty roster preserves another person's existing global work");
        // Poison the old saved boolean and salary; the original TT title must win.
        using(var z=ZipFile.Open(f,ZipArchiveMode.Update)) { var map=ExcelService.GetSheetMap(z);var doc=ExcelService.LoadSheetXml(z,map["Nhan su bo sung"]); var ss=ExcelService.LoadSharedStrings(z);int flag=0,paid=0,role=0;
            for(int c=1;c<100;c++){string h=ExcelService.GetCellString(doc,ss,TimesheetImporter.ColName(c)+"1");if(h=="QuanLyTinhMoi")flag=c;if(h=="LuongSanPham")paid=c;if(h=="LaQuanLy")role=c;}
            Check(flag>0 && paid>0 && role>0,"Legacy metadata fields found");
            ExcelService.SetCellInlineString(doc,TimesheetImporter.ColName(flag)+"2","True",false);ExcelService.SetCellInlineString(doc,TimesheetImporter.ColName(role)+"2","True",false);ExcelService.SetCellNumeric(doc,TimesheetImporter.ColName(paid)+"2",999999,false);z.GetEntry(map["Nhan su bo sung"]).Delete(); using(var stream=z.CreateEntry(map["Nhan su bo sung"]).Open()) doc.Save(stream);
        }
        var fixedData=ExcelService.LoadWorkshopsFromExcel(f);var fixedWorker=Find(fixedData[2],"171");Check(!fixedWorker.LaQuanLy && !fixedWorker.QuanLyTinhMoi && fixedWorker.LuongSanPham!=999999,"Old applied-manager snapshot cannot restore direct worker's wrong classification or pay");
    }
    static void Workbook(string source,string dir)
    {
        string file=Path.Combine(dir,"user-copy.xlsx");File.Copy(source,file,true);string sourceHash=Hash(source);var all=ExcelService.LoadWorkshopsFromExcel(file);
        foreach(var w in all) foreach(var e in w.ListNhanVien) Check(e.LaQuanLy==EmployeeClassification.IsManagementRole(e.ChucDanh),"Imported source classification matches job policy");
        var a=Find(all[2],"171");var b=Find(all[2],"177");double ca=a.CongSP,cb=b.CongSP,la=a.CongLePhep,lb=b.CongLePhep;
        Check(ca>0 && cb>0 && Find(all[4],"171").CongSP==0 && Find(all[4],"177").CongSP==0,"Reproduced active Steel and empty Khai tuyen duplicates");
        for(int i=0;i<3;i++) {
            ExcelService.SaveWorkshopsToExcel(file,all);var read=ExcelService.LoadWorkshopsFromExcel(file);a=Find(read[2],"171");b=Find(read[2],"177");
            Check(a.CongSP==ca && b.CongSP==cb && a.CongLePhep==la && b.CongLePhep==lb,"Both Steel employees retain original work and leave after save/reload");
            Check(Cell(file,"Tổng hợp công","L"+a.TongHopCongRow)==ca && Cell(file,"Tổng hợp công","L"+b.TongHopCongRow)==cb,"Global work not overwritten by later inactive department");
            using(var z=ZipFile.OpenRead(file)) { var book=PayrollFormulaBook.Load(z);
                // Isolate attendance lookup from unrelated pre-existing #N/A code cells in this source.
                foreach(var cell in book.Cells.Values) if(cell.Sheet=="Tổng hợp công" && System.Text.RegularExpressions.Regex.IsMatch(cell.Address,@"^C\d+$")) { if(cell.IsError) book.Set(cell.Sheet,cell.Address,0); else if(cell.IsNumber) book.Set(cell.Sheet,cell.Address,cell.Number); }
                Check(book.Evaluate("thép","K"+a.ExcelRow)==ca && book.Evaluate("thép","K"+b.ExcelRow)==cb,"Actual INDEX/MATCH attendance formulas also recompute correctly");
                Check(book.Evaluate("NT Thép","D10")==read[2].CongQL && book.Evaluate("NT Thép","D16")==read[2].LDQuanLy,"Excel management day/count summaries exclude direct titles on recalculation"); }
            all=read;
        }
        Check(Hash(source)==sourceHash,"Original uploaded workbook untouched");
    }
    static int Main(string[] args)
    {
        string dir=Path.Combine(Path.GetTempPath(),"classification-save-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
        try {Roles();Synthetic(dir);if(args.Length>0)Workbook(args[0],dir);} finally{Directory.Delete(dir,true);}
        Console.WriteLine("PASS: "+checks+" classification and lossless Excel checks.");return 0;
    }
}
