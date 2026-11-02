using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using CiscoLDTLApp;
class TimesheetLayoutTests
{
    static int checks;
    static void Check(bool ok,string text){checks++;if(!ok)throw new Exception(text);}
    static void Eq(double a,double b,string text){Check(Math.Abs(a-b)<0.000001,text);}
    static string Hash(string f){using(var sha=SHA256.Create())using(var s=File.OpenRead(f))return Convert.ToBase64String(sha.ComputeHash(s));}
    static XmlDocument Doc(){var d=new XmlDocument();d.LoadXml("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>");return d;}
    static void Text(XmlDocument d,string a,string t){ExcelService.SetCellInlineString(d,a,t,false);}
    static void Num(XmlDocument d,string a,double v){ExcelService.SetCellNumeric(d,a,v,false);}
    static void Entry(ZipArchive z,string p,string s){using(var w=new StreamWriter(z.CreateEntry(p).Open()))w.Write(s);}
    static void Workbook(string file,Dictionary<string,XmlDocument> sheets)
    {
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)){
            Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/></Types>");
            Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Target='xl/workbook.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'/></Relationships>");
            string wb="",rels="";int i=0;foreach(var p in sheets){i++;wb+="<sheet name='"+p.Key+"' sheetId='"+i+"' r:id='r"+i+"'/>";rels+="<Relationship Id='r"+i+"' Target='worksheets/sheet"+i+".xml'/>";using(var st=z.CreateEntry("xl/worksheets/sheet"+i+".xml").Open())p.Value.Save(st);}
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets>"+wb+"</sheets><calcPr/></workbook>");
            Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"+rels+"</Relationships>");
        }
    }
    static XmlDocument FixtureSheet(bool points,string heading="Tổng công",double work=24)
    {
        var d=Doc();Text(d,"K2",points?"Bảng chấm điểm":"Bảng chấm công");Text(d,"B6","Họ và tên");Text(d,"C6","SBD");Text(d,"D6","Chức vụ");Text(d,"E6","Ngày trong tháng");
        var merges=d.CreateElement("mergeCells",d.DocumentElement.NamespaceURI);var merge=d.CreateElement("mergeCell",d.DocumentElement.NamespaceURI);merge.SetAttribute("ref","E6:AI6");merges.AppendChild(merge);d.DocumentElement.AppendChild(merges);
        if(points){Text(d,"AJ6","Tổng điểm");Text(d,"AK6","Xếp loại");Text(d,"AL6","Hệ số");Text(d,"AM6","Tổng điểm chia lương");}
        else{Text(d,"AJ6","Quy ra công");Text(d,"AJ7",heading);Text(d,"AK7","Số công hưởng lương thời gian");Text(d,"AL7","Công Văn hóa thể thao");Text(d,"AM7","Công nghỉ việc hưởng 100% lương");Text(d,"AP7","Công K3");Text(d,"AQ7","Số công hưởng BHXH");Text(d,"AR6","Tổng cộng");}
        Text(d,"B10","Nhân sự giả");Text(d,"C10","901");Text(d,"D10","CN");
        if(points){Num(d,"AJ10",100);Text(d,"AK10","B");Num(d,"AL10",0.95);Num(d,"AM10",95);}
        else{Num(d,"AJ10",work);Num(d,"AK10",1);Num(d,"AL10",2);Num(d,"AM10",3);Num(d,"AP10",8);Num(d,"AQ10",4);Num(d,"AR10",31);}
        return d;
    }
    static EmployeeRecord Fake(){return new EmployeeRecord {HoTen="Nhân sự giả",SDB="901",ChucDanh="CN",CongSP=99,DiemSanXuat=7,HeSoHuong=7,XepLoai="A"};}
    static void Synthetic(string dir)
    {
        string f=Path.Combine(dir,"input.xlsx");var w=WorkshopData.CreateDefaultWorkshops()[1];var e=Fake();w.ListNhanVien.Add(e);
        Workbook(f,new Dictionary<string,XmlDocument>{{"Bảng công",FixtureSheet(false)},{"Bảng điểm",FixtureSheet(true)}});string log=TimesheetImporter.Import(f,w);
        Eq(e.CongSP,24,"Merged generic Total work is productive work, not 31 calendar days");Eq(e.CongHocHop,1,"Time wages mapped");Eq(e.CongVHTT,2,"Sport mapped");Eq(e.CongLePhep,3,"Paid leave mapped");Eq(e.CongBHXH,4,"BHXH mapped");Eq(e.CongK3,8,"K3 mapped");Eq(e.DiemSanXuat,100,"Generic Total points imported");Eq(e.HeSoHuong,95,"Explicit graded points preserved, not recalculated as B=90 percent");Check(log.Contains("khớp Tên+SDB: 1"),"Layout import reports successful match");
        w.RecalculateFromEmployees(true);Eq(e.HeSoHuong,95,"Source weight survives later recalculation");
        string master=Path.Combine(dir,"master.xlsx");var emp=Doc();Text(emp,"C5",e.HoTen);Text(emp,"D5",e.SDB);Text(emp,"E5","CN");Num(emp,"G5",5200000);Num(emp,"K5",99);Num(emp,"W5",7);Num(emp,"Z5",7);var company=Doc();Text(company,"C6",e.SDB);Text(company,"D6",e.HoTen);Num(company,"L6",99);Num(company,"U6",7);
        Workbook(master,new Dictionary<string,XmlDocument>{{"Gang",emp},{"Tổng hợp công",company}});var all=ExcelService.LoadWorkshopsFromExcel(master);TimesheetImporter.Import(f,all[1]);
        for(int pass=0;pass<2;pass++){ExcelService.SaveWorkshopsToExcel(master,all);all=ExcelService.LoadWorkshopsFromExcel(master);e=all[1].ListNhanVien[0];all[1].RecalculateFromEmployees(true);Eq(e.CongSP,24,"Productive days persist after reload");Eq(e.HeSoHuong,95,"Explicit graded weight persists after reload and recalculation");using(var z=ZipFile.OpenRead(master)){var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);var d=ExcelService.LoadSheetXml(z,map["Gang"]);Eq(ExcelService.GetCellDouble(d,ss,"Z5",-1),95,"Native worksheet weight agrees with source, not stale formula");}}
        File.Delete(f);Workbook(f,new Dictionary<string,XmlDocument>{{"Bảng công",FixtureSheet(false,"Tổng công",0)},{"Bảng điểm",FixtureSheet(true)}});using(var z=ZipFile.Open(f,ZipArchiveMode.Update)){var map=ExcelService.GetSheetMap(z);var d=ExcelService.LoadSheetXml(z,map["Bảng điểm"]);Num(d,"AJ10",0);Num(d,"AM10",0);var entry=z.GetEntry(map["Bảng điểm"]);entry.Delete();using(var st=z.CreateEntry(map["Bảng điểm"]).Open())d.Save(st);}TimesheetImporter.Import(f,all[1]);Eq(all[1].ListNhanVien[0].CongSP,0,"A genuine numeric zero still clears days");Eq(all[1].ListNhanVien[0].HeSoHuong,0,"Genuine zero clears stale points");ExcelService.SaveWorkshopsToExcel(master,all);all=ExcelService.LoadWorkshopsFromExcel(master);Eq(all[1].ListNhanVien[0].DiemSanXuat,0,"Zero source point survives master fallback and reload");
        using(var z=ZipFile.OpenRead(master)){var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);Eq(ExcelService.GetCellDouble(ExcelService.LoadSheetXml(z,map["Tổng hợp công"]),ss,"U6",-1),0,"Zero source points also clear company master points");}
        File.Delete(f);Workbook(f,new Dictionary<string,XmlDocument>{{"Bảng công đúng",FixtureSheet(false)},{"Bảng công lỗi",FixtureSheet(false,"Cột lạ chưa hỗ trợ")}});w=all[1];e=w.ListNhanVien[0];e.CongSP=99;bool blocked=false;try{TimesheetImporter.Import(f,w);}catch(ArgumentException ex){blocked=ex.Message.Contains("cột công");}Check(blocked && e.CongSP==99,"Missing required work column aborts before resetting existing days");
        File.Delete(f);var invalid=FixtureSheet(false);Text(invalid,"AJ10","#N/A");Workbook(f,new Dictionary<string,XmlDocument>{{"Bảng công",invalid}});log=TimesheetImporter.Import(f,w);Check(e.CongSP==99 && log.Contains("không có số công hợp lệ"),"Invalid numeric work cell preserves existing attendance and reports the source cell");
    }
    static string Id(string code,string name){long n;if(long.TryParse(code,out n))code=n.ToString(CultureInfo.InvariantCulture);return code.Trim()+"|"+System.Text.RegularExpressions.Regex.Replace(name.Trim().Normalize(),@"\s+"," ").ToUpperInvariant();}
    static void Real(string source,string dir)
    {
        string hash=Hash(source);var all=WorkshopData.CreateDefaultWorkshops();var w=all[1];var expected=new Dictionary<string,int>();
        using(var z=ZipFile.OpenRead(source)){
            var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);var days=ExcelService.LoadSheetXml(z,map["Bảng công"]);
            foreach(XmlElement row in days.GetElementsByTagName("row","http://schemas.openxmlformats.org/spreadsheetml/2006/main")){
                int r;long code;if(!int.TryParse(row.GetAttribute("r"),out r)||r<=8)continue;string sdb=ExcelService.GetCellString(days,ss,"C"+r),name=ExcelService.GetCellString(days,ss,"B"+r);if(!long.TryParse(sdb,out code)||code<=0||name=="")continue;
                string key=Id(sdb,name);if(expected.ContainsKey(key))throw new Exception("Source has repeated identity; fixture requires explicit split-row expectations");expected[key]=r;
                var e=new EmployeeRecord {HoTen=name.Trim(),SDB=sdb.Trim(),ChucDanh=ExcelService.GetCellString(days,ss,"D"+r),CongSP=99,DiemSanXuat=999,HeSoHuong=999,XepLoai="A"};EmployeeClassification.Apply(e);w.ListNhanVien.Add(e);
            }
            int prompts=0;string log=TimesheetImporter.Import(source,w,all,e=>{prompts++;return null;});Check(prompts==0 && w.ListNhanVien.Count==expected.Count,"All source employees resolve without fake approvals or duplicate profiles");int positive=0;
            foreach(var e in w.ListNhanVien){int r=expected[Id(e.SDB,e.HoTen)];double actual=ExcelService.GetCellDouble(days,ss,"AJ"+r,0);Eq(e.CongSP,actual,"Actual productive work equals source AJ");if(actual>0)positive++;
                Eq(e.CongHocHop,ExcelService.GetCellDouble(days,ss,"AK"+r,0),"Actual time-wage days equal AK");Eq(e.CongVHTT,ExcelService.GetCellDouble(days,ss,"AL"+r,0),"Actual sport days equal AL");Eq(e.CongLePhep,ExcelService.GetCellDouble(days,ss,"AM"+r,0),"Actual paid leave equals AM");Eq(e.CongK3,ExcelService.GetCellDouble(days,ss,"AP"+r,0),"Actual K3 equals AP");Eq(e.CongBHXH,ExcelService.GetCellDouble(days,ss,"AQ"+r,0),"Actual BHXH equals AQ");}
            Check(positive>0 && !log.Contains("chưa được nạp"),"Actual upload no longer imports every productive day as zero");
            foreach(string sheet in new[]{"Quyết toán ăn ca","Quyết toán độc hại","Bảng điểm"}){
                var d=ExcelService.LoadSheetXml(z,map[sheet]);var byId=new Dictionary<string,EmployeeRecord>();foreach(var e in w.ListNhanVien)byId[Id(e.SDB,e.HoTen)]=e;
                foreach(XmlElement row in d.GetElementsByTagName("row","http://schemas.openxmlformats.org/spreadsheetml/2006/main")){
                    int r;if(!int.TryParse(row.GetAttribute("r"),out r)||r<=8)continue;string key=Id(ExcelService.GetCellString(d,ss,"C"+r),ExcelService.GetCellString(d,ss,"B"+r));EmployeeRecord e;if(!byId.TryGetValue(key,out e))continue;
                    if(sheet=="Quyết toán ăn ca"){Eq(e.AnCaNhaAn,ExcelService.GetCellDouble(d,ss,"AL"+r,0),"Actual meal-house count matches AL");Eq(e.AnCaTienMat,ExcelService.GetCellDouble(d,ss,"AN"+r,0),"Actual cash meals match AN");}
                    else if(sheet=="Quyết toán độc hại"){Eq(e.BDDH_M1,ExcelService.GetCellDouble(d,ss,"AK"+r,0),"Actual M1 matches AK");Eq(e.BDDH_M2,ExcelService.GetCellDouble(d,ss,"AL"+r,0),"Actual M2 matches AL");Eq(e.BDDH_M3,ExcelService.GetCellDouble(d,ss,"AM"+r,0),"Actual M3 matches AM");}
                    else if(!e.LaQuanLy){Eq(e.DiemSanXuat,ExcelService.GetCellDouble(d,ss,"AJ"+r,0),"Actual production points match AJ");Eq(e.HeSoHuong,ExcelService.GetCellDouble(d,ss,"AM"+r,0),"Actual graded points match AM");}
                }
            }
            TimesheetImporter.Import(source,w,all);foreach(var e in w.ListNhanVien)Eq(e.CongSP,ExcelService.GetCellDouble(days,ss,"AJ"+expected[Id(e.SDB,e.HoTen)],0),"Repeated real import is stable");
            string master=Path.Combine(dir,"actual-import-copy.xlsx");var staffDoc=Doc();var masterDoc=Doc();int index=0;
            var pointsBefore=new Dictionary<string,double>();var weightsBefore=new Dictionary<string,double>();
            foreach(var e in w.ListNhanVien){e.ExcelRow=5+index;e.TongHopCongRow=6+index;int a=e.ExcelRow,b=e.TongHopCongRow;index++;
                Text(staffDoc,"C"+a,e.HoTen);Text(staffDoc,"D"+a,e.SDB);Text(staffDoc,"E"+a,e.ChucDanh);Num(staffDoc,"K"+a,99);Num(staffDoc,"W"+a,999);Num(staffDoc,"Z"+a,999);
                Text(masterDoc,"C"+b,e.SDB);Text(masterDoc,"D"+b,e.HoTen);Num(masterDoc,"L"+b,99);Num(masterDoc,"U"+b,999);
                pointsBefore[Id(e.SDB,e.HoTen)]=e.DiemSanXuat;weightsBefore[Id(e.SDB,e.HoTen)]=e.HeSoHuong;
            }
            Workbook(master,new Dictionary<string,XmlDocument>{{"Gang",staffDoc},{"Tổng hợp công",masterDoc}});
            for(int pass=0;pass<2;pass++){
                ExcelService.SaveWorkshopsToExcel(master,all);all=ExcelService.LoadWorkshopsFromExcel(master);w=all[1];Check(w.ListNhanVien.Count==expected.Count,"All actual-import identities persist to a separate test master");w.RecalculateFromEmployees(true);
                foreach(var e in w.ListNhanVien){string key=Id(e.SDB,e.HoTen);int r=expected[key];Eq(e.CongSP,ExcelService.GetCellDouble(days,ss,"AJ"+r,0),"Real-source work survives save/reload");if(!e.LaQuanLy){Eq(e.DiemSanXuat,pointsBefore[key],"Real-source production points survive save/reload");Eq(e.HeSoHuong,weightsBefore[key],"Real-source graded points survive save/reload");}Check(e.CoAnCa && e.CoBDDH,"Meal/BDDH import metadata survives real-source save/reload");}
            }
            Console.WriteLine("Actual upload verified: "+expected.Count+" staff rows; "+positive+" have positive productive work.");
        }
        Check(Hash(source)==hash,"Original uploaded workbook unchanged");
    }
    static int Main(string[] args){string dir=Path.Combine(Path.GetTempPath(),"layout-import-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);try{Synthetic(dir);if(args.Length>0)Real(args[0],dir);}finally{Directory.Delete(dir,true);}Console.WriteLine("PASS: "+checks+" timesheet layout checks.");return 0;}
}
