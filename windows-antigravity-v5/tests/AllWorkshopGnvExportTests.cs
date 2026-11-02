using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Xml;
using CiscoLDTLApp;
class AllWorkshopGnvExportTests
{
    class Spec {public string Code,Sheet,Amount="J",Rate="I",Total,QuantityTotal;public int[] Rows,Protected;public int Boundary;}
    static int checks;
    static int[] Range(int a,int b){return Enumerable.Range(a,b-a+1).ToArray();}
    // Independent expectations from the source templates, not the application's layout table.
    static Spec[] Specs={
        new Spec {Code="TK",Sheet="NT thiêu kết",Rows=Range(80,90),Protected=new[]{79},Boundary=91,Total="J79",QuantityTotal="D79"},
        new Spec {Code="LG",Sheet="NT gang",Rows=new[]{61,62}.Concat(Range(64,93)).Concat(Range(95,100)).Concat(new[]{102,103}).ToArray(),Protected=new[]{60,63,94,101},Boundary=104,Total="J59",QuantityTotal="D59"},
        new Spec {Code="LT",Sheet="NT Thép",Rows=Range(60,97),Protected=new int[0],Boundary=98,Total="J59",QuantityTotal="D59"},
        new Spec {Code="CDNL",Sheet="NT CĐNL",Rows=Range(372,393),Protected=new[]{368,369,371,394,415,416},Boundary=417,Amount="I",Rate="H",Total="I371",QuantityTotal="D371"},
        new Spec {Code="KT",Sheet="NT. K. tuyển ",Rows=Range(60,63).Concat(Range(65,79)).Concat(new[]{82,83}).Concat(Range(85,87)).Concat(Range(89,94)).Concat(Range(96,101)).ToArray(),Protected=new[]{59,64,80,81,84,88,95},Boundary=102,Total="J52",QuantityTotal="D52"}
    };
    static void Check(bool ok,string message){checks++;if(!ok)throw new Exception(message);}
    static string Hash(string f){using(var sha=SHA256.Create())using(var s=File.OpenRead(f))return Convert.ToBase64String(sha.ComputeHash(s));}
    static void Entry(ZipArchive z,string path,string text){using(var w=new StreamWriter(z.CreateEntry(path).Open()))w.Write(text);}
    static void Put(ZipArchive z,string path,XmlDocument d){if(z.Mode==ZipArchiveMode.Update){var e=z.GetEntry(path);if(e!=null)e.Delete();}using(var s=z.CreateEntry(path).Open())d.Save(s);}
    static void Fixture(string file)
    {
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)){
            Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/></Types>");
            Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='xl/workbook.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'/></Relationships>");
            string sheets="",rels="";
            for(int i=0;i<Specs.Length;i++){
                var p=Specs[i];sheets+="<sheet name='"+p.Sheet+"' sheetId='"+(i+1)+"' r:id='r"+i+"'/>";rels+="<Relationship Id='r"+i+"' Target='worksheets/sheet"+i+".xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'/>";
                var d=new XmlDocument();d.LoadXml("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>");
                foreach(int r in p.Protected)ExcelService.SetCellInlineString(d,"B"+r,"Tổng phụ giả "+r,false);
                ExcelService.SetCellInlineString(d,"B"+p.Boundary,"Tiền lương chế độ",false);
                int last=p.Rows.Last();ExcelService.SetCellInlineString(d,"B"+last,"GNV cuối mẫu "+p.Code,false);ExcelService.SetCellInlineString(d,"C"+last,"Công",false);
                ExcelService.SetCellNumeric(d,"D"+last,p.Code=="CDNL"?-2:2,false);ExcelService.SetCellNumeric(d,p.Rate+last,1000,false);ExcelService.SetCellNumeric(d,p.Amount+last,p.Code=="CDNL"?-2000:2000,false);
                Put(z,"xl/worksheets/sheet"+i+".xml",d);
            }
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets>"+sheets+"</sheets><calcPr/></workbook>");
            Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"+rels+"</Relationships>");
        }
    }
    static ArisingWorkItem Item(string code,int i,int row){return new ArisingWorkItem(i,"GNV giả "+code+" "+i,"Công",i,1000,0,"",row){TuTinhThanhTien=true};}
    static WorkshopData Find(List<WorkshopData> all,string code){return all.First(w=>w.MaPX==code);}
    static string Cell(XmlDocument d,List<string> ss,string a){return ExcelService.GetCellString(d,ss,a);}
    static double Num(XmlDocument d,List<string> ss,string a){return ExcelService.GetCellDouble(d,ss,a,-1);}
    static void Run(string file,bool checkWord)
    {
        var all=ExcelService.LoadWorkshopsFromExcel(file);
        var originalHeaders=new Dictionary<string,string>();
        using(var z=ZipFile.OpenRead(file)){
            var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);
            foreach(var p in Specs){var d=ExcelService.LoadSheetXml(z,map[p.Sheet]);var w=Find(all,p.Code);
                foreach(int r in p.Protected.Concat(new[]{p.Boundary}))originalHeaders[p.Code+"|"+r]=Cell(d,ss,"B"+r);
                foreach(var item in w.ListGNV)Check(p.Rows.Contains(item.ExcelRow),"Reader includes detail only: "+p.Code);
                foreach(int r in p.Rows){string name=Cell(d,ss,"B"+r).Trim();double qty=Num(d,ss,"D"+r),money=Num(d,ss,p.Amount+r);
                    bool has=(qty!=-1 && qty!=0 || money!=-1 && money!=0) && name!="" && !name.StartsWith("#");
                    if(has)Check(w.ListGNV.Any(g=>g.ExcelRow==r),"Every nonzero native detail, including negatives and tail rows, is loaded: "+p.Code+" "+r);
                }
            }
        }
        // Fill every permitted native slot, with highest rows first and explicit group anchors.
        foreach(var p in Specs){var w=Find(all,p.Code);w.GnvDaChinhSua=true;w.ListGNV.Clear();int i=0;
            foreach(int r in p.Rows.Reverse()){var item=Item(p.Code,++i,r);if(p.Code=="LG" && r<63 || p.Code=="CDNL" && r==393)item.KhoiLuong=-i;w.ListGNV.Add(item);}w.RecalculateDerived();}
        ExcelService.SaveWorkshopsToExcel(file,all);
        using(var z=ZipFile.OpenRead(file)){
            var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);var book=PayrollFormulaBook.Load(z);
            foreach(var p in Specs){var w=Find(all,p.Code);var d=ExcelService.LoadSheetXml(z,map[p.Sheet]);double total=0,work=0;
                foreach(var g in w.ListGNV){Check(Cell(d,ss,"B"+g.ExcelRow)==g.NoiDung && Num(d,ss,"D"+g.ExcelRow)==g.KhoiLuong && Num(d,ss,p.Amount+g.ExcelRow)==g.ThanhTien,"All native details written: "+p.Code+" "+g.ExcelRow);
                    if(!(p.Code=="LG" && g.ExcelRow<63)){total+=g.ThanhTien;work+=g.KhoiLuong;}}
                Check(book.Evaluate(p.Sheet,p.Total)==total && book.Evaluate(p.Sheet,p.QuantityTotal)==work,"Native formulas total leaf rows exactly once: "+p.Code);
                if(p.Code!="CDNL")Check(w.LuongCVKhac==total,"App fund agrees with native payroll total: "+p.Code);
                if(p.Code=="LG")Check(book.Evaluate(p.Sheet,"J60")==w.ListGNV.Where(g=>g.ExcelRow<63).Sum(g=>g.ThanhTien),"Gang external payment subtotal separate from native J59=J63 payroll fund");
                if(p.Code=="KT")Check(book.Evaluate(p.Sheet,"J80")==w.ListGNV.Where(g=>new[]{82,83,85,86}.Contains(g.ExcelRow)).Sum(g=>g.ThanhTien),"Khai tuyen parent totals do not duplicate their children in J52");
                foreach(int r in p.Protected.Concat(new[]{p.Boundary}))Check(Cell(d,ss,"B"+r)==originalHeaders[p.Code+"|"+r],"Section and subtotal labels preserved: "+p.Code+" "+r);
            }
        }
        if(checkWord) foreach(var p in Specs) {
            var w=Find(all,p.Code);string preview=WordReportService.GeneratePreviewText(w),path=Path.Combine(Path.GetDirectoryName(file),p.Code+"-full.docx");
            WordReportService.ExportWordDocx(w,path);string xml;using(var z=ZipFile.OpenRead(path))using(var r=new StreamReader(z.GetEntry("word/document.xml").Open()))xml=r.ReadToEnd();
            foreach(var item in w.ListGNV) {Check(preview.Contains(item.NoiDung),"Preview contains every long-list item: "+p.Code);Check(xml.Contains(item.NoiDung),"DOCX contains every long-list item: "+p.Code);}
        }
        all=ExcelService.LoadWorkshopsFromExcel(file);foreach(var p in Specs)Check(Find(all,p.Code).ListGNV.Count==p.Rows.Length,"Full capacity roundtrip: "+p.Code);
        // Overflows at any unit must leave all five native sheets and metadata byte-identical.
        foreach(var p in Specs){var w=Find(all,p.Code);w.ListGNV.Add(Item(p.Code,999,0));w.RecalculateDerived();string hash=Hash(file);bool blocked=false;
            try{ExcelService.SaveWorkshopsToExcel(file,all);}catch(ArgumentException ex){blocked=ex.Message.Contains("GNV") && ex.Message.Contains(p.Rows.Length.ToString());}
            Check(blocked && Hash(file)==hash,"Overflow is explicit and cannot partially save: "+p.Code);w.ListGNV.RemoveAt(w.ListGNV.Count-1);w.RecalculateDerived();}
        // New, duplicate and external anchors cannot overwrite later native records.
        foreach(var p in Specs){var w=Find(all,p.Code);w.ListGNV.Clear();int last=p.Rows.Last();w.ListGNV.Add(Item(p.Code,1,last));w.ListGNV.Add(Item(p.Code,2,0));w.ListGNV.Add(Item(p.Code,3,last));w.ListGNV.Add(Item(p.Code,4,9000));w.ListGNV.Add(Item(p.Code,5,0));w.ListGNV[1].KhoiLuong=0;w.ListGNV[2].KhoiLuong=-3;w.ListGNV[4].TuTinhThanhTien=false;w.ListGNV[4].ThanhTien=12345;w.RecalculateDerived();}
        ExcelService.SaveWorkshopsToExcel(file,all);
        using(var z=ZipFile.OpenRead(file)){var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);foreach(var p in Specs){var w=Find(all,p.Code);var d=ExcelService.LoadSheetXml(z,map[p.Sheet]);Check(w.ListGNV.Select(g=>g.ExcelRow).Distinct().Count()==5,"Allocated unique rows: "+p.Code);foreach(var g in w.ListGNV)Check(Cell(d,ss,"B"+g.ExcelRow)==g.NoiDung && Num(d,ss,p.Amount+g.ExcelRow)==g.ThanhTien,"New/zero/negative/custom-money item exported: "+p.Code);}}
        all=ExcelService.LoadWorkshopsFromExcel(file);foreach(var p in Specs){var w=Find(all,p.Code);Check(w.ListGNV.Count==5 && w.ListGNV[1].KhoiLuong==0 && w.ListGNV[4].ThanhTien==12345,"Snapshot matches actual native rows after reload: "+p.Code);w.ListGNV.RemoveAt(0);w.RecalculateDerived();}
        ExcelService.SaveWorkshopsToExcel(file,all);
        using(var z=ZipFile.OpenRead(file)){var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);foreach(var p in Specs){var d=ExcelService.LoadSheetXml(z,map[p.Sheet]);Check(Cell(d,ss,"B"+p.Rows.Last())=="" && Num(d,ss,p.Amount+p.Rows.Last())==0,"Deleted tail cleared: "+p.Code);}}
        foreach(var w in all){w.ListGNV.Clear();w.RecalculateDerived();}ExcelService.SaveWorkshopsToExcel(file,all);
        using(var z=ZipFile.OpenRead(file)){var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);var book=PayrollFormulaBook.Load(z);foreach(var p in Specs){var d=ExcelService.LoadSheetXml(z,map[p.Sheet]);foreach(int r in p.Rows)Check(Cell(d,ss,"B"+r)=="" && Num(d,ss,p.Amount+r)==0,"All deleted detail rows cleared: "+p.Code);Check(book.Evaluate(p.Sheet,p.Total)==0,"Native cleared total recalculates to zero: "+p.Code);}}
        Check(ExcelService.LoadWorkshopsFromExcel(file).All(w=>w.ListGNV.Count==0),"Deleted lists stay empty at all units");
    }
    static void GuardTests(string file)
    {
        Fixture(file);using(var z=ZipFile.Open(file,ZipArchiveMode.Update)){var map=ExcelService.GetSheetMap(z);var d=ExcelService.LoadSheetXml(z,map["NT CĐNL"]);ExcelService.SetCellNumeric(d,"E394",1,false);Put(z,map["NT CĐNL"],d);}
        bool blocked=false;try{ExcelService.LoadWorkshopsFromExcel(file);}catch(ArgumentException ex){blocked=ex.Message.Contains("394") && ex.Message.Contains("I371");}Check(blocked,"CDNL active work outside source GNV fund cannot be silently omitted or moved to another fund");
        File.Delete(file);Fixture(file);using(var z=ZipFile.Open(file,ZipArchiveMode.Update)){var map=ExcelService.GetSheetMap(z);var d=ExcelService.LoadSheetXml(z,map["NT gang"]);ExcelService.SetCellInlineString(d,"B104","GNV thêm do chèn dòng",false);Put(z,map["NT gang"],d);}
        blocked=false;try{ExcelService.LoadWorkshopsFromExcel(file);}catch(ArgumentException ex){blocked=ex.Message.Contains("vị trí");}Check(blocked,"Inserted-row template rejected instead of overwriting its moved statutory boundary");
        File.Delete(file);Fixture(file);using(var z=ZipFile.Open(file,ZipArchiveMode.Update)){var d=new XmlDocument();using(var st=z.GetEntry("xl/workbook.xml").Open())d.Load(st);foreach(XmlElement sh in d.GetElementsByTagName("sheet","http://schemas.openxmlformats.org/spreadsheetml/2006/main"))if(sh.GetAttribute("name")=="NT. K. tuyển ")sh.SetAttribute("name","NT. K. tuyển");Put(z,"xl/workbook.xml",d);}
        var all=ExcelService.LoadWorkshopsFromExcel(file);var kt=Find(all,"KT");Check(kt.ListGNV.Count==1,"Khai tuyen native sheet without trailing space is read");kt.GnvDaChinhSua=true;ExcelService.SaveWorkshopsToExcel(file,all);
        using(var z=ZipFile.OpenRead(file)){var map=ExcelService.GetSheetMap(z);var ss=ExcelService.LoadSharedStrings(z);Check(Cell(ExcelService.LoadSheetXml(z,map["NT. K. tuyển"]),ss,"B101")==kt.ListGNV[0].NoiDung,"Trimmed Khai tuyen sheet also receives native export");}
        File.Delete(file);Fixture(file);all=ExcelService.LoadWorkshopsFromExcel(file);var gang=Find(all,"LG");gang.GnvDaChinhSua=true;gang.ListGNV.Clear();for(int i=0;i<39;i++)gang.ListGNV.Add(Item("LG",i+1,0));gang.RecalculateDerived();string before=Hash(file);blocked=false;
        try{ExcelService.SaveWorkshopsToExcel(file,all);}catch(ArgumentException ex){blocked=ex.Message.Contains("nhóm thanh toán");}Check(blocked && Hash(file)==before,"New Gang items cannot spill into the two external-payment slots and change payroll classification");
        gang.ListGNV.Clear();gang.ListGNV.Add(Item("LG",1,61));gang.ListGNV.Add(Item("LG",2,61));blocked=false;
        try{ExcelService.SaveWorkshopsToExcel(file,all);}catch(ArgumentException ex){blocked=ex.Message.Contains("trùng");}Check(blocked && Hash(file)==before,"Duplicate external-payment anchors cannot silently move an expense into payroll");
        File.Delete(file);Fixture(file);using(var z=ZipFile.Open(file,ZipArchiveMode.Update)){var map=ExcelService.GetSheetMap(z);var d=ExcelService.LoadSheetXml(z,map["NT thiêu kết"]);ExcelService.SetCellNumeric(d,"D80",1,false);Put(z,map["NT thiêu kết"],d);}
        blocked=false;try{ExcelService.LoadWorkshopsFromExcel(file);}catch(ArgumentException ex){blocked=ex.Message.Contains("nội dung");}Check(blocked,"Nonzero work without usable description is reported instead of omitted");
        File.Delete(file);Fixture(file);using(var z=ZipFile.Open(file,ZipArchiveMode.Update)){
            string path=(string)typeof(ExcelService).GetMethod("EnsureWorksheet",System.Reflection.BindingFlags.Static|System.Reflection.BindingFlags.NonPublic).Invoke(null,new object[]{z,"GNV v5","xl/worksheets/oldGnv.xml"});
            var d=new XmlDocument();d.LoadXml("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>");
            foreach(int r in new[]{2,3}){ExcelService.SetCellInlineString(d,"A"+r,"LG",false);ExcelService.SetCellInlineString(d,"B"+r,"True",false);ExcelService.SetCellInlineString(d,"C"+r,r==2?"Tổng phụ giả 60":"GNV cuối mẫu LG",false);ExcelService.SetCellInlineString(d,"D"+r,"Công",false);ExcelService.SetCellNumeric(d,"E"+r,r==2?99:2,false);ExcelService.SetCellNumeric(d,"F"+r,1000,false);ExcelService.SetCellNumeric(d,"G"+r,r==2?99000:2000,false);ExcelService.SetCellNumeric(d,"H"+r,r==2?60:103,false);}Put(z,path,d);
        }
        all=ExcelService.LoadWorkshopsFromExcel(file);Check(Find(all,"LG").ListGNV.Count==1 && Find(all,"LG").ListGNV[0].ExcelRow==103,"Old snapshot subtotal is not restored as a duplicate payable detail");

    }
    static int Main(string[] args)
    {
        string dir=Path.Combine(Path.GetTempPath(),"all-gnv-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
        try{string f=Path.Combine(dir,"copy.xlsx");if(args.Length>0){string hash=Hash(args[0]);File.Copy(args[0],f);Run(f,false);Check(Hash(args[0])==hash,"User workbook untouched");}else{Fixture(f);Run(f,true);GuardTests(Path.Combine(dir,"guard.xlsx"));}Check(Directory.GetFiles(dir,".emga-save-*.tmp").Length==0,"No staged save files left after conflicts");}
        finally{Directory.Delete(dir,true);}Console.WriteLine("PASS: "+checks+" all-workshop native GNV checks.");return 0;
    }
}
