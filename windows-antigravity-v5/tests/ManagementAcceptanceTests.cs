using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Xml;
using CiscoLDTLApp;
class ManagementAcceptanceTests
{
    class Spec {public string Code,Sheet,Employee,Fund,Other,Total,All,Direct,DirectFund;public int Boundary;}
    static Spec[] Specs={
        new Spec{Code="TK",Sheet="NT thiêu kết",Employee="T kết",Fund="J96",Other="J114",Total="J95",All="J126",Direct="J56",DirectFund="J57",Boundary=91},
        new Spec{Code="LG",Sheet="NT gang",Employee="Gang",Fund="J109",Other="J140",Total="J108",All="J161",Direct="J47",DirectFund="J48",Boundary=104},
        new Spec{Code="LT",Sheet="NT Thép",Employee="thép",Fund="J103",Other="J130",Total="J102",All="J137",Direct="J34",DirectFund="J35",Boundary=98},
        new Spec{Code="CDNL",Sheet="NT CĐNL",Employee="CĐNL",Fund="I422",Other="I432",Total="I421",All="I440",Direct="I42",DirectFund="I43",Boundary=417},
        new Spec{Code="KT",Sheet="NT. K. tuyển ",Employee="K. tuyển",Fund="J107",Other="J128",Total="J106",All="J131",Direct="J38",DirectFund="J40",Boundary=102}
    };
    static int checks;
    static void Check(bool b,string message){checks++;if(!b)throw new Exception(message);}
    static XmlDocument Doc(){var d=new XmlDocument();d.LoadXml("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>");return d;}
    static void Num(XmlDocument d,string a,double n){ExcelService.SetCellNumeric(d,a,n,false);}
    static void Text(XmlDocument d,string a,string s){ExcelService.SetCellInlineString(d,a,s,false);}
    static void Formula(XmlDocument d,string a,string f,double value){Num(d,a,value);var ns=new XmlNamespaceManager(d.NameTable);ns.AddNamespace("s",d.DocumentElement.NamespaceURI);var c=d.SelectSingleNode("//s:c[@r='"+a+"']",ns);var fn=d.CreateElement("f",d.DocumentElement.NamespaceURI);fn.InnerText=f;c.PrependChild(fn);}
    static void Entry(ZipArchive z,string p,string t){using(var w=new StreamWriter(z.CreateEntry(p).Open()))w.Write(t);}
    static string Hash(string p){using(var sha=SHA256.Create())using(var s=File.OpenRead(p))return Convert.ToBase64String(sha.ComputeHash(s));}
    static void Fixture(string file)
    {
        var docs=new Dictionary<string,XmlDocument>();
        foreach(var p in Specs){var d=Doc();Text(d,"B"+p.Boundary,"Tiền lương chế độ");Formula(d,p.Fund,"1000000+0",1000000);Formula(d,p.Other,"200000+0",200000);Formula(d,p.Total,p.Fund+"+"+p.Other,1200000);Num(d,p.DirectFund,5000000);Formula(d,p.Direct,p.DirectFund,5000000);Formula(d,p.All,p.Direct+"+"+p.Total,6200000);docs[p.Sheet]=d;
            var emp=Doc();Text(emp,"C5","Quản lý giả "+p.Code);Text(emp,"D5",p.Code+"901");Text(emp,"E5","QĐ");Num(emp,"G5",5200000);Num(emp,"I5",2);Num(emp,"K5",26);Num(emp,"M5",1);Num(emp,"P5",1);Num(emp,"R5",1);Text(emp,p.Code=="CDNL"?"Z5":"X5","A");
            Text(emp,"C6","Tổ trưởng giả "+p.Code);Text(emp,"D6",p.Code+"902");Text(emp,"E6","TT");Num(emp,"G6",5200000);Num(emp,"K6",26);Num(emp,"AB6",555000);Num(emp,"AD6",10000);Num(emp,"AH6",0);Num(emp,"AI6",0);Num(emp,"AK6",0);
            if(p.Code!="CDNL"){
                Formula(emp,"AA5","AB5+AD5+AH5+AI5+AK5",1200000);var ns=new XmlNamespaceManager(emp.NameTable);ns.AddNamespace("s",emp.DocumentElement.NamespaceURI);var f=(XmlElement)emp.SelectSingleNode("//s:c[@r='AA5']/s:f",ns);f.SetAttribute("t","shared");f.SetAttribute("si","1");f.SetAttribute("ref","AA5:AA6");Formula(emp,"AA6","",565000);f=(XmlElement)emp.SelectSingleNode("//s:c[@r='AA6']/s:f",ns);f.SetAttribute("t","shared");f.SetAttribute("si","1");
            }
            docs[p.Employee]=emp;
        }
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)){
            Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/></Types>");
            Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Target='xl/workbook.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'/></Relationships>");string wb="",rels="";int i=0;
            foreach(var p in docs){i++;wb+="<sheet name='"+p.Key+"' sheetId='"+i+"' r:id='r"+i+"'/>";rels+="<Relationship Id='r"+i+"' Target='worksheets/sheet"+i+".xml'/>";using(var st=z.CreateEntry("xl/worksheets/sheet"+i+".xml").Open())p.Value.Save(st);}
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets>"+wb+"</sheets><calcPr/></workbook>");Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"+rels+"</Relationships>");
        }
    }
    static void Verify(string file,List<WorkshopData> all,bool personal,string only=null,bool neighbor=true)
    {
        using(var z=ZipFile.OpenRead(file)){
            var book=PayrollFormulaBook.Load(z);foreach(var p in Specs){if(only!=null && p.Code!=only)continue;var w=all.First(x=>x.MaPX==p.Code);double expected=w.LuongSPQL+w.LuongCheDoQL+w.DieuChinhLuongQL;
                Check(Math.Abs(book.Evaluate(p.Sheet,p.Total)-expected)<0.001,"NT management formula recomputes applied totals: "+p.Code);
                // Bind the separate direct-labor amount to isolate source defects outside this change.
                book.Set(p.Sheet,p.Direct,w.TongLuongTT);book.Reset();Check(Math.Abs(book.Evaluate(p.Sheet,p.All)-(w.TongLuongTT+expected))<0.001,"Whole workshop includes updated management exactly once: "+p.Code);
                if(personal)foreach(var e in w.ListNhanVien)if(e.LaQuanLy){string row=e.ExcelRow.ToString();Check(book.Evaluate(p.Employee,(p.Code=="CDNL"?"AI":"AH")+row)==e.LuongLamLe,"Approved holiday money survives native formula calculation: "+p.Code);Check(book.Evaluate(p.Employee,(p.Code=="CDNL"?"AL":"AK")+row)==e.PhuCapTien,"Approved allowance survives native formula calculation: "+p.Code);Check(book.Evaluate(p.Employee,(p.Code=="CDNL"?"AC":"AA")+row)==e.TongTienLuong,"Personal total agrees with current inputs: "+p.Code);}
                if(personal && neighbor && p.Code!="CDNL")Check(book.Evaluate(p.Employee,"AA6")==565000,"Direct neighbor's shared formula preserved after replacing management anchor: "+p.Code);
            }
        }
    }
    static void Applied(string file)
    {
        var all=ExcelService.LoadWorkshopsFromExcel(file);foreach(var w in all){w.TinhQuyQLTuSanLuong=true;w.SanLuongLoai1=10;w.DonGiaQLLoai1=100003;w.SanLuongLoai2=w.SanLuongLe=0;w.DieuChinhLuongQL=123456;foreach(var e in w.ListNhanVien)if(e.LaQuanLy){e.HeSoCoDinh=2;e.CongSP=26;e.CongLePhep=1;e.CongHocHop=1;e.CongVHTT=1;e.LuongLamLe=345600;e.PhuCapTien=78900;e.QuanLyTinhMoi=true;}ManagementPayroll.Recalculate(w);}
        for(int i=0;i<3;i++){ExcelService.SaveWorkshopsToExcel(file,all);Verify(file,all,true);var next=ExcelService.LoadWorkshopsFromExcel(file);foreach(var w in all){var n=next.First(x=>x.MaPX==w.MaPX);Check(n.TongLuongQL==w.TongLuongQL && n.DieuChinhLuongQL==123456,"Applied total and common adjustment persist after reload: "+w.MaPX);}all=next;}
        foreach(var w in all){w.DieuChinhLuongQL=-65432;foreach(var e in w.ListNhanVien)if(e.LaQuanLy){e.LuongLamLe=0;e.PhuCapTien=0;e.CongLePhep=e.CongHocHop=e.CongVHTT=0;}ManagementPayroll.Recalculate(w);}ExcelService.SaveWorkshopsToExcel(file,all);Verify(file,all,true);
        all=ExcelService.LoadWorkshopsFromExcel(file);foreach(var w in all){w.DieuChinhLuongQL=0;ManagementPayroll.Recalculate(w);}ExcelService.SaveWorkshopsToExcel(file,all);Verify(file,all,true);
    }
    static void Edit(DataGridView grid,int row,string value){grid.Rows[row].Cells[2].Value=value;typeof(DataGridView).GetMethod("OnCellEndEdit",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(grid,new object[]{new DataGridViewCellEventArgs(2,row)});Application.DoEvents();}
    static void Ui(string file)
    {
        var all=ExcelService.LoadWorkshopsFromExcel(file);var w=all[0];w.LuongSPQL=1000000;w.LuongCheDoQL=750000;
        using(var main=new MainForm()){
            typeof(MainForm).GetField("_workshops",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(main,all);typeof(MainForm).GetField("_currentIndex",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(main,0);typeof(MainForm).GetMethod("LoadCurrentWorkshopToUI",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(main,null);main.Show();Application.DoEvents();
            var grid=(DataGridView)typeof(MainForm).GetField("dgvTienLuong",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(main);var auto=(CheckBox)typeof(MainForm).GetField("chkTuDongLuyKe",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(main);auto.Checked=false;Application.DoEvents();Check(grid.Rows[8].Cells[2].ReadOnly,"Computed management total cannot accept an edit that would be discarded");Edit(grid,11,"150.000");Check(w.DieuChinhLuongQL==150000 && w.TongLuongQL==1900000 && Convert.ToString(grid.Rows[8].Cells[2].Value)=="1.900.000","Main summary edit updates model and visible total even with automatic rollups off");
            Edit(grid,10,"333.333");Check(w.LuongCheDoQL==333333 && w.TongLuongQL==1483333,"Approved aggregate outside-product edit is respected");w.RecalculateFromEmployees(false);Check(w.LuongCheDoQL==333333,"Generic recalculation does not silently erase an explicit approved aggregate");
            Edit(grid,9,"1.500.000");Check(w.LuongSPQL==1500000 && w.TongLuongQL==1983333,"Manual approved product fund updates when production pricing is off");
            ExcelService.SaveWorkshopsToExcel(file,all);Verify(file,all,false,"TK");all=ExcelService.LoadWorkshopsFromExcel(file);w=all[0];Check(w.LuongSPQL==1500000 && w.LuongCheDoQL==333333 && w.DieuChinhLuongQL==150000 && w.TongLuongQL==1983333,"Main-grid inputs persist without a prior management-dialog apply");
            using(var dialog=new ManagementPayrollForm(w)){dialog.Show();Application.DoEvents();typeof(ManagementPayrollForm).GetMethod("Preview",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(dialog,new object[]{true});}
            Check(w.LuongSPQL==1500000 && w.LuongCheDoQL!=333333 && w.DieuChinhLuongQL==150000,"Applying personal detail intentionally replaces the separate aggregate while keeping the common adjustment");
            ExcelService.SaveWorkshopsToExcel(file,all);Verify(file,all,false,"TK");

        }
    }
    static void Source(string source,string dir)
    {
        string original=Hash(source),file=Path.Combine(dir,"source-copy.xlsx");File.Copy(source,file);var all=ExcelService.LoadWorkshopsFromExcel(file);int index=0;
        foreach(var w in all){w.TinhQuyQLTuSanLuong=false;w.LuongSPQL=w.CongQL>0?1000000+index:0;w.DieuChinhLuongQL=4000;foreach(var e in w.ListNhanVien)if(e.LaQuanLy){e.HeSoCoDinh=2;e.XepLoai="A";e.LuongLamLe=100000+index;e.PhuCapTien=20000;e.QuanLyTinhMoi=true;}ManagementPayroll.Recalculate(w);index++;}
        for(int i=0;i<2;i++){ExcelService.SaveWorkshopsToExcel(file,all);Verify(file,all,true,null,false);var next=ExcelService.LoadWorkshopsFromExcel(file);foreach(var w in all)Check(next.First(x=>x.MaPX==w.MaPX).TongLuongQL==w.TongLuongQL,"Actual template management total stable after reload: "+w.MaPX);all=next;}
        Check(Hash(source)==original,"Original user workbook untouched; source-copy monetary inputs are synthetic test values only");
    }
    [STAThread] static int Main(string[] args){string dir=Path.Combine(Path.GetTempPath(),"management-acceptance-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);try{string f=Path.Combine(dir,"test.xlsx");Fixture(f);Applied(f);if(args.Length>0 && args[0]!="--ui")Source(args[0],dir);if(args.Length>0 && args[0]=="--ui"){File.Delete(f);Fixture(f);Application.EnableVisualStyles();Ui(f);}Check(Directory.GetFiles(dir,".emga-save-*.tmp").Length==0,"No staged files remain");}finally{Directory.Delete(dir,true);}Console.WriteLine("PASS: "+checks+" management acceptance/recalculation checks.");return 0;}
}
