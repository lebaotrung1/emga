using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Xml;
using System.Reflection;
using System.Windows.Forms;
using CiscoLDTLApp;
class CranePayrollTests
{
    static int checks;
    static void Check(bool ok,string message) { checks++; if(!ok) throw new Exception(message); }
    static WorkshopData Steel() { var ws=Find(WorkshopData.CreateDefaultWorkshops()); ws.LuongSPTT=5000000; ws.SanLuongPhuTT=1000000; ws.CauPhoiKhoanKhac=100000; return ws; }
    static CraneStage[] Stages() {
        var stages=CranePayroll.CreateStages();
        for(int i=0;i<3;i++) { stages[i].Regular=100*(i+1); stages[i].Holiday=10*(i+1); stages[i].Rate=1000*(i+1); }
        return stages;
    }
    static void Model()
    {
        var ws=Steel(); var stages=Stages();
        CranePayroll.Apply(ws,stages);
        Check(ws.CauPhoi[0].Amount==120000 && ws.CauPhoi[1].Amount==480000 && ws.CauPhoi[2].Amount==1080000,"Three separate rates and doubled holidays");
        Check(ws.SanLuongPhuTT==1780000 && ws.LuongSPTT==5780000,"Replace old total, retain other crane work once");
        Check(ws.SanLuongPhuKL==660 && ws.SanLuongLoai1==0 && ws.SanLuongLe==0,"Holiday counted once in crane volume; production quantities untouched");
        for(int i=0;i<5;i++) { CranePayroll.Apply(ws,stages); ws.RecalculateDerived(); }
        Check(ws.LuongSPTT==5780000 && ws.TongLuongTT==5780000,"Repeated apply/recalculate does not accumulate fund");
        stages[1].Regular=0; stages[1].Holiday=0; CranePayroll.Apply(ws,stages);
        Check(ws.LuongSPTT==5300000,"Clearing one stage removes its previous amount");
        foreach(var s in stages) { s.Regular=0;s.Holiday=0; } CranePayroll.Apply(ws,stages);
        Check(ws.SanLuongPhuTT==100000 && ws.LuongSPTT==4100000,"Zero all stages retains unrelated fund only");
        double before=ws.LuongSPTT; stages[2].Rate=double.NaN;
        bool rejected=false; try { CranePayroll.Apply(ws,stages); } catch(ArgumentException) { rejected=true; }
        Check(rejected && ws.LuongSPTT==before && ws.CauPhoi[2].Rate==3000,"Invalid input cannot partially update model");
        rejected=false; try { CranePayroll.Apply(new WorkshopData { MaPX="CDNL" },Stages()); } catch(ArgumentException) { rejected=true; }
        Check(rejected,"Crane formula never applied to CDNL");
        stages=Stages(); stages[0].Regular=1.125; stages[0].Holiday=.25; stages[0].Rate=1234.5;
        Check(stages[0].Amount==1.625*1234.5,"Fractional volume and rates match requested formula");
        ws=Steel(); ws.TinhQuyQLTuSanLuong=true; ws.SanLuongLoai1=10; ws.DonGiaQLLoai1=2000;
        ws.ListGNV.Add(new ArisingWorkItem(1,"GNV giả","Công",2,300000,0) { TuTinhThanhTien=true });
        CranePayroll.Apply(ws,Stages());
        Check(ws.LuongSPQL==20000 && ws.ListGNV[0].ThanhTien==600000,"Independent management rate and GNV recalculation retained");
    }
    static string Text(string file) {
        using(var zip=ZipFile.OpenRead(file)) using(var reader=new StreamReader(zip.GetEntry("word/document.xml").Open())) {
            var doc=new XmlDocument(); doc.LoadXml(reader.ReadToEnd()); return doc.InnerText;
        }
    }
    static void Word(string dir) {
        foreach(LoaiKyBaoCao period in Enum.GetValues(typeof(LoaiKyBaoCao))) {
            var ws=Steel(); ws.LoaiKy=period; CranePayroll.Apply(ws,Stages());
            string file=Path.Combine(dir,"crane-"+period+".docx"), updated=Path.Combine(dir,"updated-"+period+".docx");
            WordReportService.ExportWordDocx(ws,file); WordReportService.UpdateExistingMasterDocx(ws,file,updated);
            foreach(string text in new[] {WordReportService.GeneratePreviewText(ws),Text(file),Text(updated)}) {
                foreach(var stage in ws.CauPhoi) Check(text.Contains(stage.Name),"All three stages in preview/new/template Word");
                Check(text.Contains("1,8 triệu đồng") && text.Contains("không cộng thêm"),"Million currency with one decimal and fund inclusion explicit");
            }
        }
    }
    static WorkshopData Find(List<WorkshopData> all) { foreach(var w in all) if(w.MaPX=="LT") return w; throw new Exception("No steel workshop"); }
    static double Number(XmlDocument doc,string address) { var ns=new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("s","http://schemas.openxmlformats.org/spreadsheetml/2006/main"); return double.Parse(doc.SelectSingleNode("//s:c[@r='"+address+"']/s:v",ns).InnerText,System.Globalization.CultureInfo.InvariantCulture); }
    static Dictionary<string,string> Formulas(string file)
    {
        var result=new Dictionary<string,string>();
        using(var zip=ZipFile.OpenRead(file)) {
            var map=ExcelService.GetSheetMap(zip);
            foreach(string name in new[] {"NT CĐNL","CĐNL"}) {
                var doc=ExcelService.LoadSheetXml(zip,map[name]);
                foreach(XmlElement cell in doc.GetElementsByTagName("c","http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                    foreach(XmlNode child in cell.ChildNodes) if(child.LocalName=="f") result[name+"!"+cell.GetAttribute("r")]=child.InnerText;
            }
        }
        return result;
    }
    static void Workbook(string source,string dir) {
        string file=Path.Combine(dir,"crane-roundtrip.xlsx"); File.Copy(source,file,true);
        var formulas=Formulas(source);
        var all=ExcelService.LoadWorkshopsFromExcel(file); var ws=Find(all);
        double baseFund=ws.LuongSPTT-ws.SanLuongPhuTT, other=ws.CauPhoiKhoanKhac;
        var stages=Stages(); stages[0].Regular=100.123; stages[0].Holiday=10.456;
        CranePayroll.Apply(ws,stages); double expected=ws.LuongSPTT;
        Check(expected==baseFund+other+ws.CauPhoi[0].Amount+ws.CauPhoi[1].Amount+ws.CauPhoi[2].Amount,"Workbook replaces existing three stages, keeps ancillary work");
        for(int pass=0;pass<3;pass++) {
            ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); ws=Find(all);
            Check(ws.CauPhoiDaApDung && ws.LuongSPTT==expected,"Stable repeated workbook save/reload");
            Check(ws.CauPhoi[0].Regular==100.123 && ws.CauPhoi[0].Holiday==10.456 && ws.CauPhoiKhoanKhac==other,"Input precision and ancillary work persist");
            using(var zip=ZipFile.OpenRead(file)) {
                var map=ExcelService.GetSheetMap(zip); var doc=ExcelService.LoadSheetXml(zip,map["NT Thép"]);
                Check(Number(doc,"D50")==100.123 && Number(doc,"D46")==10.456 && Number(doc,"I46")==2000,"Original NT rows updated with separate holiday input and doubled rate");
                Check(Number(doc,"J44")==ws.SanLuongPhuTT && Number(doc,"J35")==expected,"Acceptance fund cells updated");
            }
        }
        stages[0].Regular=100.125; stages[0].Holiday=10.25; stages[0].Rate=1234.5;
        CranePayroll.Apply(ws,stages); expected=ws.LuongSPTT;
        ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); ws=Find(all);
        Check(Math.Abs(ws.LuongSPTT-expected)<0.000001 && ws.CauPhoi[0].Amount==120.625*1234.5,"Fractional rates and money remain stable on reopen");
        foreach(var stage in stages) { stage.Regular=0; stage.Holiday=0; } CranePayroll.Apply(ws,stages);
        ExcelService.SaveWorkshopsToExcel(file,all); ws=Find(ExcelService.LoadWorkshopsFromExcel(file));
        Check(ws.CauPhoiDaApDung && ws.CauPhoi[0].Amount==0 && ws.SanLuongPhuTT==other,"Cleared inputs stay zero on reopen");
        var savedFormulas=Formulas(file);
        Check(formulas.Count>0 && formulas.Count==savedFormulas.Count,"CDNL formula count unchanged");
        foreach(var pair in formulas) Check(savedFormulas.ContainsKey(pair.Key) && savedFormulas[pair.Key]==pair.Value,"CDNL formula preserved: "+pair.Key);

    }
    static void Ui()
    {
        var ws=Steel();
        using(var form=new CranePayrollForm(ws)) {
            var grid=(DataGridView)typeof(CranePayrollForm).GetField("grid",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            form.Show(); Application.DoEvents();
            Check(grid.Rows.Count==3 && grid.Columns[4].ReadOnly,"Three stages, amount read-only");
            grid.Rows[0].Cells[1].Value="100,5"; grid.Rows[0].Cells[2].Value="10,25"; grid.Rows[0].Cells[3].Value="1.000";
            Check(!ws.CauPhoiDaApDung && ws.LuongSPTT==5000000,"Editing alone leaves model untouched");
            form.Close();
        }
        Check(!ws.CauPhoiDaApDung && ws.LuongSPTT==5000000,"Cancel isolates all changes");
        using(var form=new CranePayrollForm(ws)) {
            var grid=(DataGridView)typeof(CranePayrollForm).GetField("grid",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            form.Show(); Application.DoEvents();
            grid.Rows[0].Cells[1].Value="100,5"; grid.Rows[0].Cells[2].Value="10,25"; grid.Rows[0].Cells[3].Value="1.000";
            foreach(Control c in form.Controls) if(c is FlowLayoutPanel) ((Button)c.Controls[0]).PerformClick();
            Check(ws.CauPhoiDaApDung && ws.CauPhoi[0].Amount==121000 && ws.LuongSPTT==4221000,"Apply parses VN decimals and replaces old fund");
        }
        using(var form=new MainForm()) {
            typeof(MainForm).GetField("_workshops",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(form,new List<WorkshopData>{ws});
            typeof(MainForm).GetField("_currentIndex",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(form,0);
            typeof(MainForm).GetMethod("LoadCurrentWorkshopToUI",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null);
            form.Show(); Application.DoEvents();
            var summary=(TextBox)typeof(MainForm).GetField("txtSLPhuTT",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            Check(summary.ReadOnly && summary.Text==WordReportService.FmtInt(ws.SanLuongPhuTT),"Main acceptance summary updated and protected");
            typeof(MainForm).GetMethod("SaveUIToCurrentWorkshop",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null);
            Check(ws.LuongSPTT==4221000,"Main UI save does not add fund again");
            form.Close();
        }
    }
    [STAThread]
    static int Main(string[] args) {
        string dir=Path.Combine(Path.GetTempPath(),"crane-payroll-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        try { Model(); Word(dir); if(args.Length>0 && args[0]=="--ui") { Application.EnableVisualStyles(); Ui(); }
            else if(args.Length>0) Workbook(args[0],dir); }
        finally { Directory.Delete(dir,true); }
        Console.WriteLine("PASS: "+checks+" crane payroll checks."); return 0;
    }
}
