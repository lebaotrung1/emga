using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using CiscoLDTLApp;
class CdnlPayrollTests
{
    static int checks;
    static void Check(bool ok,string message) { checks++; if(!ok) throw new Exception(message); }
    static void Number(PayrollFormulaBook b,string sheet,string cell,double n) { b.Sheets.Add(sheet); b.Cells[sheet+"!"+cell]=new PayrollCell { Sheet=sheet,Address=cell,Number=n,Text=n.ToString(CultureInfo.InvariantCulture),IsNumber=true }; }
    static void Formula(PayrollFormulaBook b,string sheet,string cell,string f) { b.Sheets.Add(sheet); b.Cells[sheet+"!"+cell]=new PayrollCell { Sheet=sheet,Address=cell,Formula=f }; }
    static WorkshopData Example()
    {
        var all=WorkshopData.CreateDefaultWorkshops(); var w=all[3]; var b=new PayrollFormulaBook(); all[0].LuongSPTT=10000000;
        Number(b,"NT thiêu kết","J57",1); Number(b,"NT gang","J48",0); Number(b,"NT Thép","J35",0);
        Formula(b,"NT CĐNL","D37","'NT thiêu kết'!J57+'NT gang'!J48+'NT Thép'!J35");
        Number(b,"NT CĐNL","H44",.01); Formula(b,"NT CĐNL","I44","ROUND(D37*H44,-2)");
        Number(b,"NT CĐNL","H46",.05); Number(b,"NT CĐNL","D75",100); Number(b,"NT CĐNL","H75",200); Number(b,"NT CĐNL","D76",10);
        Formula(b,"NT CĐNL","I45","ROUND(D37*H46,-2)+D75*H75+D76*H75*2");
        Number(b,"NT CĐNL","D82",20); Number(b,"NT CĐNL","E82",2); Number(b,"NT CĐNL","H82",10000); Number(b,"NT CĐNL","I340",-1000);
        Formula(b,"NT CĐNL","I81","ROUND(D82/E82*H82*0.8,-2)+I340");
        Number(b,"NT CĐNL","I342",10000); Number(b,"NT CĐNL","I343",-1000); Formula(b,"NT CĐNL","I341","SUM(I342:I343)");
        Number(b,"NT CĐNL","I361",30000); Number(b,"NT CĐNL","I371",20000); Number(b,"NT CĐNL","I367",-5000); Formula(b,"NT CĐNL","I356","I361+I371+I367");
        Formula(b,"NT CĐNL","I422","ROUND(D37*1%,-2)");
        Formula(b,"NT CĐNL","I43","I44+I45+I81+I341+I356");
        Number(b,"NT CĐNL","I416",1234); Formula(b,"NT CĐNL","I418","SUM(CĐNL!AE12:AE12)"); Number(b,"NT CĐNL","I419",0); Number(b,"NT CĐNL","I420",0);
        Formula(b,"NT CĐNL","I417","I418+I419+I420"); Number(b,"NT CĐNL","I432",0);
        Formula(b,"NT CĐNL","I42","I43+I416+I417"); Formula(b,"NT CĐNL","I421","I422+I432"); Formula(b,"NT CĐNL","I440","I42+I421");
        Number(b,"CĐNL","AB12",10); Formula(b,"CĐNL","AJ146","'NT CĐNL'!I45/10");
        Formula(b,"CĐNL","AD12","ROUND(AB12*$AJ$146,-2)"); Formula(b,"CĐNL","AE12","ROUND(M12/26*G12,-2)");
        Number(b,"CĐNL","AH12",1000); Formula(b,"CĐNL","AI12","ROUND(AH12*Q12,-3)");
        Formula(b,"CĐNL","AJ12","ROUND(R12/26*G12,-2)"); Number(b,"CĐNL","AK12",0); Number(b,"CĐNL","AL12",0);
        Formula(b,"CĐNL","AC12","AD12+AE12+AI12+AJ12+AK12+AL12");
        w.ListNhanVien.Add(new EmployeeRecord { ExcelRow=12,SDB="FAKE001",HoTen="Nhân sự giả",CongSP=26,CongLePhep=1,CongVHTT=2,CongLamLe=3,MucLuongBHXH=5200000,HeSoHuong=10,DiemSanXuat=10,XepLoai="B",LuongSanPham=999 });
        w.Cdnl=new CdnlPayroll(b,all); return w;
    }
    static void Engine()
    {
        var b=new PayrollFormulaBook(); Number(b,"X","A1",10); Number(b,"X","A2",20);
        Formula(b,"X","B1","SUM(A1:A2)+ROUND(-1.25,1)+5%*100");
        Check(Math.Abs(b.Evaluate("X","B1")-33.7)<.000001,"SUM, percent and Excel negative midpoint rounding");
        Formula(b,"X","B2","IF(A1>0,42,1/0)"); Check(b.Evaluate("X","B2")==42,"Lazy IF does not evaluate unused error branch");
        Formula(b,"X","B3","IFERROR(1/0,7)"); Check(b.Evaluate("X","B3")==7,"IFERROR matches explicit source rule");
        Formula(b,"X","B4","SUMIF(A1:A2,\">10\",A1:A2)"); Check(b.Evaluate("X","B4")==20,"SUMIF evaluates matching range");
        Formula(b,"X","C1","COUNT(A1:A4)"); Check(b.Evaluate("X","C1")==2,"Missing/blank cells do not count as numeric zero");
        Formula(b,"X","C2","COUNTA(A1:A4)"); Check(b.Evaluate("X","C2")==2,"Missing cells do not count as nonempty");
        Formula(b,"X","B5","B6"); Formula(b,"X","B6","B5");
        bool rejected=false; try { b.Evaluate("X","B5"); } catch(PayrollFormulaError) { rejected=true; } Check(rejected,"Circular formulas rejected");
        Formula(b,"X","B7","WEBSERVICE(\"https://example.com\")"); rejected=false; try { b.Evaluate("X","B7"); } catch(PayrollFormulaError) { rejected=true; } Check(rejected,"No network-capable formula execution");
    }
    static void Model()
    {
        var w=Example(); var data=w.Cdnl; var result=data.Preview(w);
        Check(result.Valid && w.ListNhanVien[0].LuongSanPham==999 && !data.Applied,"Preview is isolated");
        Check(result.Groups["I44"]==100000 && result.Groups["I45"]==524000,"Percentage fund and regular/holiday tonnage independent");
        Check(result.Groups["I81"]==79000 && result.Groups["I341"]==9000 && result.Groups["I356"]==45000,"Transport distance/80%, screening SUM and repair deductions remain distinct");
        data.Apply(w); Check(w.LuongSPTT==757000 && w.LuongCVKhac==1234,"Five funds reach acceptance; auxiliary work counted once");
        var e=w.ListNhanVien[0]; Check(e.LuongSanPham==524000 && e.HeSoHuong==10,"Source AB and AJ146 used, no blanket B=90% adjustment");
        Check(e.LuongHocVHTT==400000 && e.LuongLePhepCD==200000 && e.LuongLamLe==3000,"Personal wage components follow own source formulas, not generic constants");
        double before=e.LuongSanPham; data.Apply(w); w.RecalculateFromEmployees(true); Check(e.LuongSanPham==before,"Repeat application does not add funds or use generic redistribution");
        data.Workshops[0].LuongSPTT=20000000; CdnlPayroll.EnsureCurrent(w); Check(w.LuongSPTT==1357000 && e.LuongSanPham==1024000,"Other workshop fund changes propagate to CDNL and correct allocation group");
        w.GnvDaChinhSua=true; w.ListGNV.Add(new ArisingWorkItem(1,"GNV giả","Công",2,30000,60000)); w.RecalculateFromEmployees(true);
        Check(data.Last.Groups["I356"]==85000 && w.LuongCVKhac==1234,"Edited GNV flows through I371 once, not additional whole-workshop wage");
        var draft=data.Draft(); draft.SetOverride("NT CĐNL!H46",.1,"Giá giả để kiểm thử"); draft.Preview(w);
        Check(data.Overrides.Count==0 && w.Cdnl==data,"Cancel/draft input isolation");
        bool rejected=false; try { draft.SetOverride("NT CĐNL!H46",.1,""); } catch(ArgumentException) { rejected=true; } Check(rejected,"Override needs reason");
        Formula(data.Book,"CĐNL","AB12","#REF!"); before=e.LuongSanPham;
        rejected=false; try { data.Apply(w); } catch(ArgumentException) { rejected=true; } Check(rejected && e.LuongSanPham==before,"Missing source aborts entire apply without partial mutation");
        rejected=false; try { WordReportService.ExportWordDocx(w,Path.Combine(Path.GetTempPath(),"must-not-write-cdnl-"+Guid.NewGuid().ToString("N")+".docx")); } catch(ArgumentException) { rejected=true; } Check(rejected,"Word blocks stale wage results");
        data.Distribute=false; data.Apply(w); Check(e.LuongSanPham==before,"Fund-only mode preserves individual payroll");
        w.DinhKemBangChiTietNhanVien=true; rejected=false; try { WordReportService.GeneratePreviewText(w); } catch(ArgumentException) { rejected=true; } Check(rejected,"Fund-only cannot masquerade as newly calculated salary appendix");
        var missing=Example(); missing.Cdnl.Book.Cells.Remove("CĐNL!AD12"); Check(!missing.Cdnl.Preview(missing).Valid,"Working employee without source pay formula must not silently receive zero");
        var legacy=Example(); legacy.Cdnl=null; var original=legacy.ListNhanVien[0]; legacy.RecalculateFromEmployees(true); Check(original.LuongSanPham==999 && original.HeSoHuong==10,"Unapplied CDNL also bypasses old generic allocation");
    }
    static void SpecialCases()
    {
        var w=Example(); var b=w.Cdnl.Book;
        var repair=new EmployeeRecord { ExcelRow=68,SDB="FAKE002",HoTen="Nhân sự giả 2",CongSP=26,HeSoHuong=20 };
        w.ListNhanVien.Add(repair); Number(b,"CĐNL","AB68",20); Formula(b,"CĐNL","AJ152","'NT CĐNL'!I356/20"); Formula(b,"CĐNL","AD68","ROUND(AB68*$AJ$152,-2)");
        foreach(string c in new[] {"AE","AI","AJ","AK","AL"}) Number(b,"CĐNL",c+"68",0);
        Formula(b,"CĐNL","AC68","AD68+AE68+AI68+AJ68+AK68+AL68");
        w.Cdnl.Apply(w);
        Check(repair.LuongSanPham==45000 && w.ListNhanVien[0].LuongSanPham==524000,"Two departments allocate with their own independent rates");
        w.Cdnl.KeepSource("CĐNL!AD68"); w.Cdnl.KeepSource("CĐNL!AB68");
        w.Cdnl.Rules[CdnlPayroll.EmployeeKey(repair)]=new CdnlEmployeeRule { Group="I356",Weight=20,RateCell="AJ152",Reason="Synthetic source rate rule" };
        w.Cdnl.Apply(w); Check(repair.LuongSanPham==45000,"Explicit membership can reuse the correct department source rate");
        Number(b,"CĐNL","Z12",2); Formula(b,"CĐNL","AB12","Z12*5");
        w.Cdnl.Apply(w); Check(w.ListNhanVien[0].HeSoHuong==10 && !w.Cdnl.IsRatingRow(12),"Numeric Z must not be replaced by grade B");
        var fresh=new EmployeeRecord { SDB="FAKE003",HoTen="Nhân sự mới giả",CongSP=26 }; w.ListNhanVien.Add(fresh);
        Check(!w.Cdnl.Preview(w).Valid,"New attendance import requires an explicit allocation rule");
        w.Cdnl.Rules[CdnlPayroll.EmployeeKey(fresh)]=new CdnlEmployeeRule { Group="I341",Weight=90,Rate=100,Reason="Synthetic approved rule" };
        w.Cdnl.Apply(w); Check(fresh.LuongSanPham==9000,"New employee explicit department/weight/rate accepted without common pool");
        fresh.CongLePhep=1; Check(!w.Cdnl.Preview(w).Valid,"No fabricated holiday/leave formula for new employee without source row");
        w=Example(); b=w.Cdnl.Book; w.TinhQuyQLTuSanLuong=true; w.SanLuongLoai1=10; w.DonGiaQLLoai1=100;
        var manager=new EmployeeRecord { ExcelRow=5,SDB="FAKE004",HoTen="Quản lý giả",LaQuanLy=true,CongSP=26,HeSoCoDinh=2,XepLoai="A" }; w.ListNhanVien.Add(manager);
        Formula(b,"NT CĐNL","I422","#REF!"); Formula(b,"CĐNL","AB5","#REF!");
        w.Cdnl.Apply(w); Check(w.LuongSPQL==1000 && manager.LuongSanPham==1000,"Opted-in management production rate remains independent from original broken source formula");
    }
    static string WordText(string file) { using(var z=ZipFile.OpenRead(file)) using(var r=new StreamReader(z.GetEntry("word/document.xml").Open())) return r.ReadToEnd(); }
    static void Word(string dir)
    {
        var w=Example(); w.Cdnl.Apply(w);
        foreach(LoaiKyBaoCao period in Enum.GetValues(typeof(LoaiKyBaoCao))) {
            w.LoaiKy=period; string file=Path.Combine(dir,"cdnl-"+period+".docx"),updated=Path.Combine(dir,"updated-"+period+".docx");
            WordReportService.ExportWordDocx(w,file); WordReportService.UpdateExistingMasterDocx(w,file,updated);
            foreach(string text in new[] {WordReportService.GeneratePreviewText(w),WordText(file),WordText(updated)})
                foreach(string group in CdnlPayroll.GroupNames) Check(text.Contains(group),"Group details in preview/new/template Word");
        }
    }
    static void Workbook(string path,string dir)
    {
        string file=Path.Combine(dir,"cdnl-roundtrip.xlsx"); File.Copy(path,file,true);
        var all=ExcelService.LoadWorkshopsFromExcel(file); var w=all[3]; var data=w.Cdnl;
        Check(data!=null && !data.Applied,"Existing workbook gains tool without automatic payroll mutation");
        var full=data.Preview(w); Check(!full.Valid && full.Errors.ContainsKey("CĐNL!AB98"),"Real source #N/A must be reviewed, never replaced by zero silently");
        data.Distribute=false; var p=data.Preview(w); Check(p.Valid && p.Groups.Count==6,"All real source group formulas computed");
        double old=w.ListNhanVien[0].LuongSanPham; data.Apply(w); double fund=w.LuongSPTT;
        Check(old==w.ListNhanVien[0].LuongSanPham,"Applying real fund-only does not overwrite salaries");
        for(int i=0;i<2;i++) { ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); w=all[3]; data=w.Cdnl; Check(data.Applied && !data.Distribute && Math.Abs(w.LuongSPTT-fund)<1,"Real fund calculations persist on repeated save/reload"); }
        data.Distribute=true; Check(data.Preview(w).Errors.ContainsKey("CĐNL!AD8"),"Fund-only saves do not turn missing source salaries into approved numeric input"); data.Distribute=false;
        double rate=data.Book.Cells["NT CĐNL!H46"].Number;
        data.SetOverride("NT CĐNL!H46",rate*1.1,"TEST ONLY: synthetic override, temporary copy"); data.Apply(w); double changed=w.LuongSPTT;
        Check(changed!=fund,"Entered department rate changes corresponding fund");
        ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); w=all[3]; data=w.Cdnl;
        Check(data.Overrides["NT CĐNL!H46"].Active && Math.Abs(w.LuongSPTT-changed)<1,"Input, reason and active state persist");
        data.Overrides["NT CĐNL!H46"].Active=false; data.Apply(w);
        Check(Math.Abs(w.LuongSPTT-fund)<1,"Removing override restores original source input");
        ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); w=all[3]; data=w.Cdnl;
        Check(!data.Overrides["NT CĐNL!H46"].Active && Math.Abs(w.LuongSPTT-fund)<1,"Restored original input survives reload");
        data.Distribute=true;
        for(int attempt=0;attempt<10;attempt++) {
            var r=data.Preview(w); if(r.Valid) break;
            foreach(var error in r.Errors) { Check(error.Key.Contains("!"),"Source issue is a precise editable cell"); data.SetOverride(error.Key,error.Key.StartsWith("CĐNL!AB")?10:0,"TEST ONLY: fake approved value, never a user payroll result"); }
        }
        Check(data.Preview(w).Valid,"With explicitly synthetic corrections, full workflow is computable"); data.Apply(w);
        EmployeeRecord assigned=null; foreach(var e in w.ListNhanVien) if(e.ExcelRow==12) assigned=e;
        Check(assigned!=null,"Source employee available for rate rule roundtrip");
        data.KeepSource("CĐNL!AB12"); data.KeepSource("CĐNL!AD12");
        data.Rules[CdnlPayroll.EmployeeKey(assigned)]=new CdnlEmployeeRule { Group="I45",Weight=assigned.HeSoHuong,RateCell="AJ146",Reason="TEST ONLY: department rate link" }; data.Apply(w);

        double salary=w.ListNhanVien[0].LuongSanPham;
        var wages=new Dictionary<string,double>(); foreach(var e in w.ListNhanVien) wages[CdnlPayroll.EmployeeKey(e)]=e.TongTienLuong;
        ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); w=all[3];
        Check(w.Cdnl.Applied && w.Cdnl.Distribute && w.ListNhanVien[0].LuongSanPham==salary,"Explicit corrections and full allocation survive Excel roundtrip");
        Check(w.Cdnl.Rules[CdnlPayroll.EmployeeKey(assigned)].RateCell=="AJ146","Department rate reference persists with the assignment");
        foreach(var e in w.ListNhanVien) Check(wages[CdnlPayroll.EmployeeKey(e)]==e.TongTienLuong,"Every individual total stable after save/reopen");
        w.TinhQuyQLTuSanLuong=true; w.SanLuongLoai1=10; w.SanLuongLoai2=0; w.SanLuongLe=0; w.DonGiaQLLoai1=12345; w.DonGiaQLLoai2=0;
        foreach(var e in w.ListNhanVien) if(e.LaQuanLy) { e.HeSoCoDinh=2; e.XepLoai="A"; }
        w.Cdnl.Apply(w); double managerFund=w.LuongSPQL;
        var managerWages=new Dictionary<string,double>(); foreach(var e in w.ListNhanVien) if(e.LaQuanLy) managerWages[CdnlPayroll.EmployeeKey(e)]=e.TongTienLuong;
        ExcelService.SaveWorkshopsToExcel(file,all); all=ExcelService.LoadWorkshopsFromExcel(file); w=all[3];
        Check(w.TinhQuyQLTuSanLuong && w.LuongSPQL==managerFund && managerFund==123450,"Independent management production fund survives CDNL save/reload");
        foreach(var e in w.ListNhanVien) if(e.LaQuanLy) Check(managerWages[CdnlPayroll.EmployeeKey(e)]==e.TongTienLuong,"Opted-in management salary stable after source formulas saved");

    }
    static void Ui()
    {
        var w=Example(); double before=w.LuongSPTT;
        using(var form=new CdnlPayrollForm(w)) {
            form.Show(); Application.DoEvents();
            var field=typeof(CdnlPayrollForm).GetField("groups",BindingFlags.Instance|BindingFlags.NonPublic); var grid=(DataGridView)field.GetValue(form);
            Check(grid.Rows.Count==6 && grid.ReadOnly,"Six groups visible and calculated totals read-only");
            var input=(DataGridView)typeof(CdnlPayrollForm).GetField("inputs",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            Check(input.Rows.Count>0,"Actual dependent inputs visible"); form.Close();
        }
        Check(w.LuongSPTT==before && !w.Cdnl.Applied,"Closing preview preserves current worksheet model");
        using(var form=new CdnlPayrollForm(w)) {
            form.Show(); Application.DoEvents();
            var button=(Button)typeof(CdnlPayrollForm).GetField("apply",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            var people=(DataGridView)typeof(CdnlPayrollForm).GetField("people",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            Check(Convert.ToString(people.Rows[0].Cells[2].Value)=="Vận hành năng lượng","Source allocation department visible on employee row");
            people.Rows[0].Cells[7].Value=true; people.Rows[0].Cells[6].Value="Synthetic UI assignment";
            Check(Convert.ToString(people.Rows[0].Cells[5].Value)=="AJ146","Selecting explicit department rule suggests its own source rate");
            Check(button.Enabled,"Valid source enables apply"); button.PerformClick();
            Check(w.Cdnl.Applied && w.ListNhanVien[0].LuongSanPham==524000,"UI apply reaches acceptance and correct allocation");
        }
    }
    static void RealUi(string file)
    {
        var all=ExcelService.LoadWorkshopsFromExcel(file); var w=all[3]; double old=w.ListNhanVien[0].LuongSanPham;
        using(var form=new CdnlPayrollForm(w)) {
            form.Show(); Application.DoEvents();
            var apply=(Button)typeof(CdnlPayrollForm).GetField("apply",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            var errors=(DataGridView)typeof(CdnlPayrollForm).GetField("issues",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            Check(!apply.Enabled && errors.Rows.Count>=6,"Real UI identifies source defects and missing pay formulas and disables full apply");
            var people=(DataGridView)typeof(CdnlPayrollForm).GetField("people",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            bool pending=false; foreach(DataGridViewRow row in people.Rows) if(Convert.ToString(row.Cells[8].Value)=="Chưa tính được") pending=true;
            Check(pending,"Incomplete individual results are marked pending instead of displayed as zero wages");
            var distribute=(CheckBox)typeof(CdnlPayrollForm).GetField("distribute",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            distribute.Checked=false;
            typeof(CdnlPayrollForm).GetMethod("Run",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,new object[]{false});
            Check(apply.Enabled,"Fund-only mode is actionable with real source"); apply.PerformClick();
        }
        Check(w.Cdnl.Applied && !w.Cdnl.Distribute && w.ListNhanVien[0].LuongSanPham==old,"Real UI fund apply preserves existing personal pay");
        using(var main=new MainForm()) {
            typeof(MainForm).GetField("_workshops",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(main,all);
            typeof(MainForm).GetField("_currentIndex",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(main,3);
            typeof(MainForm).GetMethod("LoadCurrentWorkshopToUI",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(main,null);
            main.Show(); Application.DoEvents(); MenuStrip menu=null;
            foreach(Control c in main.Controls) if(c is MenuStrip) menu=(MenuStrip)c;
            ToolStripMenuItem entry=null; foreach(ToolStripItem i in menu.Items) if(i.Text.Contains("CĐNL")) entry=(ToolStripMenuItem)i;
            Check(entry!=null && entry.Visible && !entry.IsOnOverflow,"CDNL entry visible directly on main menu");
            var wages=(DataGridView)typeof(MainForm).GetField("dgvTienLuong",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(main);
            Check(wages.Rows[2].Cells[2].ReadOnly,"Computed monthly fund protected in main acceptance grid");
            var timer=new Timer { Interval=300 }; bool opened=false;
            timer.Tick+=(s,e)=> { foreach(Form dialog in Application.OpenForms) if(dialog is CdnlPayrollForm) { opened=true; dialog.Close(); break; } timer.Stop(); };
            timer.Start(); entry.PerformClick(); timer.Dispose(); Check(opened,"Main CDNL entry opens the real calculator"); main.Close();
        }
    }
    [STAThread] static int Main(string[] args)
    {
        string dir=Path.Combine(Path.GetTempPath(),"cdnl-tests-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        try { Engine(); Model(); SpecialCases(); Word(dir); if(args.Length>0 && args[0]=="--ui") { Application.EnableVisualStyles(); Ui(); } else if(args.Length>1 && args[0]=="--ui-workbook") { Application.EnableVisualStyles(); RealUi(args[1]); } else if(args.Length>0) Workbook(args[0],dir); }
        finally { Directory.Delete(dir,true); }
        Console.WriteLine("PASS: "+checks+" CDNL formula/payroll/Word checks."); return 0;
    }
}
