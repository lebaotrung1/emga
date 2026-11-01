using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Collections.Generic;
using CiscoLDTLApp;
class ProductionPayrollTests
{
    static int checks;
    static void Check(bool ok,string message) { checks++; if(!ok) throw new Exception(message); }
    static WorkshopData Example()
    {
        var w=new WorkshopData { MaPX="LT",TenPX="Phân xưởng giả",LDDauKy=1,LDTrucTiep=1,
            SanLuongLoai1=10,DonGiaLoai1=100000,ThanhTienLoai1=1000000,
            SanLuongLe=2,DonGiaLe=250000,ThanhTienLe=500000,LuongSPTT=1500000 };
        w.ListNhanVien.Add(new EmployeeRecord { HoTen="Nhân sự giả",SDB="001",CongSP=26,DiemSanXuat=1,XepLoai="A" });
        w.RecalculateFromEmployees(true); return w;
    }
    static void ModelChecks()
    {
        foreach(ProductionLine line in Enum.GetValues(typeof(ProductionLine)))
        {
            var w=Example(); double old=line==ProductionLine.Loai1?w.ThanhTienLoai1:line==ProductionLine.Loai2?w.ThanhTienLoai2:w.ThanhTienLe;
            double initial=w.LuongSPTT;
            ProductionPayroll.UpdateLine(w,line,0,100000);
            double amount=line==ProductionLine.Loai1?w.ThanhTienLoai1:line==ProductionLine.Loai2?w.ThanhTienLoai2:w.ThanhTienLe;
            Check(amount==0 && w.LuongSPTT==initial-old,"Zero quantity removes only that line and updates fund");
            ProductionPayroll.UpdateLine(w,line,3,100000); Check(w.LuongSPTT==initial-old+300000,"Restore quantity");
            ProductionPayroll.UpdateLine(w,line,3,0); Check(w.LuongSPTT==initial-old,"Zero price removes old amount");
            ProductionPayroll.UpdateLine(w,line,3,0); Check(w.LuongSPTT==initial-old,"Repeated calculation must not deduct twice");
        }
        var x=Example(); x.TruocTrongQuy.LuongSPTT=700000;
        ProductionPayroll.UpdateLine(x,ProductionLine.NgayLe,0,250000); x.RecalculateFromEmployees(true);
        Check(x.ListNhanVien[0].LuongSanPham==1000000,"Employee allocation uses changed fund");
        Check(x.KyQuy.LuongSPTT==1700000,"Quarter total uses changed month");
        ProductionPayroll.UpdateLine(x,ProductionLine.Loai1,0,100000); x.RecalculateFromEmployees(true);
        Check(x.LuongSPTT==0 && x.ListNhanVien[0].LuongSanPham==0,"Zero total fund clears employee salary");
        x.MaPX="LG"; ProductionPayroll.UpdateLine(x,ProductionLine.NgayLe,1,1500);
        Check(x.ThanhTienLe==2000,"Gang holiday rounds to 1000 as Excel");
        x.MaPX="LT"; ProductionPayroll.UpdateLine(x,ProductionLine.NgayLe,1,1250);
        Check(x.ThanhTienLe==1300,"Thep holiday rounds to 100 as Excel");
        double unchanged=x.LuongSPTT; x.RecalculateDerived(); Check(x.LuongSPTT==unchanged,"Summary refresh does not change fund again");
    }
    static TextBox Box(MainForm form,string name)
    { return (TextBox)typeof(MainForm).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form); }
    static void UiChecks()
    {
        using(var form=new MainForm())
        {
            var w=Example();
            typeof(MainForm).GetField("_workshops",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(form,new List<WorkshopData>{w});
            typeof(MainForm).GetField("_currentIndex",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(form,0);
            typeof(MainForm).GetMethod("LoadCurrentWorkshopToUI",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null);
            form.Show(); Application.DoEvents();
            Check(Box(form,"txtTTLe").ReadOnly && Box(form,"txtTTLoai1").ReadOnly && Box(form,"txtTTLoai2").ReadOnly,"Calculated amount fields are read-only");
            Box(form,"txtSLLe").Clear(); Application.DoEvents();
            Check(w.ThanhTienLe==0 && Box(form,"txtTTLe").Text=="0","Clearing holiday quantity updates displayed amount immediately");
            Check(w.LuongSPTT==1000000 && w.ListNhanVien[0].LuongSanPham==1000000,"UI changes update fund and employee salary");
            Box(form,"txtSLLe").Text="2"; Application.DoEvents();
            Check(w.ThanhTienLe==500000 && w.LuongSPTT==1500000,"UI restore recalculates");
            Box(form,"txtDGLe").Text="0"; Application.DoEvents();
            Check(w.ThanhTienLe==0 && w.LuongSPTT==1000000,"UI zero holiday unit price");
            Box(form,"txtSLLoai1").Clear(); Application.DoEvents();
            Check(w.ThanhTienLoai1==0 && w.LuongSPTT==0 && w.ListNhanVien[0].LuongSanPham==0,"UI clear last product clears salary");
            form.Close();
        }
    }
    static void WorkbookChecks()
    {
        string data=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Data","03.9.2026_NT_T9_TCHC.xlsx");
        var all=ExcelService.LoadWorkshopsFromExcel(data);
        foreach(var w in all)
            if(w.MaPX=="TK" || w.MaPX=="LT")
                Check(w.DonGiaLoai2>0 && w.DonGiaLe>0,"Load all unit prices from workbook: "+w.MaPX);
        string dir=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-output"); Directory.CreateDirectory(dir);
        string path=Path.Combine(dir,"production-roundtrip.xlsx"); File.Copy(data,path,true);
        foreach(var w in all)
            if(w.MaPX=="TK" || w.MaPX=="LT" || w.MaPX=="LG")
            { ProductionPayroll.UpdateLine(w,ProductionLine.NgayLe,0,w.DonGiaLe); w.RecalculateFromEmployees(true); }
        ExcelService.SaveWorkshopsToExcel(path,all);
        var loaded=ExcelService.LoadWorkshopsFromExcel(path);
        for(int i=0;i<all.Count;i++) if(all[i].MaPX=="TK" || all[i].MaPX=="LT" || all[i].MaPX=="LG")
        {
            Check(loaded[i].SanLuongLe==0 && loaded[i].ThanhTienLe==0,"Zero holiday survives workbook reload: "+all[i].MaPX);
            Check(loaded[i].LuongSPTT==all[i].LuongSPTT,"Updated fund survives workbook reload: "+all[i].MaPX);
        }
    }
    [STAThread] static int Main(string[] args)
    {
        ModelChecks();
        if(args.Length>0 && args[0]=="--ui") { Application.EnableVisualStyles(); UiChecks(); }
        if(args.Length>0 && args[0]=="--workbook") WorkbookChecks();
        Console.WriteLine("PASS: "+checks+" production payroll checks ("+(args.Length>0?args[0]:"model")+")."); return 0;
    }
}
