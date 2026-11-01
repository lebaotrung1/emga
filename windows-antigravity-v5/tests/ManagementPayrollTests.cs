using System;
using System.IO;
using System.Collections.Generic;
using CiscoLDTLApp;
class ManagementPayrollTests
{
    static int checks;
    static void Check(bool ok, string text) { checks++; if (!ok) throw new Exception(text); }
    static bool Near(double a, double b) { return Math.Abs(a-b) < .001; }
    static void Reject(WorkshopData w, string message)
    { bool rejected = false; try { ManagementPayroll.Recalculate(w); } catch (ArgumentException) { rejected = true; } Check(rejected, message); }
    static int Main(string[] args)
    {
        var w = new WorkshopData { MaPX="TK", LuongSPQL=3000000, DieuChinhLuongQL=-50000 };
        w.ListNhanVien.Clear();
        var a = new EmployeeRecord { LaQuanLy=true, HeSoCoDinh=2, CongSP=26, XepLoai="A", CongLePhep=1, MucLuongBHXH=5200000, CongHocHop=1, CongVHTT=1, PhuCapTien=100000, LuongLamLe=150000, QuanLyTinhMoi=true };
        var b = new EmployeeRecord { LaQuanLy=true, HeSoCoDinh=2, CongSP=13, XepLoai="B", QuanLyTinhMoi=true };
        var worker = new EmployeeRecord { LaQuanLy=false, CongSP=26, LuongSanPham=123456, TongTienLuong=123456 };
        w.ListNhanVien.Add(a); w.ListNhanVien.Add(b); w.ListNhanVien.Add(worker);
        ManagementPayroll.Recalculate(w);
        Check(Near(a.HeSoHuong,2) && Near(b.HeSoHuong,.9), "Attendance, grade and coefficients");
        Check(a.LuongSanPham==2069000 && b.LuongSanPham==931000, "Proportional fund allocation");
        Check(a.LuongLePhepCD==200000 && a.LuongHocVHTT==550000, "Leave, meeting and culture components");
        Check(a.TongTienLuong==3069000, "All individual components counted once");
        Check(w.LuongCheDoQL==1000000 && w.TongLuongQL==3950000, "Management summary and general adjustment");
        Check(worker.TongTienLuong==123456 && worker.LuongSanPham==123456, "Management calculation isolates direct workers");
        double total=a.TongTienLuong; ManagementPayroll.Recalculate(w);
        Check(total==a.TongTienLuong, "Repeat calculation idempotent");
        a.CongSP=0; ManagementPayroll.Recalculate(w);
        Check(a.HeSoHuong==0 && a.LuongSanPham==0 && b.LuongSanPham==3000000, "Zero attendance removes stale salary");
        w.LuongSPQL=0; ManagementPayroll.Recalculate(w);
        Check(a.LuongSanPham==0 && b.LuongSanPham==0, "Zero fund clears salary");
        a.CongLePhep=0; a.MucLuongBHXH=0; a.CongHocHop=0; a.CongVHTT=0; a.PhuCapTien=0; a.LuongLamLe=0;
        ManagementPayroll.Recalculate(w);
        Check(a.TongTienLuong==0 && w.LuongCheDoQL==0, "Zero components clear previous totals");
        Check(ManagementPayroll.RoundMoney(250,100)==300 && ManagementPayroll.RoundMoney(-250,100)==-300, "Excel midpoint rounding");
        b.XepLoai="C"; ManagementPayroll.Recalculate(w); Check(b.HeSoHuong==.8,"C grade");
        b.XepLoai="invalid"; Reject(w,"Invalid grade rejected"); b.XepLoai="A";
        b.CongSP=-1; Reject(w,"Negative attendance rejected"); b.CongSP=0;
        w.LuongSPQL=100; Reject(w,"Nonzero fund without weights rejected");
        w.LuongSPQL=-100; Reject(w,"Negative fund rejected"); w.LuongSPQL=0;
        b.CongSP=13; b.HeSoCoDinh=0; Reject(w,"Missing fixed coefficient rejected");
        b.HeSoCoDinh=2; w.LuongSPQL=1000;
        a.CongSP=26; a.HeSoCoDinh=2; a.LuongLamLe=432100; a.PhuCapTien=0;
        w.RecalculateFromEmployees(true);
        Check(a.LuongLamLe==432100 && a.PhuCapTien==0, "Main calculation preserves approved management amounts");

        if (args.Length > 0 && args[0] == "--logic-only")
        { Console.WriteLine("PASS: " + checks + " synthetic payroll checks; private workbook not required."); return 0; }

        string dir=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-output"); Directory.CreateDirectory(dir);
        string original=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Data","03.9.2026_NT_T9_TCHC.xlsx");
        string path=Path.Combine(dir,"management-roundtrip.xlsx"); File.Copy(original,path,true);
        var all=ExcelService.LoadWorkshopsFromExcel(path); int tested=0;
        foreach(var ws in all)
        {
            int managers=0;
            foreach(var e in ws.ListNhanVien) if(e.LaQuanLy) managers++;
            Check(managers>0 && managers<10,"Only management rows, no production teams: "+ws.MaPX);
            EmployeeRecord manager=null;
            foreach(var e in ws.ListNhanVien) if(e.LaQuanLy) { manager=e; break; }
            if(manager==null) continue;
            foreach(var e in ws.ListNhanVien) if(e.LaQuanLy)
            { e.HeSoCoDinh=2; e.CongSP=26; e.XepLoai="A"; e.QuanLyTinhMoi=true; }
            manager.HeSoCoDinh=3; manager.CongSP=13; manager.XepLoai="B";
            manager.LuongLamLe=123400; manager.PhuCapTien=0;
            ws.LuongSPQL=20000000+tested*1000; ws.DieuChinhLuongQL=-12300;
            ManagementPayroll.Recalculate(ws); tested++;
        }
        Check(tested==5,"All five workshops exercised");
        ExcelService.SaveWorkshopsToExcel(path,all);
        var loaded=ExcelService.LoadWorkshopsFromExcel(path);
        for(int i=0;i<all.Count;i++)
        {
            var x=all[i]; var y=loaded[i];
            Check(x.ListNhanVien.Count==y.ListNhanVien.Count,"Employee count preserved");
            Check(x.LuongSPQL==y.LuongSPQL && x.DieuChinhLuongQL==y.DieuChinhLuongQL && Near(x.TongLuongQL,y.TongLuongQL),"Management fund/summary roundtrip "+x.MaPX);
            for(int j=0;j<x.ListNhanVien.Count;j++) if(x.ListNhanVien[j].LaQuanLy)
            {
                var p=x.ListNhanVien[j]; var q=y.ListNhanVien[j];
                Check(q.QuanLyTinhMoi && p.HeSoCoDinh==q.HeSoCoDinh && p.CongSP==q.CongSP && p.XepLoai==q.XepLoai && p.LuongLamLe==q.LuongLamLe && p.PhuCapTien==q.PhuCapTien && p.TongTienLuong==q.TongTienLuong,"Management employee roundtrip "+x.MaPX);
            }
        }
        ExcelService.SaveWorkshopsToExcel(path,loaded);
        var twice=ExcelService.LoadWorkshopsFromExcel(path);
        for(int i=0;i<loaded.Count;i++) Check(Near(loaded[i].TongLuongQL,twice[i].TongLuongQL),"Repeated save/load stable");
        Console.WriteLine("PASS: "+checks+" checks; management payroll and five-workshop Excel roundtrip."); return 0;
    }
}
