using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using CiscoLDTLApp;
class GnvCalculationTests
{
    static int checks;
    static void Check(bool ok,string message) { checks++; if(!ok) throw new Exception(message); }
    static WorkshopData Example()
    {
        var w=new WorkshopData { MaPX="TK",TenPX="PX giả",LDDauKy=1,LDTrucTiep=1,GnvDaChinhSua=true };
        w.ListNhanVien.Add(new EmployeeRecord { HoTen="Nhân sự giả",CongSP=26,DiemSanXuat=1,XepLoai="A" });
        w.ListGNV.Add(new ArisingWorkItem(1,"GNV giả","Công",2,100000,200000,"Ghi chú giả",81));
        w.RecalculateDerived(); w.RecalculateFromEmployees(true); return w;
    }
    static void Model()
    {
        var w=Example(); var item=w.ListGNV[0]; item.TuTinhThanhTien=true;
        item.KhoiLuong=3; w.RecalculateDerived(); Check(item.ThanhTien==300000 && w.LuongCVKhac==300000,"Changed quantity recomputes existing money");
        item.DonGia=150000; w.RecalculateFromEmployees(true); Check(item.ThanhTien==450000 && w.ListNhanVien[0].LuongSanPham==450000,"Changed rate recomputes money before salary allocation");
        item.KhoiLuong=0; w.RecalculateDerived(); Check(item.ThanhTien==0 && w.CongCVKhac==0,"Zero quantity clears money");
        item.KhoiLuong=2; item.DonGia=0; w.RecalculateDerived(); Check(item.ThanhTien==0,"Zero rate clears money");
        item.DonGia=-100000; w.RecalculateDerived(); Check(item.ThanhTien==-200000,"Signed adjustment retains product sign");
        item.TuTinhThanhTien=false; item.ThanhTien=123456; w.RecalculateDerived(); Check(item.ThanhTien==123456,"Untouched special source amount not silently recalculated");
        w.ListGNV.Clear(); w.RecalculateDerived(); w.RecalculateFromEmployees(true);
        Check(w.LuongCVKhac==0 && w.CongCVKhac==0 && w.ListNhanVien[0].LuongSanPham==0,"Removing last row clears totals and allocation");
        w.MaPX="CDNL"; w.LuongCVKhac=999999; w.ListGNV.Add(new ArisingWorkItem(1,"GNV giả","Công",2,100000,0) { TuTinhThanhTien=true });
        w.RecalculateDerived(); Check(w.ListGNV[0].ThanhTien==200000 && w.LuongCVKhac==999999,"CDNL line computes without double-counting its separate wage fund");
    }
    static void Edit(DataGridView grid,int column,string value)
    {
        grid.Rows[0].Cells[column].Value=value;
        typeof(DataGridView).GetMethod("OnCellEndEdit",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(grid,new object[]{new DataGridViewCellEventArgs(column,0)});
        Application.DoEvents();
    }
    static void Ui()
    {
        using(var form=new MainForm())
        {
            var w=Example();
            typeof(MainForm).GetField("_workshops",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(form,new List<WorkshopData>{w});
            typeof(MainForm).GetField("_currentIndex",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(form,0);
            typeof(MainForm).GetMethod("LoadCurrentWorkshopToUI",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null);
            form.Show(); Application.DoEvents();
            var grid=(DataGridView)typeof(MainForm).GetField("dgvGNV",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            Check(grid.Columns[5].ReadOnly,"GNV amount is read-only");
            Edit(grid,3,"3"); Check(w.ListGNV[0].ThanhTien==300000 && Convert.ToString(grid.Rows[0].Cells[5].Value)=="300.000","UI recomputes after quantity edit");
            Check(w.ListGNV[0].ExcelRow==81 && w.ListGNV[0].GhiChu=="Ghi chú giả","Editing retains source row and note");
            Check(w.LuongCVKhac==300000 && w.ListNhanVien[0].LuongSanPham==300000,"UI updates totals and salary");
            Edit(grid,4,""); Check(w.ListGNV[0].ThanhTien==0 && w.ListNhanVien[0].LuongSanPham==0,"Cleared UI rate removes old amount");
            Edit(grid,4,"100.000"); Check(w.ListGNV[0].ThanhTien==300000,"Restore rate");
            grid.Rows.RemoveAt(0); Application.DoEvents();
            Check(w.ListGNV.Count==0 && w.CongCVKhac==0 && w.LuongCVKhac==0,"Deleting last UI row clears totals"); form.Close();
        }
    }
    static void Workbook()
    {
        string dir=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-output"); Directory.CreateDirectory(dir);
        string path=Path.Combine(dir,"gnv-roundtrip.xlsx");
        File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Data","03.9.2026_NT_T9_TCHC.xlsx"),path,true);
        var all=ExcelService.LoadWorkshopsFromExcel(path);
        foreach(var w in all)
        {
            w.GnvDaChinhSua=true; w.ListGNV.Clear();
            w.ListGNV.Add(new ArisingWorkItem(1,"GNV giả","Công",3.5,120000,1,"Ghi chú giả") { TuTinhThanhTien=true });
            w.RecalculateDerived();
        }
        ExcelService.SaveWorkshopsToExcel(path,all); var read=ExcelService.LoadWorkshopsFromExcel(path);
        foreach(var w in read) Check(w.GnvDaChinhSua && w.ListGNV.Count==1 && w.ListGNV[0].ThanhTien==420000 && w.ListGNV[0].TuTinhThanhTien,"Edited GNV roundtrip: "+w.MaPX);
        foreach(var w in read) { w.ListGNV.Clear(); w.RecalculateDerived(); }
        ExcelService.SaveWorkshopsToExcel(path,read);
        foreach(var w in ExcelService.LoadWorkshopsFromExcel(path)) Check(w.ListGNV.Count==0 && w.GnvDaChinhSua,"Deleted GNV stays deleted after reload: "+w.MaPX);
    }
    [STAThread] static int Main(string[] args)
    {
        Model();
        if(args.Length>0 && args[0]=="--ui") { Application.EnableVisualStyles(); Ui(); }
        if(args.Length>0 && args[0]=="--workbook") Workbook();
        Console.WriteLine("PASS: "+checks+" GNV checks ("+(args.Length>0?args[0]:"model")+")."); return 0;
    }
}
