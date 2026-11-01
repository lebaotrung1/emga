using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using CiscoLDTLApp;
class ManagementPayrollUiTests
{
    static void Check(bool ok,string text) { if(!ok) throw new Exception(text); }
    [STAThread] static int Main()
    {
        Application.EnableVisualStyles();
        var w=new WorkshopData { TenPX="Phân xưởng kiểm thử", MaPX="TK", LuongSPQL=1000000, SanLuongLoai1=10, DonGiaQLLoai1=100000 };
        w.ListNhanVien.Clear();
        w.ListNhanVien.Add(new EmployeeRecord { HoTen="Nhân sự giả", ChucDanh="QĐ", LaQuanLy=true, HeSoCoDinh=2, CongSP=26, XepLoai="A", QuanLyTinhMoi=true });
        w.ListNhanVien.Add(new EmployeeRecord { HoTen="Trưởng ca giả", ChucDanh="TRC", LaQuanLy=true, QuanLyTinhMoi=true, CongSP=26, LuongSanPham=1234 });
        var preview=typeof(ManagementPayrollForm).GetMethod("Preview",BindingFlags.NonPublic|BindingFlags.Instance);
        using(var f=new ManagementPayrollForm(w))
        {
            f.Show(); Application.DoEvents();
            var grid=(DataGridView)typeof(ManagementPayrollForm).GetField("grid",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);
            Check(grid.Rows.Count==1 && !w.ListNhanVien[1].LaQuanLy,"Direct shift leader excluded even with stale management flag");
            preview.Invoke(f,new object[]{false}); Application.DoEvents();
            Check(w.ListNhanVien[0].LuongSanPham==0,"Preview must not mutate original");
            preview.Invoke(f,new object[]{true});
            Check(w.ListNhanVien[0].LuongSanPham==1000000 && f.DialogResult==DialogResult.OK,"Apply updates original");
        }
        using(var f=new ManagementPayrollForm(w))
        {
            f.Show(); Application.DoEvents(); f.Close();
            Check(w.ListNhanVien[0].LuongSanPham==1000000,"Close preserves saved state");
        }
        w.SanLuongLoai2=2; w.SanLuongLe=1;
        using(var f=new ManagementPayrollForm(w))
        {
            f.Show(); Application.DoEvents();
            var first=(NumericUpDown)typeof(ManagementPayrollForm).GetField("price1",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);
            var second=(NumericUpDown)typeof(ManagementPayrollForm).GetField("price2",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);
            var fund=(NumericUpDown)typeof(ManagementPayrollForm).GetField("fund",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);
            first.Value=200000; second.Value=50000; Application.DoEvents();
            Check(fund.Value==2500000 && !fund.Enabled,"UI fund is computed from 3 production lines, not editable");
            Check(w.DonGiaQLLoai1==100000,"Unapplied rate edits isolated");
            var holiday=(Label)typeof(ManagementPayrollForm).GetField("holidayPrice",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);
            Check(holiday.Text.Contains("400.000"),"Holiday management price automatically doubles L1");
            first.Value=0; Application.DoEvents(); Check(fund.Value==100000,"Zero L1 rate also clears holiday fund");
            first.Value=200000; preview.Invoke(f,new object[]{true});
            Check(w.TinhQuyQLTuSanLuong && w.LuongSPQL==2500000 && w.ListNhanVien[0].LuongSanPham==2500000,"Apply saves rates and allocates calculated fund");
        }
        using(var f=new ManagementPayrollForm(w))
        {
            var first=(NumericUpDown)typeof(ManagementPayrollForm).GetField("price1",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);
            first.Value=300000; f.Close();
            Check(w.DonGiaQLLoai1==200000 && w.LuongSPQL==2500000,"Cancel does not persist changed rate");
        }
        Check(w.ListNhanVien[1].LuongSanPham==1234 && !w.ListNhanVien[1].QuanLyTinhMoi,"Management application never reallocates shift leader salary");
        using(var main=new MainForm())
        {
            main.Show(); Application.DoEvents();
            MenuStrip menu=null;
            foreach(Control c in main.Controls) if(c is MenuStrip) menu=(MenuStrip)c;
            Check(menu!=null && menu.Visible && menu.Items.Count==3,"Management entry visible in main form");
            Check(menu.Bounds.Width>500 && menu.Bounds.Height>10,"Management menu has usable size");
            var timer=new Timer { Interval=200 };
            bool found=false;
            timer.Tick+=(s,e)=> {
                foreach(Form form in Application.OpenForms)
                    if(form is ManagementPayrollForm) { found=true; form.Close(); break; }
                timer.Stop();
            };
            timer.Start(); ((ToolStripMenuItem)menu.Items[0]).PerformClick();
            timer.Dispose(); Check(found,"Main form opens management payroll dialog");
            main.Close();
        }
        Console.WriteLine("PASS: WinForms preview, apply, cancel and main menu/dialog smoke test (Linux/Mono)."); return 0;
    }
}
