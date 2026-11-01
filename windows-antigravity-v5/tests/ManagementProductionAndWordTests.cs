using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using CiscoLDTLApp;
class ManagementProductionAndWordTests
{
    static int checks;
    static void Check(bool ok,string text) { checks++; if(!ok) throw new Exception(text); }
    static void ModelChecks()
    {
        var w=new WorkshopData { MaPX="LT", TinhQuyQLTuSanLuong=true, SanLuongLoai1=10.123,
            SanLuongLoai2=3.25, SanLuongLe=1, DonGiaQLLoai1=100000, DonGiaQLLoai2=50000 };
        w.ListNhanVien.Add(new EmployeeRecord { ChucDanh="QĐ", LaQuanLy=true, CongSP=26, HeSoCoDinh=2, XepLoai="A" });
        ManagementPayroll.Recalculate(w);
        Check(w.DonGiaQLLe==200000 && w.LuongSPQL==1374800,"Fund sums three quantity × management price lines; holiday price doubles L1");
        Check(w.ListNhanVien[0].LuongSanPham==1374800,"Fund reaches individual salary");
        w.DonGiaQLLoai1=0; ManagementPayroll.Recalculate(w);
        Check(w.DonGiaQLLe==0 && w.LuongSPQL==162500,"Zero L1 price clears both L1 and holiday");
        w.DonGiaQLLoai1=100000; w.SanLuongLe=0; ManagementPayroll.Recalculate(w);
        Check(w.ThanhTienQLLe==0 && w.LuongSPQL==1174800,"Zero holiday quantity clears holiday fund");
        w.SanLuongLoai1=0; w.SanLuongLoai2=0; ManagementPayroll.Recalculate(w);
        Check(w.LuongSPQL==0 && w.ListNhanVien[0].LuongSanPham==0,"All zero quantities clear management allocation");
        w.SanLuongLoai1=2; w.DonGiaQLLoai1=1234.5678; ManagementPayroll.Recalculate(w);
        Check(w.ThanhTienQLLoai1==2469,"Management fund rounds once per line to dong");
        w.DonGiaLoai1=999999; w.DonGiaQLLoai1=100;
        ProductionPayroll.UpdateLine(w,ProductionLine.Loai1,3,999999); w.RecalculateFromEmployees(true);
        Check(w.LuongSPQL==300 && w.DonGiaQLLoai1==100,"Production edits update management fund using independent entered management rate");
        Check(WordReportService.FmtMoney(123456789)=="123,5","Million conversion");
        Check(WordReportService.FmtMoney(1000000)=="1,0" && WordReportService.FmtMoney(0)=="0,0","Exactly one decimal including zero");
        Check(WordReportService.FmtMoney(-1250000)=="-1,3","Negative midpoint rounding");
        Check(WordReportService.FmtMoneyFromThousands(12500)=="12,5","Averages already in thousands convert correctly");
        Check(WordReportService.FmtInt(1000000)=="1.000.000","UI money formatter remains dong");
    }
    static XmlDocument Document(string path)
    { var doc=new XmlDocument(); using(var zip=ZipFile.OpenRead(path)) using(var stream=zip.GetEntry("word/document.xml").Open()) doc.Load(stream); return doc; }
    static void CheckDocument(WorkshopData w, string path)
    {
        var doc=Document(path); string text=doc.InnerText;
        Check(text.Contains("triệu đồng") && !text.Contains("nghìn đồng") && !text.Contains("1.000đ/người"),"Word currency labels: "+w.MaPX);
        Check(text.Contains(WordReportService.FmtMoney(w.TongLuongToanPX)+" triệu đồng"),"Word fund is scaled: "+w.MaPX);
        Check(text.Contains("12,5 triệu đồng/người/tháng"),"Word average is scaled from thousands");
        var ns=new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("w","http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        bool employee=false,gnv=false;
        foreach(XmlNode row in doc.SelectNodes("//w:tr",ns))
        {
            var cells=row.SelectNodes("w:tc",ns);
            if(cells.Count==12 && cells[2].InnerText=="Nhân sự giả")
            { employee=true; Check(cells[10].InnerText=="1,2" && cells[11].InnerText=="2,0","Employee money cells in millions with one decimal"); }
            if(cells.Count==6 && cells[1].InnerText=="GNV giả")
            { gnv=true; Check(cells[4].InnerText=="1.234.500" && cells[5].InnerText=="2,5","GNV rate kept explicit dong; total converted to millions"); }
        }
        Check(employee && gnv,"Detailed report tables tested");
    }
    static void WordChecks()
    {
        string dir=Path.Combine(Path.GetTempPath(),"cisco-word-million-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        try
        {
            foreach(var w in WorkshopData.CreateDefaultWorkshops())
            foreach(LoaiKyBaoCao period in Enum.GetValues(typeof(LoaiKyBaoCao)))
            {
                w.LoaiKy=period; w.TuDongCongLuyKe=false; w.LuongSPTT=1234500000; w.LuongSPQL=234500000;
                w.TLBinhQuanToanPX=12500; w.ThuNhapBQToanPX=12500; w.LDThucTeBQ=0;
                w.CongAnCa=100; w.BDDH_Muc1_Cong=100; w.BDDH_Muc2_Cong=100; w.CoMuc3BDDH=true; w.BDDH_Muc3_Cong=100;
                w.ListGNV.Clear(); w.ListGNV.Add(new ArisingWorkItem(1,"GNV giả","Công",2,1234500,2469000));
                w.ListNhanVien.Clear(); w.DinhKemBangChiTietNhanVien=true;
                w.ListNhanVien.Add(new EmployeeRecord { STT=1,SDB="001",HoTen="Nhân sự giả",BoPhan="Bộ phận giả",LuongSanPham=1234567,TongTienLuong=2000000 });
                w.KyQuy.TongLuongToanPX=2345600000; w.KyLuyKe.TongLuongToanPX=3456700000;
                string path=Path.Combine(dir,w.MaPX+"-"+period+".docx");
                WordReportService.ExportWordDocx(w,path); CheckDocument(w,path);
                string preview=WordReportService.GeneratePreviewText(w);
                Check(preview.Contains(WordReportService.FmtMoney(w.TongLuongToanPX)+" triệu đồng") && preview.Contains("12,5 triệu đồng/người/tháng") ||
                    (period!=LoaiKyBaoCao.ThangThuong && preview.Contains("12,5") && preview.Contains("triệu đồng/người/tháng")),"Preview consistent");
                string updated=Path.Combine(dir,w.MaPX+"-updated-"+period+".docx");
                WordReportService.UpdateExistingMasterDocx(w,path,updated); CheckDocument(w,updated);
                Check(Document(updated).SelectNodes("//*[local-name()='document']").Count==1,"Updated DOCX valid");
            }
        }
        finally { Directory.Delete(dir,true); }
    }
    static void WorkbookChecks()
    {
        string dir=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"test-output"); Directory.CreateDirectory(dir);
        string path=Path.Combine(dir,"management-production-roundtrip.xlsx");
        File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Data","03.9.2026_NT_T9_TCHC.xlsx"),path,true);
        var all=ExcelService.LoadWorkshopsFromExcel(path);
        foreach(var w in all)
        {
            w.TinhQuyQLTuSanLuong=true; w.DonGiaQLLoai1=1234.5678; w.DonGiaQLLoai2=987.6543;
            w.SanLuongLoai1=10.123; w.SanLuongLoai2=3.456; w.SanLuongLe=1.234;
            w.RecalculateFromEmployees(true);
        }
        ExcelService.SaveWorkshopsToExcel(path,all);
        var read=ExcelService.LoadWorkshopsFromExcel(path);
        for(int i=0;i<all.Count;i++)
        {
            var a=all[i]; var b=read[i];
            Check(b.TinhQuyQLTuSanLuong && b.DonGiaQLLoai1==a.DonGiaQLLoai1 && b.DonGiaQLLoai2==a.DonGiaQLLoai2,"Rates survive save/reload: "+a.MaPX);
            Check(b.LuongSPQL==a.LuongSPQL && b.DonGiaQLLe==2*b.DonGiaQLLoai1,"Computed fund stable: "+a.MaPX);
            Check(b.SanLuongLoai1==a.SanLuongLoai1 && b.SanLuongLoai2==a.SanLuongLoai2 && b.SanLuongLe==a.SanLuongLe,"Quantities preserve precision: "+a.MaPX);
        }
        ExcelService.SaveWorkshopsToExcel(path,read);
        var twice=ExcelService.LoadWorkshopsFromExcel(path);
        for(int i=0;i<all.Count;i++) Check(twice[i].LuongSPQL==read[i].LuongSPQL,"Repeated save stable");
    }
    static int Main(string[] args)
    {
        ModelChecks(); WordChecks();
        if(args.Length>0 && args[0]=="--workbook") WorkbookChecks();
        Console.WriteLine("PASS: "+checks+" management production/Word checks; five workshops × five periods, new and updated DOCX."); return 0;
    }
}
