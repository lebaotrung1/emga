using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Xml;
using CiscoLDTLApp;
class ImportTests
{
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static void Entry(ZipArchive zip, string path, string xml)
    { using (var writer = new StreamWriter(zip.CreateEntry(path).Open())) writer.Write(xml); }
    static string Cell(string address, string text)
    { return "<c r='" + address + "' t='inlineStr'><is><t>" + System.Security.SecurityElement.Escape(text) + "</t></is></c>"; }
    static int Main()
    {
        string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test-output");
        Directory.CreateDirectory(dir);
        string file = Path.Combine(dir, "timesheet.xlsx");
        using (var stream = File.Create(file))
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            Entry(zip, "xl/workbook.xml", "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='Bảng công' sheetId='1' r:id='rId1'/></sheets></workbook>");
            Entry(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>");
            string data = "<row r='1'>" + Cell("A1", "Bảng chấm công") + Cell("F1", "Ngày trong tháng") + "</row>";
            data += "<row r='3'>" + Cell("B3", "Họ và tên") + Cell("C3", "SDB") + Cell("F3", "Số công sản phẩm") + "</row>";
            data += "<row r='6'>" + Cell("B6", "Tổ Phối liệu A") + "</row>";
            data += "<row r='7'>" + Cell("B7", "Nguyễn Văn An") + Cell("C7", "001") + Cell("F7", "12") + "</row>";
            data += "<row r='8'>" + Cell("B8", "Trần Văn Mới") + Cell("C8", "002") + Cell("F8", "15") + "</row>";
            data += "<row r='9'>" + Cell("B9", "Người Bỏ Qua") + Cell("C9", "003") + Cell("F9", "10") + "</row>";
            Entry(zip, "xl/worksheets/sheet1.xml", "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData>" + data + "</sheetData></worksheet>");
        }
        var all = WorkshopData.CreateDefaultWorkshops();
        var target = all[0]; var source = all[1];
        foreach (var ws in all) ws.ListNhanVien.Clear();
        var original = new EmployeeRecord { HoTen = "Nguyễn Văn An", SDB = "1", BoPhan = "Tổ cũ", HeSoCoDinh = 3, MucLuongBHXH = 6000000, CongSP = 99, TongTienLuong = 999999, ExcelRow = 10, TongHopCongRow = 20 };
        source.ListNhanVien.Add(original);
        int prompts = 0;
        Func<EmployeeRecord, EmployeeRecord> resolve = delegate(EmployeeRecord employee) {
            prompts++;
            if (employee.SDB == "003") return null;
            employee.HeSoCoDinh = 2; return employee;
        };
        string log = TimesheetImporter.Import(file, target, all, resolve);
        Check(target.ListNhanVien.Count == 2, "Expected two new employees");
        var copied = target.ListNhanVien[0];
        Check(copied != original && copied.ExcelRow == 0 && copied.TongHopCongRow == 0, "Copy must not reuse source row/reference");
        Check(copied.CongSP == 12 && original.CongSP == 99, "Attendance/source isolation");
        Check(copied.HeSoCoDinh == 3 && copied.BoPhan == "Tổ Phối liệu A", "Profile/team copy");
        Check(target.ListNhanVien[1].CongSP == 15 && prompts == 2, "Manual new employee and skip");
        TimesheetImporter.Import(file, target, all, resolve);
        Check(target.ListNhanVien.Count == 2 && copied.CongSP == 12, "Repeated import must be idempotent");
        Check(log.Contains("Người Bỏ Qua"), "Skipped employee must be reported");
        string master = Path.Combine(dir, "master.xlsx");
        File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "03.9.2026_NT_T9_TCHC.xlsx"), master, true);
        all = ExcelService.LoadWorkshopsFromExcel(master);
        target = all[0];
        copied.STT = target.ListNhanVien.Count + 1;
        target.ListNhanVien.Add(copied);
        target.ListNhanVien[0].ToSanXuat = target.ListNhanVien[0].BoPhan = "Tổ kiểm thử";
        int expected = target.ListNhanVien.Count;
        ExcelService.SaveWorkshopsToExcel(master, all);
        var reloaded = ExcelService.LoadWorkshopsFromExcel(master);
        Check(reloaded[0].ListNhanVien.Count == expected, "Added employee must persist");
        var last = reloaded[0].ListNhanVien[expected - 1];
        Check(last.HoTen == copied.HoTen && last.CongSP == 12 && last.ExcelRow == 0 && last.ToSanXuat == copied.ToSanXuat, "Roundtrip profile/data");
        Check(reloaded[0].ListNhanVien[0].ToSanXuat == "Tổ kiểm thử", "Existing team must persist");
        ExcelService.SaveWorkshopsToExcel(master, reloaded);
        Check(ExcelService.LoadWorkshopsFromExcel(master)[0].ListNhanVien.Count == expected, "Repeated save/load must not duplicate");
        Console.WriteLine("PASS: copy, correct team, manual addition, skip log, repeat import, workbook roundtrip.");
        return 0;
    }
}