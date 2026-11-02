using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using CiscoLDTLApp;
class SteelGnvExportTests
{
    static int checks;
    static void Check(bool ok,string message) { checks++; if(!ok) throw new Exception(message); }
    static void Entry(ZipArchive z,string path,string value) { using(var w=new StreamWriter(z.CreateEntry(path).Open())) w.Write(value); }
    static void Fixture(string file)
    {
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)) {
            Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/></Types>");
            Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/></Relationships>");
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='NT Thép' sheetId='1' r:id='rId1'/></sheets><calcPr/></workbook>");
            Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>");
            Entry(z,"xl/worksheets/sheet1.xml","<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData><row r='97'><c r='B97' t='inlineStr'><is><t>GNV cuối mẫu giả</t></is></c><c r='C97' t='inlineStr'><is><t>Công</t></is></c><c r='D97'><v>2</v></c><c r='I97'><v>1000</v></c><c r='J97'><v>2000</v></c></row><row r='98'><c r='B98' t='inlineStr'><is><t>Tiền lương chế độ</t></is></c><c r='J98'><v>123</v></c></row></sheetData></worksheet>");
        }
    }
    static string Hash(string file) { using(var sha=SHA256.Create()) using(var s=File.OpenRead(file)) return Convert.ToBase64String(sha.ComputeHash(s)); }
    static string Text(string file,string address) { using(var z=ZipFile.OpenRead(file)) return ExcelService.GetCellString(ExcelService.LoadSheetXml(z,ExcelService.GetSheetMap(z)["NT Thép"]),ExcelService.LoadSharedStrings(z),address); }
    static double Number(string file,string address) { using(var z=ZipFile.OpenRead(file)) return ExcelService.GetCellDouble(ExcelService.LoadSheetXml(z,ExcelService.GetSheetMap(z)["NT Thép"]),ExcelService.LoadSharedStrings(z),address,-1); }
    static ArisingWorkItem Item(int i,int row) { return new ArisingWorkItem(i,"GNV giả "+i,"Công",i,1000,0,"",row) { TuTinhThanhTien=true }; }
    static void Run(string file)
    {
        var all=ExcelService.LoadWorkshopsFromExcel(file);var steel=all[2];
        bool last=false;foreach(var item in steel.ListGNV) if(item.ExcelRow==97) last=true;
        Check(last==(Number(file,"D97")>0 || Number(file,"J97")>0),"Native row 97 imported whenever source has work or money");
        var sourceItems=new List<ArisingWorkItem>(steel.ListGNV);
        ExcelService.SaveWorkshopsToExcel(file,all);
        foreach(var item in sourceItems) Check(Text(file,"B"+item.ExcelRow)==item.NoiDung && Number(file,"D"+item.ExcelRow)==item.KhoiLuong && Number(file,"J"+item.ExcelRow)==item.ThanhTien,"Untouched native items exported through the last source item");
        // Former break at row 80 dropped both this item and later items in lower rows.
        steel.GnvDaChinhSua=true;steel.ListGNV.Clear();
        steel.ListGNV.Add(Item(1,92));steel.ListGNV.Add(Item(2,60));steel.ListGNV.Add(Item(3,97));
        steel.ListGNV.Add(Item(4,0));steel.ListGNV.Add(Item(5,92));steel.ListGNV.Add(Item(6,9999));
        steel.ListGNV[2].KhoiLuong=0;steel.ListGNV[4].DonGia=-1000;
        steel.RecalculateDerived();ExcelService.SaveWorkshopsToExcel(file,all);
        Check(Text(file,"B92")=="GNV giả 1" && Text(file,"B60")=="GNV giả 2" && Text(file,"B97")=="GNV giả 3","Early high row cannot stop later export, first and last native slots included");
        Check(Text(file,"B61")=="GNV giả 4" && Text(file,"B62")=="GNV giả 5" && Text(file,"B63")=="GNV giả 6","New, duplicate and external anchors allocate distinct free slots without overwriting later anchors");
        Check(Number(file,"J97")==0 && Number(file,"J62")==-5000,"Zero and signed adjustment exported");
        using(var z=ZipFile.OpenRead(file)) {var b=PayrollFormulaBook.Load(z);Check(b.Evaluate("NT Thép","J59")==steel.LuongCVKhac,"Native total recalculates to all exported details, not stale source formulas");}
        Check(Text(file,"B98")=="Tiền lương chế độ","Statutory section below GNV untouched");
        all=ExcelService.LoadWorkshopsFromExcel(file);steel=all[2];
        Check(steel.ListGNV.Count==6 && steel.ListGNV[2].KhoiLuong==0,"Snapshot retains edited list, including zero rows");
        steel.ListGNV.Clear();for(int i=0;i<38;i++)steel.ListGNV.Add(Item(i+1,0));
        steel.RecalculateDerived();ExcelService.SaveWorkshopsToExcel(file,all);
        for(int i=0;i<38;i++) Check(Text(file,"B"+(60+i))=="GNV giả "+(i+1) && Number(file,"J"+(60+i))==(i+1)*1000,"Every row of full 38-item section exported");
        using(var z=ZipFile.OpenRead(file)){var b=PayrollFormulaBook.Load(z);Check(b.Evaluate("NT Thép","D59")==741 && b.Evaluate("NT Thép","J59")==741000,"Full capacity totals match real Excel detail formulas");}
        all=ExcelService.LoadWorkshopsFromExcel(file);steel=all[2];
        Check(steel.ListGNV.Count==38,"Full list survives save/reload");
        steel.ListGNV.Add(Item(39,0));steel.RecalculateDerived();string before=Hash(file);bool blocked=false;
        try{ExcelService.SaveWorkshopsToExcel(file,all);}catch(ArgumentException ex){blocked=ex.Message.Contains("38") && ex.Message.Contains("GNV");}
        Check(blocked && Hash(file)==before,"Overflow explains capacity and leaves entire original unchanged, never saves half");
        steel.ListGNV.Clear();steel.ListGNV.Add(Item(1,97));steel.RecalculateDerived();ExcelService.SaveWorkshopsToExcel(file,all);
        Check(Text(file,"B60")=="" && Number(file,"J60")==0 && Text(file,"B97")=="GNV giả 1","Removed rows cleared in native sheet, surviving anchor retained");
        steel.ListGNV.Clear();steel.RecalculateDerived();ExcelService.SaveWorkshopsToExcel(file,all);
        Check(Text(file,"B97")=="" && Number(file,"J97")==0 && Number(file,"J59")==0,"Deleting all rows clears native detail and total");
        using(var z=ZipFile.OpenRead(file)){var b=PayrollFormulaBook.Load(z);Check(b.Evaluate("NT Thép","J59")==0,"Cleared total remains zero after Excel formula recalculation");}
    }
    static int Main(string[] args)
    {
        string dir=Path.Combine(Path.GetTempPath(),"steel-gnv-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
        try { string file=Path.Combine(dir,"copy.xlsx");
            if(args.Length>0) {string hash=Hash(args[0]);File.Copy(args[0],file);Run(file);Check(Hash(args[0])==hash,"Original uploaded workbook untouched");}
            else {Fixture(file);var imported=ExcelService.LoadWorkshopsFromExcel(file)[2];
                ExcelService.ImportExternalExcelToWorkshop(file,imported);
                Check(imported.GnvDaChinhSua,"Imported native GNV is authoritative when exported to a different template");Run(file);}
            Check(Directory.GetFiles(dir,".emga-save-*.tmp").Length==0,"Temporary files cleaned up after success and failure");
        } finally{Directory.Delete(dir,true);}
        Console.WriteLine("PASS: "+checks+" native Steel GNV export checks.");return 0;
    }
}
