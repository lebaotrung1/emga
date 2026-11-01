using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Reflection;
using System.Reflection.Emit;
using CiscoLDTLApp;
class FrameworkCompatibilityTests
{
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static void Entry(ZipArchive zip, string path, string xml)
    { using (var writer = new StreamWriter(zip.CreateEntry(path).Open())) writer.Write(xml); }
    static string Cell(string address, string value)
    { return "<c r='"+address+"' t='inlineStr'><is><t>"+System.Security.SecurityElement.Escape(value)+"</t></is></c>"; }
    static void VerifySplitReference()
    {
        var method = typeof(TimesheetImporter).GetMethod("Import");
        byte[] il = method.GetMethodBody().GetILAsByteArray();
        var codes = new Dictionary<short, OpCode>();
        foreach (var field in typeof(OpCodes).GetFields(BindingFlags.Public | BindingFlags.Static))
            if (field.FieldType == typeof(OpCode)) { var code=(OpCode)field.GetValue(null); codes[code.Value]=code; }
        bool found=false;
        for(int offset=0; offset<il.Length;)
        {
            short value=il[offset++];
            if(value==0xfe) value=(short)(0xfe00 | il[offset++]);
            OpCode code=codes[value]; int size=0;
            switch(code.OperandType)
            {
                case OperandType.InlineNone: break;
                case OperandType.ShortInlineBrTarget: case OperandType.ShortInlineI: case OperandType.ShortInlineVar: size=1; break;
                case OperandType.InlineVar: size=2; break;
                case OperandType.InlineI8: case OperandType.InlineR: size=8; break;
                case OperandType.InlineSwitch: size=4+4*BitConverter.ToInt32(il,offset); break;
                default: size=4; break;
            }
            if(code.OperandType==OperandType.InlineMethod)
            {
                var called=method.Module.ResolveMethod(BitConverter.ToInt32(il,offset));
                if(called.DeclaringType==typeof(string) && called.Name=="Split")
                {
                    var parameters=called.GetParameters();
                    Check(parameters.Length==2 && parameters[0].ParameterType==typeof(char[]) && parameters[1].ParameterType==typeof(StringSplitOptions), "Binary must call .NET Framework char[] Split overload");
                    found=true;
                }
            }
            offset+=size;
        }
        Check(found,"Expected String.Split call in importer");
        Check(typeof(TimesheetImporter).Assembly.GetName().Version.ToString()=="5.3.0.0","Expected v5.3 binary");
    }
    static int Main()
    {
        VerifySplitReference();
        string dir=Path.Combine(Path.GetTempPath(),"cisco-compat-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        try
        {
            string file=Path.Combine(dir,"synthetic-bddh.xlsx");
            using(var stream=File.Create(file)) using(var zip=new ZipArchive(stream,ZipArchiveMode.Create))
            {
                Entry(zip,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='BDDH' sheetId='1' r:id='rId1'/></sheets></workbook>");
                Entry(zip,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>");
                string data="<row r='1'>"+Cell("B1","Quyết toán bồi dưỡng độc hại")+Cell("F1","Ngày trong tháng")+"</row>";
                data+="<row r='3'>"+Cell("B3","Họ và tên")+Cell("C3","SDB")+Cell("E3","1")+Cell("F3","2")+Cell("G3","3")+Cell("H3","Tổng công hưởng BDĐH")+"</row>";
                data+="<row r='5'>"+Cell("B5","Nhân sự giả")+Cell("C5","001")+Cell("E5","M1,M2")+Cell("F5","M3")+Cell("G5"," M2 , M3 , ")+"</row>";
                Entry(zip,"xl/worksheets/sheet1.xml","<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData>"+data+"</sheetData></worksheet>");
            }
            var ws=new WorkshopData();
            var employee=new EmployeeRecord { HoTen="Nhân sự giả",SDB="001" }; ws.ListNhanVien.Add(employee);
            TimesheetImporter.Import(file,ws);
            Check(employee.CoBDDH && employee.BDDH_M1==1 && employee.BDDH_M2==2 && employee.BDDH_M3==2,"Comma-separated BDDH levels parsed");
            Check(employee.TienBDDH==105000,"BDDH amount");
            TimesheetImporter.Import(file,ws);
            Check(employee.BDDH_M1==1 && employee.BDDH_M2==2 && employee.BDDH_M3==2 && employee.TienBDDH==105000,"Repeated import is stable");
            Console.WriteLine("PASS: v5.3 binary references Framework-compatible Split(char[], StringSplitOptions); synthetic BDDH import and repeat import.");
            return 0;
        }
        finally { Directory.Delete(dir,true); }
    }
}
