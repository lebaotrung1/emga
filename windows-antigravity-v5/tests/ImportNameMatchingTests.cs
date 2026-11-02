using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml;
using CiscoLDTLApp;
class ImportNameMatchingTests
{
    class Row {public string Name,Code;public double Work=12;}
    static int checks;
    static void Check(bool ok,string text){checks++;if(!ok)throw new Exception(text);}
    static void Entry(ZipArchive z,string path,string value){using(var w=new StreamWriter(z.CreateEntry(path).Open()))w.Write(value);}
    static void Put(ZipArchive z,string path,XmlDocument doc){using(var s=z.CreateEntry(path).Open())doc.Save(s);}
    static void SheetFixture(string file,Row[] rows,bool allKinds=true,string mealName=null)
    {
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)){
            string sheets="",rels="";string[] titles=allKinds?new[]{"Bảng chấm công","Quyết toán ăn ca","Quyết toán bồi dưỡng độc hại","Bảng chấm điểm"}:new[]{"Bảng chấm công"};
            for(int kind=0;kind<titles.Length;kind++){
                sheets+="<sheet name='Sheet "+kind+"' sheetId='"+(kind+1)+"' r:id='r"+kind+"'/>";rels+="<Relationship Id='r"+kind+"' Target='worksheets/sheet"+kind+".xml'/>";
                var doc=new XmlDocument();doc.LoadXml("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>");
                ExcelService.SetCellInlineString(doc,"B1",titles[kind],false);ExcelService.SetCellInlineString(doc,"F1","Ngày trong tháng",false);
                ExcelService.SetCellInlineString(doc,"B3","Họ và tên",false);ExcelService.SetCellInlineString(doc,"C3","SDB",false);
                if(kind==0){string[] headers={"Số công sản phẩm","Số công hưởng lương thời gian","Công văn hóa thể thao","Hưởng 100% lương","Số công hưởng BHXH","Công K3"};for(int c=0;c<headers.Length;c++)ExcelService.SetCellInlineString(doc,TimesheetImporter.ColName(6+c)+"3",headers[c],false);}
                else if(kind==3){ExcelService.SetCellInlineString(doc,"F3","Tổng điểm sản xuất",false);ExcelService.SetCellInlineString(doc,"G3","Tổng điểm sau xếp loại",false);ExcelService.SetCellInlineString(doc,"H3","Xếp loại",false);}
                else{for(int c=5;c<=7;c++)ExcelService.SetCellNumeric(doc,TimesheetImporter.ColName(c)+"3",c-4,false);ExcelService.SetCellInlineString(doc,"H3",kind==1?"Tổng công được thanh toán ăn ca":"Tổng công hưởng BDĐH",false);}
                ExcelService.SetCellInlineString(doc,"B6","Tổ kiểm thử",false);
                for(int i=0;i<rows.Length;i++){
                    int r=7+i;var row=rows[i];ExcelService.SetCellInlineString(doc,"B"+r,kind==1 && mealName!=null?mealName:row.Name,false);ExcelService.SetCellInlineString(doc,"C"+r,row.Code,false);
                    if(kind==0){ExcelService.SetCellNumeric(doc,"F"+r,row.Work,false);ExcelService.SetCellNumeric(doc,"G"+r,1,false);ExcelService.SetCellNumeric(doc,"H"+r,2,false);ExcelService.SetCellNumeric(doc,"I"+r,3,false);ExcelService.SetCellNumeric(doc,"J"+r,4,false);ExcelService.SetCellNumeric(doc,"K"+r,5,false);}
                    else if(kind==3){ExcelService.SetCellNumeric(doc,"F"+r,2.25,false);ExcelService.SetCellNumeric(doc,"G"+r,2,false);ExcelService.SetCellInlineString(doc,"H"+r,"B",false);}
                    else{ExcelService.SetCellInlineString(doc,"E"+r,kind==1?"na":"M1",false);ExcelService.SetCellInlineString(doc,"F"+r,kind==1?"tt":"M2",false);ExcelService.SetCellInlineString(doc,"G"+r,kind==1?"na":"M1,M3",false);}
                }
                Put(z,"xl/worksheets/sheet"+kind+".xml",doc);
            }
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets>"+sheets+"</sheets></workbook>");
            Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"+rels+"</Relationships>");
        }
    }
    static WorkshopData Unit(params EmployeeRecord[] people){var w=WorkshopData.CreateDefaultWorkshops()[0];w.ListNhanVien.AddRange(people);return w;}
    static EmployeeRecord Person(string name,string code){return new EmployeeRecord {HoTen=name,SDB=code,CongSP=99,ChucDanh="CN",HeSoCoDinh=2,MucLuongBHXH=5200000};}
    static string Import(string file,WorkshopData w,out int prompts,List<WorkshopData> all=null){int calls=0;string log=TimesheetImporter.Import(file,w,all,e=>{calls++;return null;});prompts=calls;return log;}
    static void MainCases(string dir)
    {
        string file=Path.Combine(dir,"input.xlsx");int prompts;
        foreach(string sourceName in new[]{"Nguyễn Văn Kiểm (1989)","Nguyễn Văn Kiểm(1989)","  NGUYỄN  VĂN KIỂM ( 1989 )  ","Nguyễn\u00a0Văn Kiểm (1989)","Nguyễn Văn Kiểm (1989)".Normalize(NormalizationForm.FormD)}){
            SheetFixture(file,new[]{new Row{Name=sourceName,Code="0901"}});var e=Person("Nguyễn Văn Kiểm","901");var w=Unit(e);string log=Import(file,w,out prompts);
            Check(prompts==0 && w.ListNhanVien.Count==1 && e.HoTen=="Nguyễn Văn Kiểm" && e.SDB=="901","Birth-year annotation/spacing/Unicode never creates or renames an existing profile");
            Check(e.CongSP==12 && e.CongHocHop==1 && e.CongVHTT==2 && e.CongLePhep==3 && e.CongBHXH==4 && e.CongK3==5,"All working-day components imported to canonical profile");
            Check(e.CoBangCong && e.CoAnCa && e.CoBDDH && e.AnCaNhaAn==2 && e.AnCaTienMat==1 && e.BDDH_M1==2 && e.BDDH_M2==1 && e.BDDH_M3==1,"Attendance, meals and BDDH resolve same annotated identity");
            Check(e.DiemSanXuat==2.25 && e.XepLoai=="B" && !log.Contains("chưa được nạp") && !log.Contains("KHÔNG tìm thấy"),"Points imported without false missing-person warnings");
            Import(file,w,out prompts);Check(w.ListNhanVien.Count==1 && e.CongSP==12 && e.CongLePhep==3 && e.BDDH_M1==2,"Repeat import does not duplicate people or attendance");File.Delete(file);
        }
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm",Code="901"}});var annotated=Person("Nguyễn Văn Kiểm (1989)","901");var reverse=Unit(annotated);Import(file,reverse,out prompts);Check(annotated.CongSP==12 && annotated.HoTen.EndsWith("(1989)"),"Annotated roster also matches plain source name");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901"}},false);var a=Person("Nguyễn Văn Kiểm (1988)","901");var b=Person("Nguyễn Văn Kiểm (1989)","901");var same=Unit(a,b);Import(file,same,out prompts);Check(a.CongSP==99 && b.CongSP==12 && prompts==0,"Explicit matching birth year distinguishes same-name same-code profiles");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm",Code="901"}},false);a=Person("Nguyễn Văn Kiểm (1988)","901");b=Person("Nguyễn Văn Kiểm (1989)","901");same=Unit(a,b);string msg=Import(file,same,out prompts);Check(a.CongSP==99 && b.CongSP==99 && prompts==0 && msg.Contains("nhiều hồ sơ"),"No year plus multiple candidates is rejected, not first-match wins");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1988)",Code="901"}},false);b=Person("Nguyễn Văn Kiểm (1989)","901");same=Unit(b);msg=Import(file,same,out prompts);Check(b.CongSP==99 && prompts==0 && msg.Contains("năm sinh"),"Conflicting known birth year never merges profiles");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm",Code="901"}},false);a=Person("Nguyễn Văn Kiểm","901");b=Person("Nguyễn Văn Kiểm","901");same=Unit(a,b);msg=Import(file,same,out prompts);Check(a.CongSP==99 && b.CongSP==99 && msg.Contains("nhiều hồ sơ"),"Duplicate identical roster entries are reported rather than silently choosing the first");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="902"}},false);a=Person("Nguyễn Văn Kiểm","901");same=Unit(a);msg=Import(file,same,out prompts);Check(a.CongSP==99 && same.ListNhanVien.Count==1 && prompts==0 && msg.Contains("SDB khác") && msg.Contains("902") && msg.Contains("901"),"Different code is explained, never guessed or transposed by name");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Trần Văn Thử (1989)",Code="901"}},false);a=Person("Nguyễn Văn Kiểm","901");same=Unit(a);msg=Import(file,same,out prompts);Check(a.CongSP==99 && prompts==0 && msg.Contains("tên khác"),"Same code with a different name does not overwrite current employee");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (Tổ A)",Code="901"}},false);a=Person("Nguyễn Văn Kiểm","901");same=Unit(a);msg=Import(file,same,out prompts);Check(a.CongSP==99 && msg.Contains("tên khác"),"Non-year parenthetical annotations are not stripped indiscriminately");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1988)",Code="901"},new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901"}},false);a=Person("Nguyễn Văn Kiểm","901");same=Unit(a);msg=Import(file,same,out prompts);Check(a.CongSP==99 && prompts==0 && msg.Contains("file có nhiều năm sinh"),"Source preflight prevents merging two differently born workers into a plain roster profile");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1988)",Code="901"}},true,"Nguyễn Văn Kiểm (1989)");a=Person("Nguyễn Văn Kiểm","901");same=Unit(a);msg=Import(file,same,out prompts);Check(a.CongSP==99 && !a.CoAnCa && prompts==0 && msg.Contains("file có nhiều năm sinh"),"Conflicting birth years across sheet types are checked before any attendance is touched");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901"},new Row{Name="Nguyễn Văn Kiểm (1989)",Code="902",Work=15}},false);a=Person("Nguyễn Văn Kiểm","901");b=Person("Nguyễn Văn Kiểm","902");same=Unit(a,b);Import(file,same,out prompts);Check(a.CongSP==12 && b.CongSP==15,"Same name with distinct matching codes remains two separate people");File.Delete(file);

        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1988)",Code="901",Work=10},new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901",Work=15}},false);a=Person("Nguyễn Văn Kiểm (1988)","901");b=Person("Nguyễn Văn Kiểm (1989)","901");same=Unit(a,b);Import(file,same,out prompts);Check(a.CongSP==10 && b.CongSP==15 && prompts==0,"Two explicit birth years in source can resolve two explicitly dated profiles safely");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901",Work=10},new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901",Work=15}},false);a=Person("Nguyễn Văn Kiểm","901");same=Unit(a);Import(file,same,out prompts);Check(a.CongSP==25 && a.CongLePhep==6,"Same dated person on legitimate split attendance rows still aggregates");Import(file,same,out prompts);Check(a.CongSP==25 && a.CongLePhep==6,"Repeat split-row import resets before aggregating, not double counting");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Lê Văn Hồ Sơ Mới (1990)",Code="903"}});same=Unit();int created=0;TimesheetImporter.Import(file,same,null,e=>{created++;e.ChucDanh="CN";e.MucLuongBHXH=5200000;return e;});Check(created==1 && same.ListNhanVien.Count==1 && same.ListNhanVien[0].CongSP==12 && same.ListNhanVien[0].CoAnCa && same.ListNhanVien[0].CoBDDH,"Truly new annotated profile retains explicit manual-create workflow across all sheets");File.Delete(file);
        SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901",Work=10},new Row{Name="Trần Văn Thử (1987)",Code="901",Work=15}},false);a=Person("Nguyễn Văn Kiểm","901");b=Person("Trần Văn Thử","902");same=Unit(a,b);msg=Import(file,same,out prompts);Check(a.CongSP==10 && b.CongSP==99 && same.ListNhanVien.Count==2 && msg.Contains("SDB khác"),"One mistyped code cannot mix another person's attendance into a valid same-code row");File.Delete(file);
        var all=WorkshopData.CreateDefaultWorkshops();a=Person("Nguyễn Văn Kiểm","901");all[1].ListNhanVien.Add(a);SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="901"}});msg=Import(file,all[0],out prompts,all);Check(prompts==0 && all[0].ListNhanVien.Count==1 && all[0].ListNhanVien[0]!=a && a.CongSP==99 && all[0].ListNhanVien[0].CongSP==12 && all[0].ListNhanVien[0].HoTen==a.HoTen && all[0].ListNhanVien[0].SDB==a.SDB,"Cross-workshop profile copy also accepts year annotation while preserving original attendance");File.Delete(file);
        all=WorkshopData.CreateDefaultWorkshops();a=Person("Nguyễn Văn Kiểm","901");b=Person("Nguyễn Văn Kiểm","902");all[0].ListNhanVien.Add(a);all[1].ListNhanVien.Add(b);SheetFixture(file,new[]{new Row{Name="Nguyễn Văn Kiểm (1989)",Code="902"}},false);msg=Import(file,all[0],out prompts,all);Check(all[0].ListNhanVien.Count==1 && a.CongSP==99 && b.CongSP==99 && prompts==0 && msg.Contains("SDB khác"),"A conflicting local code is not bypassed by copying a same-name profile from another workshop");File.Delete(file);

    }
    static void SavedWorkbook(string dir)
    {
        string file=Path.Combine(dir,"master.xlsx"),input=Path.Combine(dir,"attendance.xlsx");
        using(var z=ZipFile.Open(file,ZipArchiveMode.Create)) {
            Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/></Types>");
            Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Target='xl/workbook.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'/></Relationships>");
            Entry(z,"xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='thép' sheetId='1' r:id='r1'/><sheet name='Tổng hợp công' sheetId='2' r:id='r2'/></sheets><calcPr/></workbook>");
            Entry(z,"xl/_rels/workbook.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='r1' Target='worksheets/sheet1.xml'/><Relationship Id='r2' Target='worksheets/sheet2.xml'/></Relationships>");
            foreach(int part in new[]{1,2}) {
                var d=new XmlDocument();d.LoadXml("<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData/></worksheet>");
                if(part==1){ExcelService.SetCellInlineString(d,"C5","Nguyễn Văn Kiểm",false);ExcelService.SetCellInlineString(d,"D5","901",false);ExcelService.SetCellInlineString(d,"E5","CN",false);ExcelService.SetCellNumeric(d,"G5",5200000,false);ExcelService.SetCellNumeric(d,"K5",99,false);}
                else {ExcelService.SetCellInlineString(d,"C6","901",false);ExcelService.SetCellInlineString(d,"D6","Nguyễn Văn Kiểm",false);ExcelService.SetCellNumeric(d,"K6",5200000,false);ExcelService.SetCellNumeric(d,"L6",99,false);}
                Put(z,"xl/worksheets/sheet"+part+".xml",d);
            }
        }
        SheetFixture(input,new[]{new Row {Name="Nguyễn Văn Kiểm (1989)",Code="0901"}});
        var all=ExcelService.LoadWorkshopsFromExcel(file);int prompts;Import(input,all[2],out prompts,all);
        for(int pass=0;pass<2;pass++) {
            ExcelService.SaveWorkshopsToExcel(file,all);all=ExcelService.LoadWorkshopsFromExcel(file);var w=all[2];
            Check(w.ListNhanVien.Count==1 && w.ListNhanVien[0].HoTen=="Nguyễn Văn Kiểm" && w.ListNhanVien[0].SDB=="901","Canonical profile survives annotated import and repeated save/reload without a duplicate");
            Check(w.ListNhanVien[0].CongSP==12 && w.ListNhanVien[0].CongLePhep==3 && w.ListNhanVien[0].CoAnCa && w.ListNhanVien[0].CoBDDH,"Imported attendance and auxiliary data persist to Excel");
            using(var z=ZipFile.OpenRead(file)){var d=ExcelService.LoadSheetXml(z,ExcelService.GetSheetMap(z)["Tổng hợp công"]);var ss=ExcelService.LoadSharedStrings(z);Check(ExcelService.GetCellDouble(d,ss,"L6",0)==12 && ExcelService.GetCellDouble(d,ss,"O6",0)==3 && ExcelService.GetCellDouble(d,ss,"S6",0)==5,"Company master still receives days/leave/K3 through original canonical identity");}
            Import(input,w,out prompts,all);Check(prompts==0 && w.ListNhanVien.Count==1 && w.ListNhanVien[0].CongSP==12,"Reimport after reload remains idempotent");
        }
    }
    static int Main(){string dir=Path.Combine(Path.GetTempPath(),"name-import-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);try{MainCases(dir);SavedWorkbook(dir);}finally{Directory.Delete(dir,true);}Console.WriteLine("PASS: "+checks+" name/year/code matching checks.");return 0;}
}
