using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml;

namespace NghiemThuStudio
{
    public sealed class SourceCell
    {
        public string Address;
        public string Raw;
        public string Formula;
        public decimal? Number;
        public string State;
    }
    public sealed class Workbook : IDisposable
    {
        const string Main="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        readonly ZipArchive zip;
        readonly List<string> strings=new List<string>();
        readonly Dictionary<string,string> sheets=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string,Dictionary<string,SourceCell>> loaded=new Dictionary<string,Dictionary<string,SourceCell>>();
        public readonly List<string> External=new List<string>();
        static XmlDocument Xml(Stream stream)
        {
            var settings=new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null };
            var doc=new XmlDocument { XmlResolver=null };
            using(var reader=XmlReader.Create(stream,settings)) doc.Load(reader); return doc;
        }
        XmlDocument Part(string name) { var entry=zip.GetEntry(name); if(entry==null) throw new InvalidDataException("Thiếu phần Excel: "+name); using(var s=entry.Open()) return Xml(s); }
        public Workbook(string path)
        {
            if(!string.Equals(Path.GetExtension(path),".xlsx",StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Chỉ đọc .xlsx; hãy chuyển .xls thành .xlsx bằng Excel.");
            zip=ZipFile.OpenRead(path);
            try
            {
                var rel=Part("xl/_rels/workbook.xml.rels"); var targets=new Dictionary<string,string>();
                foreach(XmlElement e in rel.DocumentElement.ChildNodes) targets[e.GetAttribute("Id")]=e.GetAttribute("Target");
                var wb=Part("xl/workbook.xml");
                foreach(XmlElement e in wb.GetElementsByTagName("sheet",Main))
                {
                    string target; if(!targets.TryGetValue(e.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships"),out target)) continue;
                    target=target.TrimStart('/'); if(!target.StartsWith("xl/")) target="xl/"+target;
                    sheets[e.GetAttribute("name").Trim()]=target;
                }
                if(zip.GetEntry("xl/sharedStrings.xml")!=null)
                    foreach(XmlElement si in Part("xl/sharedStrings.xml").GetElementsByTagName("si",Main))
                        strings.Add(si.InnerText);
                foreach(var part in zip.Entries.Where(e=>e.FullName.StartsWith("xl/externalLinks/_rels/",StringComparison.OrdinalIgnoreCase)))
                    using(var s=part.Open()) foreach(XmlElement e in Xml(s).DocumentElement.ChildNodes)
                        External.Add(Uri.UnescapeDataString(e.GetAttribute("Target")));
            }
            catch { zip.Dispose(); throw; }
        }
        public int SheetCount { get { return sheets.Count; } }
        static XmlElement Child(XmlElement parent,string local)
        { foreach(XmlNode n in parent.ChildNodes) if(n is XmlElement && n.LocalName==local && n.NamespaceURI==Main) return (XmlElement)n; return null; }
        public Dictionary<string,SourceCell> Cells(string sheet)
        {
            sheet=sheet.Trim(); Dictionary<string,SourceCell> cells;
            if(loaded.TryGetValue(sheet,out cells)) return cells;
            string path; if(!sheets.TryGetValue(sheet,out path)) throw new InvalidDataException("Không có sheet '"+sheet+"'. Không tự chọn sheet tên gần giống.");
            cells=new Dictionary<string,SourceCell>(StringComparer.OrdinalIgnoreCase);
            foreach(XmlElement e in Part(path).GetElementsByTagName("c",Main))
            {
                string type=e.GetAttribute("t"); var value=Child(e,"v");
                string raw=value==null?"":value.InnerText;
                if(type=="s") { int index; raw=int.TryParse(raw,out index)&&index>=0&&index<strings.Count?strings[index]:""; }
                else if(type=="inlineStr") raw=Child(e,"is")==null?"":Child(e,"is").InnerText;
                var f=Child(e,"f"); decimal n;
                var c=new SourceCell { Address=e.GetAttribute("r"),Raw=raw,Formula=f==null?null:f.InnerText };
                c.Number=type!="e" && type!="b" && decimal.TryParse(raw,NumberStyles.Float,CultureInfo.InvariantCulture,out n)?(decimal?)n:null;
                c.State=type=="e"?"Lỗi Excel: "+raw:c.Number.HasValue?(f==null?"Giá trị nguồn":"Kết quả lưu trong Excel — chưa chạy lại công thức"):(string.IsNullOrWhiteSpace(raw)?"Ô trống/thiếu kết quả":"Không phải số");
                cells[c.Address]=c;
            }
            loaded[sheet]=cells; return cells;
        }
        public SourceCell Cell(string sheet,string address)
        { SourceCell c; return Cells(sheet).TryGetValue(address,out c)?c:new SourceCell { Address=address,Raw="",State="Ô trống/không tồn tại" }; }
        public void Dispose() { zip.Dispose(); }
        public static string Hash(string path)
        { using(var s=File.OpenRead(path)) using(var h=SHA256.Create()) return BitConverter.ToString(h.ComputeHash(s)).Replace("-",""); }
        public static bool Unchanged(Dossier d)
        { return string.IsNullOrEmpty(d.SourcePath) || (File.Exists(d.SourcePath) && Hash(d.SourcePath)==d.SourceHash); }
        static void Map(Dossier d,Workbook w,string sheet,string key,string address)
        {
            var m=d.Input(key); m.SourceCell=sheet+"!"+address; var values=address.Split(new char[]{'+'},StringSplitOptions.RemoveEmptyEntries).Select(a=>w.Cell(sheet,a)).ToArray();
            m.Value=m.SourceValue=values.All(c=>c.Number.HasValue)?values.Sum(c=>c.Number.Value):(decimal?)null;
            m.Formula=string.Join(" | ",values.Select(c=>c.Address+": "+(c.Formula??"(giá trị)")));
            m.SourceState=string.Join("; ",values.Select(c=>c.State).Distinct());
        }
        static void Ref(Dossier d,Workbook w,string sheet,string key,string cell) { d.References[key]=w.Cell(sheet,cell).Number; }
        public static Dossier Import(string path,string code,int year,int month,string dependencyFolder=null)
        {
            var d=Catalog.New(code,year,month); d.SourcePath=Path.GetFullPath(path); d.SourceHash=Hash(path);
            using(var w=new Workbook(path))
            {
                string sheet=Catalog.Units.First(x=>x[0]==code)[2]; var cells=w.Cells(sheet);
                d.Diagnostics.Add(w.SheetCount+" sheet; "+w.External.Count+" liên kết ngoài. Không tự tính toàn bộ công thức Excel.");
                d.Diagnostics.Add(cells.Values.Count(c=>c.State.StartsWith("Lỗi Excel"))+" ô lỗi trong sheet NT đang chọn; ô đã ánh xạ bị lỗi sẽ để trống, không coi là 0.");
                foreach(string target in w.External)
                {
                    string leaf=target.Replace('\\','/').Split(new char[]{'/'},StringSplitOptions.RemoveEmptyEntries).LastOrDefault()??target;
                    bool found=!string.IsNullOrEmpty(dependencyFolder)&&Directory.Exists(dependencyFolder)&&Directory.EnumerateFiles(dependencyFolder).Any(f=>string.Equals(Path.GetFileName(f),leaf,StringComparison.OrdinalIgnoreCase));
                    d.Diagnostics.Add((found?"Tìm thấy tệp tham chiếu (chưa tính lại): ":"Chưa tìm thấy tệp tham chiếu: ")+leaf);
                }
                // Header is evidence of period; a mismatch is explicit and must be corrected, never inferred from filename alone.
                string header=string.Join(" ",cells.Values.Where(c=>Regex.IsMatch(c.Address,@"^[A-Z]+[1-3]$")).Select(c=>c.Raw));
                var date=Regex.Match(header,@"tháng\s*(\d{1,2}).*?năm\s*(20\d{2})",RegexOptions.IgnoreCase|RegexOptions.Singleline);
                if(date.Success && (int.Parse(date.Groups[1].Value)!=month || int.Parse(date.Groups[2].Value)!=year))
                    d.Diagnostics.Add("KỲ NGUỒN KHÁC KỲ CHỌN: "+date.Groups[1].Value+"/"+date.Groups[2].Value+". Chọn lại kỳ hoặc giải trình rõ trước khi chốt.");
                string[] labor=code=="TK"?new[]{"D33","D35","D29","D27","D31","D32"}:code=="LG"?new[]{"D15","D17","D12","D11","D14","D13"}:code=="LT"?new[]{"D14","D16","D11","D10","D12","D13"}:code=="CDNL"?new[]{"D18","D20","D14","D12","D16","D17"}:new[]{"D16","D18","D12","D11","D14","D13"};
                var keys=new[]{"head","headm","dt","dm","other","bh"}; for(int i=0;i<keys.Length;i++) Map(d,w,sheet,keys[i],labor[i]);
                string[] product=code=="TK"?new[]{"D60","D61","D63","I60","I61"}:code=="LG"?new[]{"D52","D53","D55","I52","I53"}:code=="LT"?new[]{"D37","D38","D43","I37","I38"}:null;
                if(product!=null) { keys=new[]{"q1","q2","qh","p1","p2"}; for(int i=0;i<keys.Length;i++) Map(d,w,sheet,keys[i],product[i]); }
                if(product!=null)
                {
                    string[] amountCells=code=="TK"?new[]{"J60","J61","J63"}:code=="LG"?new[]{"J52","J53","J55"}:new[]{"J37","J38","J43"};
                    string[] roundKeys={"rt1","rt2","rth"};
                    for(int i=0;i<3;i++)
                    {
                        var f=w.Cell(sheet,amountCells[i]); var match=Regex.Match(f.Formula??"",@",\s*-(\d{1,2})\s*\)");
                        if(match.Success && int.Parse(match.Groups[1].Value)<=6)
                        { var m=d.Input(roundKeys[i]); m.Value=(decimal)Math.Pow(10,int.Parse(match.Groups[1].Value)); m.SourceCell=sheet+"!"+amountCells[i]; m.SourceState="Mức làm tròn đọc từ công thức nguồn"; m.Formula=f.Formula; }
                    }
                }
                if(product==null) d.Diagnostics.Add("PX này không có cùng sơ đồ sản lượng L1/L2/lễ; nhập cơ sở khối lượng/đơn giá riêng đã xác nhận. Không biến quỹ tiền bằng đồng thành sản lượng tấn.");
                d.Diagnostics.Add("Đơn giá quản lý L1/L2 cần nhập theo phê duyệt; ngày lễ tự gấp đôi L1. Không suy đơn giá từ quỹ cũ.");
                string[] money=code=="TK"?new[]{"J91","J114+J116+J117","R126","J56","J95","J126","L126"}:code=="LG"?new[]{"J104","J140","S161","J47","J108","J161","L161"}:code=="LT"?new[]{"J98","J130","Q137","J34","J102","J137","L137"}:code=="CDNL"?new[]{"I417","I432","Q440","I42","I421","I440","K440"}:new[]{"J102","J128","Q131","J38","J106","J131","L131"};
                Map(d,w,sheet,"cdt",money[0]); Map(d,w,sheet,"cdm",money[1]); Map(d,w,sheet,"meal",money[2]);
                Ref(d,w,sheet,"direct",money[3]); Ref(d,w,sheet,"management",money[4]); Ref(d,w,sheet,"total",money[5]); Ref(d,w,sheet,"meal",money[6]);
                if(code=="TK") Map(d,w,sheet,"at","J75");
                int br=code=="TK"?180:code=="LG"?181:code=="LT"?182:code=="CDNL"?183:179;
                if(w.sheets.ContainsKey("TT BDĐH")) { Map(d,w,"TT BDĐH","b1","E"+br); Map(d,w,"TT BDĐH","b2","F"+br); Map(d,w,"TT BDĐH","b3","G"+br); Ref(d,w,"TT BDĐH","bddh","H"+br); }
                int first=code=="TK"?81:code=="LG"?60:code=="LT"?61:code=="CDNL"?372:0;
                int last=code=="TK"?89:code=="LG"?75:code=="LT"?80:code=="CDNL"?393:0;
                for(int row=first;row>0&&row<=last;row++)
                {
                    var name=w.Cell(sheet,"B"+row); var qty=w.Cell(sheet,(code=="CDNL"?"D":"D")+row);
                    var price=w.Cell(sheet,(code=="CDNL"?"H":"I")+row); var amount=w.Cell(sheet,(code=="CDNL"?"I":"J")+row);
                    if(string.IsNullOrWhiteSpace(name.Raw) || !((qty.Number??0)!=0 || (amount.Number??0)!=0)) continue;
                    decimal round=1; var rounding=Regex.Match(amount.Formula??"",@",\s*-(\d+)\s*\)"); if(rounding.Success) round=(decimal)Math.Pow(10,int.Parse(rounding.Groups[1].Value));
                    d.Works.Add(new WorkLine { Description=name.Raw,Unit=w.Cell(sheet,"C"+row).Raw,Quantity=qty.Number,Price=price.Number,SourceAmount=amount.Number,SourceCell=sheet+"!"+amount.Address,SourceFormula=amount.Formula,Factor=(amount.Formula??"").TrimStart().StartsWith("-")?-1:1,Rounding=round });
                }
                if(code=="LT")
                {
                    var qty=w.Cell(sheet,"D44"); var price=w.Cell(sheet,"I44"); var amount=w.Cell(sheet,"J44");
                    if((qty.Number??0)!=0 || (amount.Number??0)!=0)
                        d.Works.Add(new WorkLine { Kind="SP phụ",Description="Sản phẩm phụ / cẩu phôi (đối chiếu NT)",Unit=d.QuantityUnit,Quantity=qty.Number,Price=price.Number,Rounding=100,SourceAmount=amount.Number,SourceCell=sheet+"!J44",SourceFormula=amount.Formula });
                }
                if(code=="CDNL") d.Diagnostics.Add("GNV CĐNL có thể đã nằm trong quỹ bộ phận: phải đối chiếu khoản cộng thêm để tránh cộng trùng. Bản này không tự tái tạo công thức quỹ đặc thù.");
                d.Note("Nhập nguồn","Đọc chỉ xem; mã SHA256 "+d.SourceHash+"; sheet "+sheet);
            }
            return d;
        }
    }
}
