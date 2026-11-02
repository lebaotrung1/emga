using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
namespace CiscoLDTLApp
{
    // Resolve one author for each company-wide identity; never let workshop iteration order decide.
    internal class MasterAttendancePlan
    {
        internal readonly Dictionary<int,EmployeeRecord> Writers=new Dictionary<int,EmployeeRecord>();
        internal static string Identity(string code,string name)
        {
            code=(code??"").Trim(); long number;
            if(long.TryParse(code,out number)) { if(number<=0) return ""; code=number.ToString(CultureInfo.InvariantCulture); }
            if(code=="" || code.StartsWith("#") || string.IsNullOrWhiteSpace(name)) return "";
            return code.ToUpperInvariant()+"|"+Regex.Replace(name.Trim(),@"\s+"," ").ToUpperInvariant();
        }
        internal static int ResolveHint(string formula,int localRow,string identity,Dictionary<string,List<int>> index)
        {
            if(string.IsNullOrEmpty(formula) || identity=="") return 0;
            string prefix=@"'Tổng hợp công'!";
            var match=Regex.Match(formula,@"MATCH\s*\(\s*\$?D\$?"+localRow+@"\s*,\s*"+prefix+@"\$?C\$?(\d+):\$?C\$?(\d+)\s*,\s*0\s*\)",RegexOptions.IgnoreCase);
            var range=Regex.Match(formula,@"INDEX\s*\(\s*"+prefix+@"\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)\s*,",RegexOptions.IgnoreCase);
            if(!match.Success || !range.Success) return 0;
            int start=int.Parse(match.Groups[1].Value),end=int.Parse(match.Groups[2].Value),first=int.MaxValue;
            string code=identity.Substring(0,identity.IndexOf('|'))+"|";
            foreach(var pair in index) if(pair.Key.StartsWith(code,StringComparison.OrdinalIgnoreCase)) foreach(int r in pair.Value) if(r>=start && r<=end) first=Math.Min(first,r);
            if(first==int.MaxValue) return 0;
            int resolved=int.Parse(range.Groups[1].Value)+first-start;
            List<int> exact; return resolved<=int.Parse(range.Groups[2].Value) && index.TryGetValue(identity,out exact) && exact.Contains(resolved)?resolved:0;
        }
        internal static MasterAttendancePlan Build(List<WorkshopData> workshops,Dictionary<string,List<int>> index)
        {
            var result=new MasterAttendancePlan(); var groups=new Dictionary<int,List<EmployeeRecord>>();
            foreach(var ws in workshops) foreach(var e in ws.ListNhanVien) {
                List<int> rows; string key=Identity(e.SDB,e.HoTen);
                if(key=="" || !index.TryGetValue(key,out rows)) continue;
                int target=rows.Count==1?rows[0]:e.MasterRowHint;
                if(target<=0 || !rows.Contains(target)) throw new ArgumentException("Tên + SDB "+e.SDB+" có nhiều dòng trong Tổng hợp công và chưa xác định được từ vùng tra cứu gốc. Chưa lưu để tránh chọn nhầm dòng.");
                List<EmployeeRecord> candidates; if(!groups.TryGetValue(target,out candidates)) { candidates=new List<EmployeeRecord>(); groups[target]=candidates; }
                candidates.Add(e);
            }
            foreach(var group in groups) {
                var active=new List<EmployeeRecord>(); var changed=new List<EmployeeRecord>();
                foreach(var e in group.Value) { if(AttendanceStamp.Active(e)) active.Add(e); if(e.AttendanceEdited || e.AttendanceBaseline==null || e.AttendanceBaseline.Changed(e)) changed.Add(e); }
                var options=active.Count>0?active:changed;
                if(options.Count==0) continue; // Unchanged empty roster must not erase a populated company row.
                var owner=options[0];
                foreach(var other in options) if(!AttendanceStamp.Same(owner,other)) throw new ArgumentException("SDB "+owner.SDB+" có công khác nhau ở nhiều PX. Chưa lưu để tránh ghi đè; đối chiếu và giữ dữ liệu đúng trước khi lưu lại.");
                foreach(var e in changed) if(options.Contains(e)) { owner=e; break; }
                result.Writers[group.Key]=owner;
            }
            return result;
        }
        internal static Dictionary<string,List<int>> Index(XmlDocument doc,List<string> strings)
        {
            var result=new Dictionary<string,List<int>>(StringComparer.OrdinalIgnoreCase);
            if(doc==null) return result;
            foreach(XmlElement row in doc.GetElementsByTagName("row","http://schemas.openxmlformats.org/spreadsheetml/2006/main")) {
                int r; if(!int.TryParse(row.GetAttribute("r"),out r)||r<6) continue;
                Add(result,Identity(ExcelService.GetCellString(doc,strings,"C"+r),ExcelService.GetCellString(doc,strings,"D"+r)),r);
            }
            return result;
        }
        static void Add(Dictionary<string,List<int>> index,string key,int row) { if(key=="") return; List<int> rows; if(!index.TryGetValue(key,out rows)) { rows=new List<int>(); index[key]=rows; } rows.Add(row); }
        internal static MasterAttendancePlan FromBook(PayrollFormulaBook book,List<WorkshopData> workshops)
        {
            var index=new Dictionary<string,List<int>>(StringComparer.OrdinalIgnoreCase);
            foreach(var pair in book.Cells) {
                var c=pair.Value; if(!c.Sheet.Equals("Tổng hợp công",StringComparison.OrdinalIgnoreCase) || !Regex.IsMatch(c.Address,@"^C\d+$")) continue;
                PayrollCell name; if(!book.Cells.TryGetValue(c.Sheet+"!D"+c.Address.Substring(1),out name)) continue;
                Add(index,Identity(c.IsNumber?c.Number.ToString(CultureInfo.InvariantCulture):c.Text,name.Text),int.Parse(c.Address.Substring(1)));
            }
            foreach(var ws in workshops) foreach(var e in ws.ListNhanVien) {
                List<int> rows; string key=Identity(e.SDB,e.HoTen);
                if(index.TryGetValue(key,out rows) && rows.Count>1 && e.MasterRowHint<=0) { PayrollCell cell; if(book.Cells.TryGetValue(ws.SheetBangCong+"!K"+e.ExcelRow,out cell)) e.MasterRowHint=ResolveHint(cell.Formula,e.ExcelRow,key,index); if(e.MasterRowHint<=0 && book.Cells.TryGetValue(ws.SheetBangCong+"!N"+e.ExcelRow,out cell)) e.MasterRowHint=ResolveHint(cell.Formula,e.ExcelRow,key,index); }
            }
            return Build(workshops,index);
        }
        internal void Bind(PayrollFormulaBook book)
        {
            foreach(var pair in Writers) {
                var e=pair.Value; string r=pair.Key.ToString();
                book.Set("Tổng hợp công","L"+r,e.CongSP); book.Set("Tổng hợp công","M"+r,e.CongHocHop);
                book.Set("Tổng hợp công","N"+r,e.CongVHTT); book.Set("Tổng hợp công","O"+r,e.CongLePhep);
                book.Set("Tổng hợp công","T"+r,e.CongBHXH);
                if(e.CoBangCong) book.Set("Tổng hợp công","S"+r,e.CongK3);
                if(e.MucLuongBHXH>0) book.Set("Tổng hợp công","K"+r,e.MucLuongBHXH);
            }
        }
        internal void Write(XmlDocument doc,List<string> strings)
        {
            foreach(var pair in Writers) {
                var e=pair.Value; string r=pair.Key.ToString();
                if(e.MucLuongBHXH>0) ExcelService.SetCellNumeric(doc,"K"+r,e.MucLuongBHXH);
                ExcelService.SetCellNumeric(doc,"L"+r,e.CongSP); ExcelService.SetCellNumeric(doc,"M"+r,e.CongHocHop);
                ExcelService.SetCellNumeric(doc,"N"+r,e.CongVHTT); ExcelService.SetCellNumeric(doc,"O"+r,e.CongLePhep);
                double extra=ExcelService.GetCellDouble(doc,strings,"P"+r,0)+ExcelService.GetCellDouble(doc,strings,"Q"+r,0);
                ExcelService.SetCellNumeric(doc,"R"+r,e.CongHocHop+e.CongVHTT+e.CongLePhep+extra);
                ExcelService.SetCellNumeric(doc,"T"+r,e.CongBHXH);
                if(e.CoBangCong) ExcelService.SetCellNumeric(doc,"S"+r,e.CongK3);
                if(e.DiemSanXuat>0 || e.QuanLyTinhMoi || e.ImportedPoints!=null) ExcelService.SetCellNumeric(doc,"U"+r,e.DiemSanXuat);
                ExcelService.SetCellNumeric(doc,"W"+r,e.LuongSanPham); ExcelService.SetCellNumeric(doc,"AB"+r,e.LuongLePhepCD);
                ExcelService.SetCellNumeric(doc,"AF"+r,e.CongAnCa);
            }
        }
    }
}
