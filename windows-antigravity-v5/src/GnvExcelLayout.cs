using System;
using System.Collections.Generic;
using System.Linq;
namespace CiscoLDTLApp
{
    // These are detail rows in the supported source templates, never contiguous payroll sections.
    internal sealed class GnvExcelLayout
    {
        internal string Code, Quantity="D", Rate="I", Amount="J", TotalQuantity, TotalAmount;
        internal int[] Rows, DefaultRows, ExternalRows=new int[0], SubtotalRows=new int[0];
        internal readonly Dictionary<string,int[]> Summaries=new Dictionary<string,int[]>();
        static int[] Range(int a,int b) { return Enumerable.Range(a,b-a+1).ToArray(); }
        internal static GnvExcelLayout For(string code)
        {
            var l=new GnvExcelLayout { Code=code };
            if(code=="TK") {
                l.Rows=Range(80,90);l.SubtotalRows=new[] {79};l.TotalQuantity="D79";l.TotalAmount="J79";
            } else if(code=="LG") {
                l.Rows=new[] {61,62}.Concat(Range(64,93)).Concat(Range(95,100)).Concat(new[] {102,103}).ToArray();
                l.ExternalRows=new[] {61,62};l.DefaultRows=l.Rows.Where(r=>!l.ExternalRows.Contains(r)).ToArray();l.SubtotalRows=new[] {60,63,94,101};
                l.TotalQuantity="D59";l.TotalAmount="J59";l.Summaries["J60"]=l.ExternalRows;l.Summaries["J63"]=l.Rows.Where(r=>!l.ExternalRows.Contains(r)).ToArray();
            } else if(code=="LT") {l.Rows=Range(60,97);l.TotalQuantity="D59";l.TotalAmount="J59";}
            else if(code=="CDNL") {l.Rows=Range(372,393);l.Rate="H";l.Amount="I";l.TotalQuantity="D371";l.TotalAmount="I371";}
            else if(code=="KT") {
                l.Rows=Range(60,63).Concat(Range(65,79)).Concat(new[] {82,83}).Concat(Range(85,87)).Concat(Range(89,94)).Concat(Range(96,101)).ToArray();
                l.DefaultRows=l.Rows.Where(r=>r>=65).Concat(Range(60,63)).ToArray();l.SubtotalRows=new[] {59,64,80,81,84,88,95};
                l.TotalQuantity="D52";l.TotalAmount="J52";
                l.Summaries["D59"]=Range(60,63);l.Summaries["J59"]=Range(60,63);
                l.Summaries["D64"]=l.Rows.Where(r=>r>=65).ToArray();l.Summaries["J64"]=l.Summaries["D64"];
                l.Summaries["J58"]=l.Rows;l.Summaries["D80"]=new[] {82,83,85,86};l.Summaries["J80"]=l.Summaries["D80"];
                l.Summaries["D81"]=new[] {82,83};l.Summaries["J81"]=l.Summaries["D81"];
                l.Summaries["D84"]=new[] {85,86};l.Summaries["J84"]=l.Summaries["D84"];
                l.Summaries["D88"]=new[] {89,90};l.Summaries["J88"]=l.Summaries["D88"];
                l.Summaries["D95"]=Range(96,99);l.Summaries["J95"]=Range(96,99);
            } else return null;
            if(l.DefaultRows==null) l.DefaultRows=l.Rows;
            return l;
        }
        internal Dictionary<int,ArisingWorkItem> Plan(WorkshopData ws)
        {
            if(ws.ListGNV.Count>Rows.Length) throw new ArgumentException("GNV "+ws.TenPX+" có "+ws.ListGNV.Count+" dòng, vượt "+Rows.Length+" dòng chi tiết của mẫu Excel được hỗ trợ. Chưa lưu tệp để tránh thiếu dữ liệu; cần giảm số mục hoặc dùng mẫu được hỗ trợ đủ dòng.");
            var used=new Dictionary<int,ArisingWorkItem>();var pending=new List<ArisingWorkItem>();
            foreach(var item in ws.ListGNV) {
                if(Rows.Contains(item.ExcelRow) && !used.ContainsKey(item.ExcelRow)) used[item.ExcelRow]=item;
                else if(ExternalRows.Contains(item.ExcelRow)) throw new ArgumentException("GNV "+ws.TenPX+" trùng dòng chuyển thanh toán PX khác. Chưa lưu để tránh chuyển nhầm khoản sang quỹ của PX.");
                else pending.Add(item);
            }
            foreach(var item in pending) {
                int r=DefaultRows.FirstOrDefault(x=>!used.ContainsKey(x));
                if(r==0) throw new ArgumentException("GNV "+ws.TenPX+" không còn dòng chi tiết trống trong nhóm thanh toán cho PX. Chưa lưu; không chuyển mục sang nhóm thanh toán PX khác để tránh sai quỹ.");
                used[r]=item;
            }
            return used;
        }
        internal int[] PayrollRows {get {return Rows.Where(r=>!ExternalRows.Contains(r)).ToArray();}}
        internal IEnumerable<KeyValuePair<string,int[]>> Totals()
        {
            foreach(var p in Summaries) yield return p;
            yield return new KeyValuePair<string,int[]>(TotalQuantity,PayrollRows);
            yield return new KeyValuePair<string,int[]>(TotalAmount,PayrollRows);
        }
        internal string Formula(string address,int[] rows) {string col=address.StartsWith("D")?Quantity:Amount;return "SUM("+string.Join(",",rows.Select(r=>col+r).ToArray())+")";}
        internal void Bind(PayrollFormulaBook book,WorkshopData ws)
        {
            if(!ws.GnvDaChinhSua) return;
            var plan=Plan(ws);string sheet=book.Sheets.Contains(ws.SheetNT)?ws.SheetNT:ws.SheetNT.Trim();
            foreach(int r in Rows) {ArisingWorkItem item;plan.TryGetValue(r,out item);book.Set(sheet,Quantity+r,item==null?0:item.KhoiLuong);book.Set(sheet,Rate+r,item==null?0:item.DonGia);book.Set(sheet,Amount+r,item==null?0:item.ThanhTien);}
            foreach(var p in Totals()) {double sum=0;foreach(int r in p.Value){ArisingWorkItem item;if(plan.TryGetValue(r,out item))sum+=p.Key.StartsWith("D")?item.KhoiLuong:item.ThanhTien;}book.Set(sheet,p.Key,sum);}
        }
        internal static bool InPayroll(WorkshopData ws,ArisingWorkItem item) {return ws.MaPX!="LG" || (item.ExcelRow!=61 && item.ExcelRow!=62);}
    }
}
