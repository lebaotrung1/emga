using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
namespace CiscoLDTLApp
{
    public static class EmployeeClassification
    {
        static string Normalize(string role)
        {
            var result=new StringBuilder();
            foreach(char c in (role??"").Normalize(NormalizationForm.FormD))
                if(CharUnicodeInfo.GetUnicodeCategory(c)!=UnicodeCategory.NonSpacingMark) result.Append(c=='Đ'||c=='đ'?'D':char.ToUpperInvariant(c));
            return Regex.Replace(Regex.Replace(result.ToString(),@"[^A-Z0-9]+"," "),@"\s+"," ").Trim();
        }
        public static bool IsManagementRole(string role)
        {
            string r=Normalize(role);
            if(Regex.IsMatch(r,@"\b(TT|TTR|TP|TRC|TR CA|TRUONG CA|PHO TRUONG CA|PHO CA|P CA|TO TRUONG|T TRUONG|TO PHO)\b")) return false;
            return Regex.IsMatch(r,@"^(QD|PQD|Q D|P QD|P Q D)(\b|$)")
                || Regex.IsMatch(r,@"^(QUAN DOC|PHO QUAN DOC|P QUAN DOC)(\b|$)")
                || Regex.IsMatch(r,@"^(KTV|KY THUAT VIEN)(\b|$)")
                || Regex.IsMatch(r,@"^(TK|T KE|THONG KE)(\b|$)");
        }
        public static void Apply(EmployeeRecord e)
        {
            e.LaQuanLy=IsManagementRole(e.ChucDanh);
            if(!e.LaQuanLy) e.QuanLyTinhMoi=false;
        }
        public static void Apply(WorkshopData ws)
        {
            int managers=0; double work=0;
            foreach(var e in ws.ListNhanVien) { Apply(e); if(e.LaQuanLy) { managers++; work+=e.CongSP; } }
            if(ws.ListNhanVien.Count>0) { ws.LDQuanLy=managers; ws.LDTrucTiep=ws.ListNhanVien.Count-managers; ws.CongQL=work; }
        }
    }
    internal class AttendanceStamp
    {
        readonly double[] values;
        internal AttendanceStamp(EmployeeRecord e) { values=Values(e); }
        static double[] Values(EmployeeRecord e) { return new[] { e.CongSP,e.CongLePhep,e.CongHocHop,e.CongVHTT,e.CongBHXH,e.CongAnCa,e.CongK3 }; }
        internal bool Changed(EmployeeRecord e) { var current=Values(e); for(int i=0;i<values.Length;i++) if(Math.Abs(current[i]-values[i])>0.00000001) return true; return false; }
        internal static bool Active(EmployeeRecord e) { return e.CongSP!=0 || e.CongLePhep!=0 || e.CongHocHop!=0 || e.CongVHTT!=0 || e.CongBHXH!=0; }
        internal static bool Same(EmployeeRecord a,EmployeeRecord b) { var x=Values(a); var y=Values(b); for(int i=0;i<x.Length;i++) if(Math.Abs(x[i]-y[i])>0.00000001) return false; return true; }
    }
}
