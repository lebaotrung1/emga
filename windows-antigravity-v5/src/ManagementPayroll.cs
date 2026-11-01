using System;
using System.Collections.Generic;

namespace CiscoLDTLApp
{
    // Phân bổ quỹ SP đã được duyệt; không suy diễn công thức quỹ từ Excel lỗi/liên kết ngoài.
    public static class ManagementPayroll
    {
        public static void Copy(EmployeeRecord from, EmployeeRecord to)
        { foreach (var p in typeof(EmployeeRecord).GetProperties()) if (p.CanWrite) p.SetValue(to, p.GetValue(from, null), null); }

        public static double RoundMoney(double value, double unit)
        { return Math.Round(value / unit, 0, MidpointRounding.AwayFromZero) * unit; }

        public static double Weight(EmployeeRecord e)
        {
            double rate = e.XepLoai == "B" ? .9 : e.XepLoai == "C" ? .8 : 1;
            return Math.Round(Math.Round(e.CongSP * e.HeSoCoDinh / 26, 2,
                MidpointRounding.AwayFromZero) * rate, 2, MidpointRounding.AwayFromZero);
        }

        public static void Validate(WorkshopData ws)
        {
            if (double.IsNaN(ws.LuongSPQL) || double.IsInfinity(ws.LuongSPQL) || ws.LuongSPQL < 0)
                throw new ArgumentException("Quỹ lương sản phẩm quản lý phải là số không âm.");
            if (double.IsNaN(ws.DieuChinhLuongQL) || double.IsInfinity(ws.DieuChinhLuongQL))
                throw new ArgumentException("Điều chỉnh quản lý phải là số hữu hạn.");
            double weights = 0;
            foreach (var e in ws.ListNhanVien)
            {
                if (!e.LaQuanLy) continue;
                foreach (double v in new[] { e.CongSP, e.CongLePhep, e.CongHocHop, e.CongVHTT,
                    e.HeSoCoDinh, e.MucLuongBHXH, e.LuongLamLe, e.PhuCapTien })
                    if (v < 0 || double.IsNaN(v) || double.IsInfinity(v))
                        throw new ArgumentException("Công, hệ số và các khoản tiền phải là số không âm.");
                if (e.XepLoai != "A" && e.XepLoai != "B" && e.XepLoai != "C")
                    throw new ArgumentException("Xếp loại quản lý chỉ nhận A, B hoặc C.");
                if (e.CongSP > 0 && e.HeSoCoDinh <= 0)
                    throw new ArgumentException("Người có công sản phẩm cần hệ số cố định lớn hơn 0.");
                weights += Weight(e);
            }
            if (ws.LuongSPQL > 0 && weights <= 0)
                throw new ArgumentException("Có quỹ lương nhưng chưa có hệ số hưởng để phân bổ. Kiểm tra công và hệ số quản lý.");
        }

        public static void Calculate(WorkshopData ws, bool redistribute)
        {
            double weights = 0, other = 0;
            bool hasManagement = false;
            foreach (var e in ws.ListNhanVien)
            {
                if (!e.LaQuanLy) continue;
                hasManagement = true;
                e.XepLoai = (e.XepLoai ?? "A").Trim().ToUpperInvariant();
                e.DiemSanXuat = Math.Round(e.CongSP * e.HeSoCoDinh / 26, 2, MidpointRounding.AwayFromZero);
                e.HeSoHuong = Weight(e);
                weights += e.HeSoHuong;
            }
            foreach (var e in ws.ListNhanVien)
            {
                if (!e.LaQuanLy) continue;
                if (redistribute)
                    e.LuongSanPham = weights > 0 && ws.LuongSPQL > 0
                        ? RoundMoney(ws.LuongSPQL * e.HeSoHuong / weights, 100) : 0;
                e.TongCong = e.CongSP + e.CongLePhep + e.CongVHTT;
                e.LuongLePhepCD = RoundMoney(e.CongLePhep * e.MucLuongBHXH / 26, 100);
                e.LuongHocVHTT = e.CongHocHop * 250000 + e.CongVHTT * 300000;
                // Làm lễ, phụ cấp được nhập/đọc riêng, không áp một đơn giá cho mọi PX.
                e.TongTienLuong = e.LuongSanPham + e.LuongLePhepCD + e.LuongHocVHTT + e.LuongLamLe + e.PhuCapTien;
                other += e.LuongLePhepCD + e.LuongHocVHTT + e.LuongLamLe + e.PhuCapTien;
            }
            if (hasManagement) ws.LuongCheDoQL = other;
        }

        public static void Recalculate(WorkshopData ws)
        {
            Validate(ws);
            Calculate(ws, true);
            ws.CongSP = ws.CongQL = ws.CongCDQL = ws.CongLePhepCD = ws.CongBHXH = 0;
            foreach (var e in ws.ListNhanVien)
            {
                ws.CongSP += e.CongSP;
                ws.CongBHXH += e.CongBHXH;
                ws.CongLePhepCD += e.CongLePhep + e.CongHocHop + e.CongVHTT;
                if (e.LaQuanLy) { ws.CongQL += e.CongSP; ws.CongCDQL += e.CongLePhep + e.CongHocHop + e.CongVHTT; }
            }
            ws.LDThucTeBQ = Math.Round(ws.CongSP / 26, 2);
            ws.LDTTThucTeBQ = Math.Round((ws.CongSP - ws.CongQL) / 26, 2);
            ws.RecalculateDerived();
        }
    }
}
