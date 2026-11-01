using System;
namespace CiscoLDTLApp
{
    public static class ManagementProductionPayroll
    {
        public static void CopySettings(WorkshopData from, WorkshopData to)
        {
            to.TinhQuyQLTuSanLuong = from.TinhQuyQLTuSanLuong;
            to.DonGiaQLLoai1 = from.DonGiaQLLoai1;
            to.DonGiaQLLoai2 = from.DonGiaQLLoai2;
        }
        public static void UpdateFund(WorkshopData ws)
        {
            if (!ws.TinhQuyQLTuSanLuong) return;
            foreach (double value in new[] { ws.SanLuongLoai1, ws.SanLuongLoai2, ws.SanLuongLe, ws.DonGiaQLLoai1, ws.DonGiaQLLoai2 })
                if (value < 0 || double.IsNaN(value) || double.IsInfinity(value))
                    throw new ArgumentException("Sản lượng và đơn giá quản lý phải là số không âm.");
            ws.LuongSPQL = ws.ThanhTienQLLoai1 + ws.ThanhTienQLLoai2 + ws.ThanhTienQLLe;
        }
    }
}
