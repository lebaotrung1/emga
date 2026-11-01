using System;
namespace CiscoLDTLApp
{
    public enum ProductionLine { Loai1, Loai2, NgayLe }
    public static class ProductionPayroll
    {
        public static void UpdateLine(WorkshopData ws, ProductionLine line, double quantity, double unitPrice)
        {
            if (quantity < 0 || unitPrice < 0 || double.IsNaN(quantity) || double.IsNaN(unitPrice)
                || double.IsInfinity(quantity) || double.IsInfinity(unitPrice))
                throw new ArgumentException("Sản lượng và đơn giá phải là số không âm.");
            double unit = ws.MaPX == "LG" && line == ProductionLine.NgayLe ? 1000 : 100;
            double amount = ManagementPayroll.RoundMoney(quantity * unitPrice, unit);
            double old;
            if (line == ProductionLine.Loai1)
            { old=ws.ThanhTienLoai1; ws.SanLuongLoai1=quantity; ws.DonGiaLoai1=unitPrice; ws.ThanhTienLoai1=amount; }
            else if (line == ProductionLine.Loai2)
            { old=ws.ThanhTienLoai2; ws.SanLuongLoai2=quantity; ws.DonGiaLoai2=unitPrice; ws.ThanhTienLoai2=amount; }
            else
            { old=ws.ThanhTienLe; ws.SanLuongLe=quantity; ws.DonGiaLe=unitPrice; ws.ThanhTienLe=amount; }
            // Các dòng khác trong quỹ (GNV, hỗ trợ...) được giữ; chỉ thay đóng góp của dòng vừa sửa.
            ws.LuongSPTT += amount-old;
            ws.RecalculateDerived();
        }
    }
}
