using System;
using System.Collections.Generic;
namespace CiscoLDTLApp
{
    public class CraneStage
    {
        public string Name { get; set; }
        public double Regular { get; set; }
        public double Holiday { get; set; }
        public double Rate { get; set; }
        public double Amount { get { return Regular * Rate + Holiday * Rate * 2; } }
    }
    public static class CranePayroll
    {
        public static CraneStage[] CreateStages()
        {
            return new[] {
                new CraneStage { Name="Vận chuyển phôi từ SX đến xếp kiêu ngoài bãi" },
                new CraneStage { Name="Xuất bán phôi từ kiêu ngoài bãi" },
                new CraneStage { Name="Vệ sinh phôi, xuất bán hàng từ kho sản xuất" }
            };
        }
        public static void Apply(WorkshopData ws, CraneStage[] stages, bool recalculateEmployees = true)
        {
            if(ws.MaPX!="LT" || stages==null || stages.Length!=3) throw new ArgumentException("Cẩu phôi chỉ áp dụng 3 công đoạn PX Thép.");
            double total=0, quantity=0;
            foreach(var stage in stages)
            {
                if(stage==null) throw new ArgumentException("Thiếu công đoạn.");
                foreach(double value in new[] {stage.Regular,stage.Holiday,stage.Rate,stage.Amount})
                    if(value<0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("Sản lượng và đơn giá phải là số hữu hạn không âm.");
                total+=stage.Amount; quantity+=stage.Regular+stage.Holiday;
            }
            if(double.IsInfinity(total) || double.IsInfinity(quantity)) throw new ArgumentException("Số liệu vượt giới hạn.");
            // D44/J44 là thành phần cẩu phôi đã nằm trong quỹ SP, thay thế thay vì cộng thêm.
            ws.LuongSPTT += total+ws.CauPhoiKhoanKhac-ws.SanLuongPhuTT;
            for(int i=0;i<3;i++) { ws.CauPhoi[i].Regular=stages[i].Regular; ws.CauPhoi[i].Holiday=stages[i].Holiday; ws.CauPhoi[i].Rate=stages[i].Rate; }
            ws.SanLuongPhuName="Cẩu phôi (3 công đoạn)";
            ws.SanLuongPhuKL=quantity; ws.SanLuongPhuTT=total+ws.CauPhoiKhoanKhac; ws.CauPhoiDaApDung=true;
            if(recalculateEmployees) ws.RecalculateFromEmployees(true); else ws.RecalculateDerived();
        }
        public static IEnumerable<string> ReportLines(WorkshopData ws)
        {
            if(ws.MaPX!="LT" || !ws.CauPhoiDaApDung) yield break;
            yield return "Cẩu phôi tháng: SL thường không bao gồm SL lễ; tiền = SL thường × đơn giá + SL lễ × đơn giá × 2. Quỹ đã nằm trong lương sản phẩm trực tiếp.";
            foreach(var stage in ws.CauPhoi)
                yield return string.Format("+ {0}: SL thường {1}; SL lễ {2}; đơn giá {3} đồng/ĐVT; thành tiền {4} triệu đồng.",stage.Name,WordReportService.FmtDec(stage.Regular,3),WordReportService.FmtDec(stage.Holiday,3),WordReportService.FmtDec(stage.Rate,2),WordReportService.FmtMoney(stage.Amount));
            yield return "Khoản cẩu phôi khác giữ từ Excel (ngoài 3 công đoạn): "+WordReportService.FmtMoney(ws.CauPhoiKhoanKhac)+" triệu đồng.";
            yield return "Tổng quỹ cẩu phôi tháng: "+WordReportService.FmtMoney(ws.SanLuongPhuTT)+" triệu đồng (đã bao gồm trong quỹ SP, không cộng thêm).";
        }
    }
}
