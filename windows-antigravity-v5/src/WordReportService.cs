using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace CiscoLDTLApp
{
    public class WordReportService
    {
        private static readonly CultureInfo ViCulture = new CultureInfo("vi-VN");

        public static string FmtInt(double v)
        {
            return Math.Round(v, 0).ToString("#,##0", ViCulture);
        }

        public static string FmtDec(double v, int maxDecimals = 2)
        {
            double r = Math.Round(v, maxDecimals);
            if (Math.Abs(r - Math.Round(r, 0)) < 0.0001)
                return Math.Round(r, 0).ToString("#,##0", ViCulture);
            if (maxDecimals == 3)
                return r.ToString("#,##0.###", ViCulture);
            if (maxDecimals == 1)
                return r.ToString("#,##0.#", ViCulture);
            return r.ToString("#,##0.##", ViCulture);
        }

        public static string GetTitlePeriodText(WorkshopData ws)
        {
            switch (ws.LoaiKy)
            {
                case LoaiKyBaoCao.QuyI:
                    return string.Format("tháng {0}, Quý I năm {1}", ws.Thang, ws.Nam);
                case LoaiKyBaoCao.QuyII_6Thang:
                    return string.Format("tháng {0}, Quý II, 6 tháng đầu năm {1}", ws.Thang, ws.Nam);
                case LoaiKyBaoCao.QuyIII_9Thang:
                    return string.Format("tháng {0}, Quý III, 9 tháng năm {1}", ws.Thang, ws.Nam);
                case LoaiKyBaoCao.QuyIV_CaNam:
                    return string.Format("tháng {0}, Quý IV và cả năm {1}", ws.Thang, ws.Nam);
                default:
                    return string.Format("tháng {0} năm {1}", ws.Thang, ws.Nam);
            }
        }

        public static string GeneratePreviewText(WorkshopData ws)
        {
            ws.RecalculateDerived();
            var sb = new StringBuilder();
            bool isMultiPeriod = ws.LoaiKy != LoaiKyBaoCao.ThangThuong;
            string tenQuy = ws.KyQuy != null && !string.IsNullOrEmpty(ws.KyQuy.TenKy) ? ws.KyQuy.TenKy : "Quý III";
            string tenLK = ws.KyLuyKe != null && !string.IsNullOrEmpty(ws.KyLuyKe.TenKy) ? ws.KyLuyKe.TenKy : "9 tháng";

            sb.AppendLine("CÔNG TY CP GANG THÉP CAO BẰNG               CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
            sb.AppendLine("PHÒNG TỔ CHỨC HÀNH CHÍNH                        Độc lập - Tự do - Hạnh phúc");
            sb.AppendLine("--------------------------------------------------------------------------------");
            sb.AppendLine("                                   BIÊN BẢN");
            sb.AppendLine(string.Format("   ĐÁNH GIÁ CÔNG TÁC LAO ĐỘNG TIỀN LƯƠNG {0}", GetTitlePeriodText(ws).ToUpperInvariant()));
            sb.AppendLine(string.Format("                    ĐƠN VỊ: {0}", ws.TenPX.ToUpperInvariant()));
            sb.AppendLine();
            sb.AppendLine(string.Format("Hôm nay, ngày {0:00} tháng {1:00} năm {2}, tại {3} Công ty cổ phần Gang thép Cao Bằng, chúng tôi gồm:", ws.NgayHop, ws.ThangHop, ws.NamHop, ws.DiaDiem));
            sb.AppendLine("I. Thành phần:");
            sb.AppendLine("1. Phòng Tổ chức hành chính:");
            foreach (var p in ws.ThanhPhanTCHC)
                sb.AppendLine(string.Format("   - {0,-32} - Chức vụ: {1}", p.HoTen, p.ChucVu));
            sb.AppendLine(string.Format("2. {0}:", ws.TenPX));
            foreach (var p in ws.ThanhPhanPX)
                sb.AppendLine(string.Format("   - {0,-32} - Chức vụ: {1}", p.HoTen, p.ChucVu));
            sb.AppendLine();
            sb.AppendLine("II. Đánh giá kết quả thực hiện công tác LĐTL:");
            sb.AppendLine("Cùng nhau đánh giá công tác quản lý lao động, tiền lương, ăn ca, bồi dưỡng độc hại của Phân xưởng như sau:");
            sb.AppendLine();
            sb.AppendLine("1. Công tác quản lý, sử dụng lao động:");
            sb.AppendLine(string.Format("- Số lao động đầu kỳ (đầu tháng {0}/{1}): {2} người.", ws.Thang, ws.Nam, ws.LDDauKy));
            if (ws.LDTang == 0 && ws.LDGiam == 0)
            {
                sb.AppendLine(string.Format("- Biến động lao động trong kỳ: Tăng 0 người; Giảm 0 người ({0}).", ws.LyDoBienDong));
            }
            else
            {
                sb.AppendLine(string.Format("- Biến động lao động trong kỳ: Tăng {0} người; Giảm {1} người ({2}).", ws.LDTang, ws.LDGiam, ws.LyDoBienDong));
            }
            string ghiChuLD = string.IsNullOrEmpty(ws.GhiChuLaoDong) ? "" : string.Format(" ({0})", ws.GhiChuLaoDong);
            sb.AppendLine(string.Format("- Lao động danh sách tháng {0}/{1}: {2} người{3}, trong đó:", ws.Thang, ws.Nam, ws.LDCuoiKy, ghiChuLD));
            sb.AppendLine(string.Format("  + Lao động quản lý: {0} người;", ws.LDQuanLy));
            sb.AppendLine(string.Format("  + Lao động trực tiếp: {0} người.", ws.LDTrucTiep));

            if (!isMultiPeriod)
            {
                sb.AppendLine(string.Format("- Tổng số công tháng {0}/{1}: {2} công, trong đó:", ws.Thang, ws.Nam, FmtDec(ws.TongCong, 1)));
                sb.AppendLine(string.Format("  + Công sản phẩm: {0} công (Công lao động quản lý: {1} công; Công lao động trực tiếp: {2} công).", FmtDec(ws.CongSP, 1), FmtDec(ws.CongQL, 1), FmtDec(ws.CongLDTT, 1)));
                sb.AppendLine(string.Format("  + Công lễ, phép, tết, học, chế độ: {0} công.", FmtDec(ws.CongLePhepCD, 1)));
                sb.AppendLine(string.Format("  + Công nghỉ ốm, thai sản hưởng BHXH: {0} công.", FmtDec(ws.CongBHXH, 1)));
                sb.AppendLine(string.Format("- Lao động thực tế bình quân đi làm: {0} người (LĐ trực tiếp: {1} người).", FmtDec(ws.LDThucTeBQ, 2), FmtDec(ws.LDTTThucTeBQ, 2)));
                sb.AppendLine(string.Format("- Ngày công thực tế đi làm bình quân: {0} công/người (QL: {1} công/người; LĐTT: {2} công/người).", FmtDec(ws.NgayCongBQ, 2), FmtDec(ws.NgayCongBQQL, 2), FmtDec(ws.NgayCongBQLDTT, 2)));
            }
            else
            {
                sb.AppendLine(string.Format("- Lao động thực tế bình quân: Tháng {0}: {1} người; {2}: {3} người; {4}: {5} người.", ws.Thang, FmtDec(ws.LDThucTeBQ, 2), tenQuy, FmtDec(ws.KyQuy.LDThucTeBQ, 2), tenLK, FmtDec(ws.KyLuyKe.LDThucTeBQ, 2)));
                sb.AppendLine(string.Format("- Lao động trực tiếp thực tế bình quân: Tháng {0}: {1} người; {2}: {3} người; {4}: {5} người.", ws.Thang, FmtDec(ws.LDTTThucTeBQ, 2), tenQuy, FmtDec(ws.KyQuy.LDTTThucTeBQ, 2), tenLK, FmtDec(ws.KyLuyKe.LDTTThucTeBQ, 2)));
                sb.AppendLine(string.Format("- Tổng số công: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công. Trong đó:", ws.Thang, FmtDec(ws.TongCong, 1), tenQuy, FmtDec(ws.KyQuy.TongCong, 1), tenLK, FmtDec(ws.KyLuyKe.TongCong, 1)));
                sb.AppendLine(string.Format("  + Công sản phẩm: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công.", ws.Thang, FmtDec(ws.CongSP, 1), tenQuy, FmtDec(ws.KyQuy.CongSP, 1), tenLK, FmtDec(ws.KyLuyKe.CongSP, 1)));
                sb.AppendLine(string.Format("    * Công quản lý: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công.", ws.Thang, FmtDec(ws.CongQL, 1), tenQuy, FmtDec(ws.KyQuy.CongQL, 1), tenLK, FmtDec(ws.KyLuyKe.CongQL, 1)));
                sb.AppendLine(string.Format("    * Công LĐ trực tiếp: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công.", ws.Thang, FmtDec(ws.CongLDTT, 1), tenQuy, FmtDec(ws.KyQuy.CongLDTT, 1), tenLK, FmtDec(ws.KyLuyKe.CongLDTT, 1)));
                sb.AppendLine(string.Format("  + Công nghỉ phép, lễ, học, chế độ: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công.", ws.Thang, FmtDec(ws.CongLePhepCD, 1), tenQuy, FmtDec(ws.KyQuy.CongCheDo, 1), tenLK, FmtDec(ws.KyLuyKe.CongCheDo, 1)));
                sb.AppendLine(string.Format("  + Công nghỉ hưởng BHXH: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công.", ws.Thang, FmtDec(ws.CongBHXH, 1), tenQuy, FmtDec(ws.KyQuy.CongBHXH, 1), tenLK, FmtDec(ws.KyLuyKe.CongBHXH, 1)));
                sb.AppendLine(string.Format("- Ngày công thực tế bình quân: Tháng {0}: {1} công/người; {2}: {3} công/người; {4}: {5} công/người.", ws.Thang, FmtDec(ws.NgayCongBQ, 2), tenQuy, FmtDec(ws.KyQuy.NgayCongBQ, 2), tenLK, FmtDec(ws.KyLuyKe.NgayCongBQ, 2)));
            }

            sb.AppendLine();
            sb.AppendLine("2. Tiền lương, ăn ca thực hiện:");
            sb.AppendLine("2.1. Tiền lương:");
            if (ws.MaPX != "CDNL" && ws.SanLuongTH > 0)
            {
                sb.AppendLine(string.Format("- Sản lượng thực hiện ({0}): Kế hoạch tháng: {1} {2}; Thực hiện tháng {3}: {4} {5} (Sản lượng tính lương: {6} {7}, trong đó Loại 1: {8} {9}, Loại 2: {10} {11}, làm ngày Lễ: {12} {13}).", ws.TenSanPhamChinh, FmtDec(ws.SanLuongKH, 2), ws.DvtSanPham, ws.Thang, FmtDec(ws.SanLuongTH, 3), ws.DvtSanPham, FmtDec(ws.SanLuongTinhLuong, 3), ws.DvtSanPham, FmtDec(ws.SanLuongLoai1, 3), ws.DvtSanPham, FmtDec(ws.SanLuongLoai2, 3), ws.DvtSanPham, FmtDec(ws.SanLuongLe, 3), ws.DvtSanPham));
                if (isMultiPeriod && ws.KyQuy.SanLuongTH > 0)
                {
                    sb.AppendLine(string.Format("  + Lũy kế sản lượng: {0}: {1} {2}; {3}: {4} {5}.", tenQuy, FmtDec(ws.KyQuy.SanLuongTH, 2), ws.DvtSanPham, tenLK, FmtDec(ws.KyLuyKe.SanLuongTH, 2), ws.DvtSanPham));
                }
            }
            else if (ws.MaPX == "CDNL")
            {
                sb.AppendLine(string.Format("- Quỹ lương sản phẩm các phân xưởng chính làm cơ sở tính lương tháng {0}: {1} đồng.", ws.Thang, FmtInt(ws.SanLuongTH)));
            }

            if (!isMultiPeriod)
            {
                sb.AppendLine(string.Format("- Tổng quỹ tiền lương toàn Phân xưởng tháng {0}/{1}: {2} đồng. Trong đó:", ws.Thang, ws.Nam, FmtInt(ws.TongLuongToanPX)));
                sb.AppendLine(string.Format("  * Tiền lương lao động trực tiếp: {0} đồng, gồm:", FmtInt(ws.TongLuongTT)));
                sb.AppendLine(string.Format("    + Tiền lương sản phẩm trực tiếp: {0} đồng.", FmtInt(ws.LuongSPTT)));
                if (ws.LuongCVKhac > 0)
                    sb.AppendLine(string.Format("    + Tiền lương công việc khác (GNV): {0} đồng ({1} công).", FmtInt(ws.LuongCVKhac), FmtDec(ws.CongCVKhac, 1)));
                if (ws.TruLuong < 0)
                    sb.AppendLine(string.Format("    + Giảm trừ tiền lương: {0} đồng ({1}).", FmtInt(ws.TruLuong), ws.LyDoTruLuong));
                sb.AppendLine(string.Format("    + Tiền lương chế độ (lễ, phép, học): {0} đồng.", FmtInt(ws.LuongCheDoTT)));
                sb.AppendLine(string.Format("  * Tiền lương lao động quản lý: {0} đồng (Lương SP: {1} đồng; Lương chế độ: {2} đồng).", FmtInt(ws.TongLuongQL), FmtInt(ws.LuongSPQL), FmtInt(ws.LuongCheDoQL)));
            }
            else
            {
                sb.AppendLine(string.Format("- Tổng tiền lương toàn Phân xưởng: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng. Trong đó:", ws.Thang, FmtInt(ws.TongLuongToanPX), tenQuy, FmtInt(ws.KyQuy.TongLuongToanPX), tenLK, FmtInt(ws.KyLuyKe.TongLuongToanPX)));
                sb.AppendLine(string.Format("  * Tiền lương lao động trực tiếp: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng:", ws.Thang, FmtInt(ws.TongLuongTT), tenQuy, FmtInt(ws.KyQuy.TongLuongTT), tenLK, FmtInt(ws.KyLuyKe.TongLuongTT)));
                sb.AppendLine(string.Format("    + Lương sản phẩm: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng.", ws.Thang, FmtInt(ws.LuongSPTT), tenQuy, FmtInt(ws.KyQuy.LuongSPTT), tenLK, FmtInt(ws.KyLuyKe.LuongSPTT)));
                if (ws.LuongCVKhac > 0 || ws.KyQuy.LuongCVKhac > 0 || ws.KyLuyKe.LuongCVKhac > 0)
                    sb.AppendLine(string.Format("    + Lương công việc khác (GNV): Tháng {0}: {1} đồng ({2} công); {3}: {4} đồng ({5} công); {6}: {7} đồng ({8} công).", ws.Thang, FmtInt(ws.LuongCVKhac), FmtDec(ws.CongCVKhac, 1), tenQuy, FmtInt(ws.KyQuy.LuongCVKhac), FmtDec(ws.KyQuy.CongCVKhac, 1), tenLK, FmtInt(ws.KyLuyKe.LuongCVKhac), FmtDec(ws.KyLuyKe.CongCVKhac, 1)));
                if (ws.KyLuyKe.LuongTDT > 0)
                    sb.AppendLine(string.Format("    + Lương công Trung đại tu (TĐT): {0}: {1} đồng ({2} công).", tenLK, FmtInt(ws.KyLuyKe.LuongTDT), FmtDec(ws.KyLuyKe.CongTDT, 1)));
                if (ws.TruLuong < 0)
                    sb.AppendLine(string.Format("    + Trừ lương trong kỳ: {0} đồng ({1}).", FmtInt(ws.TruLuong), ws.LyDoTruLuong));
                sb.AppendLine(string.Format("    + Lương chế độ (lễ, phép, học): Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng.", ws.Thang, FmtInt(ws.LuongCheDoTT), tenQuy, FmtInt(ws.KyQuy.LuongCheDoTT), tenLK, FmtInt(ws.KyLuyKe.LuongCheDoTT)));
                sb.AppendLine(string.Format("  * Tiền lương lao động quản lý: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng.", ws.Thang, FmtInt(ws.TongLuongQL), tenQuy, FmtInt(ws.KyQuy.TongLuongQL), tenLK, FmtInt(ws.KyLuyKe.TongLuongQL)));
                sb.AppendLine(string.Format("    + Lương sản phẩm QL: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng.", ws.Thang, FmtInt(ws.LuongSPQL), tenQuy, FmtInt(ws.KyQuy.LuongSPQL), tenLK, FmtInt(ws.KyLuyKe.LuongSPQL)));
                sb.AppendLine(string.Format("    + Lương chế độ QL: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng.", ws.Thang, FmtInt(ws.LuongCheDoQL), tenQuy, FmtInt(ws.KyQuy.LuongCheDoQL), tenLK, FmtInt(ws.KyLuyKe.LuongCheDoQL)));
            }

            if (ws.ListGNV != null && ws.ListGNV.Count > 0)
            {
                sb.AppendLine(string.Format("  * Chi tiết công việc phát sinh / Giấy giao nhiệm vụ (GNV) trong tháng {0}:", ws.Thang));
                foreach (var gnv in ws.ListGNV)
                {
                    sb.AppendLine(string.Format("    [{0}] {1} | ĐVT: {2} | KL: {3} | Đơn giá: {4} đ | Thành tiền: {5} đ", gnv.STT, gnv.NoiDung, gnv.DVT, FmtDec(gnv.KhoiLuong, 2), FmtInt(gnv.DonGia), FmtInt(gnv.ThanhTien)));
                }
            }

            if (!isMultiPeriod)
            {
                sb.AppendLine(string.Format("- Tiền lương bình quân toàn PX: {0} nghìn đồng/người/tháng (LĐTT: {1} nghìn đồng/người/tháng).", FmtDec(ws.TLBinhQuanToanPX, 2), FmtDec(ws.TLSPBinhQuanLDTT, 2)));
                sb.AppendLine(string.Format("- Thu nhập bình quân toàn PX: {0} nghìn đồng/người/tháng (LĐTT: {1} nghìn đồng/người/tháng).", FmtDec(ws.ThuNhapBQToanPX, 2), FmtDec(ws.ThuNhapBQLDTT, 2)));
            }
            else
            {
                sb.AppendLine(string.Format("- Tiền lương bình quân toàn PX (1.000đ/người/tháng): Tháng {0}: {1}; {2}: {3}; {4}: {5}.", ws.Thang, FmtDec(ws.TLBinhQuanToanPX, 2), tenQuy, FmtDec(ws.KyQuy.TLBinhQuanToanPX, 2), tenLK, FmtDec(ws.KyLuyKe.TLBinhQuanToanPX, 2)));
                sb.AppendLine(string.Format("- Thu nhập bình quân toàn PX (1.000đ/người/tháng): Tháng {0}: {1}; {2}: {3}; {4}: {5}.", ws.Thang, FmtDec(ws.ThuNhapBQToanPX, 2), tenQuy, FmtDec(ws.KyQuy.ThuNhapBQToanPX, 2), tenLK, FmtDec(ws.KyLuyKe.ThuNhapBQToanPX, 2)));
                sb.AppendLine(string.Format("- Thu nhập bình quân LĐTT (1.000đ/người/tháng): Tháng {0}: {1}; {2}: {3}; {4}: {5}.", ws.Thang, FmtDec(ws.ThuNhapBQLDTT, 2), tenQuy, FmtDec(ws.KyQuy.ThuNhapBQLDTT, 2), tenLK, FmtDec(ws.KyLuyKe.ThuNhapBQLDTT, 2)));
            }

            sb.AppendLine();
            sb.AppendLine("2.2. Ăn ca (Đơn giá 30.000 đồng/công):");
            if (!isMultiPeriod)
            {
                sb.AppendLine(string.Format("- Tổng số công ăn ca tháng {0}/{1}: {2} công; Thành tiền: {3} đồng.", ws.Thang, ws.Nam, FmtDec(ws.CongAnCa, 1), FmtInt(ws.TienAnCa)));
            }
            else
            {
                sb.AppendLine(string.Format("- Tổng số công ăn ca: Tháng {0}: {1} công; {2}: {3} công; {4}: {5} công.", ws.Thang, FmtDec(ws.CongAnCa, 1), tenQuy, FmtDec(ws.KyQuy.CongAnCa, 1), tenLK, FmtDec(ws.KyLuyKe.CongAnCa, 1)));
                sb.AppendLine(string.Format("- Tổng số tiền ăn ca: Tháng {0}: {1} đồng; {2}: {3} đồng; {4}: {5} đồng.", ws.Thang, FmtInt(ws.TienAnCa), tenQuy, FmtInt(ws.KyQuy.TienAnCa), tenLK, FmtInt(ws.KyLuyKe.TienAnCa)));
            }

            sb.AppendLine();
            sb.AppendLine("2.3. Bồi dưỡng độc hại bằng hiện vật:");
            sb.AppendLine(string.Format("- Mức 1 (13.000đ/công): Tháng {0}: {1} công ({2} đ)", ws.Thang, FmtDec(ws.BDDH_Muc1_Cong, 1), FmtInt(ws.BDDH_Muc1_Tien)) +
                (isMultiPeriod ? string.Format(" | {0}: {1} công ({2} đ) | {3}: {4} công ({5} đ)", tenQuy, FmtDec(ws.KyQuy.BDDH_Muc1_Cong, 1), FmtInt(ws.KyQuy.BDDH_Muc1_Tien), tenLK, FmtDec(ws.KyLuyKe.BDDH_Muc1_Cong, 1), FmtInt(ws.KyLuyKe.BDDH_Muc1_Tien)) : ""));
            sb.AppendLine(string.Format("- Mức 2 (20.000đ/công): Tháng {0}: {1} công ({2} đ)", ws.Thang, FmtDec(ws.BDDH_Muc2_Cong, 1), FmtInt(ws.BDDH_Muc2_Tien)) +
                (isMultiPeriod ? string.Format(" | {0}: {1} công ({2} đ) | {3}: {4} công ({5} đ)", tenQuy, FmtDec(ws.KyQuy.BDDH_Muc2_Cong, 1), FmtInt(ws.KyQuy.BDDH_Muc2_Tien), tenLK, FmtDec(ws.KyLuyKe.BDDH_Muc2_Cong, 1), FmtInt(ws.KyLuyKe.BDDH_Muc2_Tien)) : ""));
            if (ws.CoMuc3BDDH || ws.BDDH_Muc3_Cong > 0)
            {
                sb.AppendLine(string.Format("- Mức 3 (26.000đ/công): Tháng {0}: {1} công ({2} đ)", ws.Thang, FmtDec(ws.BDDH_Muc3_Cong, 1), FmtInt(ws.BDDH_Muc3_Tien)) +
                    (isMultiPeriod ? string.Format(" | {0}: {1} công ({2} đ) | {3}: {4} công ({5} đ)", tenQuy, FmtDec(ws.KyQuy.BDDH_Muc3_Cong, 1), FmtInt(ws.KyQuy.BDDH_Muc3_Tien), tenLK, FmtDec(ws.KyLuyKe.BDDH_Muc3_Cong, 1), FmtInt(ws.KyLuyKe.BDDH_Muc3_Tien)) : ""));
            }
            double tongCongBDDH_T = ws.BDDH_Muc1_Cong + ws.BDDH_Muc2_Cong + ws.BDDH_Muc3_Cong;
            double tongCongBDDH_Q = ws.KyQuy.BDDH_Muc1_Cong + ws.KyQuy.BDDH_Muc2_Cong + ws.KyQuy.BDDH_Muc3_Cong;
            double tongCongBDDH_LK = ws.KyLuyKe.BDDH_Muc1_Cong + ws.KyLuyKe.BDDH_Muc2_Cong + ws.KyLuyKe.BDDH_Muc3_Cong;
            sb.AppendLine(string.Format("- TỔNG CỘNG BDĐH: Tháng {0}: {1} công ({2} đ)", ws.Thang, FmtDec(tongCongBDDH_T, 1), FmtInt(ws.TongTienBDDH)) +
                (isMultiPeriod ? string.Format(" | {0}: {1} công ({2} đ) | {3}: {4} công ({5} đ)", tenQuy, FmtDec(tongCongBDDH_Q, 1), FmtInt(ws.KyQuy.TongTienBDDH), tenLK, FmtDec(tongCongBDDH_LK, 1), FmtInt(ws.KyLuyKe.TongTienBDDH)) : ""));

            sb.AppendLine();
            sb.AppendLine("3. Công tác bình công chấm điểm, ghi chép sổ sách:");
            sb.AppendLine("- Thực hiện chấm công, ghi chép sổ sách rõ ràng, đúng quy định hiện hành.");
            sb.AppendLine();
            sb.AppendLine("III. Kiến nghị, đề xuất:");
            foreach (var kn in ws.KienNghi)
            {
                if (!string.IsNullOrWhiteSpace(kn))
                    sb.AppendLine(string.Format("- {0}", kn));
            }
            sb.AppendLine();
            sb.AppendLine(string.Format("Biên bản kết thúc vào hồi {0} cùng ngày, đã được đọc lại cho mọi người cùng nghe và thống nhất ký tên./.", ws.GioKetThuc));
            sb.AppendLine();
            sb.AppendLine("                 QUẢN ĐỐC                                     TP. TCHC");

            if (ws.ListNhanVien != null && ws.ListNhanVien.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("================================================================================");
                sb.AppendLine(string.Format("   TÓM TẮT BẢNG NGHIỆM THU CHI TIẾT TỪNG NGƯỜI ({0} LAO ĐỘNG - SHEET '{1}')", ws.ListNhanVien.Count, ws.SheetBangCong));
                sb.AppendLine("================================================================================");
                sb.AppendLine(string.Format("{0,-4} | {1,-5} | {2,-24} | {3,-6} | {4,7} | {5,7} | {6,8} | {7,12} | {8,12}",
                    "STT", "SDB", "Họ và tên", "C.Danh", "Công SP", "Lễ/Phép", "Điểm/HS", "Lương SP", "Tổng lương"));
                sb.AppendLine("----------------------------------------------------------------------------------------------------");
                int previewMax = Math.Min(ws.ListNhanVien.Count, 25);
                for (int i = 0; i < previewMax; i++)
                {
                    var e = ws.ListNhanVien[i];
                    sb.AppendLine(string.Format("{0,-4} | {1,-5} | {2,-24} | {3,-6} | {4,7} | {5,7} | {6,8} | {7,12} | {8,12}",
                        e.STT, e.SDB, TruncateStr(e.HoTen, 24), TruncateStr(e.ChucDanh, 6),
                        FmtDec(e.CongSP, 1), FmtDec(e.CongLePhep, 1), FmtDec(e.HeSoHuong, 2),
                        FmtInt(e.LuongSanPham), FmtInt(e.TongTienLuong)));
                }
                if (ws.ListNhanVien.Count > previewMax)
                {
                    sb.AppendLine(string.Format("... (Còn {0} lao động khác trong tổng số {1} người - Xem đầy đủ tại Tab 1 hoặc tích chọn Đính kèm Phụ lục khi xuất Word)",
                        ws.ListNhanVien.Count - previewMax, ws.ListNhanVien.Count));
                }
            }

            return sb.ToString();
        }

        private static string TruncateStr(string s, int maxLen)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Length <= maxLen ? s : s.Substring(0, maxLen);
        }

        public static void ExportWordDocx(WorkshopData ws, string outputPath)
        {
            ws.RecalculateDerived();
            string dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml", GetContentTypesXml());
                WriteEntry(zip, "_rels/.rels", GetRootRelsXml());
                WriteEntry(zip, "word/_rels/document.xml.rels", GetDocRelsXml());
                WriteEntry(zip, "word/styles.xml", GetStylesXml());
                WriteEntry(zip, "word/document.xml", BuildDocumentXml(ws));
            }
        }

        /// <summary>
        /// Điền / Cập nhật Biên bản đánh giá công tác LĐTL trực tiếp vào file Word mẫu (.docx) hiện có của người dùng
        /// (VD: Nt_Thieu_ket_T9.docx, Nt_thang_9_Gang.docx, NT_thep_Thang_9.docx, BB_Nt_CDNL_T8.docx).
        /// Nếu file mẫu đã có biên bản của tháng hiện tại ở cuối file thì thay thế biên bản cuối đó bằng số liệu mới nhất;
        /// nếu chưa có (VD: file đang ở Tháng 8, nay xuất Tháng 9) thì tự động nối thêm trang Biên bản mới vào cuối file!
        /// </summary>
        public static void UpdateExistingMasterDocx(WorkshopData ws, string sourceDocxPath, string targetDocxPath)
        {
            ws.RecalculateDerived();
            string targetDir = Path.GetDirectoryName(targetDocxPath);
            if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            if (!File.Exists(sourceDocxPath))
            {
                ExportWordDocx(ws, targetDocxPath);
                return;
            }

            if (!string.Equals(Path.GetFullPath(sourceDocxPath), Path.GetFullPath(targetDocxPath), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourceDocxPath, targetDocxPath, true);
            }

            using (var fs = new FileStream(targetDocxPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                var docEntry = zip.GetEntry("word/document.xml");
                if (docEntry == null) return;

                var xmlDoc = new XmlDocument();
                using (var s = docEntry.Open()) xmlDoc.Load(s);

                var ns = new XmlNamespaceManager(xmlDoc.NameTable);
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                var bodyNode = xmlDoc.SelectSingleNode("//w:body", ns) as XmlElement;
                if (bodyNode == null) return;

                // Tìm vị trí bắt đầu của Biên bản cuối cùng trong file (bắt đầu từ bảng Quốc hiệu CÔNG TY CP GANG THÉP CAO BẰNG cuối cùng)
                XmlNode lastReportStartNode = null;
                foreach (XmlNode child in bodyNode.ChildNodes)
                {
                    if (child.LocalName == "tbl" && child.InnerText.Contains("CÔNG TY CP GANG THÉP CAO BẰNG") && child.InnerText.Contains("Độc lập"))
                    {
                        lastReportStartNode = child;
                    }
                }

                // Kiểm tra xem biên bản cuối cùng có phải cùng Tháng hoặc người dùng muốn cập nhật trực tiếp không
                if (lastReportStartNode != null)
                {
                    // Thu thập toàn bộ các node từ lastReportStartNode đến trước w:sectPr cuối cùng và thay bằng nội dung Biên bản mới
                    var nodesToRemove = new List<XmlNode>();
                    XmlNode cur = lastReportStartNode;
                    while (cur != null && cur.LocalName != "sectPr")
                    {
                        nodesToRemove.Add(cur);
                        cur = cur.NextSibling;
                    }
                    foreach (var rm in nodesToRemove)
                    {
                        bodyNode.RemoveChild(rm);
                    }
                }

                // Chèn các node của Biên bản mới vào trước w:sectPr
                var sectPrNode = bodyNode.SelectSingleNode("w:sectPr", ns);
                var fragmentDoc = new XmlDocument();
                fragmentDoc.LoadXml(string.Format(@"<root xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">{0}</root>", BuildReportBodyElementsXml(ws)));

                foreach (XmlNode newChild in fragmentDoc.DocumentElement.ChildNodes)
                {
                    var imported = xmlDoc.ImportNode(newChild, true);
                    if (sectPrNode != null)
                        bodyNode.InsertBefore(imported, sectPrNode);
                    else
                        bodyNode.AppendChild(imported);
                }

                docEntry.Delete();
                var newEntry = zip.CreateEntry("word/document.xml", CompressionLevel.Optimal);
                using (var s = newEntry.Open())
                {
                    xmlDoc.Save(s);
                }
            }
        }

        private static void WriteEntry(ZipArchive zip, string path, string content)
        {
            var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            using (var sw = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                sw.Write(content);
            }
        }

        private static string Esc(string s)
        {
            return SecurityElement.Escape(s ?? "");
        }

        private static string BuildReportBodyElementsXml(WorkshopData ws)
        {
            bool isMultiPeriod = ws.LoaiKy != LoaiKyBaoCao.ThangThuong;
            string tenQuy = ws.KyQuy != null && !string.IsNullOrEmpty(ws.KyQuy.TenKy) ? ws.KyQuy.TenKy : "Quý III";
            string tenLK = ws.KyLuyKe != null && !string.IsNullOrEmpty(ws.KyLuyKe.TenKy) ? ws.KyLuyKe.TenKy : "9 tháng";

            var body = new StringBuilder();

            // 1. Quốc hiệu & Tiêu ngữ (Bảng 2 cột chuẩn CISCO)
            body.Append(BuildHeaderTable());
            body.Append(P("", false, false, "center", 40, 40));

            // 2. Tiêu đề Biên bản
            body.Append(P("BIÊN BẢN", true, false, "center", 80, 20, 28));
            body.Append(P(string.Format("ĐÁNH GIÁ CÔNG TÁC LAO ĐỘNG TIỀN LƯƠNG {0}", GetTitlePeriodText(ws).ToUpperInvariant()), true, false, "center", 0, 20, 26));
            body.Append(P(string.Format("ĐƠN VỊ: {0}", ws.TenPX.ToUpperInvariant()), true, false, "center", 0, 140, 26));

            // 3. Thời gian, Thành phần, Địa điểm
            body.Append(P(string.Format("Hôm nay, ngày {0:00} tháng {1:00} năm {2}, tại {3} Công ty cổ phần Gang thép Cao Bằng, chúng tôi gồm:", ws.NgayHop, ws.ThangHop, ws.NamHop, ws.DiaDiem), false, false, "both", 40, 40, 26, 420));
            body.Append(P("I. Thành phần:", true, false, "both", 60, 40));
            body.Append(P("1. Phòng Tổ chức hành chính:", true, false, "both", 20, 20, 26, 420));
            foreach (var p in ws.ThanhPhanTCHC)
            {
                body.Append(P(string.Format("- {0}\t\t- Chức vụ: {1}", p.HoTen, p.ChucVu), false, false, "both", 20, 20, 26, 420));
            }
            body.Append(P(string.Format("2. {0}:", ws.TenPX), true, false, "both", 20, 20, 26, 420));
            foreach (var p in ws.ThanhPhanPX)
            {
                body.Append(P(string.Format("- {0}\t\t- Chức vụ: {1}", p.HoTen, p.ChucVu), false, false, "both", 20, 20, 26, 420));
            }

            // 4. Đánh giá kết quả thực hiện công tác LĐTL
            body.Append(P("II. Đánh giá kết quả thực hiện công tác LĐTL:", true, false, "both", 60, 40));
            body.Append(P("Cùng nhau đánh giá công tác quản lý lao động, tiền lương, ăn ca, bồi dưỡng độc hại của Phân xưởng như sau:", false, false, "both", 20, 40, 26, 420));

            // Mục 1. Công tác quản lý, sử dụng lao động
            body.Append(P("1. Công tác quản lý, sử dụng lao động:", true, false, "both", 60, 40, 26, 420));
            string ghiChuLD = string.IsNullOrEmpty(ws.GhiChuLaoDong) ? "" : string.Format(" ({0})", ws.GhiChuLaoDong);
            body.Append(P(string.Format("- Lao động danh sách tháng {0}/{1}: {2} người{3}, trong đó:", ws.Thang, ws.Nam, ws.LDCuoiKy, ghiChuLD), false, false, "both", 20, 20, 26, 420));
            body.Append(P(string.Format("+ Lao động quản lý: {0} người;", ws.LDQuanLy), false, false, "both", 20, 20, 26, 560));
            body.Append(P(string.Format("+ Lao động trực tiếp: {0} người.", ws.LDTrucTiep), false, false, "both", 20, 20, 26, 560));
            if (ws.LDTang > 0 || ws.LDGiam > 0 || !string.IsNullOrEmpty(ws.LyDoBienDong))
            {
                body.Append(P(string.Format("+ Biến động lao động trong kỳ: Đầu kỳ {0} người; Tăng {1} người; Giảm {2} người ({3}).", ws.LDDauKy, ws.LDTang, ws.LDGiam, ws.LyDoBienDong), false, false, "both", 20, 20, 26, 560));
            }

            if (!isMultiPeriod)
            {
                body.Append(P(string.Format("- Lao động thực tế bình quân đi làm: {0} người (Lao động trực tiếp thực tế bình quân: {1} người).", FmtDec(ws.LDThucTeBQ, 2), FmtDec(ws.LDTTThucTeBQ, 2)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("- Tổng số công tháng {0}/{1}: {2} công, trong đó:", ws.Thang, ws.Nam, FmtDec(ws.TongCong, 1)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("+ Công sản phẩm: {0} công (Công lao động quản lý: {1} công; Công lao động trực tiếp: {2} công).", FmtDec(ws.CongSP, 1), FmtDec(ws.CongQL, 1), FmtDec(ws.CongLDTT, 1)), false, false, "both", 20, 20, 26, 560));
                body.Append(P(string.Format("+ Công nghỉ phép, lễ, tết, học, chế độ: {0} công.", FmtDec(ws.CongLePhepCD, 1)), false, false, "both", 20, 20, 26, 560));
                body.Append(P(string.Format("+ Công nghỉ hưởng BHXH: {0} công.", FmtDec(ws.CongBHXH, 1)), false, false, "both", 20, 20, 26, 560));
                body.Append(P(string.Format("- Ngày công thực tế bình quân: {0} công/người (LĐ quản lý: {1} công/người; LĐ trực tiếp: {2} công/người).", FmtDec(ws.NgayCongBQ, 2), FmtDec(ws.NgayCongBQQL, 2), FmtDec(ws.NgayCongBQLDTT, 2)), false, false, "both", 20, 40, 26, 420));
            }
            else
            {
                body.Append(P(string.Format("+ Lao động danh sách bình quân: {0}: {1} người; {2} năm {3}: {4} người.", tenQuy, FmtDec(ws.KyQuy.LDDanhSachBQ, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.LDDanhSachBQ, 2)), false, false, "both", 20, 20, 26, 560));
                body.Append(P(string.Format("- Lao động thực tế bình quân: Tháng {0}: {1} người; {2}: {3} người; {4} năm {5}: {6} người.", ws.Thang, FmtDec(ws.LDThucTeBQ, 2), tenQuy, FmtDec(ws.KyQuy.LDThucTeBQ, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.LDThucTeBQ, 2)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("- Lao động trực tiếp thực tế bình quân: Tháng {0}: {1} người; {2}: {3} người; {4} năm {5}: {6} người.", ws.Thang, FmtDec(ws.LDTTThucTeBQ, 2), tenQuy, FmtDec(ws.KyQuy.LDTTThucTeBQ, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.LDTTThucTeBQ, 2)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("- Tổng số công: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công. Trong đó:", ws.Thang, FmtDec(ws.TongCong, 1), tenQuy, FmtDec(ws.KyQuy.TongCong, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.TongCong, 1)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("+ Công sản phẩm: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công, trong đó:", ws.Thang, FmtDec(ws.CongSP, 1), tenQuy, FmtDec(ws.KyQuy.CongSP, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.CongSP, 1)), false, false, "both", 20, 20, 26, 560));
                body.Append(P(string.Format("• Công quản lý: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công;", ws.Thang, FmtDec(ws.CongQL, 1), tenQuy, FmtDec(ws.KyQuy.CongQL, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.CongQL, 1)), false, false, "both", 20, 20, 26, 720));
                body.Append(P(string.Format("• Công lao động trực tiếp: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công.", ws.Thang, FmtDec(ws.CongLDTT, 1), tenQuy, FmtDec(ws.KyQuy.CongLDTT, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.CongLDTT, 1)), false, false, "both", 20, 20, 26, 720));
                body.Append(P(string.Format("+ Công nghỉ phép, lễ, học, chế độ: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công.", ws.Thang, FmtDec(ws.CongLePhepCD, 1), tenQuy, FmtDec(ws.KyQuy.CongCheDo, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.CongCheDo, 1)), false, false, "both", 20, 20, 26, 560));
                body.Append(P(string.Format("+ Công nghỉ hưởng BHXH: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công.", ws.Thang, FmtDec(ws.CongBHXH, 1), tenQuy, FmtDec(ws.KyQuy.CongBHXH, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.CongBHXH, 1)), false, false, "both", 20, 40, 26, 560));
            }

            // Mục 2. Tiền lương, ăn ca thực hiện
            body.Append(P("2. Tiền lương, ăn ca thực hiện:", true, false, "both", 60, 40, 26, 420));
            body.Append(P("2.1. Tiền lương", true, true, "both", 20, 20, 26, 420));

            if (ws.MaPX != "CDNL" && ws.SanLuongTH > 0)
            {
                if (!isMultiPeriod)
                {
                    body.Append(P(string.Format("- Sản lượng thực hiện ({0}): Kế hoạch giao: {1} {2}; Thực hiện tháng {3}: {4} {5}; Sản lượng tính lương: {6} {7} (Loại 1: {8} {9}; Loại 2: {10} {11}; Sản phẩm làm ngày Lễ: {12} {13}).", ws.TenSanPhamChinh, FmtDec(ws.SanLuongKH, 2), ws.DvtSanPham, ws.Thang, FmtDec(ws.SanLuongTH, 3), ws.DvtSanPham, FmtDec(ws.SanLuongTinhLuong, 3), ws.DvtSanPham, FmtDec(ws.SanLuongLoai1, 3), ws.DvtSanPham, FmtDec(ws.SanLuongLoai2, 3), ws.DvtSanPham, FmtDec(ws.SanLuongLe, 3), ws.DvtSanPham), false, false, "both", 20, 20, 26, 420));
                }
                else
                {
                    body.Append(P(string.Format("- Sản lượng {0} thực hiện: Tháng {1}: {2} {3}; {4}: {5} {6}; {7} năm {8}: {9} {10}.", ws.TenSanPhamChinh, ws.Thang, FmtDec(ws.SanLuongTH, 2), ws.DvtSanPham, tenQuy, FmtDec(ws.KyQuy.SanLuongTH, 2), ws.DvtSanPham, tenLK, ws.Nam, FmtDec(ws.KyLuyKe.SanLuongTH, 2), ws.DvtSanPham), false, false, "both", 20, 20, 26, 420));
                    body.Append(P(string.Format("- Sản lượng {0} tính lương: Tháng {1}: {2} {3} (Loại I: {4} {5}; Loại II: {6} {7}; Ngày lễ: {8} {9}); {10}: {11} {12}; {13} năm {14}: {15} {16}.", ws.TenSanPhamChinh, ws.Thang, FmtDec(ws.SanLuongTinhLuong, 2), ws.DvtSanPham, FmtDec(ws.SanLuongLoai1, 2), ws.DvtSanPham, FmtDec(ws.SanLuongLoai2, 2), ws.DvtSanPham, FmtDec(ws.SanLuongLe, 2), ws.DvtSanPham, tenQuy, FmtDec(ws.KyQuy.SanLuongTinhLuong, 2), ws.DvtSanPham, tenLK, ws.Nam, FmtDec(ws.KyLuyKe.SanLuongTinhLuong, 2), ws.DvtSanPham), false, false, "both", 20, 20, 26, 420));
                }
                if (!string.IsNullOrEmpty(ws.SanLuongPhuName) && ws.SanLuongPhuKL > 0)
                {
                    body.Append(P(string.Format("+ Khối lượng {0}: {1} {2}, thành tiền: {3} đồng.", ws.SanLuongPhuName, FmtDec(ws.SanLuongPhuKL, 2), ws.DvtSanPham, FmtInt(ws.SanLuongPhuTT)), false, false, "both", 20, 20, 26, 560));
                }
            }
            else if (ws.MaPX == "CDNL")
            {
                body.Append(P(string.Format("- Quỹ lương sản phẩm các phân xưởng chính làm cơ sở tính lương tháng {0}: {1} đồng.", ws.Thang, FmtInt(ws.SanLuongTH)), false, false, "both", 20, 20, 26, 420));
            }

            if (!isMultiPeriod)
            {
                body.Append(PLabelValue(string.Format("- Tổng tiền lương toàn Phân xưởng tháng {0}/{1}: ", ws.Thang, ws.Nam), string.Format("{0} đồng. Trong đó:", FmtInt(ws.TongLuongToanPX)), "both", 20, 20, 420));
                body.Append(PLabelValue("+ Tiền lương lao động trực tiếp: ", string.Format("{0} đồng, gồm:", FmtInt(ws.TongLuongTT)), "both", 20, 20, 560));
                body.Append(P(string.Format("• Tiền lương sản phẩm trực tiếp: {0} đồng.", FmtInt(ws.LuongSPTT)), false, false, "both", 20, 20, 26, 720));
                if (ws.LuongCVKhac > 0)
                    body.Append(P(string.Format("• Tiền lương công việc khác (GNV): {0} đồng ({1} công).", FmtInt(ws.LuongCVKhac), FmtDec(ws.CongCVKhac, 1)), false, false, "both", 20, 20, 26, 720));
                if (ws.TruLuong < 0)
                    body.Append(P(string.Format("• Giảm trừ tiền lương: {0} đồng ({1}).", FmtInt(ws.TruLuong), ws.LyDoTruLuong), false, false, "both", 20, 20, 26, 720));
                body.Append(P(string.Format("• Tiền lương chế độ (lễ, phép, học): {0} đồng.", FmtInt(ws.LuongCheDoTT)), false, false, "both", 20, 20, 26, 720));
                body.Append(PLabelValue("+ Tiền lương lao động quản lý: ", string.Format("{0} đồng (Lương sản phẩm: {1} đồng; Lương chế độ: {2} đồng).", FmtInt(ws.TongLuongQL), FmtInt(ws.LuongSPQL), FmtInt(ws.LuongCheDoQL)), "both", 20, 20, 560));
            }
            else
            {
                body.Append(PLabelValue("- Tổng tiền lương của toàn Phân xưởng: ", string.Format("Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng. Trong đó:", ws.Thang, FmtInt(ws.TongLuongToanPX), tenQuy, FmtInt(ws.KyQuy.TongLuongToanPX), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.TongLuongToanPX)), "both", 20, 20, 420));
                body.Append(PLabelValue("+ Tiền lương lao động trực tiếp: ", string.Format("Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng:", ws.Thang, FmtInt(ws.TongLuongTT), tenQuy, FmtInt(ws.KyQuy.TongLuongTT), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.TongLuongTT)), "both", 20, 20, 560));
                body.Append(P(string.Format("• Tiền lương theo sản phẩm: Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng.", ws.Thang, FmtInt(ws.LuongSPTT), tenQuy, FmtInt(ws.KyQuy.LuongSPTT), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.LuongSPTT)), false, false, "both", 20, 20, 26, 720));
                if (ws.LuongCVKhac > 0 || ws.KyQuy.LuongCVKhac > 0 || ws.KyLuyKe.LuongCVKhac > 0)
                {
                    body.Append(P(string.Format("• Tiền lương công việc khác (GNV): Tháng {0}: {1} đồng ({2} công); {3}: {4} đồng ({5} công); {6} năm {7}: {8} đồng ({9} công).", ws.Thang, FmtInt(ws.LuongCVKhac), FmtDec(ws.CongCVKhac, 1), tenQuy, FmtInt(ws.KyQuy.LuongCVKhac), FmtDec(ws.KyQuy.CongCVKhac, 1), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.LuongCVKhac), FmtDec(ws.KyLuyKe.CongCVKhac, 1)), false, false, "both", 20, 20, 26, 720));
                }
                if (ws.KyLuyKe.LuongTDT > 0)
                {
                    body.Append(P(string.Format("• Tiền lương nhân công Trung đại tu (TĐT): {0} năm {1}: {2} đồng ({3} công).", tenLK, ws.Nam, FmtInt(ws.KyLuyKe.LuongTDT), FmtDec(ws.KyLuyKe.CongTDT, 1)), false, false, "both", 20, 20, 26, 720));
                }
                if (ws.TruLuong < 0)
                {
                    body.Append(P(string.Format("• Trừ tiền lương công tác VSMT, ATBHLĐ trong kỳ: {0} đồng ({1}).", FmtInt(ws.TruLuong), ws.LyDoTruLuong), false, false, "both", 20, 20, 26, 720));
                }
                body.Append(P(string.Format("• Tiền lương chế độ (lễ, phép, học): Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng.", ws.Thang, FmtInt(ws.LuongCheDoTT), tenQuy, FmtInt(ws.KyQuy.LuongCheDoTT), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.LuongCheDoTT)), false, false, "both", 20, 20, 26, 720));
                body.Append(PLabelValue("+ Tiền lương lao động quản lý: ", string.Format("Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng:", ws.Thang, FmtInt(ws.TongLuongQL), tenQuy, FmtInt(ws.KyQuy.TongLuongQL), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.TongLuongQL)), "both", 20, 20, 560));
                body.Append(P(string.Format("• Tiền lương theo sản phẩm QL: Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng;", ws.Thang, FmtInt(ws.LuongSPQL), tenQuy, FmtInt(ws.KyQuy.LuongSPQL), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.LuongSPQL)), false, false, "both", 20, 20, 26, 720));
                body.Append(P(string.Format("• Tiền lương chế độ QL (phép, học, lễ): Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng.", ws.Thang, FmtInt(ws.LuongCheDoQL), tenQuy, FmtInt(ws.KyQuy.LuongCheDoQL), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.LuongCheDoQL)), false, false, "both", 20, 20, 26, 720));
            }

            // Bảng chi tiết GNV nếu có
            if (ws.ListGNV != null && ws.ListGNV.Count > 0)
            {
                body.Append(P(string.Format("Tổng hợp khối lượng công việc thực hiện theo Giấy giao nhiệm vụ (GNV) tháng {0}/{1}:", ws.Thang, ws.Nam), false, true, "both", 30, 30, 26, 420));
                body.Append(BuildGNVTable(ws.ListGNV));
            }

            // Chỉ tiêu bình quân
            if (!isMultiPeriod)
            {
                body.Append(P(string.Format("- Tiền lương bình quân toàn Phân xưởng: {0} nghìn đồng/người/tháng (LĐ trực tiếp: {1} nghìn đồng/người/tháng).", FmtDec(ws.TLBinhQuanToanPX, 2), FmtDec(ws.TLSPBinhQuanLDTT, 2)), false, false, "both", 40, 20, 26, 420));
                body.Append(P(string.Format("- Thu nhập bình quân toàn Phân xưởng: {0} nghìn đồng/người/tháng (LĐ trực tiếp: {1} nghìn đồng/người/tháng).", FmtDec(ws.ThuNhapBQToanPX, 2), FmtDec(ws.ThuNhapBQLDTT, 2)), false, false, "both", 20, 40, 26, 420));
            }
            else
            {
                body.Append(P(string.Format("- Tiền lương bình quân của toàn Phân xưởng: Tháng {0}: {1} nghìn đồng/người/tháng; {2}: {3} nghìn đồng/người/tháng; {4} năm {5}: {6} nghìn đồng/người/tháng.", ws.Thang, FmtDec(ws.TLBinhQuanToanPX, 2), tenQuy, FmtDec(ws.KyQuy.TLBinhQuanToanPX, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.TLBinhQuanToanPX, 2)), false, false, "both", 40, 20, 26, 420));
                body.Append(P(string.Format("- Thu nhập bình quân của toàn Phân xưởng: Tháng {0}: {1} nghìn đồng/người/tháng; {2}: {3} nghìn đồng/người/tháng; {4} năm {5}: {6} nghìn đồng/người/tháng.", ws.Thang, FmtDec(ws.ThuNhapBQToanPX, 2), tenQuy, FmtDec(ws.KyQuy.ThuNhapBQToanPX, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.ThuNhapBQToanPX, 2)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("- Tiền lương sản phẩm bình quân LĐTT: Tháng {0}: {1} nghìn đồng/người/tháng; {2}: {3} nghìn đồng/người/tháng; {4} năm {5}: {6} nghìn đồng/người/tháng.", ws.Thang, FmtDec(ws.TLSPBinhQuanLDTT, 2), tenQuy, FmtDec(ws.KyQuy.TLSPBinhQuanLDTT, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.TLSPBinhQuanLDTT, 2)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("- Thu nhập bình quân của LĐTT: Tháng {0}: {1} nghìn đồng/người/tháng; {2}: {3} nghìn đồng/người/tháng; {4} năm {5}: {6} nghìn đồng/người/tháng.", ws.Thang, FmtDec(ws.ThuNhapBQLDTT, 2), tenQuy, FmtDec(ws.KyQuy.ThuNhapBQLDTT, 2), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.ThuNhapBQLDTT, 2)), false, false, "both", 20, 40, 26, 420));
            }

            // Mục 2.2. Ăn ca
            body.Append(P("2.2. Ăn ca", true, true, "both", 40, 20, 26, 420));
            if (!isMultiPeriod)
            {
                body.Append(P(string.Format("- Tổng số công ăn ca tháng {0}/{1}: {2} công; Tổng số tiền ăn ca (mức 30.000đ/công): {3} đồng.", ws.Thang, ws.Nam, FmtDec(ws.CongAnCa, 1), FmtInt(ws.TienAnCa)), false, false, "both", 20, 40, 26, 420));
            }
            else
            {
                body.Append(P(string.Format("- Tổng số công ăn ca: Tháng {0}: {1} công; {2}: {3} công; {4} năm {5}: {6} công.", ws.Thang, FmtDec(ws.CongAnCa, 1), tenQuy, FmtDec(ws.KyQuy.CongAnCa, 1), tenLK, ws.Nam, FmtDec(ws.KyLuyKe.CongAnCa, 1)), false, false, "both", 20, 20, 26, 420));
                body.Append(P(string.Format("- Tổng số tiền ăn ca: Tháng {0}: {1} đồng; {2}: {3} đồng; {4} năm {5}: {6} đồng.", ws.Thang, FmtInt(ws.TienAnCa), tenQuy, FmtInt(ws.KyQuy.TienAnCa), tenLK, ws.Nam, FmtInt(ws.KyLuyKe.TienAnCa)), false, false, "both", 20, 40, 26, 420));
            }

            // Mục 2.3. Bồi dưỡng độc hại bằng hiện vật
            body.Append(P("2.3. Bồi dưỡng độc hại bằng hiện vật", true, true, "both", 40, 30, 26, 420));
            body.Append(BuildBDDHTable(ws, isMultiPeriod, tenQuy, tenLK));
            body.Append(P("", false, false, "both", 20, 20));

            // Mục 3. Công tác bình công chấm điểm, ghi chép sổ sách
            body.Append(P("3. Công tác bình công chấm điểm, ghi chép sổ sách:", true, false, "both", 40, 20, 26, 420));
            body.Append(P("- Thực hiện chấm công, ghi chép sổ sách rõ ràng, đúng quy định hiện hành.", false, false, "both", 20, 20, 26, 420));

            // Mục III. Kiến nghị, đề xuất
            body.Append(P("III. Kiến nghị, đề xuất:", true, false, "both", 60, 20));
            if (ws.KienNghi != null && ws.KienNghi.Count > 0)
            {
                foreach (var kn in ws.KienNghi)
                {
                    if (!string.IsNullOrWhiteSpace(kn))
                        body.Append(P(string.Format("- {0}", kn), false, false, "both", 20, 20, 26, 420));
                }
            }
            body.Append(P(string.Format("Biên bản kết thúc vào hồi {0} cùng ngày, đã được đọc lại cho mọi người cùng nghe và thống nhất ký tên./.", ws.GioKetThuc), false, false, "both", 40, 100, 26, 420));

            // 5. Chữ ký
            body.Append(BuildSignatureTable(ws));

            // 6. Phụ lục Bảng nghiệm thu công, ăn ca & lương chi tiết từng người lao động (nếu chọn đính kèm)
            if (ws.DinhKemBangChiTietNhanVien && ws.ListNhanVien != null && ws.ListNhanVien.Count > 0)
            {
                body.Append(@"<w:p><w:r><w:br w:type=""page""/></w:r></w:p>");
                body.Append(P(string.Format("PHỤ LỤC: BẢNG NGHIỆM THU NGÀY CÔNG, ĂN CA & TIỀN LƯƠNG CHI TIẾT TỪNG NGƯỜI THÁNG {0}/{1}", ws.Thang, ws.Nam), true, false, "center", 80, 20, 26));
                body.Append(P(string.Format("Đơn vị: {0} (Tổng số: {1} lao động)", ws.TenPX, ws.ListNhanVien.Count), true, true, "center", 0, 80, 24));
                body.Append(BuildEmployeeDetailTable(ws));
            }

            return body.ToString();
        }

        private static string BuildDocumentXml(WorkshopData ws)
        {
            var body = new StringBuilder();
            body.Append(BuildReportBodyElementsXml(ws));

            // Page setup A4
            body.Append(@"<w:sectPr>
                <w:pgSz w:w=""11906"" w:h=""16838""/>
                <w:pgMar w:top=""1134"" w:right=""1134"" w:bottom=""1134"" w:left=""1701"" w:header=""720"" w:footer=""720"" w:gutter=""0""/>
            </w:sectPr>");

            return string.Format(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    {0}
  </w:body>
</w:document>", body);
        }

        private static string BuildEmployeeDetailTable(WorkshopData ws)
        {
            var sb = new StringBuilder();
            sb.Append(@"<w:tbl>
                <w:tblPr>
                    <w:tblW w:w=""5000"" w:type=""pct""/>
                    <w:tblBorders>
                        <w:top w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:left w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:bottom w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:right w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:insideH w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:insideV w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                    </w:tblBorders>
                </w:tblPr>");

            sb.Append("<w:tr>");
            sb.Append(TcSmall("TT", 220, true, "center"));
            sb.Append(TcSmall("SDB", 280, true, "center"));
            sb.Append(TcSmall("Họ và tên", 1050, true, "center"));
            sb.Append(TcSmall("Chức danh", 320, true, "center"));
            sb.Append(TcSmall("Tổng công", 320, true, "center"));
            sb.Append(TcSmall("Công SP", 320, true, "center"));
            sb.Append(TcSmall("Phép, Lễ", 300, true, "center"));
            sb.Append(TcSmall("Ăn ca", 300, true, "center"));
            sb.Append(TcSmall("Điểm / HS", 420, true, "center"));
            sb.Append(TcSmall("XL", 220, true, "center"));
            sb.Append(TcSmall("Lương SP (đ)", 620, true, "center"));
            sb.Append(TcSmall("Tổng lương (đ)", 630, true, "center"));
            sb.Append("</w:tr>");

            string lastBoPhan = "";
            double sTongCong = 0, sCongSP = 0, sPhep = 0, sAnCa = 0, sLuongSP = 0, sTongLuong = 0;

            foreach (var e in ws.ListNhanVien)
            {
                if (!string.IsNullOrEmpty(e.BoPhan) && e.BoPhan != lastBoPhan)
                {
                    lastBoPhan = e.BoPhan;
                    sb.Append("<w:tr>");
                    sb.Append(string.Format(@"<w:tc><w:tcPr><w:tcW w:w=""5000"" w:type=""pct""/><w:gridSpan w:val=""12""/></w:tcPr>{0}</w:tc>",
                        P(lastBoPhan, true, true, "left", 20, 20, 20)));
                    sb.Append("</w:tr>");
                }

                sTongCong += e.TongCong;
                sCongSP += e.CongSP;
                sPhep += e.CongLePhep;
                sAnCa += e.CongAnCa;
                sLuongSP += e.LuongSanPham;
                sTongLuong += e.TongTienLuong;

                sb.Append("<w:tr>");
                sb.Append(TcSmall(e.STT.ToString(), 220, false, "center"));
                sb.Append(TcSmall(e.SDB, 280, false, "center"));
                sb.Append(TcSmall(e.HoTen, 1050, false, "left"));
                sb.Append(TcSmall(e.ChucDanh, 320, false, "center"));
                sb.Append(TcSmall(FmtDec(e.TongCong, 1), 320, false, "right"));
                sb.Append(TcSmall(FmtDec(e.CongSP, 1), 320, false, "right"));
                sb.Append(TcSmall(FmtDec(e.CongLePhep, 1), 300, false, "right"));
                sb.Append(TcSmall(FmtDec(e.CongAnCa, 1), 300, false, "right"));
                sb.Append(TcSmall(FmtDec(e.HeSoHuong, 2), 420, false, "right"));
                sb.Append(TcSmall(e.XepLoai, 220, false, "center"));
                sb.Append(TcSmall(FmtInt(e.LuongSanPham), 620, false, "right"));
                sb.Append(TcSmall(FmtInt(e.TongTienLuong), 630, true, "right"));
                sb.Append("</w:tr>");
            }

            sb.Append("<w:tr>");
            sb.Append(TcSmall("", 220, true, "center"));
            sb.Append(TcSmall("", 280, true, "center"));
            sb.Append(TcSmall("TỔNG CỘNG TOÀN PHÂN XƯỞNG", 1050, true, "center"));
            sb.Append(TcSmall("", 320, true, "center"));
            sb.Append(TcSmall(FmtDec(sTongCong, 1), 320, true, "right"));
            sb.Append(TcSmall(FmtDec(sCongSP, 1), 320, true, "right"));
            sb.Append(TcSmall(FmtDec(sPhep, 1), 300, true, "right"));
            sb.Append(TcSmall(FmtDec(sAnCa, 1), 300, true, "right"));
            sb.Append(TcSmall("", 420, true, "right"));
            sb.Append(TcSmall("", 220, true, "center"));
            sb.Append(TcSmall(FmtInt(sLuongSP), 620, true, "right"));
            sb.Append(TcSmall(FmtInt(sTongLuong), 630, true, "right"));
            sb.Append("</w:tr>");

            sb.Append("</w:tbl>");
            return sb.ToString();
        }

        private static string P(string text, bool bold = false, bool italic = false, string align = "both", int before = 40, int after = 40, int fontSizeHalfPt = 26, int firstLineIndent = 0)
        {
            string bTag = bold ? "<w:b/>" : "";
            string iTag = italic ? "<w:i/>" : "";
            string indTag = firstLineIndent > 0 ? string.Format(@"<w:ind w:firstLine=""{0}""/>", firstLineIndent) : "";
            return string.Format(@"<w:p>
                <w:pPr>
                    <w:spacing w:before=""{0}"" w:after=""{1}"" w:line=""276"" w:lineRule=""auto""/>
                    {2}
                    <w:jc w:val=""{3}""/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                        {4}{5}
                        <w:sz w:val=""{6}""/>
                        <w:szCs w:val=""{7}""/>
                    </w:rPr>
                    <w:t xml:space=""preserve"">{8}</w:t>
                </w:r>
            </w:p>", before, after, indTag, align, bTag, iTag, fontSizeHalfPt, fontSizeHalfPt, Esc(text));
        }

        private static string PLabelValue(string boldPrefix, string normalText, string align = "both", int before = 40, int after = 40, int firstLineIndent = 0)
        {
            string indTag = firstLineIndent > 0 ? string.Format(@"<w:ind w:firstLine=""{0}""/>", firstLineIndent) : "";
            return string.Format(@"<w:p>
                <w:pPr>
                    <w:spacing w:before=""{0}"" w:after=""{1}"" w:line=""276"" w:lineRule=""auto""/>
                    {2}
                    <w:jc w:val=""{3}""/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                        <w:b/>
                        <w:sz w:val=""26""/><w:szCs w:val=""26""/>
                    </w:rPr>
                    <w:t xml:space=""preserve"">{4}</w:t>
                </w:r>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                        <w:sz w:val=""26""/><w:szCs w:val=""26""/>
                    </w:rPr>
                    <w:t xml:space=""preserve"">{5}</w:t>
                </w:r>
            </w:p>", before, after, indTag, align, Esc(boldPrefix), Esc(normalText));
        }

        private static string BuildHeaderTable()
        {
            return @"<w:tbl>
                <w:tblPr>
                    <w:tblW w:w=""5000"" w:type=""pct""/>
                    <w:tblBorders>
                        <w:top w:val=""none""/><w:left w:val=""none""/><w:bottom w:val=""none""/><w:right w:val=""none""/>
                        <w:insideH w:val=""none""/><w:insideV w:val=""none""/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tr>
                    <w:tc>
                        <w:tcPr><w:tcW w:w=""2350"" w:type=""pct""/></w:tcPr>
                        " + P("CÔNG TY CP GANG THÉP CAO BẰNG", false, false, "center", 0, 20, 24) + @"
                        " + P("PHÒNG TỔ CHỨC HÀNH CHÍNH", true, false, "center", 0, 20, 24) + @"
                        " + P("________________________", true, false, "center", 0, 40, 18) + @"
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w=""2650"" w:type=""pct""/></w:tcPr>
                        " + P("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, false, "center", 0, 20, 24) + @"
                        " + P("Độc lập - Tự do - Hạnh phúc", true, false, "center", 0, 20, 26) + @"
                        " + P("________________________", true, false, "center", 0, 40, 18) + @"
                    </w:tc>
                </w:tr>
            </w:tbl>";
        }

        private static string BuildGNVTable(List<ArisingWorkItem> items)
        {
            var sb = new StringBuilder();
            sb.Append(@"<w:tbl>
                <w:tblPr>
                    <w:tblW w:w=""5000"" w:type=""pct""/>
                    <w:tblBorders>
                        <w:top w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:left w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:bottom w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:right w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:insideH w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:insideV w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                    </w:tblBorders>
                </w:tblPr>");

            sb.Append("<w:tr>");
            sb.Append(Tc("STT", 350, true, "center"));
            sb.Append(Tc("Nội dung công việc", 2150, true, "center"));
            sb.Append(Tc("Đơn vị tính", 500, true, "center"));
            sb.Append(Tc("KL thực hiện", 600, true, "center"));
            sb.Append(Tc("Đơn giá (đ)", 650, true, "center"));
            sb.Append(Tc("Thành tiền (đ)", 750, true, "center"));
            sb.Append("</w:tr>");

            double sumKL = 0;
            double sumTT = 0;
            int idx = 1;
            foreach (var it in items)
            {
                sumKL += it.KhoiLuong;
                sumTT += it.ThanhTien;
                sb.Append("<w:tr>");
                sb.Append(Tc((idx++).ToString(), 350, false, "center"));
                sb.Append(Tc(it.NoiDung, 2150, false, "left"));
                sb.Append(Tc(it.DVT, 500, false, "center"));
                sb.Append(Tc(FmtDec(it.KhoiLuong, 2), 600, false, "right"));
                sb.Append(Tc(FmtInt(it.DonGia), 650, false, "right"));
                sb.Append(Tc(FmtInt(it.ThanhTien), 750, false, "right"));
                sb.Append("</w:tr>");
            }

            sb.Append("<w:tr>");
            sb.Append(Tc("", 350, true, "center"));
            sb.Append(Tc("Tổng cộng", 2150, true, "center"));
            sb.Append(Tc("Công", 500, true, "center"));
            sb.Append(Tc(FmtDec(sumKL, 2), 600, true, "right"));
            sb.Append(Tc("", 650, true, "right"));
            sb.Append(Tc(FmtInt(sumTT), 750, true, "right"));
            sb.Append("</w:tr>");

            sb.Append("</w:tbl>");
            return sb.ToString();
        }

        /// <summary>
        /// Bảng Bồi dưỡng độc hại đúng chuẩn mẫu CISCO:
        /// Stt | Mức | Đơn giá | Tháng X (Công được hưởng | Thành tiền) | Quý Y (Công được hưởng | Thành tiền) | Z tháng (Công được hưởng | Thành tiền) | Ghi chú
        /// </summary>
        private static string BuildBDDHTable(WorkshopData ws, bool isMultiPeriod, string tenQuy, string tenLK)
        {
            var sb = new StringBuilder();
            sb.Append(@"<w:tbl>
                <w:tblPr>
                    <w:tblW w:w=""5000"" w:type=""pct""/>
                    <w:tblBorders>
                        <w:top w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:left w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:bottom w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:right w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:insideH w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                        <w:insideV w:val=""single"" w:sz=""6"" w:space=""0"" w:color=""000000""/>
                    </w:tblBorders>
                </w:tblPr>");

            if (!isMultiPeriod)
            {
                sb.Append("<w:tr>");
                sb.Append(Tc("Stt", 400, true, "center"));
                sb.Append(Tc("Mức", 800, true, "center"));
                sb.Append(Tc("Đơn giá", 900, true, "center"));
                sb.Append(Tc("Công được hưởng", 1100, true, "center"));
                sb.Append(Tc("Thành tiền", 1200, true, "center"));
                sb.Append(Tc("Ghi chú", 600, true, "center"));
                sb.Append("</w:tr>");

                sb.Append("<w:tr>");
                sb.Append(Tc("1", 400, false, "center"));
                sb.Append(Tc("Mức 1", 800, false, "center"));
                sb.Append(Tc("13.000", 900, false, "right"));
                sb.Append(Tc(FmtDec(ws.BDDH_Muc1_Cong, 1), 1100, false, "right"));
                sb.Append(Tc(FmtInt(ws.BDDH_Muc1_Tien), 1200, false, "right"));
                sb.Append(Tc("", 600, false, "center"));
                sb.Append("</w:tr>");

                sb.Append("<w:tr>");
                sb.Append(Tc("2", 400, false, "center"));
                sb.Append(Tc("Mức 2", 800, false, "center"));
                sb.Append(Tc("20.000", 900, false, "right"));
                sb.Append(Tc(FmtDec(ws.BDDH_Muc2_Cong, 1), 1100, false, "right"));
                sb.Append(Tc(FmtInt(ws.BDDH_Muc2_Tien), 1200, false, "right"));
                sb.Append(Tc("", 600, false, "center"));
                sb.Append("</w:tr>");

                if (ws.CoMuc3BDDH || ws.BDDH_Muc3_Cong > 0)
                {
                    sb.Append("<w:tr>");
                    sb.Append(Tc("3", 400, false, "center"));
                    sb.Append(Tc("Mức 3", 800, false, "center"));
                    sb.Append(Tc("26.000", 900, false, "right"));
                    sb.Append(Tc(FmtDec(ws.BDDH_Muc3_Cong, 1), 1100, false, "right"));
                    sb.Append(Tc(FmtInt(ws.BDDH_Muc3_Tien), 1200, false, "right"));
                    sb.Append(Tc("", 600, false, "center"));
                    sb.Append("</w:tr>");
                }

                double tongC = ws.BDDH_Muc1_Cong + ws.BDDH_Muc2_Cong + ws.BDDH_Muc3_Cong;
                sb.Append("<w:tr>");
                sb.Append(Tc("", 400, true, "center"));
                sb.Append(Tc("Tổng", 800, true, "center"));
                sb.Append(Tc("", 900, true, "center"));
                sb.Append(Tc(FmtDec(tongC, 1), 1100, true, "right"));
                sb.Append(Tc(FmtInt(ws.TongTienBDDH), 1200, true, "right"));
                sb.Append(Tc("", 600, true, "center"));
                sb.Append("</w:tr>");
            }
            else
            {
                // Bảng 10 cột chuẩn CISCO: Stt | Mức | Đơn giá | Tháng X (Công | Tiền) | Quý (Công | Tiền) | Lũy kế (Công | Tiền) | Ghi chú
                sb.Append("<w:tr>");
                sb.Append(TcSpan("Stt", 250, true, "center", 1, "restart"));
                sb.Append(TcSpan("Mức", 450, true, "center", 1, "restart"));
                sb.Append(TcSpan("Đơn giá", 480, true, "center", 1, "restart"));
                sb.Append(TcSpan(string.Format("Tháng {0}", ws.Thang), 1140, true, "center", 2, null));
                sb.Append(TcSpan(tenQuy, 1140, true, "center", 2, null));
                sb.Append(TcSpan(tenLK, 1190, true, "center", 2, null));
                sb.Append(TcSpan("Ghi chú", 350, true, "center", 1, "restart"));
                sb.Append("</w:tr>");

                sb.Append("<w:tr>");
                sb.Append(TcSpan("", 250, true, "center", 1, "continue"));
                sb.Append(TcSpan("", 450, true, "center", 1, "continue"));
                sb.Append(TcSpan("", 480, true, "center", 1, "continue"));
                sb.Append(Tc("Công được hưởng", 480, true, "center"));
                sb.Append(Tc("Thành tiền", 660, true, "center"));
                sb.Append(Tc("Công được hưởng", 480, true, "center"));
                sb.Append(Tc("Thành tiền", 660, true, "center"));
                sb.Append(Tc("Công được hưởng", 500, true, "center"));
                sb.Append(Tc("Thành tiền", 690, true, "center"));
                sb.Append(TcSpan("", 350, true, "center", 1, "continue"));
                sb.Append("</w:tr>");

                sb.Append("<w:tr>");
                sb.Append(Tc("1", 250, false, "center"));
                sb.Append(Tc("Mức 1", 450, false, "center"));
                sb.Append(Tc("13.000", 480, false, "right"));
                sb.Append(Tc(FmtDec(ws.BDDH_Muc1_Cong, 1), 480, false, "right"));
                sb.Append(Tc(FmtInt(ws.BDDH_Muc1_Tien), 660, false, "right"));
                sb.Append(Tc(FmtDec(ws.KyQuy.BDDH_Muc1_Cong, 1), 480, false, "right"));
                sb.Append(Tc(FmtInt(ws.KyQuy.BDDH_Muc1_Tien), 660, false, "right"));
                sb.Append(Tc(FmtDec(ws.KyLuyKe.BDDH_Muc1_Cong, 1), 500, false, "right"));
                sb.Append(Tc(FmtInt(ws.KyLuyKe.BDDH_Muc1_Tien), 690, false, "right"));
                sb.Append(Tc("", 350, false, "center"));
                sb.Append("</w:tr>");

                sb.Append("<w:tr>");
                sb.Append(Tc("2", 250, false, "center"));
                sb.Append(Tc("Mức 2", 450, false, "center"));
                sb.Append(Tc("20.000", 480, false, "right"));
                sb.Append(Tc(FmtDec(ws.BDDH_Muc2_Cong, 1), 480, false, "right"));
                sb.Append(Tc(FmtInt(ws.BDDH_Muc2_Tien), 660, false, "right"));
                sb.Append(Tc(FmtDec(ws.KyQuy.BDDH_Muc2_Cong, 1), 480, false, "right"));
                sb.Append(Tc(FmtInt(ws.KyQuy.BDDH_Muc2_Tien), 660, false, "right"));
                sb.Append(Tc(FmtDec(ws.KyLuyKe.BDDH_Muc2_Cong, 1), 500, false, "right"));
                sb.Append(Tc(FmtInt(ws.KyLuyKe.BDDH_Muc2_Tien), 690, false, "right"));
                sb.Append(Tc("", 350, false, "center"));
                sb.Append("</w:tr>");

                if (ws.CoMuc3BDDH || ws.BDDH_Muc3_Cong > 0 || ws.KyLuyKe.BDDH_Muc3_Cong > 0)
                {
                    sb.Append("<w:tr>");
                    sb.Append(Tc("3", 250, false, "center"));
                    sb.Append(Tc("Mức 3", 450, false, "center"));
                    sb.Append(Tc("26.000", 480, false, "right"));
                    sb.Append(Tc(FmtDec(ws.BDDH_Muc3_Cong, 1), 480, false, "right"));
                    sb.Append(Tc(FmtInt(ws.BDDH_Muc3_Tien), 660, false, "right"));
                    sb.Append(Tc(FmtDec(ws.KyQuy.BDDH_Muc3_Cong, 1), 480, false, "right"));
                    sb.Append(Tc(FmtInt(ws.KyQuy.BDDH_Muc3_Tien), 660, false, "right"));
                    sb.Append(Tc(FmtDec(ws.KyLuyKe.BDDH_Muc3_Cong, 1), 500, false, "right"));
                    sb.Append(Tc(FmtInt(ws.KyLuyKe.BDDH_Muc3_Tien), 690, false, "right"));
                    sb.Append(Tc("", 350, false, "center"));
                    sb.Append("</w:tr>");
                }

                double cT = ws.BDDH_Muc1_Cong + ws.BDDH_Muc2_Cong + ws.BDDH_Muc3_Cong;
                double cQ = ws.KyQuy.BDDH_Muc1_Cong + ws.KyQuy.BDDH_Muc2_Cong + ws.KyQuy.BDDH_Muc3_Cong;
                double cLK = ws.KyLuyKe.BDDH_Muc1_Cong + ws.KyLuyKe.BDDH_Muc2_Cong + ws.KyLuyKe.BDDH_Muc3_Cong;

                sb.Append("<w:tr>");
                sb.Append(Tc("", 250, true, "center"));
                sb.Append(Tc("Tổng", 450, true, "center"));
                sb.Append(Tc("", 480, true, "center"));
                sb.Append(Tc(FmtDec(cT, 1), 480, true, "right"));
                sb.Append(Tc(FmtInt(ws.TongTienBDDH), 660, true, "right"));
                sb.Append(Tc(FmtDec(cQ, 1), 480, true, "right"));
                sb.Append(Tc(FmtInt(ws.KyQuy.TongTienBDDH), 660, true, "right"));
                sb.Append(Tc(FmtDec(cLK, 1), 500, true, "right"));
                sb.Append(Tc(FmtInt(ws.KyLuyKe.TongTienBDDH), 690, true, "right"));
                sb.Append(Tc("", 350, true, "center"));
                sb.Append("</w:tr>");
            }

            sb.Append("</w:tbl>");
            return sb.ToString();
        }

        private static string BuildSignatureTable(WorkshopData ws)
        {
            return @"<w:tbl>
                <w:tblPr>
                    <w:tblW w:w=""5000"" w:type=""pct""/>
                    <w:tblBorders>
                        <w:top w:val=""none""/><w:left w:val=""none""/><w:bottom w:val=""none""/><w:right w:val=""none""/>
                        <w:insideH w:val=""none""/><w:insideV w:val=""none""/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tr>
                    <w:tc>
                        <w:tcPr><w:tcW w:w=""2500"" w:type=""pct""/></w:tcPr>
                        " + P("QUẢN ĐỐC", true, false, "center", 60, 1200, 26) + @"
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w=""2500"" w:type=""pct""/></w:tcPr>
                        " + P("TP. TCHC", true, false, "center", 60, 1200, 26) + @"
                    </w:tc>
                </w:tr>
            </w:tbl>";
        }

        private static string Tc(string text, int widthPct, bool bold, string align)
        {
            return string.Format(@"<w:tc>
                <w:tcPr>
                    <w:tcW w:w=""{0}"" w:type=""pct""/>
                    <w:vAlign w:val=""center""/>
                    <w:tcMar>
                        <w:top w:w=""50"" w:type=""dxa""/>
                        <w:bottom w:w=""50"" w:type=""dxa""/>
                        <w:left w:w=""60"" w:type=""dxa""/>
                        <w:right w:w=""60"" w:type=""dxa""/>
                    </w:tcMar>
                </w:tcPr>
                {1}
            </w:tc>", widthPct, P(text, bold, false, align, 20, 20, 22));
        }

        private static string TcSpan(string text, int widthPct, bool bold, string align, int gridSpan, string vMerge)
        {
            string spanXml = gridSpan > 1 ? string.Format(@"<w:gridSpan w:val=""{0}""/>", gridSpan) : "";
            string vmXml = !string.IsNullOrEmpty(vMerge)
                ? (vMerge == "restart" ? @"<w:vMerge w:val=""restart""/>" : @"<w:vMerge/>")
                : "";
            return string.Format(@"<w:tc>
                <w:tcPr>
                    <w:tcW w:w=""{0}"" w:type=""pct""/>
                    {1}{2}
                    <w:vAlign w:val=""center""/>
                    <w:tcMar>
                        <w:top w:w=""50"" w:type=""dxa""/>
                        <w:bottom w:w=""50"" w:type=""dxa""/>
                        <w:left w:w=""60"" w:type=""dxa""/>
                        <w:right w:w=""60"" w:type=""dxa""/>
                    </w:tcMar>
                </w:tcPr>
                {3}
            </w:tc>", widthPct, spanXml, vmXml, P(text, bold, false, align, 20, 20, 22));
        }

        private static string TcSmall(string text, int widthPct, bool bold, string align)
        {
            return string.Format(@"<w:tc>
                <w:tcPr>
                    <w:tcW w:w=""{0}"" w:type=""pct""/>
                    <w:vAlign w:val=""center""/>
                    <w:tcMar>
                        <w:top w:w=""30"" w:type=""dxa""/>
                        <w:bottom w:w=""30"" w:type=""dxa""/>
                        <w:left w:w=""40"" w:type=""dxa""/>
                        <w:right w:w=""40"" w:type=""dxa""/>
                    </w:tcMar>
                </w:tcPr>
                {1}
            </w:tc>", widthPct, P(text, bold, false, align, 10, 10, 19));
        }

        private static string GetContentTypesXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
  <Override PartName=""/word/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml""/>
</Types>";
        }

        private static string GetRootRelsXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>
</Relationships>";
        }

        private static string GetDocRelsXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>";
        }

        private static string GetStylesXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:styles xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
        <w:sz w:val=""26""/>
        <w:szCs w:val=""26""/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>";
        }
    }
}
