using System;
using System.Collections.Generic;

namespace CiscoLDTLApp
{
    public enum LoaiKyBaoCao
    {
        ThangThuong = 0,       // Chỉ báo cáo Tháng (VD: Tháng 1, 2, 4, 5, 7, 8, 10, 11)
        QuyI = 1,              // Tháng 3 & Quý I
        QuyII_6Thang = 2,      // Tháng 6, Quý II & 6 tháng đầu năm
        QuyIII_9Thang = 3,     // Tháng 9, Quý III & 9 tháng năm
        QuyIV_CaNam = 4        // Tháng 12, Quý IV & Cả năm
    }

    public class ThanhPhanItem
    {
        public string HoTen { get; set; }
        public string ChucVu { get; set; }

        public ThanhPhanItem() { }
        public ThanhPhanItem(string hoTen, string chucVu)
        {
            HoTen = hoTen;
            ChucVu = chucVu;
        }
    }

    public class ArisingWorkItem
    {
        public int STT { get; set; }
        public string NoiDung { get; set; }
        public string DVT { get; set; }
        public double KhoiLuong { get; set; }
        public double DonGia { get; set; }
        public double ThanhTien { get; set; }
        public string GhiChu { get; set; }
        public int ExcelRow { get; set; }

        public ArisingWorkItem() { }
        public ArisingWorkItem(int stt, string noiDung, string dvt, double kl, double dg, double tt, string ghiChu = "", int row = 0)
        {
            STT = stt;
            NoiDung = noiDung;
            DVT = dvt;
            KhoiLuong = kl;
            DonGia = dg;
            ThanhTien = tt;
            GhiChu = ghiChu;
            ExcelRow = row;
        }
    }

    /// <summary>
    /// Hồ sơ chấm công, ăn ca, điểm sản xuất và tiền lương CHI TIẾT TỪNG NGƯỜI LAO ĐỘNG
    /// Liên kết trực tiếp với dòng dữ liệu trong sheet phân xưởng ('T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển')
    /// và sheet 'Tổng hợp công'.
    /// </summary>
    public class EmployeeRecord
    {
        public int STT { get; set; }
        public int ExcelRow { get; set; }           // Dòng trong sheet phân xưởng ('T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển')
        public int TongHopCongRow { get; set; }     // Dòng trong sheet 'Tổng hợp công'
        public string DanhXung { get; set; }        // ông / bà
        public string HoTen { get; set; }
        public string SDB { get; set; }             // Mã số / Số danh bạ (VD: 203, 310, 096...)
        public string BoPhan { get; set; }          // Tổ sản xuất / Bộ phận (VD: I. Quản lý PX, Tổ Phối liệu A, Tổ Thiêu kết A...)
        public bool QuanLyTinhMoi { get; set; } // Đã áp dụng màn hình lương quản lý v5
        public bool LaQuanLy { get; set; }          // true nếu thuộc bộ phận Quản lý PX
        public string ChucDanh { get; set; }        // QĐ, PQĐ, KTV, TK, TRC, TT, TP, CN, ATV...

        public double MucLuongBHXH { get; set; }    // Mức lương đóng BHXH (Col G)
        public double PhuCapChucVu { get; set; }    // Hệ số phụ cấp CV, trách nhiệm, ATVSV (Col H)
        public double HeSoCoDinh { get; set; }      // Hệ số lương cố định / Điểm BQ ngày (Col I hoặc V/X)

        // Ngày công chi tiết từng người
        public double TongCong { get; set; }        // Col J = CongSP + CongLePhep + CongVHTT
        public double CongSP { get; set; }          // Col K: Số công sản phẩm (làm việc thực tế)
        public double CongCaDem { get; set; }       // Col L: Công ca đêm / phát sinh
        public double CongLePhep { get; set; }      // Col M/N: Công nghỉ phép, lễ hưởng 100% lương
        public double CongViecRieng { get; set; }   // Col O: Nghỉ việc riêng
        public double CongHocHop { get; set; }      // Col P: Công học, họp (hưởng lương thời gian)
        public double CongLamLe { get; set; }       // Col Q: Công đi làm ngày lễ, tết
        public double CongVHTT { get; set; }        // Col R: Công Văn hóa thể thao, thi ATVSV
        public double CongBHXH { get; set; }        // Col S: Công nghỉ ốm, thai sản hưởng BHXH
        public double CongAnCa { get; set; }        // Col T (hoặc V ở CĐNL): Số công tính tiền ăn giữa ca

        // Điểm sản xuất / Xếp loại A-B-C
        public double DiemSanXuat { get; set; }     // Col W (hoặc Y ở CĐNL): Điểm sản xuất hoặc Hệ số lương tháng
        public string XepLoai { get; set; }         // Col X (hoặc Z ở CĐNL): Phân loại A, B, C
        public double HeSoHuong { get; set; }       // Col Z (hoặc AB ở CĐNL): Hệ số / Điểm hưởng lương sau phân loại

        // Tiền lương chi tiết từng người (đồng)
        public double LuongSanPham { get; set; }    // Col AB (hoặc AD ở CĐNL): Tiền lương sản phẩm tháng
        public double LuongLePhepCD { get; set; }   // Col AD (hoặc AE ở CĐNL): Tiền lương lễ, phép, chế độ
        public double LuongLamLe { get; set; }      // Col AH (hoặc AI ở CĐNL): Tiền lương làm ngày lễ
        public double LuongHocVHTT { get; set; }    // Col AI (hoặc AJ ở CĐNL): Tiền lương học họp, VHTT
        public double PhuCapTien { get; set; }      // Col AK (hoặc AL ở CĐNL): Phụ cấp trách nhiệm / ATVSV
        public double TongTienLuong { get; set; }   // Col AA (hoặc AC ở CĐNL): Tổng tiền lương thực lĩnh tháng

        // ===== Dữ liệu trích xuất trực tiếp từ FILE CÔNG PHÂN XƯỞNG (Bảng công, QT ăn ca, Bảng BDĐH, Bảng điểm) =====
        public string ToSanXuat { get; set; }       // Tổ theo file công PX (VD: Tổ Quản lý, Tổ Trưởng ca, Tổ Phối liệu...)
        public double CongK3 { get; set; }          // Công ca 3 (K3)
        public double AnCaNhaAn { get; set; }       // Số suất báo ăn tại nhà ăn (ký hiệu "na")
        public double AnCaTienMat { get; set; }     // Số công ăn ca thanh toán tiền (ký hiệu "tt")
        public double TienAnCa { get; set; }        // Tổng tiền quyết toán ăn ca tháng
        public double BDDH_M1 { get; set; }         // Công BDĐH mức 1 (13.000đ)
        public double BDDH_M2 { get; set; }         // Công BDĐH mức 2 (20.000đ)
        public double BDDH_M3 { get; set; }         // Công BDĐH mức 3 (26.000đ)
        public double TienBDDH { get; set; }        // Tổng tiền BDĐH theo định mức
        public bool CoBangCong { get; set; }        // Đã khớp được trên Bảng công
        public bool CoAnCa { get; set; }            // Đã khớp được trên QT ăn ca
        public bool CoBDDH { get; set; }            // Đã khớp được trên Bảng BDĐH
        public string NguonDuLieu { get; set; }     // Mô tả nguồn khớp (SDB / Họ tên)

        // ===== Công ĂN CA: so sánh Công thức vs Đơn vị gửi, người dùng chọn nguồn =====
        public double AnCaCongThuc { get; set; }    // = Công SP + Công Học (H) + Công VHTT
        public double AnCaDonVi { get; set; }       // = na + tt theo QT ăn ca đơn vị gửi
        public double AnCaThuCong { get; set; }     // Số nhập thủ công
        public string AnCaNguon { get; set; }       // "CT" = công thức | "DV" = đơn vị | "TC" = thủ công | "" = giữ số cũ trong Excel

        public double AnCaChenhLech { get { return CoAnCa ? AnCaDonVi - AnCaCongThuc : 0; } }

        public EmployeeRecord()
        {
            DanhXung = "ông";
            XepLoai = "A";
            BoPhan = "Lao động trực tiếp";
        }
    }

    /// <summary>
    /// Tổng hợp số liệu theo TỔ SẢN XUẤT (cộng từ chi tiết từng người).
    /// </summary>
    public class ToSummary
    {
        public string TenTo { get; set; }
        public bool LaQuanLy { get; set; }
        public int SoLD { get; set; }
        public double TongCong { get; set; }
        public double CongSP { get; set; }
        public double CongLePhep { get; set; }
        public double CongHocHop { get; set; }
        public double CongVHTT { get; set; }
        public double CongBHXH { get; set; }
        public double CongK3 { get; set; }
        public double CongAnCa { get; set; }
        public double AnCaNhaAn { get; set; }
        public double AnCaTienMat { get; set; }
        public double TienAnCa { get; set; }
        public double BDDH_M1 { get; set; }
        public double BDDH_M2 { get; set; }
        public double BDDH_M3 { get; set; }
        public double TienBDDH { get; set; }
        public double TongTienLuong { get; set; }
    }

    public class PeriodMetrics
    {
        public string TenKy { get; set; } // VD: "Quý III", "9 tháng"
        public double TongCong { get; set; }
        public double CongSP { get; set; }
        public double CongQL { get; set; }
        public double CongLDTT { get; set; }
        public double CongCheDoBHXH { get; set; }
        public double CongCheDo { get; set; }
        public double CongBHXH { get; set; }
        public double LDDanhSachBQ { get; set; }
        public double LDThucTeBQ { get; set; }
        public double LDTTThucTeBQ { get; set; }
        public double NgayCongBQ { get; set; }
        public double NgayCongBQQL { get; set; }
        public double NgayCongBQLDTT { get; set; }

        public double SanLuongTH { get; set; }
        public double SanLuongTinhLuong { get; set; }
        public double SanLuongPhu { get; set; }

        public double TongLuongToanPX { get; set; }
        public double TongLuongTT { get; set; }
        public double LuongSPTT { get; set; }
        public double LuongCVKhac { get; set; }
        public double CongCVKhac { get; set; }
        public double LuongCheDoTT { get; set; }
        public double CongCheDoTT { get; set; }
        public double LuongTDT { get; set; }
        public double CongTDT { get; set; }
        public double TruLuong { get; set; }

        public double TongLuongQL { get; set; }
        public double LuongSPQL { get; set; }
        public double LuongCheDoQL { get; set; }
        public double DieuChinhLuongQL { get; set; }
        public double CongCheDoQL { get; set; }

        public double TLBinhQuanToanPX { get; set; }
        public double ThuNhapBQToanPX { get; set; }
        public double TLSPBinhQuanLDTT { get; set; }
        public double ThuNhapBQLDTT { get; set; }
        public double BQNgayCongLVTT { get; set; }

        public double CongAnCa { get; set; }
        public double TienAnCa { get; set; }

        public double BDDH_Muc1_Cong { get; set; }
        public double BDDH_Muc1_Tien { get; set; }
        public double BDDH_Muc2_Cong { get; set; }
        public double BDDH_Muc2_Tien { get; set; }
        public double BDDH_Muc3_Cong { get; set; }
        public double BDDH_Muc3_Tien { get; set; }
        public double TongTienBDDH { get; set; }

        public PeriodMetrics Clone()
        {
            return (PeriodMetrics)this.MemberwiseClone();
        }
    }

    public class WorkshopData
    {
        public string MaPX { get; set; }              // TK, LG, LT, CDNL, KT
        public string SheetNT { get; set; }           // NT thiêu kết, NT gang, NT Thép, NT CĐNL, NT. K. tuyển
        public string Sheet9Thang { get; set; }       // 9 tháng Thiêu kết, ...
        public string SheetBangCong { get; set; }     // T kết, Gang, thép, CĐNL, K. tuyển
        public string TenPX { get; set; }             // Phân xưởng Thiêu kết, ...
        public int RowTTBDDH_T9 { get; set; }         // Dòng trong sheet TT BDĐH
        public int TotalRowBangCong { get; set; }     // Dòng CỘNG trong sheet T kết (119), Gang (167), thép (208), CĐNL (138), K. tuyển (43)

        // Cờ bật/tắt tự động cộng lũy kế Quý & Năm từ số liệu Tháng
        public bool TuDongCongLuyKe { get; set; }
        // Cờ đính kèm Phụ lục danh sách nghiệm thu chi tiết từng người vào cuối Biên bản Word
        public bool DinhKemBangChiTietNhanVien { get; set; }

        // Danh sách chi tiết từng người lao động trong phân xưởng
        public List<EmployeeRecord> ListNhanVien { get; set; }

        // Thông tin chung biên bản
        public int Thang { get; set; }
        public int Nam { get; set; }
        public LoaiKyBaoCao LoaiKy { get; set; }
        public int NgayHop { get; set; }
        public int ThangHop { get; set; }
        public int NamHop { get; set; }
        public string DiaDiem { get; set; }
        public string GioKetThuc { get; set; }

        public List<ThanhPhanItem> ThanhPhanTCHC { get; set; }
        public List<ThanhPhanItem> ThanhPhanPX { get; set; }

        // 1.1 Biến động lao động
        public int LDDauKy { get; set; }
        public int LDTang { get; set; }
        public int LDGiam { get; set; }
        public string LyDoBienDong { get; set; }
        public int LDCuoiKy { get; set; }
        public int LDQuanLy { get; set; }
        public int LDTrucTiep { get; set; }
        public string GhiChuLaoDong { get; set; }

        // 1.2 Ngày công tháng
        public double TongCong { get; set; }
        public double CongSP { get; set; }
        public double CongQL { get; set; }
        public double CongLDTT { get; set; }
        public double CongLePhepCD { get; set; }
        public double CongLePhepTT { get; set; }
        public double CongHocHopTT { get; set; }
        public double CongVHTTTT { get; set; }
        public double CongCDQL { get; set; }
        public double CongBHXH { get; set; }
        public double LDThucTeBQ { get; set; }
        public double LDTTThucTeBQ { get; set; }
        public double NgayCongBQ { get; set; }
        public double NgayCongBQQL { get; set; }
        public double NgayCongBQLDTT { get; set; }

        // 2.1 Sản lượng & Tiền lương tháng
        public string TenSanPhamChinh { get; set; }
        public string DvtSanPham { get; set; }
        public double SanLuongKH { get; set; }
        public double SanLuongTH { get; set; }
        public double SanLuongTinhLuong { get; set; }
        public double SanLuongLoai1 { get; set; }
        public double DonGiaLoai1 { get; set; }
        public double ThanhTienLoai1 { get; set; }
        public double SanLuongLoai2 { get; set; }
        public double DonGiaLoai2 { get; set; }
        public double ThanhTienLoai2 { get; set; }
        public double SanLuongLe { get; set; }
        public double DonGiaLe { get; set; }
        public double ThanhTienLe { get; set; }
        public string SanLuongPhuName { get; set; }
        public double SanLuongPhuKL { get; set; }
        public double SanLuongPhuDonGia { get; set; }
        public double SanLuongPhuTT { get; set; }

        public double TongLuongToanPX { get; set; }
        public double TongLuongTT { get; set; }
        public double LuongSPTT { get; set; }
        public double LuongCVKhac { get; set; }
        public double CongCVKhac { get; set; }
        public double LuongCheDoTT { get; set; }
        public double LuongLePhepTT { get; set; }
        public double LuongHocHopTT { get; set; }
        public double LuongVHTTTT { get; set; }
        public double LuongTDT { get; set; }
        public double CongTDT { get; set; }
        public double TruLuong { get; set; }
        public string LyDoTruLuong { get; set; }

        public double TongLuongQL { get; set; }
        public string CanhBaoQuyQuanLy { get; set; }
        public double LuongSPQL { get; set; }
        public double LuongCheDoQL { get; set; }
        public double DieuChinhLuongQL { get; set; }

        // Chỉ tiêu bình quân tháng (ĐVT: 1.000đ hoặc đồng)
        public double TLBinhQuanToanPX { get; set; }
        public double ThuNhapBQToanPX { get; set; }
        public double TLSPBinhQuanLDTT { get; set; }
        public double ThuNhapBQLDTT { get; set; }
        public double BQNgayCongLVTT { get; set; }

        // 2.2 Ăn giữa ca tháng
        public double CongAnCa { get; set; }
        public double TienAnCa { get; set; }

        // 2.3 Bồi dưỡng độc hại tháng
        public bool CoMuc3BDDH { get; set; }
        public double BDDH_Muc1_Cong { get; set; }
        public double BDDH_Muc1_Tien { get; set; }
        public double BDDH_Muc2_Cong { get; set; }
        public double BDDH_Muc2_Tien { get; set; }
        public double BDDH_Muc3_Cong { get; set; }
        public double BDDH_Muc3_Tien { get; set; }
        public double TongTienBDDH { get; set; }

        // Danh sách công phát sinh / GNV trong tháng
        public List<ArisingWorkItem> ListGNV { get; set; }

        // Số liệu nền (Các tháng trước trong Quý & Các Quý trước trong Năm) để tự động cộng lũy kế
        public PeriodMetrics TruocTrongQuy { get; set; }
        public PeriodMetrics TruocTrongNam { get; set; }

        // Số liệu Quý và Lũy kế (6 tháng / 9 tháng / Cả năm)
        public PeriodMetrics KyQuy { get; set; }
        public PeriodMetrics KyLuyKe { get; set; }

        // 3. Kiến nghị, đề xuất
        public List<string> KienNghi { get; set; }

        public WorkshopData()
        {
            ListNhanVien = new List<EmployeeRecord>();
            ThanhPhanTCHC = new List<ThanhPhanItem>();
            ThanhPhanPX = new List<ThanhPhanItem>();
            ListGNV = new List<ArisingWorkItem>();
            KienNghi = new List<string>();
            TruocTrongQuy = new PeriodMetrics { TenKy = "Tháng trước trong Quý" };
            TruocTrongNam = new PeriodMetrics { TenKy = "Các Quý trước" };
            KyQuy = new PeriodMetrics { TenKy = "Quý III" };
            KyLuyKe = new PeriodMetrics { TenKy = "9 tháng" };
            TuDongCongLuyKe = true;
            DinhKemBangChiTietNhanVien = false;
            DiaDiem = "Phòng Tổ chức Hành chính";
            GioKetThuc = "11 giờ 00 phút";
            Thang = 9;
            Nam = 2026;
            LoaiKy = LoaiKyBaoCao.QuyIII_9Thang;
            NgayHop = 6;
            ThangHop = 10;
            NamHop = 2026;
        }

        /// <summary>
        /// Tính toán lại tiền lương, ngày công của từng người lao động trong ListNhanVien
        /// và tự động tổng hợp lên các chỉ tiêu toàn Phân xưởng (Bảng nghiệm thu & Biên bản Word).
        /// </summary>
        public void RecalculateFromEmployees(bool redistributeProductSalary = true)
        {
            if (ListNhanVien == null || ListNhanVien.Count == 0)
            {
                RecalculateDerived();
                return;
            }

            double sumCongSP = 0;
            double sumCongQL = 0;
            double sumCongLDTT = 0;
            double sumCongLePhepTT = 0;
            double sumCongHocHopTT = 0;
            double sumCongVHTTTT = 0;
            double sumCongCDQL = 0;
            double sumCongBHXH = 0;
            double sumCongAnCa = 0;

            double sumLuongLePhepTT = 0;
            double sumLuongHocHopTT = 0;
            double sumLuongCheDoQL = 0;

            double sumHeSoQL = 0;
            double sumDiemTT = 0;

            int countQL = 0;
            int countTT = 0;

            foreach (var emp in ListNhanVien)
            {
                string xl = (emp.XepLoai ?? "A").Trim().ToUpperInvariant();
                double rateXL = xl == "A" ? 1.0 : (xl == "B" ? 0.9 : (xl == "C" ? 0.8 : 1.0));

                if (emp.LaQuanLy)
                {
                    countQL++;
                    if (emp.HeSoCoDinh > 0 && emp.CongSP > 0)
                    {
                        emp.DiemSanXuat = Math.Round(emp.CongSP * emp.HeSoCoDinh / 26.0, 2);
                        emp.HeSoHuong = Math.Round(emp.DiemSanXuat * rateXL, 2);
                    }
                    else if (emp.HeSoHuong <= 0 && emp.DiemSanXuat > 0)
                    {
                        emp.HeSoHuong = Math.Round(emp.DiemSanXuat * rateXL, 2);
                    }
                    sumHeSoQL += emp.HeSoHuong;
                }
                else
                {
                    countTT++;
                    if (emp.DiemSanXuat > 0)
                    {
                        emp.HeSoHuong = Math.Round(emp.DiemSanXuat * rateXL, 2);
                    }
                    sumDiemTT += emp.HeSoHuong;
                }

                // Tính tổng công từng người = Công SP + Công Lễ phép + Công VHTT
                emp.TongCong = emp.CongSP + emp.CongLePhep + emp.CongVHTT;

                // Tính tiền lương lễ phép chế độ theo công thức Excel: =ROUND(M/26*G, -2)
                if (!emp.QuanLyTinhMoi && emp.MucLuongBHXH > 0)
                {
                    emp.LuongLePhepCD = Math.Round((emp.CongLePhep / 26.0 * emp.MucLuongBHXH) / 100.0, 0) * 100.0;
                    if (emp.CongLamLe > 0)
                        emp.LuongLamLe = Math.Round((emp.CongLamLe * 2.0 / 26.0 * emp.MucLuongBHXH) / 1000.0, 0) * 1000.0;
                }

                // Tiền học họp, VHTT (theo NT: 250.000đ/công hoặc 200.000đ/công trên bảng lương)
                emp.LuongHocVHTT = emp.CongHocHop * 250000.0 + emp.CongVHTT * 300000.0;

                // Phụ cấp chức vụ / trách nhiệm / ATVSV
                if (!emp.QuanLyTinhMoi && emp.PhuCapChucVu > 0 && emp.PhuCapTien == 0)
                {
                    emp.PhuCapTien = Math.Round((emp.PhuCapChucVu * 1600000.0) / 1000.0, 0) * 1000.0;
                }

                // Cộng dồn công
                sumCongSP += emp.CongSP;
                if (emp.LaQuanLy)
                {
                    sumCongQL += emp.CongSP;
                    sumCongCDQL += emp.CongLePhep + emp.CongHocHop + emp.CongVHTT;
                    sumLuongCheDoQL += emp.LuongLePhepCD + emp.CongHocHop * 250000.0;
                }
                else
                {
                    sumCongLDTT += emp.CongSP;
                    sumCongLePhepTT += emp.CongLePhep;
                    sumCongHocHopTT += emp.CongHocHop;
                    sumCongVHTTTT += emp.CongVHTT;
                    sumLuongLePhepTT += emp.LuongLePhepCD;
                    sumLuongHocHopTT += emp.CongHocHop * 250000.0 + emp.CongVHTT * 300000.0;
                }

                sumCongBHXH += emp.CongBHXH;
                emp.AnCaCongThuc = emp.CongSP + emp.CongHocHop + emp.CongVHTT;
                if (!string.IsNullOrEmpty(emp.AnCaNguon))
                {
                    if (emp.AnCaNguon == "DV") emp.CongAnCa = emp.AnCaDonVi;
                    else if (emp.AnCaNguon == "TC") emp.CongAnCa = emp.AnCaThuCong;
                    else emp.CongAnCa = emp.AnCaCongThuc;
                    emp.TienAnCa = emp.CongAnCa * 30000.0;
                }
                else if (emp.CongAnCa <= 0 && emp.CongSP > 0 && MaPX != "KT" && !emp.CoAnCa)
                    emp.CongAnCa = emp.CongSP + emp.CongHocHop;
                sumCongAnCa += emp.CongAnCa;
            }

            // Phân bổ quỹ lương sản phẩm cho từng người nếu bật redistributeProductSalary
            double quyLuongSPTT = LuongSPTT + LuongCVKhac + TruLuong - ListNhanVienTotalPhuCap(false);
            if (redistributeProductSalary)
            {
                foreach (var emp in ListNhanVien)
                {
                    if (emp.LaQuanLy)
                    {
                        if (sumHeSoQL > 0 && LuongSPQL > 0)
                            emp.LuongSanPham = Math.Round((LuongSPQL * emp.HeSoHuong / sumHeSoQL) / 100.0, 0) * 100.0;
                    }
                    else
                    {
                        if (sumDiemTT > 0 && quyLuongSPTT > 0)
                            emp.LuongSanPham = Math.Round((quyLuongSPTT * emp.HeSoHuong / sumDiemTT) / 100.0, 0) * 100.0;
                    }
                    emp.TongTienLuong = emp.LuongSanPham + emp.LuongLePhepCD + emp.LuongLamLe + emp.LuongHocVHTT + emp.PhuCapTien;
                }
            }
            else
            {
                foreach (var emp in ListNhanVien)
                {
                    emp.TongTienLuong = emp.LuongSanPham + emp.LuongLePhepCD + emp.LuongLamLe + emp.LuongHocVHTT + emp.PhuCapTien;
                }
            }

            // Cập nhật các chỉ tiêu tổng hợp của Phân xưởng từ danh sách nhân viên
            if (sumCongSP > 0)
            {
                CongSP = sumCongSP;
                CongQL = sumCongQL;
                CongLDTT = sumCongLDTT;
                CongLePhepTT = sumCongLePhepTT;
                CongHocHopTT = sumCongHocHopTT;
                CongCDQL = sumCongCDQL;
                CongLePhepCD = sumCongLePhepTT + sumCongHocHopTT + sumCongVHTTTT + sumCongCDQL;
                CongBHXH = sumCongBHXH;
                TongCong = CongSP + CongLePhepCD + CongBHXH;

                LDThucTeBQ = Math.Round(CongSP / 26.0, 2);
                LDTTThucTeBQ = Math.Round(CongLDTT / 26.0, 2);

                // Số công ăn ca theo chuẩn NT = Công SP + Công học họp
                CongAnCa = sumCongSP + sumCongHocHopTT;
                ApplyFileBasedTotals();
                TienAnCa = CongAnCa * 30000.0;

                if (sumLuongLePhepTT > 0)
                {
                    LuongLePhepTT = sumLuongLePhepTT;
                    LuongHocHopTT = sumCongHocHopTT * 250000.0 + sumCongVHTTTT * 300000.0;
                    LuongCheDoTT = LuongLePhepTT + LuongHocHopTT;
                }
                if (sumLuongCheDoQL > 0)
                {
                    LuongCheDoQL = sumLuongCheDoQL;
                }
            }

            ManagementPayroll.Calculate(this, redistributeProductSalary);
            RecalculateDerived();
        }

        /// <summary>
        /// Nếu dữ liệu từng người được trích từ file công PX (QT ăn ca, Bảng BDĐH) thì ghi đè tổng ăn ca / BDĐH của PX
        /// bằng tổng cộng từ chi tiết từng người.
        /// </summary>
        private void ApplyFileBasedTotals()
        {
            double anCa = 0, m1 = 0, m2 = 0, m3 = 0;
            bool hasAnCa = false, hasBDDH = false;
            foreach (var e in ListNhanVien)
            {
                if (e.CoAnCa || !string.IsNullOrEmpty(e.AnCaNguon)) hasAnCa = true;
                if (e.CoBDDH) hasBDDH = true;
                anCa += e.CongAnCa; m1 += e.BDDH_M1; m2 += e.BDDH_M2; m3 += e.BDDH_M3;
            }
            if (hasAnCa) CongAnCa = anCa;
            if (hasBDDH)
            {
                BDDH_Muc1_Cong = m1;
                BDDH_Muc2_Cong = m2;
                BDDH_Muc3_Cong = m3;
            }
        }

        /// <summary>Đặt nguồn công ăn ca cho cả PX: "CT" công thức, "DV" đơn vị gửi, "TC" thủ công. onlyDiff = chỉ người đang chênh lệch.</summary>
        public int SetAnCaNguon(string nguon, bool onlyDiff)
        {
            int n = 0;
            foreach (var e in ListNhanVien)
            {
                if (nguon == "DV" && !e.CoAnCa) continue;
                e.AnCaCongThuc = e.CongSP + e.CongHocHop + e.CongVHTT;
                if (onlyDiff && Math.Abs(e.AnCaChenhLech) <= 0.001) continue;
                if (nguon == "TC" && e.AnCaThuCong <= 0) e.AnCaThuCong = e.CongAnCa;
                e.AnCaNguon = nguon;
                n++;
            }
            return n;
        }

        /// <summary>Tổng hợp theo TỔ sản xuất từ chi tiết từng người.</summary>
        public List<ToSummary> GetToSummaries()
        {
            var list = new List<ToSummary>();
            var map = new Dictionary<string, ToSummary>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in ListNhanVien)
            {
                string to = !string.IsNullOrEmpty(e.ToSanXuat) ? e.ToSanXuat : (e.BoPhan ?? "");
                ToSummary s;
                if (!map.TryGetValue(to, out s))
                {
                    s = new ToSummary { TenTo = to, LaQuanLy = e.LaQuanLy };
                    map[to] = s;
                    list.Add(s);
                }
                s.SoLD++;
                s.TongCong += e.TongCong;
                s.CongSP += e.CongSP;
                s.CongLePhep += e.CongLePhep;
                s.CongHocHop += e.CongHocHop;
                s.CongVHTT += e.CongVHTT;
                s.CongBHXH += e.CongBHXH;
                s.CongK3 += e.CongK3;
                s.CongAnCa += e.CongAnCa;
                s.AnCaNhaAn += e.AnCaNhaAn;
                s.AnCaTienMat += e.AnCaTienMat;
                s.TienAnCa += e.TienAnCa;
                s.BDDH_M1 += e.BDDH_M1;
                s.BDDH_M2 += e.BDDH_M2;
                s.BDDH_M3 += e.BDDH_M3;
                s.TienBDDH += e.TienBDDH;
                s.TongTienLuong += e.TongTienLuong;
            }
            return list;
        }

        private double ListNhanVienTotalPhuCap(bool laQuanLy)
        {
            double s = 0;
            if (ListNhanVien == null) return 0;
            foreach (var e in ListNhanVien)
            {
                if (e.LaQuanLy == laQuanLy) s += e.PhuCapTien + e.LuongLamLe;
            }
            return s;
        }

        public void SyncBaselineFromCurrentTotals()
        {
            if (KyQuy == null || KyLuyKe == null) return;
            TruocTrongQuy = new PeriodMetrics
            {
                TenKy = "Các tháng trước trong Quý",
                TongCong = Math.Max(0, KyQuy.TongCong - TongCong),
                CongSP = Math.Max(0, KyQuy.CongSP - CongSP),
                CongQL = Math.Max(0, KyQuy.CongQL - CongQL),
                CongLDTT = Math.Max(0, KyQuy.CongLDTT - CongLDTT),
                CongCheDo = Math.Max(0, KyQuy.CongCheDo - CongLePhepCD),
                CongCheDoTT = Math.Max(0, KyQuy.CongCheDoTT - CongLePhepTT - CongHocHopTT),
                CongCheDoQL = Math.Max(0, KyQuy.CongCheDoQL - CongCDQL),
                CongBHXH = Math.Max(0, KyQuy.CongBHXH - CongBHXH),
                LDDanhSachBQ = KyQuy.LDDanhSachBQ > 0 ? KyQuy.LDDanhSachBQ : LDCuoiKy,
                SanLuongTH = Math.Max(0, KyQuy.SanLuongTH - SanLuongTH),
                SanLuongTinhLuong = Math.Max(0, KyQuy.SanLuongTinhLuong - SanLuongTinhLuong),
                TongLuongToanPX = Math.Max(0, KyQuy.TongLuongToanPX - TongLuongToanPX),
                TongLuongTT = Math.Max(0, KyQuy.TongLuongTT - TongLuongTT),
                LuongSPTT = Math.Max(0, KyQuy.LuongSPTT - LuongSPTT),
                LuongCVKhac = Math.Max(0, KyQuy.LuongCVKhac - LuongCVKhac),
                CongCVKhac = Math.Max(0, KyQuy.CongCVKhac - CongCVKhac),
                LuongCheDoTT = Math.Max(0, KyQuy.LuongCheDoTT - LuongCheDoTT),
                LuongTDT = Math.Max(0, KyQuy.LuongTDT - LuongTDT),
                CongTDT = Math.Max(0, KyQuy.CongTDT - CongTDT),
                TruLuong = KyQuy.TruLuong - TruLuong,
                TongLuongQL = Math.Max(0, KyQuy.TongLuongQL - TongLuongQL),
                LuongSPQL = Math.Max(0, KyQuy.LuongSPQL - LuongSPQL),
                LuongCheDoQL = Math.Max(0, KyQuy.LuongCheDoQL - LuongCheDoQL),
                CongAnCa = Math.Max(0, KyQuy.CongAnCa - CongAnCa),
                TienAnCa = Math.Max(0, KyQuy.TienAnCa - TienAnCa),
                BDDH_Muc1_Cong = Math.Max(0, KyQuy.BDDH_Muc1_Cong - BDDH_Muc1_Cong),
                BDDH_Muc1_Tien = Math.Max(0, KyQuy.BDDH_Muc1_Tien - BDDH_Muc1_Tien),
                BDDH_Muc2_Cong = Math.Max(0, KyQuy.BDDH_Muc2_Cong - BDDH_Muc2_Cong),
                BDDH_Muc2_Tien = Math.Max(0, KyQuy.BDDH_Muc2_Tien - BDDH_Muc2_Tien),
                BDDH_Muc3_Cong = Math.Max(0, KyQuy.BDDH_Muc3_Cong - BDDH_Muc3_Cong),
                BDDH_Muc3_Tien = Math.Max(0, KyQuy.BDDH_Muc3_Tien - BDDH_Muc3_Tien),
                TongTienBDDH = Math.Max(0, KyQuy.TongTienBDDH - TongTienBDDH)
            };

            TruocTrongNam = new PeriodMetrics
            {
                TenKy = "Các Quý trước trong năm",
                TongCong = Math.Max(0, KyLuyKe.TongCong - KyQuy.TongCong),
                CongSP = Math.Max(0, KyLuyKe.CongSP - KyQuy.CongSP),
                CongQL = Math.Max(0, KyLuyKe.CongQL - KyQuy.CongQL),
                CongLDTT = Math.Max(0, KyLuyKe.CongLDTT - KyQuy.CongLDTT),
                CongCheDo = Math.Max(0, KyLuyKe.CongCheDo - KyQuy.CongCheDo),
                CongCheDoTT = Math.Max(0, KyLuyKe.CongCheDoTT - KyQuy.CongCheDoTT),
                CongCheDoQL = Math.Max(0, KyLuyKe.CongCheDoQL - KyQuy.CongCheDoQL),
                CongBHXH = Math.Max(0, KyLuyKe.CongBHXH - KyQuy.CongBHXH),
                LDDanhSachBQ = KyLuyKe.LDDanhSachBQ > 0 ? KyLuyKe.LDDanhSachBQ : LDCuoiKy,
                SanLuongTH = Math.Max(0, KyLuyKe.SanLuongTH - KyQuy.SanLuongTH),
                SanLuongTinhLuong = Math.Max(0, KyLuyKe.SanLuongTinhLuong - KyQuy.SanLuongTinhLuong),
                TongLuongToanPX = Math.Max(0, KyLuyKe.TongLuongToanPX - KyQuy.TongLuongToanPX),
                TongLuongTT = Math.Max(0, KyLuyKe.TongLuongTT - KyQuy.TongLuongTT),
                LuongSPTT = Math.Max(0, KyLuyKe.LuongSPTT - KyQuy.LuongSPTT),
                LuongCVKhac = Math.Max(0, KyLuyKe.LuongCVKhac - KyQuy.LuongCVKhac),
                CongCVKhac = Math.Max(0, KyLuyKe.CongCVKhac - KyQuy.CongCVKhac),
                LuongCheDoTT = Math.Max(0, KyLuyKe.LuongCheDoTT - KyQuy.LuongCheDoTT),
                LuongTDT = Math.Max(0, KyLuyKe.LuongTDT - KyQuy.LuongTDT),
                CongTDT = Math.Max(0, KyLuyKe.CongTDT - KyQuy.CongTDT),
                TruLuong = KyLuyKe.TruLuong - KyQuy.TruLuong,
                TongLuongQL = Math.Max(0, KyLuyKe.TongLuongQL - KyQuy.TongLuongQL),
                LuongSPQL = Math.Max(0, KyLuyKe.LuongSPQL - KyQuy.LuongSPQL),
                LuongCheDoQL = Math.Max(0, KyLuyKe.LuongCheDoQL - KyQuy.LuongCheDoQL),
                CongAnCa = Math.Max(0, KyLuyKe.CongAnCa - KyQuy.CongAnCa),
                TienAnCa = Math.Max(0, KyLuyKe.TienAnCa - KyQuy.TienAnCa),
                BDDH_Muc1_Cong = Math.Max(0, KyLuyKe.BDDH_Muc1_Cong - KyQuy.BDDH_Muc1_Cong),
                BDDH_Muc1_Tien = Math.Max(0, KyLuyKe.BDDH_Muc1_Tien - KyQuy.BDDH_Muc1_Tien),
                BDDH_Muc2_Cong = Math.Max(0, KyLuyKe.BDDH_Muc2_Cong - KyQuy.BDDH_Muc2_Cong),
                BDDH_Muc2_Tien = Math.Max(0, KyLuyKe.BDDH_Muc2_Tien - KyQuy.BDDH_Muc2_Tien),
                BDDH_Muc3_Cong = Math.Max(0, KyLuyKe.BDDH_Muc3_Cong - KyQuy.BDDH_Muc3_Cong),
                BDDH_Muc3_Tien = Math.Max(0, KyLuyKe.BDDH_Muc3_Tien - KyQuy.BDDH_Muc3_Tien),
                TongTienBDDH = Math.Max(0, KyLuyKe.TongTienBDDH - KyQuy.TongTienBDDH)
            };
        }

        public int GetTotalMonthsForLuyKe()
        {
            switch (LoaiKy)
            {
                case LoaiKyBaoCao.QuyI: return 3;
                case LoaiKyBaoCao.QuyII_6Thang: return 6;
                case LoaiKyBaoCao.QuyIII_9Thang: return 9;
                case LoaiKyBaoCao.QuyIV_CaNam: return 12;
                default: return Math.Max(1, Thang);
            }
        }

        public void RecalculateDerived()
        {
            LDCuoiKy = LDDauKy + LDTang - LDGiam;
            CongLDTT = CongSP - CongQL;
            TongCong = CongSP + CongLePhepCD + CongBHXH;

            if (LDCuoiKy > 0)
                NgayCongBQ = Math.Round(CongSP / LDCuoiKy, 2);
            if (LDQuanLy > 0)
                NgayCongBQQL = Math.Round(CongQL / LDQuanLy, 2);
            if (LDTrucTiep > 0)
                NgayCongBQLDTT = Math.Round(CongLDTT / LDTrucTiep, 2);

            double sumGNVCong = 0;
            double sumGNVTien = 0;
            if (ListGNV != null && ListGNV.Count > 0)
            {
                foreach (var item in ListGNV)
                {
                    if (item.KhoiLuong > 0 && item.DonGia > 0 && item.ThanhTien == 0)
                        item.ThanhTien = Math.Round(item.KhoiLuong * item.DonGia, 0);
                    sumGNVCong += item.KhoiLuong;
                    sumGNVTien += item.ThanhTien;
                }
                if (MaPX != "CDNL")
                {
                    CongCVKhac = sumGNVCong;
                    LuongCVKhac = sumGNVTien;
                }
            }

            if (SanLuongLoai1 > 0 && DonGiaLoai1 > 0)
                ThanhTienLoai1 = Math.Round(SanLuongLoai1 * DonGiaLoai1, 0);
            if (SanLuongLoai2 > 0 && DonGiaLoai2 > 0)
                ThanhTienLoai2 = Math.Round(SanLuongLoai2 * DonGiaLoai2, 0);
            if (SanLuongLe > 0 && DonGiaLe > 0)
                ThanhTienLe = Math.Round(SanLuongLe * DonGiaLe, 0);

            BDDH_Muc1_Tien = BDDH_Muc1_Cong * 13000;
            BDDH_Muc2_Tien = BDDH_Muc2_Cong * 20000;
            BDDH_Muc3_Tien = BDDH_Muc3_Cong * 26000;
            TongTienBDDH = BDDH_Muc1_Tien + BDDH_Muc2_Tien + BDDH_Muc3_Tien;

            if (CongAnCa > 0)
                TienAnCa = CongAnCa * 30000;

            LuongCheDoTT = LuongLePhepTT + LuongHocHopTT + LuongVHTTTT;
            TongLuongTT = LuongSPTT + LuongCVKhac + LuongCheDoTT + LuongTDT + TruLuong;
            TongLuongQL = LuongSPQL + LuongCheDoQL + DieuChinhLuongQL;
            TongLuongToanPX = TongLuongTT + TongLuongQL;

            if (LDThucTeBQ > 0)
            {
                TLBinhQuanToanPX = Math.Round((LuongSPTT + LuongCVKhac + LuongSPQL + TruLuong) / LDThucTeBQ / 1000.0, 2);
                ThuNhapBQToanPX = Math.Round(TongLuongToanPX / LDThucTeBQ / 1000.0, 2);
            }
            if (LDTTThucTeBQ > 0)
            {
                TLSPBinhQuanLDTT = Math.Round((LuongSPTT + LuongCVKhac + TruLuong) / LDTTThucTeBQ / 1000.0, 2);
                ThuNhapBQLDTT = Math.Round(TongLuongTT / LDTTThucTeBQ / 1000.0, 2);
            }
            if (CongLDTT > 0)
            {
                BQNgayCongLVTT = Math.Round((LuongSPTT + LuongCVKhac + TruLuong) / CongLDTT, 0);
            }

            // TỰ ĐỘNG CỘNG DỒN LŨY KẾ QUÝ & NĂM (6 THÁNG / 9 THÁNG / CẢ NĂM)
            if (TuDongCongLuyKe && TruocTrongQuy != null && TruocTrongNam != null)
            {
                RecalculateCumulative();
            }
        }

        public void RecalculateCumulative()
        {
            // 1. Cộng dồn Quý = Các tháng trước trong Quý + Tháng hiện tại
            KyQuy.CongQL = TruocTrongQuy.CongQL + CongQL;
            KyQuy.CongLDTT = TruocTrongQuy.CongLDTT + CongLDTT;
            KyQuy.CongSP = KyQuy.CongQL + KyQuy.CongLDTT;
            KyQuy.CongCheDo = TruocTrongQuy.CongCheDo + CongLePhepCD;
            KyQuy.CongCheDoTT = TruocTrongQuy.CongCheDoTT + CongLePhepTT + CongHocHopTT;
            KyQuy.CongCheDoQL = TruocTrongQuy.CongCheDoQL + CongCDQL;
            KyQuy.CongBHXH = TruocTrongQuy.CongBHXH + CongBHXH;
            KyQuy.CongCheDoBHXH = KyQuy.CongCheDo + KyQuy.CongBHXH;
            KyQuy.TongCong = KyQuy.CongSP + KyQuy.CongCheDoBHXH;

            if (KyQuy.LDDanhSachBQ <= 0) KyQuy.LDDanhSachBQ = LDCuoiKy;
            KyQuy.LDThucTeBQ = Math.Round(KyQuy.CongSP / 26.0 / 3.0, 2);
            KyQuy.LDTTThucTeBQ = Math.Round(KyQuy.CongLDTT / 26.0 / 3.0, 2);
            if (KyQuy.LDDanhSachBQ > 0)
                KyQuy.NgayCongBQ = Math.Round(KyQuy.CongSP / KyQuy.LDDanhSachBQ / 3.0, 2);
            if (LDQuanLy > 0)
                KyQuy.NgayCongBQQL = Math.Round(KyQuy.CongQL / LDQuanLy / 3.0, 2);
            if (LDTrucTiep > 0)
                KyQuy.NgayCongBQLDTT = Math.Round(KyQuy.CongLDTT / LDTrucTiep / 3.0, 2);

            KyQuy.SanLuongTH = Math.Round(TruocTrongQuy.SanLuongTH + SanLuongTH, 3);
            KyQuy.SanLuongTinhLuong = Math.Round(TruocTrongQuy.SanLuongTinhLuong + SanLuongTinhLuong, 3);

            KyQuy.LuongSPTT = TruocTrongQuy.LuongSPTT + LuongSPTT;
            KyQuy.LuongCVKhac = TruocTrongQuy.LuongCVKhac + LuongCVKhac;
            KyQuy.CongCVKhac = TruocTrongQuy.CongCVKhac + CongCVKhac;
            KyQuy.LuongCheDoTT = TruocTrongQuy.LuongCheDoTT + LuongCheDoTT;
            KyQuy.LuongTDT = TruocTrongQuy.LuongTDT + LuongTDT;
            KyQuy.CongTDT = TruocTrongQuy.CongTDT + CongTDT;
            KyQuy.TruLuong = TruocTrongQuy.TruLuong + TruLuong;
            KyQuy.TongLuongTT = KyQuy.LuongSPTT + KyQuy.LuongCVKhac + KyQuy.LuongCheDoTT + KyQuy.LuongTDT + KyQuy.TruLuong;

            KyQuy.LuongSPQL = TruocTrongQuy.LuongSPQL + LuongSPQL;
            KyQuy.LuongCheDoQL = TruocTrongQuy.LuongCheDoQL + LuongCheDoQL;
            KyQuy.TongLuongQL = KyQuy.LuongSPQL + KyQuy.LuongCheDoQL + KyQuy.DieuChinhLuongQL;
            KyQuy.TongLuongToanPX = KyQuy.TongLuongTT + KyQuy.TongLuongQL;

            if (KyQuy.LDThucTeBQ > 0)
            {
                KyQuy.TLBinhQuanToanPX = Math.Round((KyQuy.LuongSPTT + KyQuy.LuongCVKhac + KyQuy.LuongSPQL + KyQuy.TruLuong) / KyQuy.LDThucTeBQ / 3.0 / 1000.0, 2);
                KyQuy.ThuNhapBQToanPX = Math.Round(KyQuy.TongLuongToanPX / KyQuy.LDThucTeBQ / 3.0 / 1000.0, 2);
            }
            if (KyQuy.LDTTThucTeBQ > 0)
            {
                KyQuy.TLSPBinhQuanLDTT = Math.Round((KyQuy.LuongSPTT + KyQuy.LuongCVKhac + KyQuy.TruLuong) / KyQuy.LDTTThucTeBQ / 3.0 / 1000.0, 2);
                KyQuy.ThuNhapBQLDTT = Math.Round(KyQuy.TongLuongTT / KyQuy.LDTTThucTeBQ / 3.0 / 1000.0, 2);
            }
            if (KyQuy.CongLDTT > 0)
            {
                KyQuy.BQNgayCongLVTT = Math.Round((KyQuy.LuongSPTT + KyQuy.LuongCVKhac + KyQuy.TruLuong) / KyQuy.CongLDTT, 0);
            }

            KyQuy.CongAnCa = TruocTrongQuy.CongAnCa + CongAnCa;
            KyQuy.TienAnCa = KyQuy.CongAnCa * 30000;

            KyQuy.BDDH_Muc1_Cong = TruocTrongQuy.BDDH_Muc1_Cong + BDDH_Muc1_Cong;
            KyQuy.BDDH_Muc1_Tien = KyQuy.BDDH_Muc1_Cong * 13000;
            KyQuy.BDDH_Muc2_Cong = TruocTrongQuy.BDDH_Muc2_Cong + BDDH_Muc2_Cong;
            KyQuy.BDDH_Muc2_Tien = KyQuy.BDDH_Muc2_Cong * 20000;
            KyQuy.BDDH_Muc3_Cong = TruocTrongQuy.BDDH_Muc3_Cong + BDDH_Muc3_Cong;
            KyQuy.BDDH_Muc3_Tien = KyQuy.BDDH_Muc3_Cong * 26000;
            KyQuy.TongTienBDDH = KyQuy.BDDH_Muc1_Tien + KyQuy.BDDH_Muc2_Tien + KyQuy.BDDH_Muc3_Tien;

            // 2. Cộng dồn Lũy kế (6 tháng / 9 tháng / Cả năm) = Các Quý trước + Quý hiện tại
            int soThangLK = GetTotalMonthsForLuyKe();
            KyLuyKe.CongQL = TruocTrongNam.CongQL + KyQuy.CongQL;
            KyLuyKe.CongLDTT = TruocTrongNam.CongLDTT + KyQuy.CongLDTT;
            KyLuyKe.CongSP = KyLuyKe.CongQL + KyLuyKe.CongLDTT;
            KyLuyKe.CongCheDo = TruocTrongNam.CongCheDo + KyQuy.CongCheDo;
            KyLuyKe.CongCheDoTT = TruocTrongNam.CongCheDoTT + KyQuy.CongCheDoTT;
            KyLuyKe.CongCheDoQL = TruocTrongNam.CongCheDoQL + KyQuy.CongCheDoQL;
            KyLuyKe.CongBHXH = TruocTrongNam.CongBHXH + KyQuy.CongBHXH;
            KyLuyKe.CongCheDoBHXH = KyLuyKe.CongCheDo + KyLuyKe.CongBHXH;
            KyLuyKe.TongCong = KyLuyKe.CongSP + KyLuyKe.CongCheDoBHXH;

            if (KyLuyKe.LDDanhSachBQ <= 0) KyLuyKe.LDDanhSachBQ = LDCuoiKy;
            KyLuyKe.LDThucTeBQ = Math.Round(KyLuyKe.CongSP / 25.0 / soThangLK, 2);
            KyLuyKe.LDTTThucTeBQ = Math.Round(KyLuyKe.CongLDTT / 25.0 / soThangLK, 2);
            if (KyLuyKe.LDDanhSachBQ > 0)
                KyLuyKe.NgayCongBQ = Math.Round(KyLuyKe.CongSP / KyLuyKe.LDDanhSachBQ / soThangLK, 2);
            if (LDQuanLy > 0)
                KyLuyKe.NgayCongBQQL = Math.Round(KyLuyKe.CongQL / LDQuanLy / soThangLK, 2);
            if (LDTrucTiep > 0)
                KyLuyKe.NgayCongBQLDTT = Math.Round(KyLuyKe.CongLDTT / LDTrucTiep / soThangLK, 2);

            KyLuyKe.SanLuongTH = Math.Round(TruocTrongNam.SanLuongTH + KyQuy.SanLuongTH, 3);
            KyLuyKe.SanLuongTinhLuong = Math.Round(TruocTrongNam.SanLuongTinhLuong + KyQuy.SanLuongTinhLuong, 3);

            KyLuyKe.LuongSPTT = TruocTrongNam.LuongSPTT + KyQuy.LuongSPTT;
            KyLuyKe.LuongCVKhac = TruocTrongNam.LuongCVKhac + KyQuy.LuongCVKhac;
            KyLuyKe.CongCVKhac = TruocTrongNam.CongCVKhac + KyQuy.CongCVKhac;
            KyLuyKe.LuongCheDoTT = TruocTrongNam.LuongCheDoTT + KyQuy.LuongCheDoTT;
            KyLuyKe.LuongTDT = TruocTrongNam.LuongTDT + KyQuy.LuongTDT;
            KyLuyKe.CongTDT = TruocTrongNam.CongTDT + KyQuy.CongTDT;
            KyLuyKe.TruLuong = TruocTrongNam.TruLuong + KyQuy.TruLuong;
            KyLuyKe.TongLuongTT = KyLuyKe.LuongSPTT + KyLuyKe.LuongCVKhac + KyLuyKe.LuongCheDoTT + KyLuyKe.LuongTDT + KyLuyKe.TruLuong;

            KyLuyKe.LuongSPQL = TruocTrongNam.LuongSPQL + KyQuy.LuongSPQL;
            KyLuyKe.LuongCheDoQL = TruocTrongNam.LuongCheDoQL + KyQuy.LuongCheDoQL;
            KyLuyKe.TongLuongQL = KyLuyKe.LuongSPQL + KyLuyKe.LuongCheDoQL + KyLuyKe.DieuChinhLuongQL;
            KyLuyKe.TongLuongToanPX = KyLuyKe.TongLuongTT + KyLuyKe.TongLuongQL;

            if (KyLuyKe.LDThucTeBQ > 0)
            {
                KyLuyKe.TLBinhQuanToanPX = Math.Round((KyLuyKe.LuongSPTT + KyLuyKe.LuongCVKhac + KyLuyKe.LuongSPQL + KyLuyKe.LuongTDT + KyLuyKe.TruLuong) / KyLuyKe.LDThucTeBQ / soThangLK / 1000.0, 2);
                KyLuyKe.ThuNhapBQToanPX = Math.Round(KyLuyKe.TongLuongToanPX / KyLuyKe.LDThucTeBQ / soThangLK / 1000.0, 2);
            }
            if (KyLuyKe.LDTTThucTeBQ > 0)
            {
                KyLuyKe.TLSPBinhQuanLDTT = Math.Round((KyLuyKe.LuongSPTT + KyLuyKe.LuongCVKhac + KyLuyKe.LuongTDT + KyLuyKe.TruLuong) / KyLuyKe.LDTTThucTeBQ / soThangLK / 1000.0, 2);
                KyLuyKe.ThuNhapBQLDTT = Math.Round(KyLuyKe.TongLuongTT / KyLuyKe.LDTTThucTeBQ / soThangLK / 1000.0, 2);
            }
            if (KyLuyKe.CongLDTT > 0)
            {
                KyLuyKe.BQNgayCongLVTT = Math.Round((KyLuyKe.LuongSPTT + KyLuyKe.LuongCVKhac + KyLuyKe.TruLuong) / KyLuyKe.CongLDTT, 0);
            }

            KyLuyKe.CongAnCa = TruocTrongNam.CongAnCa + KyQuy.CongAnCa;
            KyLuyKe.TienAnCa = KyLuyKe.CongAnCa * 30000;

            KyLuyKe.BDDH_Muc1_Cong = TruocTrongNam.BDDH_Muc1_Cong + KyQuy.BDDH_Muc1_Cong;
            KyLuyKe.BDDH_Muc1_Tien = KyLuyKe.BDDH_Muc1_Cong * 13000;
            KyLuyKe.BDDH_Muc2_Cong = TruocTrongNam.BDDH_Muc2_Cong + KyQuy.BDDH_Muc2_Cong;
            KyLuyKe.BDDH_Muc2_Tien = KyLuyKe.BDDH_Muc2_Cong * 20000;
            KyLuyKe.BDDH_Muc3_Cong = TruocTrongNam.BDDH_Muc3_Cong + KyQuy.BDDH_Muc3_Cong;
            KyLuyKe.BDDH_Muc3_Tien = KyLuyKe.BDDH_Muc3_Cong * 26000;
            KyLuyKe.TongTienBDDH = KyLuyKe.BDDH_Muc1_Tien + KyLuyKe.BDDH_Muc2_Tien + KyLuyKe.BDDH_Muc3_Tien;
        }

        public static List<WorkshopData> CreateDefaultWorkshops()
        {
            // Chỉ có ánh xạ cấu trúc bảng. Không nhúng hồ sơ, tên người hoặc số lương mẫu.
            var list = new List<WorkshopData>();
            list.Add(new WorkshopData {
                MaPX = "TK", TenPX = "Phân xưởng Thiêu kết", SheetNT = "NT thiêu kết",
                Sheet9Thang = "9 tháng Thiêu kết", SheetBangCong = "T kết",
                RowTTBDDH_T9 = 180, TotalRowBangCong = 119,
                CanhBaoQuyQuanLy = "Chưa nạp dữ liệu. Chọn Excel và kiểm tra quỹ lương được duyệt."
            });
            list.Add(new WorkshopData {
                MaPX = "LG", TenPX = "Phân xưởng Luyện gang", SheetNT = "NT gang",
                Sheet9Thang = "9 tháng Luyện gang", SheetBangCong = "Gang",
                RowTTBDDH_T9 = 181, TotalRowBangCong = 167,
                CanhBaoQuyQuanLy = "Chưa nạp dữ liệu. Chọn Excel và kiểm tra quỹ lương được duyệt."
            });
            list.Add(new WorkshopData {
                MaPX = "LT", TenPX = "Phân xưởng Luyện thép", SheetNT = "NT Thép",
                Sheet9Thang = "9 tháng Luyện thép", SheetBangCong = "thép",
                RowTTBDDH_T9 = 182, TotalRowBangCong = 208,
                CanhBaoQuyQuanLy = "Chưa nạp dữ liệu. Chọn Excel và kiểm tra quỹ lương được duyệt."
            });
            list.Add(new WorkshopData {
                MaPX = "CDNL", TenPX = "Phân xưởng Cơ điện Năng lượng", SheetNT = "NT CĐNL",
                Sheet9Thang = "9 tháng CĐNL", SheetBangCong = "CĐNL",
                RowTTBDDH_T9 = 183, TotalRowBangCong = 138,
                CanhBaoQuyQuanLy = "Chưa nạp dữ liệu. Chọn Excel và kiểm tra quỹ lương được duyệt."
            });
            list.Add(new WorkshopData {
                MaPX = "KT", TenPX = "Phân xưởng Khai tuyển", SheetNT = "NT. K. tuyển ",
                Sheet9Thang = "9 tháng K.tuyển", SheetBangCong = "K. tuyển",
                RowTTBDDH_T9 = 179, TotalRowBangCong = 43,
                CanhBaoQuyQuanLy = "Chưa nạp dữ liệu. Chọn Excel và kiểm tra quỹ lương được duyệt."
            });
            return list;
        }
    }
}
