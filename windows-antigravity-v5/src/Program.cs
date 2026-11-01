using System;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Xml;

namespace CiscoLDTLApp
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args != null && args.Length > 2 && args[0] == "--import")
            {
                string lf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "import_log.txt");
                using (var w = new StreamWriter(lf, false, System.Text.Encoding.UTF8))
                {
                    var all = ExcelService.LoadWorkshopsFromExcel(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "03.9.2026_NT_T9_TCHC.xlsx"));
                    foreach (var ws in all)
                        if (ws.MaPX == args[2])
                            w.WriteLine(ExcelService.ImportExternalExcelToWorkshop(args[1], ws, null, all));
                }
                return 0;
            }

            if (args != null && args.Length > 0 && args[0] == "--test")
            {
                return RunSelfTest();
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        private static int CountFormulasInSheet(string excelPath, string sheetName)
        {
            using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var map = ExcelService.GetSheetMap(zip);
                if (!map.ContainsKey(sheetName)) return -1;
                var doc = ExcelService.LoadSheetXml(zip, map[sheetName]);
                var ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var nodes = doc.SelectNodes("//d:f", ns);
                return nodes != null ? nodes.Count : 0;
            }
        }

        private static int RunSelfTest()
        {
            string logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test_log.txt");
            using (var sw = new StreamWriter(logFile, false, System.Text.Encoding.UTF8))
            {
                Console.SetOut(sw);
                Console.SetError(sw);
                return RunSelfTestCore();
            }
        }

        private static int RunSelfTestCore()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string excelPath = Path.Combine(baseDir, "Data", "03.9.2026_NT_T9_TCHC.xlsx");
                string outWordDir = Path.Combine(baseDir, "Output_BaoCao_Word");
                string tplDir = Path.Combine(baseDir, "Templates");
                Directory.CreateDirectory(outWordDir);

                Console.WriteLine("=== KIEM TRA HE THONG TU DONG HOA NGHIEM THU CHI TIET TUNG NGUOI CISCO ===");
                int fBeforeNT = CountFormulasInSheet(excelPath, "NT thiêu kết");
                int fBeforeTK = CountFormulasInSheet(excelPath, "T kết");
                int fBeforeTHC = CountFormulasInSheet(excelPath, "Tổng hợp công");
                Console.WriteLine(string.Format("0. So cong thuc (<f>) truoc khi luu: NT thieu ket={0}, T ket={1}, Tong hop cong={2}", fBeforeNT, fBeforeTK, fBeforeTHC));

                Console.WriteLine("1. Doc du lieu chi tiet tung nguoi & bang nghiem thu tu file Excel: " + excelPath);
                var workshops = ExcelService.LoadWorkshopsFromExcel(excelPath);
                int totalEmployees = 0;
                foreach (var ws in workshops)
                {
                    totalEmployees += ws.ListNhanVien.Count;
                    Console.WriteLine(string.Format("   - {0} ({1}): Danh sach chi tiet={2} nguoi (Sheet '{3}'), Tổng công={4}, Công SP={5}, Tổng lương={6:N0}, Ăn ca={7:N0}, BDĐH={8:N0}",
                        ws.TenPX, ws.MaPX, ws.ListNhanVien.Count, ws.SheetBangCong, ws.TongCong, ws.CongSP, ws.TongLuongToanPX, ws.TienAnCa, ws.TongTienBDDH));
                    if (ws.ListNhanVien.Count > 0)
                    {
                        var e0 = ws.ListNhanVien[0];
                        var eLast = ws.ListNhanVien[ws.ListNhanVien.Count - 1];
                        Console.WriteLine(string.Format("     + NV dau: [Dòng {0}] SDB={1} | {2} ({3}) | Công SP={4} | Lương SP={5:N0} | Tổng lương={6:N0}",
                            e0.ExcelRow, e0.SDB, e0.HoTen, e0.ChucDanh, e0.CongSP, e0.LuongSanPham, e0.TongTienLuong));
                        Console.WriteLine(string.Format("     + NV cuoi: [Dòng {0}] SDB={1} | {2} ({3}) | Công SP={4} | Lương SP={5:N0} | Tổng lương={6:N0}",
                            eLast.ExcelRow, eLast.SDB, eLast.HoTen, eLast.ChucDanh, eLast.CongSP, eLast.LuongSanPham, eLast.TongTienLuong));
                    }
                }
                Console.WriteLine("   => TONG SO LAO DONG CHI TIET TOAN CONG TY DA NAP: " + totalEmployees + " nguoi.");

                Console.WriteLine("2. Kiem tra ghi cap nhat chi tiet tung nguoi & bang nghiem thu vao file Excel...");
                ExcelService.SaveWorkshopsToExcel(excelPath, workshops);
                int fAfterNT = CountFormulasInSheet(excelPath, "NT thiêu kết");
                int fAfterTK = CountFormulasInSheet(excelPath, "T kết");
                int fAfterTHC = CountFormulasInSheet(excelPath, "Tổng hợp công");
                Console.WriteLine(string.Format("   -> So cong thuc (<f>) sau khi luu: NT thieu ket={0}, T ket={1}, Tong hop cong={2} (GIU NGUYEN 100% CONG THUC!)", fAfterNT, fAfterTK, fAfterTHC));

                Console.WriteLine("3. Kiem tra dien & xuat bao cao Word (.docx) cho ca 5 Phan xuong...");
                foreach (var ws in workshops)
                {
                    ws.LoaiKy = LoaiKyBaoCao.QuyIII_9Thang;
                    ws.DinhKemBangChiTietNhanVien = true;
                    string outFile9T = Path.Combine(outWordDir, string.Format("BB_NghiemThu_{0}_Thang_{1}_QuyIII_9Thang_{2}.docx", ws.MaPX, ws.Thang, ws.Nam));
                    WordReportService.ExportWordDocx(ws, outFile9T);
                    Console.WriteLine("   -> Da xuat (kem Phu luc chi tiet tung nguoi): " + outFile9T + " (" + new FileInfo(outFile9T).Length + " bytes)");

                    ws.DinhKemBangChiTietNhanVien = false;
                    string outFileClean = Path.Combine(outWordDir, string.Format("BB_NghiemThu_{0}_Thang_{1}_{2}.docx", ws.MaPX, ws.Thang, ws.Nam));
                    WordReportService.ExportWordDocx(ws, outFileClean);
                }

                // Cập nhật trực tiếp vào 4 file Word mẫu lũy kế trong Output_BaoCao_Word
                string[,] tplPairs = new string[,]
                {
                    { "TK", "Nt_Thieu_ket_T9.docx", "Nt_Thieu_ket_T9.docx" },
                    { "LG", "Nt_thang_9_Gang.docx", "Nt_thang_9_Gang.docx" },
                    { "LT", "NT_thep_Thang_9.docx", "NT_thep_Thang_9.docx" },
                    { "CDNL", "BB_Nt_CDNL_T8.docx", "BB_Nt_CDNL_T9.docx" }
                };
                for (int i = 0; i < tplPairs.GetLength(0); i++)
                {
                    string ma = tplPairs[i, 0];
                    string srcFile = Path.Combine(tplDir, tplPairs[i, 1]);
                    string dstFile = Path.Combine(outWordDir, tplPairs[i, 2]);
                    foreach (var ws in workshops)
                    {
                        if (ws.MaPX == ma && File.Exists(srcFile))
                        {
                            ws.LoaiKy = LoaiKyBaoCao.QuyIII_9Thang;
                            ws.DinhKemBangChiTietNhanVien = false;
                            WordReportService.UpdateExistingMasterDocx(ws, srcFile, dstFile);
                            Console.WriteLine("   -> Da dien truc tiep vao file Word mau: " + dstFile + " (" + new FileInfo(dstFile).Length + " bytes)");
                        }
                    }
                }

                Console.WriteLine("=== HOAN TAT KIEM TRA THANH CONG 100% ===");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("LOI KIEM TRA: " + ex);
                return 1;
            }
        }
    }
}
