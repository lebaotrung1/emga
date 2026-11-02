using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace CiscoLDTLApp
{
    /// <summary>
    /// Trích xuất công, ăn ca (ký hiệu na / tt), BDĐH (M1 / M2 / M3), điểm SX của TỪNG NGƯỜI
    /// từ file công phân xưởng. Quy tắc khớp CHUẨN: phải trùng đồng thời HỌ TÊN + SỐ DANH BẠ (SDB).
    /// Sau khi khớp, ws.RecalculateFromEmployees() cộng lên Tổ -> Phân xưởng -> Bảng nghiệm thu.
    /// </summary>
    public static class TimesheetImporter
    {
        private const double GiaAnCa = 30000;
        private const double GiaM1 = 13000, GiaM2 = 20000, GiaM3 = 26000;

        private enum SheetKind { None, Cong, AnCa, BDDH, Diem }

        private class Layout
        {
            public SheetKind Kind;
            public int HeaderRow, NameCol, SdbCol, DayStart, DayEnd;
            public int CSP, CTG, CVH, CPL, CBH, CK3, CDiemSX, CDiemHuong, CXepLoai;
        }

        public static string ColName(int idx)
        {
            string s = "";
            while (idx > 0) { int m = (idx - 1) % 26; s = (char)('A' + m) + s; idx = (idx - 1) / 26; }
            return s;
        }

        private static string NormSdb(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            string c = s.Trim();
            int n;
            if (int.TryParse(c, out n)) return n.ToString("000");
            return c.ToUpperInvariant();
        }

        private static string NormName(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return Regex.Replace(s.Normalize(NormalizationForm.FormC).Trim(), @"\s+", " ").ToLowerInvariant();
        }

        private static string BirthYear(string name)
        { return Regex.Match(NormName(name),@"\(\s*((?:19|20)\d{2})\s*\)\s*$").Groups[1].Value; }
        private static string BaseName(string name)
        { return Regex.Replace(NormName(name),@"\s*\(\s*(?:19|20)\d{2}\s*\)\s*$", "").Trim(); }
        private static string Key(string sdb,string name) {return NormSdb(sdb)+"|"+BaseName(name);}
        private static string ExactKey(string sdb,string name) {return Key(sdb,name)+"|"+BirthYear(name);}
        private sealed class ImportSheet {internal string Key;internal XmlDocument Document;internal Layout Layout;}
        private sealed class EmployeeIndex
        {
            readonly Dictionary<string,List<EmployeeRecord>> identities=new Dictionary<string,List<EmployeeRecord>>();
            readonly Dictionary<string,List<EmployeeRecord>> names=new Dictionary<string,List<EmployeeRecord>>();
            readonly Dictionary<string,List<EmployeeRecord>> codes=new Dictionary<string,List<EmployeeRecord>>();
            static void AddTo(Dictionary<string,List<EmployeeRecord>> index,string key,EmployeeRecord e)
            {List<EmployeeRecord> rows;if(!index.TryGetValue(key,out rows)) {rows=new List<EmployeeRecord>();index[key]=rows;}if(!rows.Contains(e))rows.Add(e);}
            internal void Add(EmployeeRecord e) {AddTo(identities,TimesheetImporter.Key(e.SDB,e.HoTen),e);AddTo(names,BaseName(e.HoTen),e);AddTo(codes,NormSdb(e.SDB),e);}
            internal EmployeeRecord Find(string code,string name,Dictionary<string,HashSet<string>> sourceYears,out string problem)
            {
                problem=null;List<EmployeeRecord> rows;
                if(!identities.TryGetValue(TimesheetImporter.Key(code,name),out rows)) return null;
                string year=BirthYear(name);HashSet<string> years;bool sourceAmbiguous=sourceYears.TryGetValue(TimesheetImporter.Key(code,name),out years) && years.Count>1;
                var candidates=new List<EmployeeRecord>();
                if(year!="") foreach(var e in rows) if(BirthYear(e.HoTen)==year)candidates.Add(e);
                if(candidates.Count==0 && !(sourceAmbiguous && year=="")) foreach(var e in rows)
                    if(year=="" || BirthYear(e.HoTen)=="")candidates.Add(e);
                if(sourceAmbiguous && (year=="" || candidates.Count==1 && BirthYear(candidates[0].HoTen)==""))
                    problem="file có nhiều năm sinh cho cùng Tên + SDB; hồ sơ chưa đủ thông tin để phân biệt";
                else if(candidates.Count==1) return candidates[0];
                else if(candidates.Count>1) problem="nhiều hồ sơ cùng Tên + SDB; cần phân biệt năm sinh/hồ sơ trước khi nạp";
                else problem="năm sinh trong file và hồ sơ không khớp hoặc chưa đủ để phân biệt";
                return null;
            }
            internal string Conflict(string code,string name)
            {
                List<EmployeeRecord> rows;
                if(names.TryGetValue(BaseName(name),out rows)) {
                    var values=new HashSet<string>();foreach(var e in rows)values.Add(e.SDB);
                    return "tên có trong danh sách nhưng SDB khác: file "+code+", danh sách "+string.Join(", ",new List<string>(values).ToArray())+"; kiểm tra mã, không tự đổi SDB";
                }
                if(codes.TryGetValue(NormSdb(code),out rows)) {
                    var values=new HashSet<string>();foreach(var e in rows)values.Add(e.HoTen);
                    return "SDB "+code+" đang thuộc tên khác trong danh sách: "+string.Join(", ",new List<string>(values).ToArray())+"; kiểm tra họ tên và mã";
                }
                return null;
            }
        }

        private static string Cell(XmlDocument d, List<string> ss, int col, int row)
        {
            return ExcelService.GetCellString(d, ss, ColName(col) + row).Trim();
        }

        private static double Num(XmlDocument d, List<string> ss, int col, int row)
        {
            if (col <= 0) return 0;
            return ExcelService.GetCellDouble(d, ss, ColName(col) + row, 0);
        }

        private static string HeaderText(XmlDocument d, List<string> ss, int col, int hRow)
        {
            var sb = new StringBuilder();
            for (int r = Math.Max(1, hRow - 1); r <= hRow + 2; r++)
                sb.Append(Cell(d, ss, col, r)).Append(' ');
            return Regex.Replace(sb.ToString().ToLowerInvariant(), @"\s+", " ");
        }

        private static bool ExactHeader(XmlDocument doc,List<string> strings,int col,int row,string text)
        {for(int r=Math.Max(1,row-1);r<=row+2;r++)if(NormName(Cell(doc,strings,col,r))==text)return true;return false;}
        private static int ColumnNumber(string text)
        {int result=0;foreach(char c in text)result=result*26+c-'A'+1;return result;}
        private static int MergedDayEnd(XmlDocument doc,Layout layout)
        {
            foreach(XmlElement merge in doc.GetElementsByTagName("mergeCell","http://schemas.openxmlformats.org/spreadsheetml/2006/main")) {
                var m=Regex.Match(merge.GetAttribute("ref"),@"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");if(!m.Success)continue;
                int a=ColumnNumber(m.Groups[1].Value),b=ColumnNumber(m.Groups[3].Value),r=int.Parse(m.Groups[2].Value);
                if(a==layout.DayStart && b-a>=27 && b-a<=30 && r>=layout.HeaderRow && r<=layout.HeaderRow+2)return b;
            }
            return 0;
        }
        private static void ValidateLayout(Layout layout,string sheet)
        {
            if(layout.Kind==SheetKind.Cong && layout.CSP<=0)
                throw new ArgumentException("Sheet '"+sheet+"': chưa xác định được cột công sản phẩm/làm việc. Chưa nạp để giữ nguyên công đang có; cần kiểm tra tiêu đề cột.");
            if(layout.Kind==SheetKind.Diem && layout.CDiemSX<=0)
                throw new ArgumentException("Sheet '"+sheet+"': chưa xác định được cột tổng điểm. Chưa nạp để tránh ghi dữ liệu sai.");
        }

        private static Layout Detect(XmlDocument d, List<string> ss)
        {
            var sbTitle = new StringBuilder();
            for (int r = 1; r <= 8; r++)
                for (int c = 1; c <= 60; c++)
                    sbTitle.Append(Cell(d, ss, c, r)).Append(' ');
            string t = Regex.Replace(sbTitle.ToString().ToLowerInvariant(), @"\s+", " ");

            var lay = new Layout();
            bool hasDays = t.Contains("ngày trong tháng");
            if (t.Contains("bảng chấm điểm")) lay.Kind = SheetKind.Diem;
            else if (hasDays && t.Contains("bảng chấm công")) lay.Kind = SheetKind.Cong;
            else if (hasDays && t.Contains("quyết toán ăn ca")) lay.Kind = SheetKind.AnCa;
            else if (hasDays && t.Contains("quyết toán bồi dưỡng độc hại")) lay.Kind = SheetKind.BDDH;
            else return null;

            // Tìm dòng tiêu đề có "Họ và tên"
            for (int r = 1; r <= 12 && lay.HeaderRow == 0; r++)
                for (int c = 1; c <= 8; c++)
                    if (Cell(d, ss, c, r).ToLowerInvariant().Replace("  ", " ") == "họ và tên")
                    { lay.HeaderRow = r; lay.NameCol = c; break; }
            if (lay.HeaderRow == 0) return null;

            lay.SdbCol = 0;
            for (int c = 1; c <= 8; c++)
            {
                string h = Cell(d, ss, c, lay.HeaderRow).ToLowerInvariant();
                if (h == "sbd" || h == "sdb" || h.Contains("số danh bạ")) { lay.SdbCol = c; break; }
            }
            if (lay.SdbCol == 0) lay.SdbCol = lay.NameCol + 1;
            lay.DayStart = lay.SdbCol + 2;

            int firstSummary = 200, genericWork=0, genericPoints=0;
            for (int c = lay.DayStart; c <= 90; c++)
            {
                string h = HeaderText(d, ss, c, lay.HeaderRow);
                bool hit = false;
                switch (lay.Kind)
                {
                    case SheetKind.Cong:
                        if (h.Contains("số công sản phẩm") && lay.CSP == 0) { lay.CSP = c; hit = true; }
                        if (h.Contains("số công hưởng lương thời gian") && lay.CTG == 0) { lay.CTG = c; hit = true; }
                        if (h.Contains("công văn hóa thể thao") && lay.CVH == 0) { lay.CVH = c; hit = true; }
                        if (h.Contains("hưởng 100% lương") && lay.CPL == 0) { lay.CPL = c; hit = true; }
                        if (h.Contains("số công hưởng bhxh") && lay.CBH == 0) { lay.CBH = c; hit = true; }
                        if (h.Contains("công k3") && lay.CK3 == 0) { lay.CK3 = c; hit = true; }
                        if(ExactHeader(d,ss,c,lay.HeaderRow,"tổng công") && genericWork==0)genericWork=c;
                        if (h.Contains("tổng công") && !hit) hit = true;
                        break;
                    case SheetKind.AnCa:
                        if (h.Contains("tổng công được thanh toán ăn ca")) hit = true;
                        break;
                    case SheetKind.BDDH:
                        if (h.Contains("tổng công hưởng bdđh") || h.Contains("tổng công bdđh")) hit = true;
                        break;
                    case SheetKind.Diem:
                        if (h.Contains("tổng điểm sản xuất") && lay.CDiemSX == 0) { lay.CDiemSX = c; hit = true; }
                        if ((h.Contains("tổng điểm sau xếp loại") || h.Contains("tổng điểm chia lương")) && lay.CDiemHuong == 0) { lay.CDiemHuong = c; hit = true; }
                        if (h.Contains("xếp loại") && !h.Contains("sau xếp loại") && lay.CXepLoai == 0) { lay.CXepLoai = c; hit = true; }
                        if(ExactHeader(d,ss,c,lay.HeaderRow,"tổng điểm") && genericPoints==0)genericPoints=c;
                        if (h.Contains("tổng điểm") && !hit) hit = true;
                        break;
                }
                if (hit && c < firstSummary) firstSummary = c;
            }
            if(lay.CSP==0 && genericWork>0) {
                int firstComponent=200;foreach(int col in new[] {lay.CTG,lay.CVH,lay.CPL,lay.CBH})if(col>0)firstComponent=Math.Min(firstComponent,col);
                if(HeaderText(d,ss,genericWork,lay.HeaderRow).Contains("quy ra công") || firstComponent<200 && genericWork<firstComponent)lay.CSP=genericWork;
            }
            if(lay.CDiemSX==0)lay.CDiemSX=genericPoints;
            int dayEnd=MergedDayEnd(d,lay);
            lay.DayEnd=dayEnd>0?Math.Min(dayEnd,firstSummary-1):(firstSummary==200?40:firstSummary-1);
            return lay;
        }

        private static bool IsSkipName(string n)
        {
            string x = NormName(n);
            return x == "" || x == "nam" || x == "nữ" || x == "tổng" || x == "cộng" || x == "tổng cộng";
        }

        public static string Import(string filePath, WorkshopData ws, List<WorkshopData> workshops = null, Func<EmployeeRecord, EmployeeRecord> requestEmployee = null)
        {
            var log = new StringBuilder();

            var local=new EmployeeIndex();foreach(var e in ws.ListNhanVien)local.Add(e);
            var declined=new HashSet<string>();var external=new EmployeeIndex();
            if(workshops!=null) foreach(var unit in workshops) if(unit!=ws) foreach(var e in unit.ListNhanVien)external.Add(e);
            var touchedCong = new HashSet<EmployeeRecord>();
            var matchedCong = new HashSet<EmployeeRecord>();
            var matchedAnCa = new HashSet<EmployeeRecord>();
            var matchedBDDH = new HashSet<EmployeeRecord>();
            var warnings = new List<string>();
            int sheetsUsed = 0;

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var map = ExcelService.GetSheetMap(zip);
                var ss = ExcelService.LoadSharedStrings(zip);
                var done = new HashSet<string>();

                var sheets=new List<ImportSheet>();
                var sourceYears=new Dictionary<string,HashSet<string>>();
                foreach(var kv in map) {
                    if(!done.Add(kv.Value)) continue;
                    var doc=ExcelService.LoadSheetXml(zip,kv.Value);if(doc==null)continue;
                    var lay=Detect(doc,ss);if(lay==null)continue;
                    ValidateLayout(lay,kv.Key);
                    if(lay.Kind==SheetKind.Cong && kv.Key.ToLowerInvariant().Contains("sàng")) {
                        log.AppendLine(string.Format("- Sheet '{0}': bảng công chi tiết theo công đoạn phụ (đã nằm trong 'Bảng công' chính) -> không cộng thêm để tránh trùng.",kv.Key));continue;
                    }
                    sheets.Add(new ImportSheet {Key=kv.Key,Document=doc,Layout=lay});
                    foreach(XmlElement row in doc.GetElementsByTagName("row","http://schemas.openxmlformats.org/spreadsheetml/2006/main")) {
                        int r;if(!int.TryParse(row.GetAttribute("r"),out r)||r<=lay.HeaderRow)continue;
                        string name=Cell(doc,ss,lay.NameCol,r),code=Cell(doc,ss,lay.SdbCol,r),year=BirthYear(name);
                        if(code=="" || year=="") continue;
                        string key=Key(code,name);HashSet<string> years;if(!sourceYears.TryGetValue(key,out years)){years=new HashSet<string>();sourceYears[key]=years;}years.Add(year);
                    }
                }
                foreach(var kv in sheets) {
                    var doc=kv.Document;var lay=kv.Layout;
                    sheetsUsed++;
                    int rows = 0, ok = 0;

                    string currentTeam = "";
                    int lastRow = lay.HeaderRow;
                    foreach (XmlElement row in doc.GetElementsByTagName("row", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                    {
                        int nr;
                        if (int.TryParse(row.GetAttribute("r"), out nr)) lastRow = Math.Max(lastRow, nr);
                    }
                    for (int r = lay.HeaderRow + 1; r <= lastRow; r++)
                    {
                        string name = Cell(doc, ss, lay.NameCol, r);
                        string sdb = Cell(doc, ss, lay.SdbCol, r);
                        if (string.IsNullOrEmpty(sdb))
                        {
                            for (int c = 1; c <= lay.SdbCol; c++)
                            {
                                string heading = Cell(doc, ss, c, r);
                                string n = NormName(heading);
                                if (!n.Contains("tổng") && !n.StartsWith("cộng") && (n.Contains("tổ ") || n.Contains("bộ phận") || n.Contains("quản lý")))
                                { currentTeam = heading; break; }
                            }
                            continue;
                        }
                        if (IsSkipName(name)) continue;
                        if(lay.Kind==SheetKind.BDDH && Regex.IsMatch(sdb.Trim(),@"^M[123]$",RegexOptions.IgnoreCase) && Regex.IsMatch(NormName(name),@"^[-+•]?\s*mức\s*[123]\s*:?$"))continue;
                        double dummy;
                        if (double.TryParse(name, out dummy)) continue;
                        rows++;
                        if(lay.Kind==SheetKind.Cong) {
                            double primary=ExcelService.GetCellDouble(doc,ss,ColName(lay.CSP)+r,double.NaN);
                            if(double.IsNaN(primary) || double.IsInfinity(primary)) {
                                warnings.Add(string.Format("[{0}] dòng {1}: {2} (SDB {3}), ô {4} không có số công hợp lệ -> giữ công đang có, chưa nạp dòng này.",kv.Key,r,name,sdb,ColName(lay.CSP)+r));continue;
                            }
                        }

                        string problem;EmployeeRecord emp=local.Find(sdb,name,sourceYears,out problem);
                        if(problem!=null) {warnings.Add(string.Format("[{0}] dòng {1}: {2} (SDB {3}): {4} -> chưa nạp.",kv.Key,r,name,sdb,problem));continue;}
                        if(emp==null)
                        {
                            string identity=ExactKey(sdb,name);
                            if(declined.Contains(identity))continue;
                            problem=local.Conflict(sdb,name);
                            if(problem!=null){warnings.Add(string.Format("[{0}] dòng {1}: {2} (SDB {3}): {4} -> chưa nạp.",kv.Key,r,name,sdb,problem));continue;}
                            EmployeeRecord source=external.Find(sdb,name,sourceYears,out problem);
                            if(problem!=null){warnings.Add(string.Format("[{0}] dòng {1}: {2} (SDB {3}): {4} -> chưa nạp.",kv.Key,r,name,sdb,problem));continue;}
                            var candidate = new EmployeeRecord {
                                HoTen = source==null?name:source.HoTen, SDB = source==null?sdb:source.SDB, BoPhan = currentTeam, ToSanXuat = currentTeam,
                                DanhXung = source == null ? "ông" : source.DanhXung,
                                ChucDanh = source == null ? "" : source.ChucDanh,
                                MucLuongBHXH = source == null ? 0 : source.MucLuongBHXH,
                                PhuCapChucVu = source == null ? 0 : source.PhuCapChucVu,
                                HeSoCoDinh = source == null ? 0 : source.HeSoCoDinh,
                                LaQuanLy = source != null && EmployeeClassification.IsManagementRole(source.ChucDanh)
                            };
                            if (source == null || string.IsNullOrWhiteSpace(currentTeam))
                                candidate = requestEmployee == null ? null : requestEmployee(candidate);
                            if (candidate == null)
                            {
                                declined.Add(identity);
                                warnings.Add(string.Format("[{0}] dòng {1}: {2} (SDB {3}) chưa bổ sung hồ sơ/tổ -> chưa nạp.", kv.Key, r, name, sdb));
                                continue;
                            }
                            emp = candidate;
                            EmployeeClassification.Apply(emp);
                            emp.STT = ws.ListNhanVien.Count + 1;
                            ws.ListNhanVien.Add(emp);
                            local.Add(emp);
                            log.AppendLine(string.Format("- {0}: {1} (SDB {2}), tổ {3}.", source == null ? "Đã thêm hồ sơ" : "Đã sao chép hồ sơ", name, sdb, emp.ToSanXuat));
                        }
                        if (lay.Kind == SheetKind.Cong && !string.IsNullOrWhiteSpace(currentTeam))
                        {
                            emp.ToSanXuat = emp.BoPhan = currentTeam;
                            EmployeeClassification.Apply(emp);
                        }
                        ok++;

                        switch (lay.Kind)
                        {
                            case SheetKind.Cong:
                                emp.AttendanceEdited=true;
                                if (touchedCong.Add(emp))
                                {
                                    emp.CongSP = 0; emp.CongHocHop = 0; emp.CongVHTT = 0;
                                    emp.CongLePhep = 0; emp.CongBHXH = 0; emp.CongK3 = 0;
                                }
                                emp.CongSP += Num(doc, ss, lay.CSP, r);
                                emp.CongHocHop += Num(doc, ss, lay.CTG, r);
                                emp.CongVHTT += Num(doc, ss, lay.CVH, r);
                                emp.CongLePhep += Num(doc, ss, lay.CPL, r);
                                emp.CongBHXH += Num(doc, ss, lay.CBH, r);
                                emp.CongK3 += Num(doc, ss, lay.CK3, r);
                                emp.CoBangCong = true;
                                matchedCong.Add(emp);
                                break;

                            case SheetKind.AnCa:
                                double na = 0, tt = 0;
                                for (int c = lay.DayStart; c <= lay.DayEnd; c++)
                                {
                                    string v = Cell(doc, ss, c, r).ToLowerInvariant();
                                    if (v == "na") na++;
                                    else if (v == "tt") tt++;
                                }
                                emp.AnCaNhaAn = na;
                                emp.AnCaTienMat = tt;
                                emp.AnCaDonVi = na + tt;
                                emp.CongAnCa = na + tt;
                                if (string.IsNullOrEmpty(emp.AnCaNguon)) emp.AnCaNguon = "CT";
                                emp.TienAnCa = emp.CongAnCa * GiaAnCa;
                                emp.CoAnCa = true;
                                matchedAnCa.Add(emp);
                                break;

                            case SheetKind.BDDH:
                                double m1 = 0, m2 = 0, m3 = 0;
                                for (int c = lay.DayStart; c <= lay.DayEnd; c++)
                                {
                                    string v = Cell(doc, ss, c, r).ToUpperInvariant();
                                    if (v == "") continue;
                                    foreach (string tk in v.Split(new char[] { ',' }, StringSplitOptions.None))
                                    {
                                        string x = tk.Trim();
                                        if (x == "M1") m1++;
                                        else if (x == "M2") m2++;
                                        else if (x == "M3") m3++;
                                    }
                                }
                                emp.BDDH_M1 = m1; emp.BDDH_M2 = m2; emp.BDDH_M3 = m3;
                                emp.TienBDDH = m1 * GiaM1 + m2 * GiaM2 + m3 * GiaM3;
                                emp.CoBDDH = true;
                                matchedBDDH.Add(emp);
                                break;

                            case SheetKind.Diem:
                                double dsx=ExcelService.GetCellDouble(doc,ss,ColName(lay.CDiemSX)+r,double.NaN);
                                double dh=lay.CDiemHuong>0?ExcelService.GetCellDouble(doc,ss,ColName(lay.CDiemHuong)+r,double.NaN):double.NaN;
                                string xl=lay.CXepLoai>0?Cell(doc,ss,lay.CXepLoai,r).ToUpperInvariant():"";
                                if(!double.IsNaN(dsx) && !double.IsInfinity(dsx))emp.DiemSanXuat=dsx;
                                if(!double.IsNaN(dh) && !double.IsInfinity(dh))emp.HeSoHuong=dh;
                                if(xl=="A" || xl=="B" || xl=="C")emp.XepLoai=xl;
                                if(!emp.LaQuanLy && ws.MaPX!="CDNL" && !double.IsNaN(dsx) && !double.IsInfinity(dsx)) {
                                    if(double.IsNaN(dh) || double.IsInfinity(dh)) {
                                        string grade=(emp.XepLoai??"A").ToUpperInvariant();double factor=grade=="B"?0.9:grade=="C"?0.8:1;
                                        emp.HeSoHuong=Math.Round(dsx*factor,2);
                                    }
                                    emp.ImportedPoints=new ImportedPointValues(emp);
                                }
                                break;
                        }
                        emp.NguonDuLieu = "Khớp Tên + SDB";
                    }

                    string kind = lay.Kind == SheetKind.Cong ? "Bảng công" : lay.Kind == SheetKind.AnCa ? "QT ăn ca (na/tt)" : lay.Kind == SheetKind.BDDH ? "Bảng BDĐH (M1/M2/M3)" : "Bảng điểm";
                    log.AppendLine(string.Format("- Sheet '{0}' [{1}]: {2} dòng có SDB, khớp Tên+SDB: {3}.", kv.Key, kind, rows, ok));
                }
            }

            if (sheetsUsed == 0)
                return "- Không nhận diện được sheet Bảng công / QT ăn ca / Bảng BDĐH / Bảng điểm trong file này.\n";

            // Gán tổ sản xuất = Tổ/Bộ phận trong danh sách
            foreach (var e in ws.ListNhanVien)
                if (string.IsNullOrEmpty(e.ToSanXuat)) e.ToSanXuat = e.BoPhan;

            log.AppendLine(string.Format("- Danh sách PX: {0} người | khớp Bảng công: {1} | khớp ăn ca: {2} | khớp BDĐH: {3}.",
                ws.ListNhanVien.Count, matchedCong.Count, matchedAnCa.Count, matchedBDDH.Count));

            var missing = new List<string>();
            foreach (var e in ws.ListNhanVien)
            {
                var miss = new List<string>();
                if (touchedCong.Count > 0 && !matchedCong.Contains(e)) miss.Add("công");
                bool anyAn = matchedAnCa.Count > 0, anyBd = matchedBDDH.Count > 0;
                if (anyAn && !matchedAnCa.Contains(e)) miss.Add("ăn ca");
                if (anyBd && !matchedBDDH.Contains(e)) miss.Add("BDĐH");
                if (miss.Count > 0) missing.Add(string.Format("{0} (SDB {1}): thiếu {2}", e.HoTen, e.SDB, string.Join(", ", miss.ToArray())));
            }
            if (missing.Count > 0)
            {
                log.AppendLine(string.Format("- {0} người trong danh sách KHÔNG tìm thấy (Tên+SDB) trong file:", missing.Count));
                for (int i = 0; i < missing.Count && i < 25; i++) log.AppendLine("    + " + missing[i]);
                if (missing.Count > 25) log.AppendLine("    ... và " + (missing.Count - 25) + " người khác.");
            }
            if (warnings.Count > 0)
            {
                log.AppendLine(string.Format("- {0} dòng trong file chưa được nạp:", warnings.Count));
                for (int i = 0; i < warnings.Count && i < 25; i++) log.AppendLine("    ! " + warnings[i]);
            }

            ws.RecalculateFromEmployees(true);

            // SO SÁNH công ăn ca theo CÔNG THỨC (Công SP + Học H + VHTT) với số của ĐƠN VỊ gửi (na + tt)
            if (matchedAnCa.Count > 0)
            {
                double sumCT = 0, sumDV = 0;
                var diffs = new List<string>();
                foreach (var e in ws.ListNhanVien)
                {
                    if (!e.CoAnCa) continue;
                    sumCT += e.AnCaCongThuc; sumDV += e.AnCaDonVi;
                    if (Math.Abs(e.AnCaChenhLech) > 0.001)
                        diffs.Add(string.Format("{0} (SDB {1}): công thức {2:N1} = SP {3:N1} + H {4:N1} + VHTT {5:N1} | đơn vị {6:N1} | chênh {7:+0.0;-0.0}",
                            e.HoTen, e.SDB, e.AnCaCongThuc, e.CongSP, e.CongHocHop, e.CongVHTT, e.AnCaDonVi, e.AnCaChenhLech));
                }
                log.AppendLine(string.Format("- SO SÁNH CÔNG ĂN CA: theo công thức = {0:N1} | đơn vị gửi (na+tt) = {1:N1} | chênh lệch = {2:+0.0;-0.0;0.0} | {3} người lệch.",
                    sumCT, sumDV, sumDV - sumCT, diffs.Count));
                for (int i = 0; i < diffs.Count && i < 25; i++) log.AppendLine("    ≠ " + diffs[i]);
                if (diffs.Count > 25) log.AppendLine("    ... và " + (diffs.Count - 25) + " người khác (xem cột Chênh lệch ở Tab 1).");
                log.AppendLine("  (Mặc định đang dùng CÔNG THỨC. Ở Tab 1 bạn chọn lại từng người: Công thức / Đơn vị / Thủ công.)");
            }

            log.AppendLine("- Tổng hợp theo TỔ (từ chi tiết từng người):");
            foreach (var s in ws.GetToSummaries())
                log.AppendLine(string.Format("    {0}: {1} LĐ | công SP {2:N1} | ăn ca {3:N1} ({4:N0}đ) | BDĐH M1/M2/M3 {5:N0}/{6:N0}/{7:N0} ({8:N0}đ)",
                    s.TenTo, s.SoLD, s.CongSP, s.CongAnCa, s.TienAnCa, s.BDDH_M1, s.BDDH_M2, s.BDDH_M3, s.TienBDDH));
            log.AppendLine(string.Format("- TOÀN PX: Công SP {0:N1} | Ăn ca {1:N1} công = {2:N0}đ | BDĐH {3:N0}đ  -> đã đẩy sang Bảng nghiệm thu.",
                ws.CongSP, ws.CongAnCa, ws.TienAnCa, ws.TongTienBDDH));
            return log.ToString();
        }
    }
}
