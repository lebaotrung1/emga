using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace CiscoLDTLApp
{
    public class ExcelService
    {
        private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static Dictionary<string, string> GetSheetMap(ZipArchive zip)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var wbEntry = zip.GetEntry("xl/workbook.xml");
            var relEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (wbEntry == null || relEntry == null) return map;

            var relDoc = new XmlDocument();
            using (var s = relEntry.Open()) relDoc.Load(s);
            var relDict = new Dictionary<string, string>();
            foreach (XmlNode node in relDoc.DocumentElement.ChildNodes)
            {
                if (node.Attributes != null && node.Attributes["Id"] != null && node.Attributes["Target"] != null)
                {
                    string target = node.Attributes["Target"].Value.Replace("\\", "/");
                    if (!target.StartsWith("xl/")) target = "xl/" + target.TrimStart('/');
                    relDict[node.Attributes["Id"].Value] = target;
                }
            }

            var wbDoc = new XmlDocument();
            using (var s = wbEntry.Open()) wbDoc.Load(s);
            var ns = new XmlNamespaceManager(wbDoc.NameTable);
            ns.AddNamespace("d", NsMain);
            ns.AddNamespace("r", NsRel);

            var sheets = wbDoc.SelectNodes("//d:sheet", ns);
            if (sheets != null)
            {
                foreach (XmlNode sh in sheets)
                {
                    string name = (sh.Attributes["name"] != null ? sh.Attributes["name"].Value : null);
                    string rId = (sh.Attributes["r:id"] != null ? sh.Attributes["r:id"].Value : null);
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(rId) && relDict.ContainsKey(rId))
                    {
                        map[name] = relDict[rId];
                        map[name.Trim()] = relDict[rId];
                    }
                }
            }
            return map;
        }

        public static List<string> LoadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return list;

            var doc = new XmlDocument();
            using (var s = entry.Open()) doc.Load(s);
            foreach (XmlNode si in doc.DocumentElement.ChildNodes)
            {
                list.Add(si.InnerText ?? "");
            }
            return list;
        }

        private class SheetIndex
        {
            public XmlElement SheetData;
            public Dictionary<int, XmlElement> Rows = new Dictionary<int, XmlElement>();
            public Dictionary<string, XmlElement> Cells = new Dictionary<string, XmlElement>(StringComparer.OrdinalIgnoreCase);
        }

        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<XmlDocument, SheetIndex> _sheetCache =
            new System.Runtime.CompilerServices.ConditionalWeakTable<XmlDocument, SheetIndex>();

        private static SheetIndex GetSheetIndex(XmlDocument doc)
        {
            if (doc == null) return null;
            SheetIndex idx;
            if (_sheetCache.TryGetValue(doc, out idx)) return idx;

            idx = new SheetIndex();
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("d", NsMain);
            idx.SheetData = doc.SelectSingleNode("//d:sheetData", ns) as XmlElement;
            if (idx.SheetData != null)
            {
                foreach (XmlNode rNode in idx.SheetData.ChildNodes)
                {
                    var rEl = rNode as XmlElement;
                    if (rEl == null || rEl.LocalName != "row") continue;
                    int rNum;
                    if (int.TryParse(rEl.GetAttribute("r"), out rNum))
                        idx.Rows[rNum] = rEl;

                    foreach (XmlNode cNode in rEl.ChildNodes)
                    {
                        var cEl = cNode as XmlElement;
                        if (cEl == null || cEl.LocalName != "c") continue;
                        string cRef = cEl.GetAttribute("r");
                        if (!string.IsNullOrEmpty(cRef))
                            idx.Cells[cRef] = cEl;
                    }
                }
            }
            _sheetCache.Add(doc, idx);
            return idx;
        }

        private static XmlElement FindChild(XmlElement parent, string localName)
        {
            if (parent == null) return null;
            foreach (XmlNode ch in parent.ChildNodes)
            {
                var el = ch as XmlElement;
                if (el != null && el.LocalName == localName) return el;
            }
            return null;
        }

        public static XmlDocument LoadSheetXml(ZipArchive zip, string entryPath)
        {
            var entry = zip.GetEntry(entryPath);
            if (entry == null) return null;
            var doc = new XmlDocument();
            using (var s = entry.Open()) doc.Load(s);
            return doc;
        }

        public static string GetCellString(XmlDocument sheetDoc, List<string> sharedStrings, string cellRef)
        {
            if (sheetDoc == null) return "";
            var sIdx = GetSheetIndex(sheetDoc);
            XmlElement c;
            if (sIdx == null || !sIdx.Cells.TryGetValue(cellRef, out c) || c == null) return "";

            string t = c.GetAttribute("t");
            var vNode = FindChild(c, "v");
            if (vNode == null)
            {
                var isNode = FindChild(c, "is");
                return isNode != null ? isNode.InnerText : "";
            }
            string raw = vNode.InnerText ?? "";
            if (t == "s")
            {
                int idx;
                if (int.TryParse(raw, out idx) && idx >= 0 && idx < sharedStrings.Count)
                    return sharedStrings[idx];
            }
            return raw;
        }

        public static string GetCellFormula(XmlDocument sheetDoc, string cellRef)
        {
            if (sheetDoc == null) return "";
            var sIdx = GetSheetIndex(sheetDoc);
            XmlElement c;
            if (sIdx == null || !sIdx.Cells.TryGetValue(cellRef, out c) || c == null) return "";
            var fNode = FindChild(c, "f");
            return fNode != null ? (fNode.InnerText ?? "") : "";
        }

        public static double GetCellDouble(XmlDocument sheetDoc, List<string> sharedStrings, string cellRef, double fallback = 0)
        {
            string s = GetCellString(sheetDoc, sharedStrings, cellRef);
            if (string.IsNullOrWhiteSpace(s) || s.StartsWith("#")) return fallback;
            double val;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out val))
                return val;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out val))
                return val;
            return fallback;
        }

        /// <summary>
        /// Ghi giá trị số vào ô Excel. Mặc định keepFormula = true để GIỮ NGUYÊN 100% CÔNG THỨC (<f>...</f>)
        /// và chỉ cập nhật giá trị bộ nhớ đệm (<v>...</v>), tuyệt đối không biến ô công thức thành "số chết".
        /// </summary>
        public static void SetCellNumeric(XmlDocument sheetDoc, string cellRef, double value, bool keepFormula = true)
        {
            if (sheetDoc == null) return;
            var sIdx = GetSheetIndex(sheetDoc);
            if (sIdx == null || sIdx.SheetData == null) return;

            XmlElement cNode;
            if (!sIdx.Cells.TryGetValue(cellRef, out cNode) || cNode == null)
            {
                string rowStr = Regex.Match(cellRef, @"\d+").Value;
                int rowNum = int.Parse(rowStr);

                XmlElement rowNode;
                if (!sIdx.Rows.TryGetValue(rowNum, out rowNode) || rowNode == null)
                {
                    rowNode = sheetDoc.CreateElement("row", NsMain);
                    rowNode.SetAttribute("r", rowNum.ToString());
                    sIdx.SheetData.AppendChild(rowNode);
                    sIdx.Rows[rowNum] = rowNode;
                }

                cNode = sheetDoc.CreateElement("c", NsMain);
                cNode.SetAttribute("r", cellRef);
                InsertCellInOrder(rowNode, cNode, cellRef);
                sIdx.Cells[cellRef] = cNode;
            }

            if (cNode.HasAttribute("t"))
                cNode.RemoveAttribute("t");

            if (!keepFormula)
            {
                var fNode = FindChild(cNode, "f");
                if (fNode != null) cNode.RemoveChild(fNode);
            }

            var vNode = FindChild(cNode, "v");
            if (vNode == null)
            {
                vNode = sheetDoc.CreateElement("v", NsMain);
                cNode.AppendChild(vNode);
            }
            vNode.InnerText = value.ToString(CultureInfo.InvariantCulture);
        }

        public static void SetCellInlineString(XmlDocument sheetDoc, string cellRef, string text, bool skipIfHasFormula = true)
        {
            if (sheetDoc == null) return;
            var sIdx = GetSheetIndex(sheetDoc);
            if (sIdx == null || sIdx.SheetData == null) return;

            XmlElement cNode;
            if (!sIdx.Cells.TryGetValue(cellRef, out cNode) || cNode == null)
            {
                string rowStr = Regex.Match(cellRef, @"\d+").Value;
                int rowNum = int.Parse(rowStr);

                XmlElement rowNode;
                if (!sIdx.Rows.TryGetValue(rowNum, out rowNode) || rowNode == null)
                {
                    rowNode = sheetDoc.CreateElement("row", NsMain);
                    rowNode.SetAttribute("r", rowNum.ToString());
                    sIdx.SheetData.AppendChild(rowNode);
                    sIdx.Rows[rowNum] = rowNode;
                }

                cNode = sheetDoc.CreateElement("c", NsMain);
                cNode.SetAttribute("r", cellRef);
                InsertCellInOrder(rowNode, cNode, cellRef);
                sIdx.Cells[cellRef] = cNode;
            }

            var fNode = FindChild(cNode, "f");
            if (fNode != null)
            {
                if (skipIfHasFormula) return;
                cNode.RemoveChild(fNode);
            }

            cNode.SetAttribute("t", "inlineStr");
            var vNode = FindChild(cNode, "v");
            if (vNode != null) cNode.RemoveChild(vNode);

            var isNode = FindChild(cNode, "is");
            if (isNode == null)
            {
                isNode = sheetDoc.CreateElement("is", NsMain);
                cNode.AppendChild(isNode);
            }
            isNode.RemoveAll();
            var tNode = sheetDoc.CreateElement("t", NsMain);
            tNode.InnerText = text ?? "";
            isNode.AppendChild(tNode);
        }

        private static int ColNameToIndex(string cellRef)
        {
            string col = Regex.Match(cellRef, @"^[A-Za-z]+").Value.ToUpperInvariant();
            int sum = 0;
            foreach (char ch in col)
            {
                sum = sum * 26 + (ch - 'A' + 1);
            }
            return sum;
        }

        private static void InsertCellInOrder(XmlElement rowNode, XmlElement newCell, string newRef)
        {
            int targetCol = ColNameToIndex(newRef);
            foreach (XmlNode child in rowNode.ChildNodes)
            {
                var el = child as XmlElement;
                if (el != null && el.LocalName == "c")
                {
                    string r = el.GetAttribute("r");
                    if (ColNameToIndex(r) > targetCol)
                    {
                        rowNode.InsertBefore(newCell, el);
                        return;
                    }
                }
            }
            rowNode.AppendChild(newCell);
        }

        private class TongHopCongEntry
        {
            public int Row;
            public string SDB;
            public string HoTen;
            public string ChucDanh;
            public double MucLuong;
            public double CongSP;
            public double CongHoc;
            public double CongVHTT;
            public double CongPhepLe;
            public double CongBHXH;
            public double DiemSX;
            public double LuongSP;
            public double LuongPhepLe;
            public double CongAnCa;
        }

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return Regex.Replace(s.Trim(), @"\s+", " ").ToLowerInvariant();
        }

        private static string NormalizeSDB(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            string clean = s.Trim();
            int n;
            if (int.TryParse(clean, out n))
                return n.ToString("000");
            return clean.ToUpperInvariant();
        }

        public static List<WorkshopData> LoadWorkshopsFromExcel(string excelPath)
        {
            var workshops = WorkshopData.CreateDefaultWorkshops();
            if (!File.Exists(excelPath)) return workshops;

            using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var sheetMap = GetSheetMap(zip);
                var ss = LoadSharedStrings(zip);

                // 0. Đọc bảng 'Tổng hợp công' (danh sách gốc toàn công ty) để ánh xạ SDB -> Dòng trong Tổng hợp công
                var thcBySdb = new Dictionary<string, TongHopCongEntry>(StringComparer.OrdinalIgnoreCase);
                var thcByName = new Dictionary<string, TongHopCongEntry>(StringComparer.OrdinalIgnoreCase);
                if (sheetMap.ContainsKey("Tổng hợp công"))
                {
                    var thcDoc = LoadSheetXml(zip, sheetMap["Tổng hợp công"]);
                    if (thcDoc != null)
                    {
                        for (int r = 6; r <= 820; r++)
                        {
                            string sdb = GetCellString(thcDoc, ss, "C" + r).Trim();
                            string name = GetCellString(thcDoc, ss, "D" + r).Trim();
                            if (string.IsNullOrEmpty(name)) continue;

                            var entry = new TongHopCongEntry
                            {
                                Row = r,
                                SDB = sdb,
                                HoTen = name,
                                ChucDanh = GetCellString(thcDoc, ss, "I" + r).Trim(),
                                MucLuong = GetCellDouble(thcDoc, ss, "K" + r, 0),
                                CongSP = GetCellDouble(thcDoc, ss, "L" + r, 0),
                                CongHoc = GetCellDouble(thcDoc, ss, "M" + r, 0),
                                CongVHTT = GetCellDouble(thcDoc, ss, "N" + r, 0),
                                CongPhepLe = GetCellDouble(thcDoc, ss, "O" + r, 0),
                                CongBHXH = GetCellDouble(thcDoc, ss, "T" + r, 0),
                                DiemSX = GetCellDouble(thcDoc, ss, "U" + r, 0),
                                LuongSP = GetCellDouble(thcDoc, ss, "W" + r, 0),
                                LuongPhepLe = GetCellDouble(thcDoc, ss, "AB" + r, 0),
                                CongAnCa = GetCellDouble(thcDoc, ss, "AF" + r, 0)
                            };
                            string normSdb = NormalizeSDB(sdb);
                            if (!string.IsNullOrEmpty(normSdb) && !thcBySdb.ContainsKey(normSdb))
                                thcBySdb[normSdb] = entry;
                            string normName = NormalizeKey(name);
                            if (!string.IsNullOrEmpty(normName) && !thcByName.ContainsKey(normName))
                                thcByName[normName] = entry;
                        }
                    }
                }

                // 1. Load NT thiêu kết
                var tk = workshops[0];
                if (sheetMap.ContainsKey("NT thiêu kết"))
                {
                    var doc = LoadSheetXml(zip, sheetMap["NT thiêu kết"]);
                    tk.TongCong = GetCellDouble(doc, ss, "D24", tk.TongCong);
                    tk.CongSP = GetCellDouble(doc, ss, "D25", tk.CongSP);
                    tk.CongQL = GetCellDouble(doc, ss, "D27", tk.CongQL);
                    tk.CongLDTT = GetCellDouble(doc, ss, "D29", tk.CongLDTT);
                    tk.CongLePhepCD = GetCellDouble(doc, ss, "D31", tk.CongLePhepCD);
                    tk.CongBHXH = GetCellDouble(doc, ss, "D32", tk.CongBHXH);
                    tk.LDCuoiKy = (int)GetCellDouble(doc, ss, "D33", tk.LDCuoiKy);
                    tk.LDTrucTiep = (int)GetCellDouble(doc, ss, "D34", tk.LDTrucTiep);
                    tk.LDQuanLy = (int)GetCellDouble(doc, ss, "D35", tk.LDQuanLy);
                    tk.LDThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D37", tk.LDThucTeBQ), 2);
                    tk.LDTTThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D38", tk.LDTTThucTeBQ), 2);
                    tk.NgayCongBQ = Math.Round(GetCellDouble(doc, ss, "D39", tk.NgayCongBQ), 2);
                    tk.NgayCongBQQL = Math.Round(GetCellDouble(doc, ss, "D40", tk.NgayCongBQQL), 2);
                    tk.NgayCongBQLDTT = Math.Round(GetCellDouble(doc, ss, "D41", tk.NgayCongBQLDTT), 2);

                    tk.SanLuongKH = GetCellDouble(doc, ss, "C43", tk.SanLuongKH);
                    tk.SanLuongTH = Math.Round(GetCellDouble(doc, ss, "C46", tk.SanLuongTH), 3);
                    tk.SanLuongTinhLuong = Math.Round(GetCellDouble(doc, ss, "D57", tk.SanLuongTinhLuong), 2);
                    tk.SanLuongLoai1 = GetCellDouble(doc, ss, "D60", tk.SanLuongLoai1);
                    tk.DonGiaLoai1 = Math.Round(GetCellDouble(doc, ss, "I60", tk.DonGiaLoai1), 2);
                    tk.ThanhTienLoai1 = Math.Round(GetCellDouble(doc, ss, "J60", tk.ThanhTienLoai1), 0);
                    tk.SanLuongLoai2 = GetCellDouble(doc, ss, "D61", tk.SanLuongLoai2);
                    tk.DonGiaLoai2 = GetCellDouble(doc, ss, "I61", 0);
                    tk.ThanhTienLoai2 = GetCellDouble(doc, ss, "J61", 0);
                    tk.SanLuongLe = GetCellDouble(doc, ss, "D63", tk.SanLuongLe);
                    tk.DonGiaLe = Math.Round(GetCellDouble(doc, ss, "I63", tk.DonGiaLe), 2);
                    tk.ThanhTienLe = Math.Round(GetCellDouble(doc, ss, "J63", tk.ThanhTienLe), 0);

                    tk.TongLuongTT = Math.Round(GetCellDouble(doc, ss, "J56", tk.TongLuongTT), 0);
                    tk.LuongSPTT = Math.Round(GetCellDouble(doc, ss, "J57", tk.LuongSPTT), 0);
                    tk.LuongCVKhac = Math.Round(GetCellDouble(doc, ss, "J74", tk.LuongCVKhac), 0);
                    tk.TruLuong = Math.Round(GetCellDouble(doc, ss, "J75", tk.TruLuong), 0);
                    tk.LuongCheDoTT = Math.Round(GetCellDouble(doc, ss, "J91", tk.LuongCheDoTT), 0);
                    tk.CongLePhepTT = GetCellDouble(doc, ss, "D92", tk.CongLePhepTT);
                    tk.LuongLePhepTT = Math.Round(GetCellDouble(doc, ss, "J92", tk.LuongLePhepTT), 0);
                    tk.CongHocHopTT = GetCellDouble(doc, ss, "D94", tk.CongHocHopTT);
                    tk.LuongHocHopTT = Math.Round(GetCellDouble(doc, ss, "J94", tk.LuongHocHopTT), 0);

                    tk.TongLuongQL = Math.Round(GetCellDouble(doc, ss, "J95", tk.TongLuongQL), 0);
                    tk.LuongSPQL = Math.Round(GetCellDouble(doc, ss, "J96", tk.LuongSPQL), 0);
                    tk.CongCDQL = GetCellDouble(doc, ss, "D114", tk.CongCDQL);
                    tk.LuongCheDoQL = Math.Round(GetCellDouble(doc, ss, "J114", tk.LuongCheDoQL), 0);

                    tk.TongLuongToanPX = Math.Round(GetCellDouble(doc, ss, "J126", tk.TongLuongToanPX), 0);
                    tk.TienAnCa = Math.Round(GetCellDouble(doc, ss, "L126", tk.TienAnCa), 0);
                    tk.CongAnCa = Math.Round(GetCellDouble(doc, ss, "R126", tk.CongAnCa), 1);
                    tk.TLBinhQuanToanPX = Math.Round(GetCellDouble(doc, ss, "O127", tk.TLBinhQuanToanPX), 2);
                    tk.ThuNhapBQToanPX = Math.Round(GetCellDouble(doc, ss, "O131", tk.ThuNhapBQToanPX), 2);
                    tk.TLSPBinhQuanLDTT = Math.Round(GetCellDouble(doc, ss, "O134", tk.TLSPBinhQuanLDTT), 2);
                    tk.ThuNhapBQLDTT = Math.Round(GetCellDouble(doc, ss, "O146", tk.ThuNhapBQLDTT), 2);
                    tk.BQNgayCongLVTT = Math.Round(GetCellDouble(doc, ss, "O149", tk.BQNgayCongLVTT), 0);

                    var gnvList = ReadGNVRows(doc, ss, 80, 90);
                    if (gnvList.Count > 0) tk.ListGNV = gnvList;
                }

                // 2. Load NT gang
                var lg = workshops[1];
                if (sheetMap.ContainsKey("NT gang"))
                {
                    var doc = LoadSheetXml(zip, sheetMap["NT gang"]);
                    lg.TongCong = GetCellDouble(doc, ss, "D9", lg.TongCong);
                    lg.CongSP = GetCellDouble(doc, ss, "D10", lg.CongSP);
                    lg.CongQL = GetCellDouble(doc, ss, "D11", lg.CongQL);
                    lg.CongLDTT = GetCellDouble(doc, ss, "D12", lg.CongLDTT);
                    lg.CongBHXH = GetCellDouble(doc, ss, "D13", lg.CongBHXH);
                    lg.CongLePhepCD = GetCellDouble(doc, ss, "D14", lg.CongLePhepCD);
                    lg.LDCuoiKy = (int)GetCellDouble(doc, ss, "D15", lg.LDCuoiKy);
                    lg.LDTrucTiep = (int)GetCellDouble(doc, ss, "D16", lg.LDTrucTiep);
                    lg.LDQuanLy = (int)GetCellDouble(doc, ss, "D17", lg.LDQuanLy);
                    lg.LDThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D19", lg.LDThucTeBQ), 2);
                    lg.LDTTThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D20", lg.LDTTThucTeBQ), 2);
                    lg.NgayCongBQ = Math.Round(GetCellDouble(doc, ss, "D22", lg.NgayCongBQ), 2);

                    lg.SanLuongKH = Math.Round(GetCellDouble(doc, ss, "C24", lg.SanLuongKH), 2);
                    lg.SanLuongTH = Math.Round(GetCellDouble(doc, ss, "D51", lg.SanLuongTH), 2);
                    lg.SanLuongTinhLuong = lg.SanLuongTH;
                    lg.SanLuongLoai1 = GetCellDouble(doc, ss, "D52", lg.SanLuongLoai1);
                    lg.DonGiaLoai1 = Math.Round(GetCellDouble(doc, ss, "I52", lg.DonGiaLoai1), 2);
                    lg.ThanhTienLoai1 = Math.Round(GetCellDouble(doc, ss, "J52", lg.ThanhTienLoai1), 0);
                    lg.SanLuongLoai2 = GetCellDouble(doc, ss, "D53", lg.SanLuongLoai2);
                    lg.DonGiaLoai2 = Math.Round(GetCellDouble(doc, ss, "I53", lg.DonGiaLoai2), 2);
                    lg.ThanhTienLoai2 = Math.Round(GetCellDouble(doc, ss, "J53", lg.ThanhTienLoai2), 0);
                    lg.SanLuongLe = GetCellDouble(doc, ss, "D55", lg.SanLuongLe);
                    lg.DonGiaLe = Math.Round(GetCellDouble(doc, ss, "I55", lg.DonGiaLe), 2);
                    lg.ThanhTienLe = Math.Round(GetCellDouble(doc, ss, "J55", lg.ThanhTienLe), 0);

                    lg.TongLuongTT = Math.Round(GetCellDouble(doc, ss, "J47", lg.TongLuongTT), 0);
                    lg.LuongSPTT = Math.Round(GetCellDouble(doc, ss, "J48", lg.LuongSPTT), 0);
                    lg.LuongCVKhac = Math.Round(GetCellDouble(doc, ss, "J59", lg.LuongCVKhac), 0);
                    lg.LuongCheDoTT = Math.Round(GetCellDouble(doc, ss, "J104", lg.LuongCheDoTT), 0);
                    lg.CongLePhepTT = GetCellDouble(doc, ss, "D105", lg.CongLePhepTT);
                    lg.LuongLePhepTT = Math.Round(GetCellDouble(doc, ss, "J105", lg.LuongLePhepTT), 0);
                    lg.CongHocHopTT = GetCellDouble(doc, ss, "D107", lg.CongHocHopTT);
                    lg.LuongHocHopTT = Math.Round(GetCellDouble(doc, ss, "J107", lg.LuongHocHopTT), 0);

                    lg.TongLuongQL = Math.Round(GetCellDouble(doc, ss, "J108", lg.TongLuongQL), 0);
                    lg.LuongSPQL = Math.Round(GetCellDouble(doc, ss, "J109", lg.LuongSPQL), 0);
                    lg.CongCDQL = GetCellDouble(doc, ss, "D140", lg.CongCDQL);
                    lg.LuongCheDoQL = Math.Round(GetCellDouble(doc, ss, "J140", lg.LuongCheDoQL), 0);

                    lg.TongLuongToanPX = Math.Round(GetCellDouble(doc, ss, "J161", lg.TongLuongToanPX), 0);
                    lg.TienAnCa = Math.Round(GetCellDouble(doc, ss, "L161", lg.TienAnCa), 0);
                    lg.CongAnCa = Math.Round(GetCellDouble(doc, ss, "S161", lg.CongAnCa), 1);
                    lg.TLBinhQuanToanPX = Math.Round(GetCellDouble(doc, ss, "O163", lg.TLBinhQuanToanPX), 2);
                    lg.ThuNhapBQToanPX = Math.Round(GetCellDouble(doc, ss, "O167", lg.ThuNhapBQToanPX), 2);

                    var gnvList = ReadGNVRows(doc, ss, 60, 102);
                    if (gnvList.Count > 0) lg.ListGNV = gnvList;
                }

                // 3. Load NT Thép
                var lt = workshops[2];
                if (sheetMap.ContainsKey("NT Thép"))
                {
                    var doc = LoadSheetXml(zip, sheetMap["NT Thép"]);
                    lt.TongCong = GetCellDouble(doc, ss, "D8", lt.TongCong);
                    lt.CongSP = GetCellDouble(doc, ss, "D9", lt.CongSP);
                    lt.CongQL = GetCellDouble(doc, ss, "D10", lt.CongQL);
                    lt.CongLDTT = GetCellDouble(doc, ss, "D11", lt.CongLDTT);
                    lt.CongLePhepCD = GetCellDouble(doc, ss, "D12", lt.CongLePhepCD);
                    lt.CongBHXH = GetCellDouble(doc, ss, "D13", lt.CongBHXH);
                    lt.LDCuoiKy = (int)GetCellDouble(doc, ss, "D14", lt.LDCuoiKy);
                    lt.LDTrucTiep = (int)GetCellDouble(doc, ss, "D15", lt.LDTrucTiep);
                    lt.LDQuanLy = (int)GetCellDouble(doc, ss, "D16", lt.LDQuanLy);
                    lt.LDThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D18", lt.LDThucTeBQ), 2);
                    lt.LDTTThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D19", lt.LDTTThucTeBQ), 2);
                    lt.NgayCongBQ = Math.Round(GetCellDouble(doc, ss, "D20", lt.NgayCongBQ), 2);

                    lt.SanLuongKH = Math.Round(GetCellDouble(doc, ss, "C24", lt.SanLuongKH), 2);
                    lt.SanLuongTH = Math.Round(GetCellDouble(doc, ss, "C28", lt.SanLuongTH), 2);
                    lt.SanLuongTinhLuong = Math.Round(GetCellDouble(doc, ss, "D36", lt.SanLuongTinhLuong), 2);
                    lt.SanLuongLoai1 = GetCellDouble(doc, ss, "D37", lt.SanLuongLoai1);
                    lt.DonGiaLoai1 = Math.Round(GetCellDouble(doc, ss, "I37", lt.DonGiaLoai1), 2);
                    lt.ThanhTienLoai1 = Math.Round(GetCellDouble(doc, ss, "J37", lt.ThanhTienLoai1), 0);
                    lt.SanLuongLoai2 = GetCellDouble(doc, ss, "D38", lt.SanLuongLoai2);
                    lt.DonGiaLoai2 = GetCellDouble(doc, ss, "I38", 0);
                    lt.ThanhTienLoai2 = GetCellDouble(doc, ss, "J38", 0);
                    lt.ThanhTienLoai2 = Math.Round(GetCellDouble(doc, ss, "J38", lt.ThanhTienLoai2), 0);
                    lt.SanLuongLe = GetCellDouble(doc, ss, "D43", lt.SanLuongLe);
                    lt.DonGiaLe = GetCellDouble(doc, ss, "I43", 0);
                    lt.ThanhTienLe = Math.Round(GetCellDouble(doc, ss, "J43", lt.ThanhTienLe), 0);
                    lt.SanLuongPhuKL = Math.Round(GetCellDouble(doc, ss, "D44", lt.SanLuongPhuKL), 2);
                    lt.SanLuongPhuTT = Math.Round(GetCellDouble(doc, ss, "J44", lt.SanLuongPhuTT), 0);

                    lt.TongLuongTT = Math.Round(GetCellDouble(doc, ss, "J34", lt.TongLuongTT), 0);
                    lt.LuongSPTT = Math.Round(GetCellDouble(doc, ss, "J35", lt.LuongSPTT), 0);
                    lt.CongCVKhac = GetCellDouble(doc, ss, "D59", lt.CongCVKhac);
                    lt.LuongCVKhac = Math.Round(GetCellDouble(doc, ss, "J59", lt.LuongCVKhac), 0);
                    lt.LuongCheDoTT = Math.Round(GetCellDouble(doc, ss, "J98", lt.LuongCheDoTT), 0);
                    lt.CongLePhepTT = GetCellDouble(doc, ss, "D99", lt.CongLePhepTT);
                    lt.LuongLePhepTT = Math.Round(GetCellDouble(doc, ss, "J99", lt.LuongLePhepTT), 0);
                    lt.CongHocHopTT = GetCellDouble(doc, ss, "D100", lt.CongHocHopTT);
                    lt.LuongHocHopTT = Math.Round(GetCellDouble(doc, ss, "J100", lt.LuongHocHopTT), 0);

                    lt.TongLuongQL = Math.Round(GetCellDouble(doc, ss, "J102", lt.TongLuongQL), 0);
                    lt.LuongSPQL = Math.Round(GetCellDouble(doc, ss, "J103", lt.LuongSPQL), 0);
                    lt.CongCDQL = GetCellDouble(doc, ss, "D130", lt.CongCDQL);
                    lt.LuongCheDoQL = Math.Round(GetCellDouble(doc, ss, "J130", lt.LuongCheDoQL), 0);

                    lt.TongLuongToanPX = Math.Round(GetCellDouble(doc, ss, "J137", lt.TongLuongToanPX), 0);
                    lt.TienAnCa = Math.Round(GetCellDouble(doc, ss, "L137", lt.TienAnCa), 0);
                    lt.CongAnCa = Math.Round(GetCellDouble(doc, ss, "Q137", lt.CongAnCa), 1);
                    lt.TLBinhQuanToanPX = Math.Round(GetCellDouble(doc, ss, "O143", lt.TLBinhQuanToanPX), 2);
                    lt.ThuNhapBQToanPX = Math.Round(GetCellDouble(doc, ss, "O144", lt.ThuNhapBQToanPX), 2);

                    var gnvList = ReadGNVRows(doc, ss, 60, 96);
                    if (gnvList.Count > 0) lt.ListGNV = gnvList;
                }

                // 4. Load NT CĐNL
                var cdnl = workshops[3];
                if (sheetMap.ContainsKey("NT CĐNL"))
                {
                    var doc = LoadSheetXml(zip, sheetMap["NT CĐNL"]);
                    cdnl.TongCong = GetCellDouble(doc, ss, "D9", cdnl.TongCong);
                    cdnl.CongSP = GetCellDouble(doc, ss, "D10", cdnl.CongSP);
                    cdnl.CongQL = GetCellDouble(doc, ss, "D12", cdnl.CongQL);
                    cdnl.CongLDTT = GetCellDouble(doc, ss, "D14", cdnl.CongLDTT);
                    cdnl.CongLePhepCD = GetCellDouble(doc, ss, "D16", cdnl.CongLePhepCD);
                    cdnl.CongBHXH = GetCellDouble(doc, ss, "D17", cdnl.CongBHXH);
                    cdnl.LDCuoiKy = (int)GetCellDouble(doc, ss, "D18", cdnl.LDCuoiKy);
                    cdnl.LDTrucTiep = (int)GetCellDouble(doc, ss, "D19", cdnl.LDTrucTiep);
                    cdnl.LDQuanLy = (int)GetCellDouble(doc, ss, "D20", cdnl.LDQuanLy);
                    cdnl.LDThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D22", cdnl.LDThucTeBQ), 2);
                    cdnl.LDTTThucTeBQ = Math.Round(GetCellDouble(doc, ss, "D23", cdnl.LDTTThucTeBQ), 2);
                    cdnl.NgayCongBQ = Math.Round(GetCellDouble(doc, ss, "D24", cdnl.NgayCongBQ), 2);

                    cdnl.SanLuongTH = Math.Round(GetCellDouble(doc, ss, "D37", cdnl.SanLuongTH), 0);
                    cdnl.TongLuongTT = Math.Round(GetCellDouble(doc, ss, "I42", cdnl.TongLuongTT), 0);
                    cdnl.LuongSPTT = Math.Round(GetCellDouble(doc, ss, "I43", cdnl.LuongSPTT), 0);
                    cdnl.LuongCheDoTT = Math.Round(GetCellDouble(doc, ss, "I417", cdnl.LuongCheDoTT), 0);
                    cdnl.CongLePhepTT = GetCellDouble(doc, ss, "E418", cdnl.CongLePhepTT);
                    cdnl.LuongLePhepTT = Math.Round(GetCellDouble(doc, ss, "I418", cdnl.LuongLePhepTT), 0);
                    cdnl.CongHocHopTT = GetCellDouble(doc, ss, "E420", cdnl.CongHocHopTT);
                    cdnl.LuongHocHopTT = Math.Round(GetCellDouble(doc, ss, "I420", cdnl.LuongHocHopTT), 0);

                    cdnl.TongLuongQL = Math.Round(GetCellDouble(doc, ss, "I421", cdnl.TongLuongQL), 0);
                    cdnl.LuongSPQL = Math.Round(GetCellDouble(doc, ss, "I422", cdnl.LuongSPQL), 0);
                    cdnl.CongCDQL = GetCellDouble(doc, ss, "E432", cdnl.CongCDQL);
                    cdnl.LuongCheDoQL = Math.Round(GetCellDouble(doc, ss, "I432", cdnl.LuongCheDoQL), 0);

                    cdnl.TongLuongToanPX = Math.Round(GetCellDouble(doc, ss, "I440", cdnl.TongLuongToanPX), 0);
                    cdnl.TienAnCa = Math.Round(GetCellDouble(doc, ss, "K440", cdnl.TienAnCa), 0);
                    cdnl.CongAnCa = Math.Round(GetCellDouble(doc, ss, "Q440", cdnl.CongAnCa), 1);
                    cdnl.TongTienBDDH = Math.Round(GetCellDouble(doc, ss, "M440", cdnl.TongTienBDDH), 0);
                    cdnl.TLBinhQuanToanPX = Math.Round(GetCellDouble(doc, ss, "N441", cdnl.TLBinhQuanToanPX), 2);
                    cdnl.ThuNhapBQToanPX = Math.Round(GetCellDouble(doc, ss, "N443", cdnl.ThuNhapBQToanPX), 2);
                    cdnl.ThuNhapBQLDTT = Math.Round(GetCellDouble(doc, ss, "N454", cdnl.ThuNhapBQLDTT), 2);

                    var gnvList = ReadGNVRowsCDNL(doc, ss, 372, 393);
                    if (gnvList.Count > 0) cdnl.ListGNV = gnvList;
                }

                // Không dùng số mẫu hard-code làm quỹ QL khi bảng nguồn bị lỗi/không có ô quỹ rõ ràng.
                foreach (var ws in workshops)
                {
                    string sheet = ws.MaPX == "TK" ? "NT thiêu kết" : ws.MaPX == "LG" ? "NT gang"
                        : ws.MaPX == "LT" ? "NT Thép" : ws.MaPX == "CDNL" ? "NT CĐNL" : "";
                    string cell = ws.MaPX == "TK" ? "J96" : ws.MaPX == "LG" ? "J109"
                        : ws.MaPX == "LT" ? "J103" : "I422";
                    double value;
                    string raw = sheetMap.ContainsKey(sheet) ? GetCellString(LoadSheetXml(zip, sheetMap[sheet]), ss, cell) : "";
                    if (sheet == "" || !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                    {
                        ws.LuongSPQL = 0;
                        ws.CanhBaoQuyQuanLy = "Chưa có quỹ SP quản lý hợp lệ từ Excel. Hãy nhập quỹ đã duyệt; không dùng số mẫu của bản cũ.";
                    }
                }

                // 5. Load TT BDĐH
                if (sheetMap.ContainsKey("TT BDĐH"))
                {
                    var doc = LoadSheetXml(zip, sheetMap["TT BDĐH"]);
                    foreach (var ws in workshops)
                    {
                        if (ws.RowTTBDDH_T9 > 0)
                        {
                            int r = ws.RowTTBDDH_T9;
                            double m1 = GetCellDouble(doc, ss, "E" + r, ws.BDDH_Muc1_Cong);
                            double m2 = GetCellDouble(doc, ss, "F" + r, ws.BDDH_Muc2_Cong);
                            double m3 = GetCellDouble(doc, ss, "G" + r, ws.BDDH_Muc3_Cong);
                            if (m1 > 0 || m2 > 0 || m3 > 0)
                            {
                                ws.BDDH_Muc1_Cong = m1;
                                ws.BDDH_Muc1_Tien = m1 * 13000;
                                ws.BDDH_Muc2_Cong = m2;
                                ws.BDDH_Muc2_Tien = m2 * 20000;
                                ws.BDDH_Muc3_Cong = m3;
                                ws.BDDH_Muc3_Tien = m3 * 26000;
                                ws.TongTienBDDH = ws.BDDH_Muc1_Tien + ws.BDDH_Muc2_Tien + ws.BDDH_Muc3_Tien;
                            }
                        }
                    }
                }

                // 6. ĐỌC DANH SÁCH CHI TIẾT TỪNG NGƯỜI LAO ĐỘNG TRONG CÁC SHEET 'T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển'
                foreach (var ws in workshops)
                {
                    if (!string.IsNullOrEmpty(ws.SheetBangCong) && sheetMap.ContainsKey(ws.SheetBangCong))
                    {
                        var empDoc = LoadSheetXml(zip, sheetMap[ws.SheetBangCong]);
                        LoadWorkshopEmployees(empDoc, ss, ws, thcBySdb, thcByName);
                    }
                    ws.SyncBaselineFromCurrentTotals();
                }
            }

            using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                LoadAdditionalEmployees(zip, workshops);
                LoadManagementProductionSettings(zip, workshops);
                LoadEditedGnv(zip, workshops);
            }
            return workshops;
        }

        private static void LoadWorkshopEmployees(
            XmlDocument doc,
            List<string> ss,
            WorkshopData ws,
            Dictionary<string, TongHopCongEntry> thcBySdb,
            Dictionary<string, TongHopCongEntry> thcByName)
        {
            if (doc == null) return;
            ws.ListNhanVien.Clear();

            bool isCDNL = (ws.MaPX == "CDNL");
            int maxRow = ws.TotalRowBangCong > 0 ? ws.TotalRowBangCong - 1 : 220;
            string currentBoPhan = "I. Quản lý phân xưởng";
            bool isQuanLySection = true;
            int sttCounter = 1;
            bool needSalaryCalc = false;

            for (int r = 4; r <= maxRow; r++)
            {
                string colA = GetCellString(doc, ss, "A" + r).Trim();
                string colB = GetCellString(doc, ss, "B" + r).Trim();
                string colC = GetCellString(doc, ss, "C" + r).Trim();
                string colD = GetCellString(doc, ss, "D" + r).Trim();
                string colE = GetCellString(doc, ss, "E" + r).Trim();
                double colG = GetCellDouble(doc, ss, "G" + r, 0);

                if (string.IsNullOrEmpty(colC)) continue;
                double numCheck;
                if (double.TryParse(colC, out numCheck)) continue;
                string cLower = colC.ToLowerInvariant();
                if (cLower == "họ và tên" || cLower.Contains("tổng cộng") || cLower == "cộng") continue;

                // Nhận diện dòng tiêu đề Tổ / Bộ phận
                bool teamHeading = cLower.StartsWith("tổ ") || cLower.StartsWith("tổ:")
                    || cLower.StartsWith("bộ phận") || cLower.StartsWith("quản lý");
                if (colG <= 0 && (teamHeading ||
                    (string.IsNullOrEmpty(colD) && string.IsNullOrEmpty(colE))))
                {
                    currentBoPhan = colC;
                    if (cLower.Contains("quản lý"))
                        isQuanLySection = true;
                    else if (teamHeading || cLower.Contains("trực tiếp"))
                        isQuanLySection = false;
                    continue;
                }

                // Nếu cả SDB và mức lương đều trống nhưng không phải dòng nhân viên thì bỏ qua
                if (string.IsNullOrEmpty(colD) && colG <= 0) continue;

                var emp = new EmployeeRecord
                {
                    STT = sttCounter++,
                    ExcelRow = r,
                    DanhXung = string.IsNullOrEmpty(colB) ? "ông" : colB,
                    HoTen = colC,
                    SDB = colD,
                    BoPhan = currentBoPhan,
                    LaQuanLy = isQuanLySection,
                    ChucDanh = colE,
                    MucLuongBHXH = colG,
                    PhuCapChucVu = GetCellDouble(doc, ss, "H" + r, 0),
                    HeSoCoDinh = GetCellDouble(doc, ss, "I" + r, 0),
                    TongCong = GetCellDouble(doc, ss, "J" + r, 0),
                    CongSP = GetCellDouble(doc, ss, "K" + r, 0),
                    CongCaDem = GetCellDouble(doc, ss, "L" + r, 0),
                    CongLePhep = GetCellDouble(doc, ss, "M" + r, 0),
                    CongViecRieng = GetCellDouble(doc, ss, "O" + r, 0),
                    CongHocHop = GetCellDouble(doc, ss, "P" + r, 0),
                    CongLamLe = GetCellDouble(doc, ss, "Q" + r, 0),
                    CongVHTT = GetCellDouble(doc, ss, "R" + r, 0),
                    CongBHXH = GetCellDouble(doc, ss, "S" + r, 0),
                    CongAnCa = GetCellDouble(doc, ss, (isCDNL ? "V" : "T") + r, 0)
                };

                if (emp.CongLePhep == 0)
                    emp.CongLePhep = GetCellDouble(doc, ss, "N" + r, 0);

                if (!isCDNL)
                {
                    if (emp.HeSoCoDinh == 0)
                        emp.HeSoCoDinh = GetCellDouble(doc, ss, "V" + r, 0);
                    emp.DiemSanXuat = Math.Round(GetCellDouble(doc, ss, "W" + r, 0), 2);
                    string xl = GetCellString(doc, ss, "X" + r).Trim();
                    emp.XepLoai = string.IsNullOrEmpty(xl) || xl == "0" ? "A" : xl;
                    emp.HeSoHuong = Math.Round(GetCellDouble(doc, ss, "Z" + r, 0), 2);
                    if (emp.HeSoHuong == 0)
                        emp.HeSoHuong = Math.Round(GetCellDouble(doc, ss, "Y" + r, 0), 2);

                    emp.TongTienLuong = Math.Round(GetCellDouble(doc, ss, "AA" + r, 0), 0);
                    emp.LuongSanPham = Math.Round(GetCellDouble(doc, ss, "AB" + r, 0), 0);
                    emp.LuongLePhepCD = Math.Round(GetCellDouble(doc, ss, "AD" + r, 0), 0);
                    emp.LuongLamLe = Math.Round(GetCellDouble(doc, ss, "AH" + r, 0), 0);
                    emp.LuongHocVHTT = Math.Round(GetCellDouble(doc, ss, "AI" + r, 0), 0);
                    emp.PhuCapTien = Math.Round(GetCellDouble(doc, ss, "AK" + r, 0), 0);
                }
                else
                {
                    if (emp.HeSoCoDinh == 0)
                        emp.HeSoCoDinh = GetCellDouble(doc, ss, "X" + r, 0);
                    emp.DiemSanXuat = Math.Round(GetCellDouble(doc, ss, "Y" + r, 0), 2);
                    string xl = GetCellString(doc, ss, "Z" + r).Trim();
                    emp.XepLoai = string.IsNullOrEmpty(xl) || xl == "0" ? "A" : xl;
                    emp.HeSoHuong = Math.Round(GetCellDouble(doc, ss, "AB" + r, 0), 2);
                    if (emp.HeSoHuong == 0)
                        emp.HeSoHuong = Math.Round(GetCellDouble(doc, ss, "AA" + r, 0), 2);

                    emp.TongTienLuong = Math.Round(GetCellDouble(doc, ss, "AC" + r, 0), 0);
                    emp.LuongSanPham = Math.Round(GetCellDouble(doc, ss, "AD" + r, 0), 0);
                    emp.LuongLePhepCD = Math.Round(GetCellDouble(doc, ss, "AE" + r, 0), 0);
                    emp.LuongLamLe = Math.Round(GetCellDouble(doc, ss, "AI" + r, 0), 0);
                    emp.LuongHocVHTT = Math.Round(GetCellDouble(doc, ss, "AJ" + r, 0), 0);
                    emp.PhuCapTien = Math.Round(GetCellDouble(doc, ss, "AL" + r, 0), 0);
                }

                // Đối chiếu với sheet 'Tổng hợp công' qua SDB hoặc Họ tên
                TongHopCongEntry thc = null;
                string normSdb = NormalizeSDB(emp.SDB);
                if (!string.IsNullOrEmpty(normSdb) && thcBySdb.ContainsKey(normSdb))
                    thc = thcBySdb[normSdb];
                else
                {
                    string normName = NormalizeKey(emp.HoTen);
                    if (!string.IsNullOrEmpty(normName) && thcByName.ContainsKey(normName))
                        thc = thcByName[normName];
                }

                if (thc != null)
                {
                    emp.TongHopCongRow = thc.Row;
                    if (string.IsNullOrEmpty(emp.SDB)) emp.SDB = thc.SDB;
                    if (emp.MucLuongBHXH <= 0 && thc.MucLuong > 0) emp.MucLuongBHXH = thc.MucLuong;
                    if (emp.DiemSanXuat <= 0 && thc.DiemSX > 0) emp.DiemSanXuat = Math.Round(thc.DiemSX, 2);
                    if (emp.LuongSanPham <= 0 && thc.LuongSP > 0) emp.LuongSanPham = Math.Round(thc.LuongSP, 0);
                    if (emp.LuongLePhepCD <= 0 && thc.LuongPhepLe > 0) emp.LuongLePhepCD = Math.Round(thc.LuongPhepLe, 0);
                    if (emp.CongAnCa <= 0 && thc.CongAnCa > 0) emp.CongAnCa = thc.CongAnCa;
                }

                if (emp.HeSoHuong <= 0 && emp.DiemSanXuat > 0)
                {
                    string xlUpper = (emp.XepLoai ?? "A").Trim().ToUpperInvariant();
                    double rate = xlUpper == "A" ? 1.0 : (xlUpper == "B" ? 0.9 : 0.8);
                    emp.HeSoHuong = Math.Round(emp.DiemSanXuat * rate, 2);
                }

                if (emp.CongSP > 0 && emp.LuongSanPham <= 0 && (ws.LuongSPTT > 0 || ws.LuongSPQL > 0))
                {
                    needSalaryCalc = true;
                }

                if (emp.TongTienLuong <= 0)
                {
                    emp.TongTienLuong = emp.LuongSanPham + emp.LuongLePhepCD + emp.LuongLamLe + emp.LuongHocVHTT + emp.PhuCapTien;
                }

                ws.ListNhanVien.Add(emp);
            }

            // Nếu trong sheet có những dòng chưa lưu giá trị bộ nhớ đệm của Lương SP thì tự động phân bổ theo đúng công thức Điểm/Hệ số
            if (needSalaryCalc && ws.ListNhanVien.Count > 0)
            {
                PopulateMissingEmployeeSalaries(ws);
            }
        }

        private static void PopulateMissingEmployeeSalaries(WorkshopData ws)
        {
            double sumHeSoQL = 0;
            double sumDiemTT = 0;
            double totalPhuCapTT = 0;

            foreach (var emp in ws.ListNhanVien)
            {
                if (emp.PhuCapChucVu > 0 && emp.PhuCapTien == 0)
                    emp.PhuCapTien = Math.Round((emp.PhuCapChucVu * 1600000.0) / 1000.0, 0) * 1000.0;
                if (emp.MucLuongBHXH > 0 && emp.LuongLePhepCD == 0 && emp.CongLePhep > 0)
                    emp.LuongLePhepCD = Math.Round((emp.CongLePhep / 26.0 * emp.MucLuongBHXH) / 100.0, 0) * 100.0;
                if (emp.LuongHocVHTT == 0 && (emp.CongHocHop > 0 || emp.CongVHTT > 0))
                    emp.LuongHocVHTT = emp.CongHocHop * 250000.0 + emp.CongVHTT * 300000.0;

                if (emp.LaQuanLy)
                {
                    if (emp.HeSoHuong <= 0 && emp.HeSoCoDinh > 0 && emp.CongSP > 0)
                    {
                        emp.DiemSanXuat = Math.Round(emp.CongSP * emp.HeSoCoDinh / 26.0, 2);
                        emp.HeSoHuong = emp.DiemSanXuat;
                    }
                    sumHeSoQL += emp.HeSoHuong;
                }
                else
                {
                    if (emp.HeSoHuong <= 0 && emp.DiemSanXuat > 0)
                        emp.HeSoHuong = emp.DiemSanXuat;
                    sumDiemTT += emp.HeSoHuong;
                    totalPhuCapTT += emp.PhuCapTien + emp.LuongLamLe;
                }
            }

            double quySPTT = Math.Max(0, ws.LuongSPTT + ws.LuongCVKhac + ws.TruLuong - totalPhuCapTT);
            foreach (var emp in ws.ListNhanVien)
            {
                if (emp.LuongSanPham <= 0 && emp.CongSP > 0)
                {
                    if (emp.LaQuanLy && sumHeSoQL > 0 && ws.LuongSPQL > 0)
                        emp.LuongSanPham = Math.Round((ws.LuongSPQL * emp.HeSoHuong / sumHeSoQL) / 100.0, 0) * 100.0;
                    else if (!emp.LaQuanLy && sumDiemTT > 0 && quySPTT > 0)
                        emp.LuongSanPham = Math.Round((quySPTT * emp.HeSoHuong / sumDiemTT) / 100.0, 0) * 100.0;
                }
                emp.TongTienLuong = emp.LuongSanPham + emp.LuongLePhepCD + emp.LuongLamLe + emp.LuongHocVHTT + emp.PhuCapTien;
            }
        }

        private static List<ArisingWorkItem> ReadGNVRows(XmlDocument doc, List<string> ss, int startRow, int endRow)
        {
            var list = new List<ArisingWorkItem>();
            int stt = 1;
            for (int r = startRow; r <= endRow; r++)
            {
                string name = GetCellString(doc, ss, "B" + r).Trim();
                string dvt = GetCellString(doc, ss, "C" + r).Trim();
                double kl = GetCellDouble(doc, ss, "D" + r, 0);
                double dg = GetCellDouble(doc, ss, "I" + r, 0);
                double tt = GetCellDouble(doc, ss, "J" + r, 0);
                if (!string.IsNullOrEmpty(name) && (kl > 0 || tt > 0))
                {
                    list.Add(new ArisingWorkItem(stt++, name, string.IsNullOrEmpty(dvt) ? "Công" : dvt, Math.Round(kl, 2), Math.Round(dg, 0), Math.Round(tt, 0), "", r));
                }
            }
            return list;
        }

        private static List<ArisingWorkItem> ReadGNVRowsCDNL(XmlDocument doc, List<string> ss, int startRow, int endRow)
        {
            var list = new List<ArisingWorkItem>();
            int stt = 1;
            for (int r = startRow; r <= endRow; r++)
            {
                string name = GetCellString(doc, ss, "B" + r).Trim();
                string dvt = GetCellString(doc, ss, "C" + r).Trim();
                double kl = GetCellDouble(doc, ss, "D" + r, 0);
                double dg = GetCellDouble(doc, ss, "H" + r, 0);
                double tt = GetCellDouble(doc, ss, "I" + r, 0);
                if (!string.IsNullOrEmpty(name) && (kl > 0 || tt > 0))
                {
                    list.Add(new ArisingWorkItem(stt++, name, string.IsNullOrEmpty(dvt) ? "Công" : dvt, Math.Round(kl, 2), Math.Round(dg, 0), Math.Round(tt, 0), "Tổ SCCĐ", r));
                }
            }
            return list;
        }

        /// <summary>
        /// Ghi toàn bộ dữ liệu (bao gồm CHI TIẾT TỪNG NGƯỜI LAO ĐỘNG ở sheet 'Tổng hợp công', 'T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển'
        /// và các bảng nghiệm thu 'NT ...', 'TT BDĐH', '9 tháng ...') vào file Excel gốc.
        /// Giữ nguyên 100% các công thức Excel (<f>...</f>)!
        /// </summary>
        public static void SaveWorkshopsToExcel(string excelPath, List<WorkshopData> workshops)
        {
            if (!File.Exists(excelPath))
                throw new FileNotFoundException("Không tìm thấy file Excel gốc: " + excelPath);

            using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                var sheetMap = GetSheetMap(zip);
                var managementStrings = LoadSharedStrings(zip);

                // 1. Mở sheet 'Tổng hợp công' để cập nhật công & lương chi tiết từng người ở bảng gốc toàn công ty
                XmlDocument thcDoc = null;
                string thcPath = null;
                if (sheetMap.ContainsKey("Tổng hợp công"))
                {
                    thcPath = sheetMap["Tổng hợp công"];
                    thcDoc = LoadSheetXml(zip, thcPath);
                }

                foreach (var ws in workshops)
                {
                    ws.RecalculateDerived();

                    // 2. Cập nhật chi tiết từng người vào sheet phân xưởng ('T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển') và 'Tổng hợp công'
                    if (!string.IsNullOrEmpty(ws.SheetBangCong) && sheetMap.ContainsKey(ws.SheetBangCong) && ws.ListNhanVien != null && ws.ListNhanVien.Count > 0)
                    {
                        string empPath = sheetMap[ws.SheetBangCong];
                        var empDoc = LoadSheetXml(zip, empPath);
                        SaveWorkshopEmployeesToSheet(empDoc, thcDoc, ws);
                        SaveSheetXml(zip, empPath, empDoc);
                    }

                    // 3. Cập nhật các sheet NT ... (Giữ nguyên 100% công thức <f> bằng keepFormula = true)
                    if (ws.MaPX == "TK" && sheetMap.ContainsKey("NT thiêu kết"))
                    {
                        string path = sheetMap["NT thiêu kết"];
                        var doc = LoadSheetXml(zip, path);
                        SetCellInlineString(doc, "I2", string.Format("Tháng {0} năm {1}", ws.Thang, ws.Nam), false);
                        SetCellNumeric(doc, "D24", ws.TongCong);
                        SetCellNumeric(doc, "D25", ws.CongSP);
                        SetCellNumeric(doc, "D27", ws.CongQL);
                        SetCellNumeric(doc, "D29", ws.CongLDTT);
                        SetCellNumeric(doc, "D31", ws.CongLePhepCD);
                        SetCellNumeric(doc, "D32", ws.CongBHXH);
                        SetCellNumeric(doc, "D33", ws.LDCuoiKy);
                        SetCellNumeric(doc, "D34", ws.LDTrucTiep);
                        SetCellNumeric(doc, "D35", ws.LDQuanLy);
                        SetCellNumeric(doc, "D37", ws.LDThucTeBQ);
                        SetCellNumeric(doc, "D38", ws.LDTTThucTeBQ);
                        SetCellNumeric(doc, "D39", ws.NgayCongBQ);
                        SetCellNumeric(doc, "D40", ws.NgayCongBQQL);
                        SetCellNumeric(doc, "D41", ws.NgayCongBQLDTT);

                        SetCellNumeric(doc, "C43", ws.SanLuongKH);
                        SetCellNumeric(doc, "C46", ws.SanLuongTH);
                        SetCellNumeric(doc, "D57", ws.SanLuongTinhLuong);
                        SetCellNumeric(doc, "D60", ws.SanLuongLoai1);
                        SetCellNumeric(doc, "I60", ws.DonGiaLoai1);
                        SetCellNumeric(doc, "J60", ws.ThanhTienLoai1);
                        SetCellNumeric(doc, "D61", ws.SanLuongLoai2);
                        SetCellNumeric(doc, "I61", ws.DonGiaLoai2);
                        SetCellNumeric(doc, "J61", ws.ThanhTienLoai2);
                        SetCellNumeric(doc, "D63", ws.SanLuongLe);
                        SetCellNumeric(doc, "I63", ws.DonGiaLe);
                        SetCellNumeric(doc, "J63", ws.ThanhTienLe);

                        SetCellNumeric(doc, "J56", ws.TongLuongTT);
                        SetCellNumeric(doc, "J57", ws.LuongSPTT);
                        SetCellNumeric(doc, "J74", ws.LuongCVKhac);
                        SetCellNumeric(doc, "J75", ws.TruLuong);
                        SetCellNumeric(doc, "J91", ws.LuongCheDoTT);
                        SetCellNumeric(doc, "D92", ws.CongLePhepTT);
                        SetCellNumeric(doc, "J92", ws.LuongLePhepTT);
                        SetCellNumeric(doc, "D94", ws.CongHocHopTT);
                        SetCellNumeric(doc, "J94", ws.LuongHocHopTT);

                        SetCellNumeric(doc, "J95", ws.TongLuongQL);
                        SetCellNumeric(doc, "J96", ws.LuongSPQL);
                        SetCellNumeric(doc, "D114", ws.CongCDQL);
                        SetCellNumeric(doc, "J114", ws.LuongCheDoQL);

                        SetCellNumeric(doc, "J126", ws.TongLuongToanPX);
                        SetCellNumeric(doc, "L126", ws.TienAnCa);
                        SetCellNumeric(doc, "R126", ws.CongAnCa);
                        SetCellNumeric(doc, "N126", ws.TongTienBDDH);
                        SetCellNumeric(doc, "O127", ws.TLBinhQuanToanPX);
                        SetCellNumeric(doc, "O131", ws.ThuNhapBQToanPX);
                        SetCellNumeric(doc, "O134", ws.TLSPBinhQuanLDTT);
                        SetCellNumeric(doc, "O146", ws.ThuNhapBQLDTT);
                        SetCellNumeric(doc, "O149", ws.BQNgayCongLVTT);

                        // Cập nhật bảng chi tiết Quản lý trong NT thiêu kết (dòng 102..106)
                        UpdateNTManagementRows(doc, managementStrings, ws, 102, 106);

                        WriteGNVRows(doc, ws.ListGNV, 81, 89, "D", "I", "J");
                        SaveSheetXml(zip, path, doc);

                        if (sheetMap.ContainsKey("9 tháng Thiêu kết"))
                            UpdateCumulativeSheetStandard(zip, sheetMap["9 tháng Thiêu kết"], ws, 8);
                    }
                    else if (ws.MaPX == "LG" && sheetMap.ContainsKey("NT gang"))
                    {
                        string path = sheetMap["NT gang"];
                        var doc = LoadSheetXml(zip, path);
                        SetCellInlineString(doc, "I2", string.Format("Tháng {0} năm {1}", ws.Thang, ws.Nam), false);
                        SetCellNumeric(doc, "D9", ws.TongCong);
                        SetCellNumeric(doc, "D10", ws.CongSP);
                        SetCellNumeric(doc, "D11", ws.CongQL);
                        SetCellNumeric(doc, "D12", ws.CongLDTT);
                        SetCellNumeric(doc, "D13", ws.CongBHXH);
                        SetCellNumeric(doc, "D14", ws.CongLePhepCD);
                        SetCellNumeric(doc, "D15", ws.LDCuoiKy);
                        SetCellNumeric(doc, "D16", ws.LDTrucTiep);
                        SetCellNumeric(doc, "D17", ws.LDQuanLy);
                        SetCellNumeric(doc, "D19", ws.LDThucTeBQ);
                        SetCellNumeric(doc, "D20", ws.LDTTThucTeBQ);
                        SetCellNumeric(doc, "D22", ws.NgayCongBQ);

                        SetCellNumeric(doc, "C24", ws.SanLuongKH);
                        SetCellNumeric(doc, "D51", ws.SanLuongTH);
                        SetCellNumeric(doc, "D52", ws.SanLuongLoai1);
                        SetCellNumeric(doc, "I52", ws.DonGiaLoai1);
                        SetCellNumeric(doc, "J52", ws.ThanhTienLoai1);
                        SetCellNumeric(doc, "D53", ws.SanLuongLoai2);
                        SetCellNumeric(doc, "I53", ws.DonGiaLoai2);
                        SetCellNumeric(doc, "J53", ws.ThanhTienLoai2);
                        SetCellNumeric(doc, "D55", ws.SanLuongLe);
                        SetCellNumeric(doc, "I55", ws.DonGiaLe);
                        SetCellNumeric(doc, "J55", ws.ThanhTienLe);

                        SetCellNumeric(doc, "J47", ws.TongLuongTT);
                        SetCellNumeric(doc, "J48", ws.LuongSPTT);
                        SetCellNumeric(doc, "J59", ws.LuongCVKhac);
                        SetCellNumeric(doc, "J104", ws.LuongCheDoTT);
                        SetCellNumeric(doc, "D105", ws.CongLePhepTT);
                        SetCellNumeric(doc, "J105", ws.LuongLePhepTT);
                        SetCellNumeric(doc, "D107", ws.CongHocHopTT);
                        SetCellNumeric(doc, "J107", ws.LuongHocHopTT);

                        SetCellNumeric(doc, "J108", ws.TongLuongQL);
                        SetCellNumeric(doc, "J109", ws.LuongSPQL);
                        SetCellNumeric(doc, "D140", ws.CongCDQL);
                        SetCellNumeric(doc, "J140", ws.LuongCheDoQL);

                        SetCellNumeric(doc, "J161", ws.TongLuongToanPX);
                        SetCellNumeric(doc, "L161", ws.TienAnCa);
                        SetCellNumeric(doc, "S161", ws.CongAnCa);
                        SetCellNumeric(doc, "N161", ws.TongTienBDDH);
                        SetCellNumeric(doc, "O163", ws.TLBinhQuanToanPX);
                        SetCellNumeric(doc, "O167", ws.ThuNhapBQToanPX);

                        UpdateNTManagementRows(doc, managementStrings, ws, 115, 121);

                        WriteGNVRows(doc, ws.ListGNV, 60, 75, "D", "I", "J");
                        SaveSheetXml(zip, path, doc);

                        if (sheetMap.ContainsKey("9 tháng Luyện gang"))
                            UpdateCumulativeSheetStandard(zip, sheetMap["9 tháng Luyện gang"], ws, 9);
                    }
                    else if (ws.MaPX == "LT" && sheetMap.ContainsKey("NT Thép"))
                    {
                        string path = sheetMap["NT Thép"];
                        var doc = LoadSheetXml(zip, path);
                        SetCellInlineString(doc, "I2", string.Format("Tháng {0} năm {1}", ws.Thang, ws.Nam), false);
                        SetCellNumeric(doc, "D8", ws.TongCong);
                        SetCellNumeric(doc, "D9", ws.CongSP);
                        SetCellNumeric(doc, "D10", ws.CongQL);
                        SetCellNumeric(doc, "D11", ws.CongLDTT);
                        SetCellNumeric(doc, "D12", ws.CongLePhepCD);
                        SetCellNumeric(doc, "D13", ws.CongBHXH);
                        SetCellNumeric(doc, "D14", ws.LDCuoiKy);
                        SetCellNumeric(doc, "D15", ws.LDTrucTiep);
                        SetCellNumeric(doc, "D16", ws.LDQuanLy);
                        SetCellNumeric(doc, "D18", ws.LDThucTeBQ);
                        SetCellNumeric(doc, "D19", ws.LDTTThucTeBQ);
                        SetCellNumeric(doc, "D20", ws.NgayCongBQ);

                        SetCellNumeric(doc, "C24", ws.SanLuongKH);
                        SetCellNumeric(doc, "C28", ws.SanLuongTH);
                        SetCellNumeric(doc, "D36", ws.SanLuongTinhLuong);
                        SetCellNumeric(doc, "D37", ws.SanLuongLoai1);
                        SetCellNumeric(doc, "I37", ws.DonGiaLoai1);
                        SetCellNumeric(doc, "J37", ws.ThanhTienLoai1);
                        SetCellNumeric(doc, "D38", ws.SanLuongLoai2);
                        SetCellNumeric(doc, "I38", ws.DonGiaLoai2);
                        SetCellNumeric(doc, "J38", ws.ThanhTienLoai2);
                        SetCellNumeric(doc, "D43", ws.SanLuongLe);
                        SetCellNumeric(doc, "I43", ws.DonGiaLe);
                        SetCellNumeric(doc, "J43", ws.ThanhTienLe);
                        SetCellNumeric(doc, "D44", ws.SanLuongPhuKL);
                        SetCellNumeric(doc, "J44", ws.SanLuongPhuTT);

                        SetCellNumeric(doc, "J34", ws.TongLuongTT);
                        SetCellNumeric(doc, "J35", ws.LuongSPTT);
                        SetCellNumeric(doc, "D59", ws.CongCVKhac);
                        SetCellNumeric(doc, "J59", ws.LuongCVKhac);
                        SetCellNumeric(doc, "J98", ws.LuongCheDoTT);
                        SetCellNumeric(doc, "D99", ws.CongLePhepTT);
                        SetCellNumeric(doc, "J99", ws.LuongLePhepTT);
                        SetCellNumeric(doc, "D100", ws.CongHocHopTT);
                        SetCellNumeric(doc, "J100", ws.LuongHocHopTT);

                        SetCellNumeric(doc, "J102", ws.TongLuongQL);
                        SetCellNumeric(doc, "J103", ws.LuongSPQL);
                        SetCellNumeric(doc, "D130", ws.CongCDQL);
                        SetCellNumeric(doc, "J130", ws.LuongCheDoQL);

                        SetCellNumeric(doc, "J137", ws.TongLuongToanPX);
                        SetCellNumeric(doc, "L137", ws.TienAnCa);
                        SetCellNumeric(doc, "Q137", ws.CongAnCa);
                        SetCellNumeric(doc, "N137", ws.TongTienBDDH);
                        SetCellNumeric(doc, "O143", ws.TLBinhQuanToanPX);
                        SetCellNumeric(doc, "O144", ws.ThuNhapBQToanPX);

                        UpdateNTManagementRows(doc, managementStrings, ws, 108, 113);

                        WriteGNVRows(doc, ws.ListGNV, 61, 80, "D", "I", "J");
                        SaveSheetXml(zip, path, doc);

                        if (sheetMap.ContainsKey("9 tháng Luyện thép"))
                            UpdateCumulativeSheetStandard(zip, sheetMap["9 tháng Luyện thép"], ws, 9);
                    }
                    else if (ws.MaPX == "CDNL" && sheetMap.ContainsKey("NT CĐNL"))
                    {
                        string path = sheetMap["NT CĐNL"];
                        var doc = LoadSheetXml(zip, path);
                        SetCellInlineString(doc, "H2", string.Format("Tháng {0} năm {1}", ws.Thang, ws.Nam), false);
                        SetCellNumeric(doc, "D9", ws.TongCong);
                        SetCellNumeric(doc, "D10", ws.CongSP);
                        SetCellNumeric(doc, "D12", ws.CongQL);
                        SetCellNumeric(doc, "D14", ws.CongLDTT);
                        SetCellNumeric(doc, "D16", ws.CongLePhepCD);
                        SetCellNumeric(doc, "D17", ws.CongBHXH);
                        SetCellNumeric(doc, "D18", ws.LDCuoiKy);
                        SetCellNumeric(doc, "D19", ws.LDTrucTiep);
                        SetCellNumeric(doc, "D20", ws.LDQuanLy);
                        SetCellNumeric(doc, "D22", ws.LDThucTeBQ);
                        SetCellNumeric(doc, "D23", ws.LDTTThucTeBQ);
                        SetCellNumeric(doc, "D24", ws.NgayCongBQ);

                        SetCellNumeric(doc, "I42", ws.TongLuongTT);
                        SetCellNumeric(doc, "I43", ws.LuongSPTT);
                        SetCellNumeric(doc, "I417", ws.LuongCheDoTT);
                        SetCellNumeric(doc, "E418", ws.CongLePhepTT);
                        SetCellNumeric(doc, "I418", ws.LuongLePhepTT);
                        SetCellNumeric(doc, "E420", ws.CongHocHopTT);
                        SetCellNumeric(doc, "I420", ws.LuongHocHopTT);

                        SetCellNumeric(doc, "I421", ws.TongLuongQL);
                        SetCellNumeric(doc, "I422", ws.LuongSPQL);
                        SetCellNumeric(doc, "E432", ws.CongCDQL);
                        SetCellNumeric(doc, "I432", ws.LuongCheDoQL);

                        SetCellNumeric(doc, "I440", ws.TongLuongToanPX);
                        SetCellNumeric(doc, "K440", ws.TienAnCa);
                        SetCellNumeric(doc, "Q440", ws.CongAnCa);
                        SetCellNumeric(doc, "M440", ws.TongTienBDDH);
                        SetCellNumeric(doc, "N441", ws.TLBinhQuanToanPX);
                        SetCellNumeric(doc, "N443", ws.ThuNhapBQToanPX);
                        SetCellNumeric(doc, "N454", ws.ThuNhapBQLDTT);

                        UpdateNTManagementRows(doc, managementStrings, ws, 426, 430);

                        SaveSheetXml(zip, path, doc);

                        if (sheetMap.ContainsKey("9 tháng CĐNL"))
                            UpdateCumulativeSheetCDNL(zip, sheetMap["9 tháng CĐNL"], ws);
                    }
                    if (ws.MaPX == "KT" && sheetMap.ContainsKey("NT. K. tuyển "))
                    {
                        bool applied = false;
                        foreach (var employee in ws.ListNhanVien) if (employee.QuanLyTinhMoi) applied = true;
                        if (applied)
                        {
                            string path = sheetMap["NT. K. tuyển "];
                            var doc = LoadSheetXml(zip, path);
                            SetCellNumeric(doc, "J106", ws.TongLuongQL);
                            SetCellNumeric(doc, "J131", ws.TongLuongToanPX);
                            SaveSheetXml(zip, path, doc);
                        }
                    }
                }

                // Lưu sheet 'Tổng hợp công' sau khi đã cập nhật toàn bộ nhân viên các phân xưởng
                if (thcDoc != null && !string.IsNullOrEmpty(thcPath))
                {
                    SaveSheetXml(zip, thcPath, thcDoc);
                }

                // Update TT BDĐH
                if (sheetMap.ContainsKey("TT BDĐH"))
                {
                    string path = sheetMap["TT BDĐH"];
                    var doc = LoadSheetXml(zip, path);
                    foreach (var ws in workshops)
                    {
                        if (ws.RowTTBDDH_T9 > 0)
                        {
                            int r = ws.RowTTBDDH_T9;
                            SetCellNumeric(doc, "D" + r, ws.CongSP);
                            SetCellNumeric(doc, "E" + r, ws.BDDH_Muc1_Cong);
                            SetCellNumeric(doc, "F" + r, ws.BDDH_Muc2_Cong);
                            SetCellNumeric(doc, "G" + r, ws.BDDH_Muc3_Cong);
                            SetCellNumeric(doc, "H" + r, ws.TongTienBDDH);
                        }
                    }
                    SaveSheetXml(zip, path, doc);
                }

                // Bật cờ fullCalcOnLoad="1" trong xl/workbook.xml để khi mở bằng Excel, mọi công thức tự động tính toán lại 100%
                SaveAdditionalEmployees(zip, workshops);
                SaveManagementProductionSettings(zip, workshops);
                SaveEditedGnv(zip, workshops);
                SaveManagementPayrollReport(zip, workshops);
                EnsureFullCalcOnLoad(zip);
            }
        }

        // A regular worksheet keeps added employees portable when the workbook is opened in Excel.
        private const string AdditionalSheet = "Nhan su bo sung";
        private static string EnsureWorksheet(ZipArchive zip, string sheetName, string preferredPath)
        {
            var map = GetSheetMap(zip);
            string path;
            if (!map.TryGetValue(sheetName, out path))
            {
                path = preferredPath;
                var wb = new XmlDocument();
                using (var stream = zip.GetEntry("xl/workbook.xml").Open()) wb.Load(stream);
                var rel = new XmlDocument();
                using (var stream = zip.GetEntry("xl/_rels/workbook.xml.rels").Open()) rel.Load(stream);
                string relId = "rIdCisco" + Regex.Replace(sheetName, "[^A-Za-z0-9]", "");
                while (rel.SelectSingleNode("//*[@Id='" + relId + "']") != null) relId += "X";
                var relationship = rel.CreateElement("Relationship", rel.DocumentElement.NamespaceURI);
                relationship.SetAttribute("Id", relId);
                relationship.SetAttribute("Type", NsRel + "/worksheet");
                relationship.SetAttribute("Target", path.Substring(3));
                rel.DocumentElement.AppendChild(relationship);
                int maxId = 0;
                foreach (XmlElement sheet in wb.GetElementsByTagName("sheet", NsMain))
                    maxId = Math.Max(maxId, int.Parse(sheet.GetAttribute("sheetId")));
                var added = wb.CreateElement("sheet", NsMain);
                added.SetAttribute("name", sheetName);
                added.SetAttribute("sheetId", (maxId + 1).ToString());
                added.SetAttribute("id", NsRel, relId);
                wb.GetElementsByTagName("sheets", NsMain)[0].AppendChild(added);
                var types = new XmlDocument();
                using (var stream = zip.GetEntry("[Content_Types].xml").Open()) types.Load(stream);
                var over = types.CreateElement("Override", types.DocumentElement.NamespaceURI);
                over.SetAttribute("PartName", "/" + path);
                over.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                types.DocumentElement.AppendChild(over);
                SaveSheetXml(zip, "xl/workbook.xml", wb);
                SaveSheetXml(zip, "xl/_rels/workbook.xml.rels", rel);
                SaveSheetXml(zip, "[Content_Types].xml", types);
            }
            return path;
        }

        private static void SaveAdditionalEmployees(ZipArchive zip, List<WorkshopData> workshops)
        {
            string path = EnsureWorksheet(zip, AdditionalSheet, "xl/worksheets/ciscoEmployees.xml");
            var doc = new XmlDocument();
            doc.LoadXml("<worksheet xmlns='" + NsMain + "'><sheetData/></worksheet>");
            var props = typeof(EmployeeRecord).GetProperties();
            SetCellInlineString(doc, "A1", "MaPX", false);
            int col = 2;
            foreach (var prop in props)
                if (prop.CanWrite) SetCellInlineString(doc, TimesheetImporter.ColName(col++) + "1", prop.Name, false);
            int fundColumn = col;
            SetCellInlineString(doc, TimesheetImporter.ColName(fundColumn) + "1", "QuyLuongSPQL_v5", false);
            SetCellInlineString(doc, TimesheetImporter.ColName(fundColumn + 1) + "1", "DieuChinhQL_v5", false);
            int rowNumber = 2;
            foreach (var ws in workshops)
                foreach (var employee in ws.ListNhanVien)
                {
                    SetCellInlineString(doc, "A" + rowNumber, ws.MaPX, false);
                    col = 2;
                    foreach (var prop in props)
                        if (prop.CanWrite)
                            SetCellInlineString(doc, TimesheetImporter.ColName(col++) + rowNumber,
                                Convert.ToString(prop.GetValue(employee, null), CultureInfo.InvariantCulture), false);
                    if (employee.QuanLyTinhMoi)
                    {
                        SetCellNumeric(doc, TimesheetImporter.ColName(fundColumn) + rowNumber, ws.LuongSPQL, false);
                        SetCellNumeric(doc, TimesheetImporter.ColName(fundColumn + 1) + rowNumber, ws.DieuChinhLuongQL, false);
                    }
                    rowNumber++;
                }
            SaveSheetXml(zip, path, doc);
        }

        private static void SaveEditedGnv(ZipArchive zip, List<WorkshopData> workshops)
        {
            bool edited=false; foreach(var ws in workshops) if(ws.GnvDaChinhSua) edited=true;
            if(!edited) return;
            string path=EnsureWorksheet(zip,"GNV v5","xl/worksheets/ciscoGnv.xml");
            var doc=new XmlDocument(); doc.LoadXml("<worksheet xmlns='"+NsMain+"'><sheetData/></worksheet>");
            string[] names={"MaPX","HasItem","NoiDung","DVT","KhoiLuong","DonGia","ThanhTien","ExcelRow","GhiChu","TuTinhThanhTien"};
            for(int c=0;c<names.Length;c++) SetCellInlineString(doc,TimesheetImporter.ColName(c+1)+"1",names[c],false);
            int row=2;
            foreach(var ws in workshops)
            {
                if(!ws.GnvDaChinhSua) continue;
                if(ws.ListGNV.Count==0)
                { SetCellInlineString(doc,"A"+row,ws.MaPX,false); SetCellInlineString(doc,"B"+row,"False",false); row++; }
                foreach(var item in ws.ListGNV)
                {
                    SetCellInlineString(doc,"A"+row,ws.MaPX,false); SetCellInlineString(doc,"B"+row,"True",false);
                    SetCellInlineString(doc,"C"+row,item.NoiDung,false); SetCellInlineString(doc,"D"+row,item.DVT,false);
                    SetCellNumeric(doc,"E"+row,item.KhoiLuong,false); SetCellNumeric(doc,"F"+row,item.DonGia,false);
                    SetCellNumeric(doc,"G"+row,item.ThanhTien,false); SetCellNumeric(doc,"H"+row,item.ExcelRow,false);
                    SetCellInlineString(doc,"I"+row,item.GhiChu,false);
                    SetCellInlineString(doc,"J"+row,item.TuTinhThanhTien?"True":"False",false); row++;
                }
            }
            SaveSheetXml(zip,path,doc);
        }

        private static void LoadEditedGnv(ZipArchive zip, List<WorkshopData> workshops)
        {
            var map=GetSheetMap(zip); string path;
            if(!map.TryGetValue("GNV v5",out path)) return;
            var doc=LoadSheetXml(zip,path); var strings=LoadSharedStrings(zip);
            var seen=new HashSet<string>();
            foreach(XmlElement row in doc.GetElementsByTagName("row",NsMain))
            {
                int r=int.Parse(row.GetAttribute("r")); if(r<2) continue;
                string code=GetCellString(doc,strings,"A"+r);
                foreach(var ws in workshops)
                {
                    if(ws.MaPX!=code) continue;
                    if(seen.Add(code)) { ws.ListGNV.Clear(); ws.GnvDaChinhSua=true; }
                    if(GetCellString(doc,strings,"B"+r)!="True") continue;
                    var item=new ArisingWorkItem(ws.ListGNV.Count+1,GetCellString(doc,strings,"C"+r),GetCellString(doc,strings,"D"+r),
                        GetCellDouble(doc,strings,"E"+r,0),GetCellDouble(doc,strings,"F"+r,0),GetCellDouble(doc,strings,"G"+r,0),
                        GetCellString(doc,strings,"I"+r),(int)GetCellDouble(doc,strings,"H"+r,0))
                        { TuTinhThanhTien=GetCellString(doc,strings,"J"+r)=="True" };
                    ws.ListGNV.Add(item);
                }
            }
            foreach(var ws in workshops) if(seen.Contains(ws.MaPX)) ws.RecalculateDerived();
        }

        private static void SaveManagementProductionSettings(ZipArchive zip, List<WorkshopData> workshops)
        {
            bool enabled=false;
            foreach(var ws in workshops) if(ws.TinhQuyQLTuSanLuong) enabled=true;
            if(!enabled) return;
            string path=EnsureWorksheet(zip,"Quy QL san luong","xl/worksheets/ciscoManagementProduction.xml");
            var doc=new XmlDocument();
            doc.LoadXml("<worksheet xmlns='"+NsMain+"'><sheetData/></worksheet>");
            string[] names={"MaPX","TinhTuSanLuong","DonGiaQLLoai1","DonGiaQLLoai2","SanLuongLoai1","SanLuongLoai2","SanLuongLe"};
            for(int col=0;col<names.Length;col++) SetCellInlineString(doc,TimesheetImporter.ColName(col+1)+"1",names[col],false);
            int row=2;
            foreach(var ws in workshops)
            {
                SetCellInlineString(doc,"A"+row,ws.MaPX,false);
                SetCellInlineString(doc,"B"+row,ws.TinhQuyQLTuSanLuong?"True":"False",false);
                double[] values={ws.DonGiaQLLoai1,ws.DonGiaQLLoai2,ws.SanLuongLoai1,ws.SanLuongLoai2,ws.SanLuongLe};
                for(int col=0;col<values.Length;col++) SetCellNumeric(doc,TimesheetImporter.ColName(col+3)+row,values[col],false);
                row++;
            }
            SaveSheetXml(zip,path,doc);
        }

        private static void LoadManagementProductionSettings(ZipArchive zip, List<WorkshopData> workshops)
        {
            var map=GetSheetMap(zip); string path;
            if(!map.TryGetValue("Quy QL san luong",out path)) return;
            var doc=LoadSheetXml(zip,path); var strings=LoadSharedStrings(zip);
            foreach(XmlElement row in doc.GetElementsByTagName("row",NsMain))
            {
                int number=int.Parse(row.GetAttribute("r")); if(number<2) continue;
                string code=GetCellString(doc,strings,"A"+number);
                foreach(var ws in workshops)
                {
                    if(ws.MaPX!=code) continue;
                    ws.TinhQuyQLTuSanLuong=GetCellString(doc,strings,"B"+number)=="True";
                    ws.DonGiaQLLoai1=GetCellDouble(doc,strings,"C"+number,0);
                    ws.DonGiaQLLoai2=GetCellDouble(doc,strings,"D"+number,0);
                    // PX có ba dòng trong NT đọc sản lượng từ chính sheet NT; đơn vị khác dùng đầu vào v5.3 đã lưu.
                    if(ws.MaPX!="TK" && ws.MaPX!="LG" && ws.MaPX!="LT")
                    {
                        ws.SanLuongLoai1=GetCellDouble(doc,strings,"E"+number,0);
                        ws.SanLuongLoai2=GetCellDouble(doc,strings,"F"+number,0);
                        ws.SanLuongLe=GetCellDouble(doc,strings,"G"+number,0);
                    }
                    if(ws.TinhQuyQLTuSanLuong)
                    {
                        ManagementPayroll.Calculate(ws,true);
                        ws.RecalculateDerived();
                        ws.CanhBaoQuyQuanLy="Quỹ quản lý tự tính từ sản lượng và đơn giá quản lý đã lưu.";
                    }
                }
            }
        }

        private static void SaveManagementPayrollReport(ZipArchive zip, List<WorkshopData> workshops)
        {
            bool applied = false;
            foreach (var ws in workshops) foreach (var e in ws.ListNhanVien) if (e.QuanLyTinhMoi) applied = true;
            if (!applied) return;
            string path = EnsureWorksheet(zip, "Luong quan ly v5", "xl/worksheets/ciscoManagementPayroll.xml");
            var doc = new XmlDocument();
            doc.LoadXml("<worksheet xmlns='" + NsMain + "'><sheetData/></worksheet>");
            SetCellInlineString(doc, "A1", "LƯƠNG QUẢN LÝ PHÂN XƯỞNG — dữ liệu v5; đối chiếu quỹ đã duyệt trước thanh toán", false);
            string[] headers = { "Phân xưởng", "Tháng/năm", "Quỹ SP", "Đã phân bổ SP", "Chênh làm tròn", "Ngoài SP", "Điều chỉnh chung", "Tổng NT quản lý", "Tổng cá nhân" };
            for (int c=0; c<headers.Length; c++) SetCellInlineString(doc, TimesheetImporter.ColName(c+1) + "2", headers[c], false);
            int row = 3;
            foreach (var ws in workshops)
            {
                bool has = false; double product = 0, personal = 0;
                foreach (var e in ws.ListNhanVien) if (e.LaQuanLy)
                { has |= e.QuanLyTinhMoi; product += e.LuongSanPham; personal += e.TongTienLuong; }
                if (!has) continue;
                SetCellInlineString(doc, "A" + row, ws.TenPX, false);
                SetCellInlineString(doc, "B" + row, ws.Thang + "/" + ws.Nam, false);
                double[] values = { ws.LuongSPQL, product, ws.LuongSPQL-product, ws.LuongCheDoQL, ws.DieuChinhLuongQL, ws.TongLuongQL, personal };
                for (int c=0; c<values.Length; c++) SetCellNumeric(doc, TimesheetImporter.ColName(c+3) + row, values[c], false);
                row++;
            }
            row += 2;
            headers = new[] { "Phân xưởng", "SDB", "Họ tên", "Chức danh", "Công SP", "Hệ số", "Xếp loại", "HS hưởng", "Lương SP", "Lương phép", "Học/họp/VHTT", "Làm lễ", "Phụ cấp", "Tổng cá nhân" };
            for (int c=0; c<headers.Length; c++) SetCellInlineString(doc, TimesheetImporter.ColName(c+1) + row, headers[c], false);
            row++;
            foreach (var ws in workshops) foreach (var e in ws.ListNhanVien)
            {
                if (!e.LaQuanLy || !e.QuanLyTinhMoi) continue;
                string[] text = { ws.TenPX, e.SDB, e.HoTen, e.ChucDanh };
                for (int c=0; c<text.Length; c++) SetCellInlineString(doc, TimesheetImporter.ColName(c+1) + row, text[c] ?? "", false);
                SetCellNumeric(doc, "E"+row, e.CongSP, false); SetCellNumeric(doc, "F"+row, e.HeSoCoDinh, false);
                SetCellInlineString(doc, "G"+row, e.XepLoai, false);
                double[] values = { e.HeSoHuong, e.LuongSanPham, e.LuongLePhepCD, e.LuongHocVHTT, e.LuongLamLe, e.PhuCapTien, e.TongTienLuong };
                for (int c=0; c<values.Length; c++) SetCellNumeric(doc, TimesheetImporter.ColName(c+8) + row, values[c], false);
                row++;
            }
            SaveSheetXml(zip, path, doc);
        }

        private static void LoadAdditionalEmployees(ZipArchive zip, List<WorkshopData> workshops)
        {
            var map = GetSheetMap(zip);
            if (!map.ContainsKey(AdditionalSheet)) return;
            var doc = LoadSheetXml(zip, map[AdditionalSheet]);
            var strings = LoadSharedStrings(zip);
            foreach (XmlElement row in doc.GetElementsByTagName("row", NsMain))
            {
                int r = int.Parse(row.GetAttribute("r"));
                if (r < 2) continue;
                string unit = GetCellString(doc, strings, "A" + r);
                var employee = new EmployeeRecord();
                for (int col = 2; col <= typeof(EmployeeRecord).GetProperties().Length + 1; col++)
                {
                    string heading = GetCellString(doc, strings, TimesheetImporter.ColName(col) + "1");
                    var prop = typeof(EmployeeRecord).GetProperty(heading);
                    if (prop == null || !prop.CanWrite) continue;
                    string value = GetCellString(doc, strings, TimesheetImporter.ColName(col) + r);
                    if (value == "" && prop.PropertyType != typeof(string)) continue;
                    prop.SetValue(employee, Convert.ChangeType(value, prop.PropertyType, CultureInfo.InvariantCulture), null);
                }
                foreach (var ws in workshops)
                    if (ws.MaPX == unit)
                    {
                        if (employee.ExcelRow <= 0)
                        {
                            employee.ExcelRow = employee.TongHopCongRow = 0;
                            ws.ListNhanVien.Add(employee);
                        }
                        else
                            foreach (var original in ws.ListNhanVien)
                                if (original.ExcelRow == employee.ExcelRow && original.SDB == employee.SDB && original.HoTen == employee.HoTen)
                                {
                                    // Bản v4 có thể kéo nhầm cả tổ SX vào QL khi tiêu đề tổ có chức danh/SDB=0.
                                    // Khi snapshot cũ mâu thuẫn với phân nhóm đã đọc đúng từ sheet gốc,
                                    // giữ nhóm của sheet thay vì khôi phục lại lỗi cũ.
                                    if (employee.QuanLyTinhMoi || original.LaQuanLy == employee.LaQuanLy)
                                    {
                                        original.ToSanXuat = employee.ToSanXuat;
                                        original.BoPhan = employee.BoPhan;
                                        original.LaQuanLy = employee.LaQuanLy;
                                    }
                                    if (employee.QuanLyTinhMoi)
                                        ManagementPayroll.Copy(employee, original);
                                }
                        if (employee.QuanLyTinhMoi)
                        {
                            int fundCol = 2;
                            foreach (var property in typeof(EmployeeRecord).GetProperties())
                                if (property.CanWrite) fundCol++;
                            ws.LuongSPQL = GetCellDouble(doc, strings, TimesheetImporter.ColName(fundCol) + r, ws.LuongSPQL);
                            ws.DieuChinhLuongQL = GetCellDouble(doc, strings, TimesheetImporter.ColName(fundCol + 1) + r, ws.DieuChinhLuongQL);
                            ws.CanhBaoQuyQuanLy = "Quỹ quản lý lấy từ dữ liệu v5 đã lưu. Đối chiếu lại quỹ được duyệt cho kỳ hiện tại.";
                        }
                    }
            }
            foreach (var ws in workshops)
            {
                bool calculated = false;
                foreach (var e in ws.ListNhanVien) if (e.QuanLyTinhMoi) calculated = true;
                if (calculated) { ManagementPayroll.Calculate(ws, false); ws.RecalculateDerived(); }
            }
        }

        private static void SaveWorkshopEmployeesToSheet(XmlDocument empDoc, XmlDocument thcDoc, WorkshopData ws)
        {
            if (empDoc == null || ws.ListNhanVien == null || ws.ListNhanVien.Count == 0) return;

            bool isCDNL = (ws.MaPX == "CDNL");
            double sumJ = 0, sumK = 0, sumM = 0, sumN = 0, sumO = 0, sumP = 0, sumQ = 0, sumR = 0, sumS = 0, sumT = 0;
            double sumW = 0, sumZ = 0, sumAA = 0, sumAB = 0, sumAD = 0, sumAH = 0, sumAI = 0, sumAK = 0;

            foreach (var emp in ws.ListNhanVien)
            {
                int r = emp.ExcelRow;
                if (r <= 0) continue;

                // Cập nhật trong sheet phân xưởng ('T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển')
                // Giữ nguyên 100% công thức (<f>) bằng keepFormula = true
                if (emp.MucLuongBHXH > 0)
                    SetCellNumeric(empDoc, "G" + r, emp.MucLuongBHXH, true);
                if (emp.PhuCapChucVu > 0 || emp.QuanLyTinhMoi)
                    SetCellNumeric(empDoc, "H" + r, emp.PhuCapChucVu, true);
                if (emp.QuanLyTinhMoi) SetCellNumeric(empDoc, "I" + r, emp.HeSoCoDinh, true);

                SetCellNumeric(empDoc, "J" + r, emp.TongCong, true);
                SetCellNumeric(empDoc, "K" + r, emp.CongSP, true);
                SetCellNumeric(empDoc, "M" + r, emp.CongLePhep, true);
                SetCellNumeric(empDoc, "N" + r, emp.CongLePhep, true);
                if (emp.CongViecRieng > 0)
                    SetCellNumeric(empDoc, "O" + r, emp.CongViecRieng, true);
                SetCellNumeric(empDoc, "P" + r, emp.CongHocHop, true);
                if (emp.CongLamLe > 0 || emp.QuanLyTinhMoi)
                    SetCellNumeric(empDoc, "Q" + r, emp.CongLamLe, true);
                SetCellNumeric(empDoc, "R" + r, emp.CongVHTT, true);
                SetCellNumeric(empDoc, "S" + r, emp.CongBHXH, true);

                if (!isCDNL)
                {
                    SetCellNumeric(empDoc, "T" + r, emp.CongAnCa, true);
                    if (emp.DiemSanXuat > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "W" + r, emp.DiemSanXuat, true);
                    if (!string.IsNullOrEmpty(emp.XepLoai))
                        SetCellInlineString(empDoc, "X" + r, emp.XepLoai, true);
                    if (emp.HeSoHuong > 0 || emp.QuanLyTinhMoi)
                    {
                        SetCellNumeric(empDoc, "Y" + r, emp.HeSoHuong, true);
                        SetCellNumeric(empDoc, "Z" + r, emp.HeSoHuong, true);
                    }
                    SetCellNumeric(empDoc, "AA" + r, emp.TongTienLuong, true);
                    SetCellNumeric(empDoc, "AB" + r, emp.LuongSanPham, true);
                    SetCellNumeric(empDoc, "AD" + r, emp.LuongLePhepCD, true);
                    if (emp.LuongLamLe > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "AH" + r, emp.LuongLamLe, true);
                    if (emp.LuongHocVHTT > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "AI" + r, emp.LuongHocVHTT, true);
                    if (emp.PhuCapTien > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "AK" + r, emp.PhuCapTien, true);
                }
                else
                {
                    SetCellNumeric(empDoc, "T" + r, emp.CongAnCa, true);
                    SetCellNumeric(empDoc, "V" + r, emp.CongAnCa, true);
                    if (emp.DiemSanXuat > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "Y" + r, emp.DiemSanXuat, true);
                    if (!string.IsNullOrEmpty(emp.XepLoai))
                        SetCellInlineString(empDoc, "Z" + r, emp.XepLoai, true);
                    if (emp.HeSoHuong > 0 || emp.QuanLyTinhMoi)
                    {
                        SetCellNumeric(empDoc, "AA" + r, emp.HeSoHuong, true);
                        SetCellNumeric(empDoc, "AB" + r, emp.HeSoHuong, true);
                    }
                    SetCellNumeric(empDoc, "AC" + r, emp.TongTienLuong, true);
                    SetCellNumeric(empDoc, "AD" + r, emp.LuongSanPham, true);
                    SetCellNumeric(empDoc, "AE" + r, emp.LuongLePhepCD, true);
                    if (emp.LuongLamLe > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "AI" + r, emp.LuongLamLe, true);
                    if (emp.LuongHocVHTT > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "AJ" + r, emp.LuongHocVHTT, true);
                    if (emp.PhuCapTien > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(empDoc, "AL" + r, emp.PhuCapTien, true);
                }

                // Đồng thời cập nhật dòng tương ứng trong sheet 'Tổng hợp công' (nếu có) để các công thức INDEX/MATCH/SUMIF tự động khớp 100%!
                if (thcDoc != null && emp.TongHopCongRow > 0)
                {
                    int tr = emp.TongHopCongRow;
                    if (emp.MucLuongBHXH > 0)
                        SetCellNumeric(thcDoc, "K" + tr, emp.MucLuongBHXH, true);
                    SetCellNumeric(thcDoc, "L" + tr, emp.CongSP, true);
                    SetCellNumeric(thcDoc, "M" + tr, emp.CongHocHop, true);
                    SetCellNumeric(thcDoc, "N" + tr, emp.CongVHTT, true);
                    SetCellNumeric(thcDoc, "O" + tr, emp.CongLePhep, true);
                    SetCellNumeric(thcDoc, "R" + tr, emp.CongHocHop + emp.CongVHTT + emp.CongLePhep, true);
                    SetCellNumeric(thcDoc, "T" + tr, emp.CongBHXH, true);
                    if (emp.DiemSanXuat > 0 || emp.QuanLyTinhMoi)
                        SetCellNumeric(thcDoc, "U" + tr, emp.DiemSanXuat, true);
                    if (emp.LuongSanPham > 0)
                        SetCellNumeric(thcDoc, "W" + tr, emp.LuongSanPham, true);
                    if (emp.LuongLePhepCD > 0)
                        SetCellNumeric(thcDoc, "AB" + tr, emp.LuongLePhepCD, true);
                    SetCellNumeric(thcDoc, "AF" + tr, emp.CongAnCa, true);
                }

                sumJ += emp.TongCong;
                sumK += emp.CongSP;
                sumM += emp.CongLePhep;
                sumN += emp.CongLePhep;
                sumO += emp.CongViecRieng;
                sumP += emp.CongHocHop;
                sumQ += emp.CongLamLe;
                sumR += emp.CongVHTT;
                sumS += emp.CongBHXH;
                sumT += emp.CongAnCa;
                if (!emp.LaQuanLy)
                {
                    sumW += emp.DiemSanXuat;
                    sumZ += emp.HeSoHuong;
                }
                sumAA += emp.TongTienLuong;
                sumAB += emp.LuongSanPham;
                sumAD += emp.LuongLePhepCD;
                sumAH += emp.LuongLamLe;
                sumAI += emp.LuongHocVHTT;
                sumAK += emp.PhuCapTien;
            }

            // Cập nhật dòng CỘNG của sheet phân xưởng (giữ nguyên 100% thẻ <f> SUM(...))
            int totRow = ws.TotalRowBangCong;
            if (totRow > 0)
            {
                SetCellNumeric(empDoc, "J" + totRow, sumJ, true);
                SetCellNumeric(empDoc, "K" + totRow, sumK, true);
                SetCellNumeric(empDoc, "M" + totRow, sumM, true);
                SetCellNumeric(empDoc, "N" + totRow, sumN, true);
                SetCellNumeric(empDoc, "P" + totRow, sumP, true);
                SetCellNumeric(empDoc, "R" + totRow, sumR, true);
                SetCellNumeric(empDoc, "S" + totRow, sumS, true);
                if (!isCDNL)
                {
                    SetCellNumeric(empDoc, "T" + totRow, sumT, true);
                    if (sumW > 0) SetCellNumeric(empDoc, "W" + totRow, sumW, true);
                    if (sumZ > 0) SetCellNumeric(empDoc, "Z" + totRow, sumZ, true);
                    SetCellNumeric(empDoc, "AA" + totRow, sumAA, true);
                    SetCellNumeric(empDoc, "AB" + totRow, sumAB, true);
                    SetCellNumeric(empDoc, "AD" + totRow, sumAD, true);
                }
                else
                {
                    SetCellNumeric(empDoc, "V" + totRow, sumT, true);
                    if (sumW > 0) SetCellNumeric(empDoc, "Y" + totRow, sumW, true);
                    if (sumZ > 0) SetCellNumeric(empDoc, "AB" + totRow, sumZ, true);
                    SetCellNumeric(empDoc, "AC" + totRow, sumAA, true);
                    SetCellNumeric(empDoc, "AD" + totRow, sumAB, true);
                    SetCellNumeric(empDoc, "AE" + totRow, sumAD, true);
                }
            }
        }

        private static void UpdateNTManagementRows(XmlDocument ntDoc, List<string> strings, WorkshopData ws, int startRow, int endRow)
        {
            if (ntDoc == null || ws.ListNhanVien == null) return;
            for (int r = startRow; r <= endRow; r++)
            {
                string name = Regex.Replace(GetCellString(ntDoc, strings, "B" + r).Trim(), @"\s+", " ");
                EmployeeRecord match = null;
                bool ambiguous = false;
                foreach (var e in ws.ListNhanVien)
                    if (e.LaQuanLy && string.Equals(name,
                        Regex.Replace((e.HoTen ?? "").Trim(), @"\s+", " "), StringComparison.OrdinalIgnoreCase))
                    {
                        if (match != null) ambiguous = true;
                        match = e;
                    }
                // Không ghép theo thứ tự dòng: mẫu NT có người nghỉ, thiếu người hoặc #REF!.
                if (match == null || ambiguous) continue;
                SetCellNumeric(ntDoc, "E" + r, match.CongSP, true);
                if (!string.IsNullOrEmpty(match.XepLoai))
                    SetCellInlineString(ntDoc, "F" + r, match.XepLoai, true);
                SetCellNumeric(ntDoc, "G" + r, match.HeSoHuong, true);
            }
        }

        private static void EnsureFullCalcOnLoad(ZipArchive zip)
        {
            var wbEntry = zip.GetEntry("xl/workbook.xml");
            if (wbEntry == null) return;
            var doc = new XmlDocument();
            using (var s = wbEntry.Open()) doc.Load(s);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("d", NsMain);
            var calcPr = doc.SelectSingleNode("//d:calcPr", ns) as XmlElement;
            if (calcPr != null)
            {
                calcPr.SetAttribute("fullCalcOnLoad", "1");
                SaveSheetXml(zip, "xl/workbook.xml", doc);
            }
        }

        private static void UpdateCumulativeSheetStandard(ZipArchive zip, string path, WorkshopData ws, int rTongCong)
        {
            var doc = LoadSheetXml(zip, path);
            if (doc == null) return;
            // Cột G = Quý III, Cột H = 9 tháng (Ngày công)
            SetCellNumeric(doc, "G" + rTongCong, ws.KyQuy.TongCong);
            SetCellNumeric(doc, "H" + rTongCong, ws.KyLuyKe.TongCong);
            SetCellNumeric(doc, "G" + (rTongCong + 1), ws.KyQuy.CongSP);
            SetCellNumeric(doc, "H" + (rTongCong + 1), ws.KyLuyKe.CongSP);
            SetCellNumeric(doc, "G" + (rTongCong + 2), ws.KyQuy.CongQL);
            SetCellNumeric(doc, "H" + (rTongCong + 2), ws.KyLuyKe.CongQL);
            SetCellNumeric(doc, "G" + (rTongCong + 3), ws.KyQuy.CongLDTT);
            SetCellNumeric(doc, "H" + (rTongCong + 3), ws.KyLuyKe.CongLDTT);
            SetCellNumeric(doc, "G" + (rTongCong + 4), ws.KyQuy.CongCheDoBHXH);
            SetCellNumeric(doc, "H" + (rTongCong + 4), ws.KyLuyKe.CongCheDoBHXH);

            int rLuongTT = rTongCong == 8 ? 21 : 22;
            int rLuongSP = rTongCong == 8 ? 27 : 28;
            int rLuongQL = rTongCong == 8 ? 30 : 31;
            int rTongLuong = rTongCong == 8 ? 33 : 34;

            SetCellNumeric(doc, "N" + rLuongTT, ws.KyQuy.TongLuongTT);
            SetCellNumeric(doc, "O" + rLuongTT, ws.KyLuyKe.TongLuongTT);
            SetCellNumeric(doc, "N" + rLuongSP, ws.KyQuy.LuongSPTT);
            SetCellNumeric(doc, "O" + rLuongSP, ws.KyLuyKe.LuongSPTT);
            SetCellNumeric(doc, "N" + (rLuongSP + 1), ws.KyQuy.LuongCVKhac);
            SetCellNumeric(doc, "O" + (rLuongSP + 1), ws.KyLuyKe.LuongCVKhac);
            SetCellNumeric(doc, "N" + (rLuongSP + 2), ws.KyQuy.LuongCheDoTT);
            SetCellNumeric(doc, "O" + (rLuongSP + 2), ws.KyLuyKe.LuongCheDoTT);

            SetCellNumeric(doc, "N" + rLuongQL, ws.KyQuy.TongLuongQL);
            SetCellNumeric(doc, "O" + rLuongQL, ws.KyLuyKe.TongLuongQL);
            SetCellNumeric(doc, "N" + (rLuongQL + 1), ws.KyQuy.LuongSPQL);
            SetCellNumeric(doc, "O" + (rLuongQL + 1), ws.KyLuyKe.LuongSPQL);
            SetCellNumeric(doc, "N" + (rLuongQL + 2), ws.KyQuy.LuongCheDoQL);
            SetCellNumeric(doc, "O" + (rLuongQL + 2), ws.KyLuyKe.LuongCheDoQL);

            SetCellNumeric(doc, "N" + rTongLuong, ws.KyQuy.TongLuongToanPX);
            SetCellNumeric(doc, "O" + rTongLuong, ws.KyLuyKe.TongLuongToanPX);
            SetCellNumeric(doc, "U" + rTongLuong, ws.KyQuy.TienAnCa);
            SetCellNumeric(doc, "V" + rTongLuong, ws.KyLuyKe.TienAnCa);
            SetCellNumeric(doc, "AB" + rTongLuong, ws.KyQuy.TongTienBDDH);
            SetCellNumeric(doc, "AC" + rTongLuong, ws.KyLuyKe.TongTienBDDH);

            SetCellNumeric(doc, "G" + (rTongLuong + 1), ws.KyQuy.TLBinhQuanToanPX);
            SetCellNumeric(doc, "H" + (rTongLuong + 1), ws.KyLuyKe.TLBinhQuanToanPX);
            SetCellNumeric(doc, "G" + (rTongLuong + 2), ws.KyQuy.ThuNhapBQToanPX);
            SetCellNumeric(doc, "H" + (rTongLuong + 2), ws.KyLuyKe.ThuNhapBQToanPX);
            SetCellNumeric(doc, "G" + (rTongLuong + 3), ws.KyQuy.TLSPBinhQuanLDTT);
            SetCellNumeric(doc, "H" + (rTongLuong + 3), ws.KyLuyKe.TLSPBinhQuanLDTT);
            SetCellNumeric(doc, "G" + (rTongLuong + 4), ws.KyQuy.ThuNhapBQLDTT);
            SetCellNumeric(doc, "H" + (rTongLuong + 4), ws.KyLuyKe.ThuNhapBQLDTT);

            SaveSheetXml(zip, path, doc);
        }

        private static void UpdateCumulativeSheetCDNL(ZipArchive zip, string path, WorkshopData ws)
        {
            var doc = LoadSheetXml(zip, path);
            if (doc == null) return;
            SetCellNumeric(doc, "H10", ws.KyQuy.TongCong);
            SetCellNumeric(doc, "I10", ws.KyLuyKe.TongCong);
            SetCellNumeric(doc, "H11", ws.KyQuy.CongSP);
            SetCellNumeric(doc, "I11", ws.KyLuyKe.CongSP);
            SetCellNumeric(doc, "H12", ws.KyQuy.CongQL);
            SetCellNumeric(doc, "I12", ws.KyLuyKe.CongQL);
            SetCellNumeric(doc, "H13", ws.KyQuy.CongLDTT);
            SetCellNumeric(doc, "I13", ws.KyLuyKe.CongLDTT);
            SetCellNumeric(doc, "H14", ws.KyQuy.CongCheDoBHXH);
            SetCellNumeric(doc, "I14", ws.KyLuyKe.CongCheDoBHXH);

            SetCellNumeric(doc, "N22", ws.KyQuy.TongLuongTT);
            SetCellNumeric(doc, "P22", ws.KyLuyKe.TongLuongTT);
            SetCellNumeric(doc, "N28", ws.KyQuy.LuongSPTT);
            SetCellNumeric(doc, "P28", ws.KyLuyKe.LuongSPTT);
            SetCellNumeric(doc, "N29", ws.KyQuy.LuongCVKhac);
            SetCellNumeric(doc, "P29", ws.KyLuyKe.LuongCVKhac);
            SetCellNumeric(doc, "N30", ws.KyQuy.LuongCheDoTT);
            SetCellNumeric(doc, "P30", ws.KyLuyKe.LuongCheDoTT);

            SetCellNumeric(doc, "N31", ws.KyQuy.TongLuongQL);
            SetCellNumeric(doc, "P31", ws.KyLuyKe.TongLuongQL);
            SetCellNumeric(doc, "N32", ws.KyQuy.LuongSPQL);
            SetCellNumeric(doc, "P32", ws.KyLuyKe.LuongSPQL);
            SetCellNumeric(doc, "N33", ws.KyQuy.LuongCheDoQL);
            SetCellNumeric(doc, "P33", ws.KyLuyKe.LuongCheDoQL);

            SetCellNumeric(doc, "N34", ws.KyQuy.TongLuongToanPX);
            SetCellNumeric(doc, "P34", ws.KyLuyKe.TongLuongToanPX);
            SetCellNumeric(doc, "S34", ws.KyQuy.TienAnCa);
            SetCellNumeric(doc, "U34", ws.KyLuyKe.TienAnCa);
            SetCellNumeric(doc, "X34", ws.KyQuy.TongTienBDDH);
            SetCellNumeric(doc, "Z34", ws.KyLuyKe.TongTienBDDH);

            SetCellNumeric(doc, "H35", ws.KyQuy.TLBinhQuanToanPX);
            SetCellNumeric(doc, "I35", ws.KyLuyKe.TLBinhQuanToanPX);
            SetCellNumeric(doc, "H36", ws.KyQuy.ThuNhapBQToanPX);
            SetCellNumeric(doc, "I36", ws.KyLuyKe.ThuNhapBQToanPX);
            SetCellNumeric(doc, "H37", ws.KyQuy.TLSPBinhQuanLDTT);
            SetCellNumeric(doc, "I37", ws.KyLuyKe.TLSPBinhQuanLDTT);
            SetCellNumeric(doc, "H38", ws.KyQuy.ThuNhapBQLDTT);
            SetCellNumeric(doc, "I38", ws.KyLuyKe.ThuNhapBQLDTT);

            SaveSheetXml(zip, path, doc);
        }

        private static void WriteGNVRows(XmlDocument doc, List<ArisingWorkItem> items, int startRow, int maxRow, string colKL, string colDG, string colTT)
        {
            if (items == null || items.Count == 0) return;
            int idx = 0;
            foreach (var item in items)
            {
                int r = item.ExcelRow > 0 ? item.ExcelRow : (startRow + idx);
                if (r > maxRow) break;
                SetCellInlineString(doc, "B" + r, item.NoiDung, false);
                SetCellInlineString(doc, "C" + r, string.IsNullOrEmpty(item.DVT) ? "Công" : item.DVT, false);
                SetCellNumeric(doc, colKL + r, item.KhoiLuong, true);
                if (item.DonGia > 0 || item.TuTinhThanhTien) SetCellNumeric(doc, colDG + r, item.DonGia, true);
                SetCellNumeric(doc, colTT + r, item.ThanhTien, true);
                idx++;
            }
        }

        private static void SaveSheetXml(ZipArchive zip, string entryPath, XmlDocument doc)
        {
            var oldEntry = zip.GetEntry(entryPath);
            if (oldEntry != null) oldEntry.Delete();
            var newEntry = zip.CreateEntry(entryPath, CompressionLevel.Optimal);
            using (var s = newEntry.Open())
            {
                doc.Save(s);
            }
        }

        /// <summary>
        /// Nạp dữ liệu từ file Excel bên ngoài:
        /// 1. Nếu là file tổng hợp có sheet NT hoặc sheet Bảng công phân xưởng ('T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển'): nạp cả chi tiết từng người và chỉ tiêu tổng hợp.
        /// 2. Nếu là file Bảng chấm công / Ăn ca / Điểm sản xuất từng người rời: tự động dò khớp theo Mã số SDB hoặc Họ và tên từng người trong ws.ListNhanVien,
        ///    cập nhật Công SP, Công phép lễ, Công học, Công BHXH, Công ăn ca, Điểm SX, Xếp loại và tự động cộng tổng lên Bảng nghiệm thu!
        /// </summary>
        public static string ImportExternalExcelToWorkshop(string externalFilePath, WorkshopData ws, string masterExcelPath = null, List<WorkshopData> workshops = null, Func<EmployeeRecord, EmployeeRecord> requestEmployee = null)
        {
            if (!File.Exists(externalFilePath))
                throw new FileNotFoundException("Không tìm thấy file: " + externalFilePath);

            var log = new StringBuilder();

            // ƯU TIÊN: file công phân xưởng (Bảng công / QT ăn ca na-tt / Bảng BDĐH M1-M2-M3 / Bảng điểm)
            // -> trích từng người theo chuẩn HỌ TÊN + SỐ DANH BẠ (SDB), rồi cộng lên Tổ -> PX -> Bảng nghiệm thu.
            string tsLog = TimesheetImporter.Import(externalFilePath, ws, workshops, requestEmployee);
            if (!tsLog.StartsWith("- Không nhận diện được sheet"))
                return tsLog;

            using (var fs = new FileStream(externalFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var sheetMap = GetSheetMap(zip);
                var ss = LoadSharedStrings(zip);

                if (sheetMap.ContainsKey(ws.SheetNT) || sheetMap.ContainsKey(ws.SheetBangCong))
                {
                    var allLoaded = LoadWorkshopsFromExcel(externalFilePath);
                    foreach (var loaded in allLoaded)
                    {
                        if (loaded.MaPX == ws.MaPX)
                        {
                            CopyWorkshopValues(loaded, ws);
                            log.AppendLine(string.Format("- Đã nạp đầy đủ danh sách chi tiết {0} người lao động (sheet '{1}') và bảng nghiệm thu '{2}'.", ws.ListNhanVien.Count, ws.SheetBangCong, ws.SheetNT));
                            break;
                        }
                    }
                    return log.ToString();
                }

                // Dò khớp danh sách nhân viên theo SDB hoặc Họ và tên
                var empBySdb = new Dictionary<string, EmployeeRecord>(StringComparer.OrdinalIgnoreCase);
                var empByName = new Dictionary<string, EmployeeRecord>(StringComparer.OrdinalIgnoreCase);
                if (ws.ListNhanVien != null)
                {
                    foreach (var e in ws.ListNhanVien)
                    {
                        string sdbKey = NormalizeSDB(e.SDB);
                        if (!string.IsNullOrEmpty(sdbKey) && !empBySdb.ContainsKey(sdbKey))
                            empBySdb[sdbKey] = e;
                        string nameKey = NormalizeKey(e.HoTen);
                        if (!string.IsNullOrEmpty(nameKey) && !empByName.ContainsKey(nameKey))
                            empByName[nameKey] = e;
                    }
                }

                int matchedEmployees = 0;
                int importedGNV = 0;

                foreach (var kv in sheetMap)
                {
                    var doc = LoadSheetXml(zip, kv.Value);
                    if (doc == null) continue;

                    var ns = new XmlNamespaceManager(doc.NameTable);
                    ns.AddNamespace("d", NsMain);
                    var rows = doc.SelectNodes("//d:row", ns);
                    if (rows == null) continue;

                    foreach (XmlNode rNode in rows)
                    {
                        var rEl = rNode as XmlElement;
                        if (rEl == null) continue;
                        int rNum;
                        if (!int.TryParse(rEl.GetAttribute("r"), out rNum)) continue;

                        // Kiểm tra xem dòng này có khớp với Nhân viên nào trong phân xưởng không
                        // Hỗ trợ cả cấu trúc bảng 'Tổng hợp công' (C=SDB, D=Họ tên) và cấu trúc sheet PX (C=Họ tên, D=SDB)
                        string cellB = GetCellString(doc, ss, "B" + rNum).Trim();
                        string cellC = GetCellString(doc, ss, "C" + rNum).Trim();
                        string cellD = GetCellString(doc, ss, "D" + rNum).Trim();

                        EmployeeRecord matchedEmp = null;
                        bool isTHCFormat = false;

                        string sdbTry1 = NormalizeSDB(cellD);
                        string sdbTry2 = NormalizeSDB(cellC);
                        string sdbTry3 = NormalizeSDB(cellB);

                        if (!string.IsNullOrEmpty(sdbTry1) && empBySdb.ContainsKey(sdbTry1))
                        {
                            matchedEmp = empBySdb[sdbTry1];
                            isTHCFormat = false;
                        }
                        else if (!string.IsNullOrEmpty(sdbTry2) && empBySdb.ContainsKey(sdbTry2))
                        {
                            matchedEmp = empBySdb[sdbTry2];
                            isTHCFormat = true;
                        }
                        else if (!string.IsNullOrEmpty(sdbTry3) && empBySdb.ContainsKey(sdbTry3))
                        {
                            matchedEmp = empBySdb[sdbTry3];
                            isTHCFormat = false;
                        }
                        else
                        {
                            string nameTryC = NormalizeKey(cellC);
                            string nameTryD = NormalizeKey(cellD);
                            string nameTryB = NormalizeKey(cellB);
                            if (!string.IsNullOrEmpty(nameTryC) && empByName.ContainsKey(nameTryC))
                                matchedEmp = empByName[nameTryC];
                            else if (!string.IsNullOrEmpty(nameTryD) && empByName.ContainsKey(nameTryD))
                            {
                                matchedEmp = empByName[nameTryD];
                                isTHCFormat = true;
                            }
                            else if (!string.IsNullOrEmpty(nameTryB) && empByName.ContainsKey(nameTryB))
                                matchedEmp = empByName[nameTryB];
                        }

                        if (matchedEmp != null)
                        {
                            if (isTHCFormat)
                            {
                                // Cấu trúc giống sheet 'Tổng hợp công': K=Lương BHXH, L=Công SP, M=Công học, N=VHTT, O=Lễ phép, T=BHXH, U=Điểm SX, AF=Ăn ca
                                double luongBH = GetCellDouble(doc, ss, "K" + rNum, -1);
                                double congSP = GetCellDouble(doc, ss, "L" + rNum, -1);
                                double congHoc = GetCellDouble(doc, ss, "M" + rNum, -1);
                                double congVH = GetCellDouble(doc, ss, "N" + rNum, -1);
                                double congPL = GetCellDouble(doc, ss, "O" + rNum, -1);
                                double congBH = GetCellDouble(doc, ss, "T" + rNum, -1);
                                double diem = GetCellDouble(doc, ss, "U" + rNum, -1);
                                double anCa = GetCellDouble(doc, ss, "AF" + rNum, -1);

                                if (luongBH > 0) matchedEmp.MucLuongBHXH = luongBH;
                                if (congSP >= 0) matchedEmp.CongSP = congSP;
                                if (congHoc >= 0) matchedEmp.CongHocHop = congHoc;
                                if (congVH >= 0) matchedEmp.CongVHTT = congVH;
                                if (congPL >= 0) matchedEmp.CongLePhep = congPL;
                                if (congBH >= 0) matchedEmp.CongBHXH = congBH;
                                if (diem > 0) matchedEmp.DiemSanXuat = diem;
                                if (anCa >= 0) matchedEmp.CongAnCa = anCa;
                            }
                            else
                            {
                                // Cấu trúc sheet phân xưởng ('T kết', 'Gang', 'thép', 'CĐNL') hoặc mẫu xuất từ phần mềm
                                double congSP = GetCellDouble(doc, ss, "K" + rNum, -1);
                                double congPL = GetCellDouble(doc, ss, "M" + rNum, -1);
                                double congHoc = GetCellDouble(doc, ss, "P" + rNum, -1);
                                double congBH = GetCellDouble(doc, ss, "S" + rNum, -1);
                                double anCa = GetCellDouble(doc, ss, "T" + rNum, -1);
                                double diem = GetCellDouble(doc, ss, "W" + rNum, -1);
                                string xl = GetCellString(doc, ss, "X" + rNum).Trim();

                                if (congSP >= 0)
                                {
                                    matchedEmp.CongSP = congSP;
                                    if (congPL >= 0) matchedEmp.CongLePhep = congPL;
                                    if (congHoc >= 0) matchedEmp.CongHocHop = congHoc;
                                    if (congBH >= 0) matchedEmp.CongBHXH = congBH;
                                    if (anCa >= 0) matchedEmp.CongAnCa = anCa;
                                    if (diem > 0) matchedEmp.DiemSanXuat = diem;
                                    if (xl == "A" || xl == "B" || xl == "C") matchedEmp.XepLoai = xl;
                                }
                                else
                                {
                                    // Cấu trúc bảng đơn giản: E=Công SP, F=Công phép lễ, G=Công học, H=Công BHXH, I=Công ăn ca, J=Điểm SX, K=Xếp loại
                                    double cSPSimple = GetCellDouble(doc, ss, "E" + rNum, -1);
                                    double cPLSimple = GetCellDouble(doc, ss, "F" + rNum, -1);
                                    double cHocSimple = GetCellDouble(doc, ss, "G" + rNum, -1);
                                    double cBHSimple = GetCellDouble(doc, ss, "H" + rNum, -1);
                                    double cAnCaSimple = GetCellDouble(doc, ss, "I" + rNum, -1);
                                    double diemSimple = GetCellDouble(doc, ss, "J" + rNum, -1);
                                    string xlSimple = GetCellString(doc, ss, "K" + rNum).Trim();

                                    if (cSPSimple >= 0) matchedEmp.CongSP = cSPSimple;
                                    if (cPLSimple >= 0) matchedEmp.CongLePhep = cPLSimple;
                                    if (cHocSimple >= 0) matchedEmp.CongHocHop = cHocSimple;
                                    if (cBHSimple >= 0) matchedEmp.CongBHXH = cBHSimple;
                                    if (cAnCaSimple >= 0) matchedEmp.CongAnCa = cAnCaSimple;
                                    if (diemSimple > 0) matchedEmp.DiemSanXuat = diemSimple;
                                    if (xlSimple == "A" || xlSimple == "B" || xlSimple == "C") matchedEmp.XepLoai = xlSimple;
                                }
                            }
                            matchedEmployees++;
                            continue;
                        }

                        string combined = (cellB + " " + cellC).ToLowerInvariant();
                        if (combined.Contains("tổng công") || combined.Contains("tổng số công"))
                        {
                            double v = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (v > 0) { ws.TongCong = v; log.AppendLine(string.Format("- Tổng công: {0:N1}", v)); }
                        }
                        else if (combined.Contains("công sản phẩm"))
                        {
                            double v = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (v > 0) { ws.CongSP = v; log.AppendLine(string.Format("- Công sản phẩm: {0:N1}", v)); }
                        }
                        else if (combined.Contains("lao động quản lý") && combined.Contains("công"))
                        {
                            double v = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (v > 0) { ws.CongQL = v; log.AppendLine(string.Format("- Công quản lý: {0:N1}", v)); }
                        }
                        else if (combined.Contains("ăn ca") || combined.Contains("ăn giữa ca"))
                        {
                            double c = GetCellDouble(doc, ss, "D" + rNum, 0);
                            double t = GetCellDouble(doc, ss, "E" + rNum, 0);
                            if (c > 0) { ws.CongAnCa = c; ws.TienAnCa = t > 0 ? t : c * 30000; log.AppendLine(string.Format("- Ăn ca: {0:N1} công ({1:N0} đ)", c, ws.TienAnCa)); }
                        }
                        else if (combined.Contains("mức 1") && (combined.Contains("13.000") || combined.Contains("13000") || combined.Contains("độc hại")))
                        {
                            double c = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (c > 0) { ws.BDDH_Muc1_Cong = c; log.AppendLine(string.Format("- BDĐH Mức 1: {0:N1} công", c)); }
                        }
                        else if (combined.Contains("mức 2") && (combined.Contains("20.000") || combined.Contains("20000") || combined.Contains("độc hại")))
                        {
                            double c = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (c > 0) { ws.BDDH_Muc2_Cong = c; log.AppendLine(string.Format("- BDĐH Mức 2: {0:N1} công", c)); }
                        }
                        else if (combined.Contains("mức 3") && (combined.Contains("26.000") || combined.Contains("26000") || combined.Contains("độc hại")))
                        {
                            double c = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (c > 0) { ws.BDDH_Muc3_Cong = c; log.AppendLine(string.Format("- BDĐH Mức 3: {0:N1} công", c)); }
                        }
                        else if (combined.Contains("gnv") || combined.Contains("giấy giao nhiệm vụ") || combined.Contains("nhân công"))
                        {
                            double kl = GetCellDouble(doc, ss, "D" + rNum, 0);
                            if (kl == 0) kl = GetCellDouble(doc, ss, "E" + rNum, 0);
                            double dg = GetCellDouble(doc, ss, "F" + rNum, 0);
                            if (dg == 0) dg = GetCellDouble(doc, ss, "I" + rNum, 0);
                            double tt = GetCellDouble(doc, ss, "G" + rNum, 0);
                            if (tt == 0) tt = GetCellDouble(doc, ss, "J" + rNum, 0);
                            if (kl > 0 || tt > 0)
                            {
                                ws.ListGNV.Add(new ArisingWorkItem(ws.ListGNV.Count + 1, cellB, string.IsNullOrEmpty(cellC) ? "Công" : cellC, kl, dg, tt > 0 ? tt : Math.Round(kl * dg, 0)));
                                importedGNV++;
                            }
                        }
                    }
                }

                if (matchedEmployees > 0)
                {
                    ws.RecalculateFromEmployees(true);
                    log.AppendLine(string.Format("- Đã nạp & cập nhật chi tiết từng người cho {0} lao động, tự động tính lương từng người và cộng tổng sang Bảng nghiệm thu + Biên bản!", matchedEmployees));
                }
                else
                {
                    ws.RecalculateDerived();
                }

                if (importedGNV > 0)
                    log.AppendLine(string.Format("- Đã nạp {0} khoản công việc phát sinh (GNV).", importedGNV));
            }

            if (log.Length == 0)
                log.AppendLine("- Đã kiểm tra file Excel và đồng bộ các chỉ tiêu hiện hành.");

            return log.ToString();
        }

        private static void CopyWorkshopValues(WorkshopData src, WorkshopData dst)
        {
            dst.TongCong = src.TongCong;
            dst.CongSP = src.CongSP;
            dst.CongQL = src.CongQL;
            dst.CongLDTT = src.CongLDTT;
            dst.CongLePhepCD = src.CongLePhepCD;
            dst.CongLePhepTT = src.CongLePhepTT;
            dst.CongHocHopTT = src.CongHocHopTT;
            dst.CongCDQL = src.CongCDQL;
            dst.CongBHXH = src.CongBHXH;
            dst.LDCuoiKy = src.LDCuoiKy;
            dst.LDTrucTiep = src.LDTrucTiep;
            dst.LDQuanLy = src.LDQuanLy;
            dst.LDThucTeBQ = src.LDThucTeBQ;
            dst.LDTTThucTeBQ = src.LDTTThucTeBQ;
            dst.NgayCongBQ = src.NgayCongBQ;
            dst.NgayCongBQQL = src.NgayCongBQQL;
            dst.NgayCongBQLDTT = src.NgayCongBQLDTT;
            dst.SanLuongKH = src.SanLuongKH;
            dst.SanLuongTH = src.SanLuongTH;
            dst.SanLuongTinhLuong = src.SanLuongTinhLuong;
            dst.SanLuongLoai1 = src.SanLuongLoai1;
            dst.DonGiaLoai1 = src.DonGiaLoai1;
            dst.ThanhTienLoai1 = src.ThanhTienLoai1;
            dst.SanLuongLoai2 = src.SanLuongLoai2;
            dst.DonGiaLoai2 = src.DonGiaLoai2;
            dst.ThanhTienLoai2 = src.ThanhTienLoai2;
            dst.SanLuongLe = src.SanLuongLe;
            dst.DonGiaLe = src.DonGiaLe;
            dst.ThanhTienLe = src.ThanhTienLe;
            dst.SanLuongPhuKL = src.SanLuongPhuKL;
            dst.SanLuongPhuTT = src.SanLuongPhuTT;
            dst.TongLuongToanPX = src.TongLuongToanPX;
            dst.TongLuongTT = src.TongLuongTT;
            dst.LuongSPTT = src.LuongSPTT;
            dst.LuongCVKhac = src.LuongCVKhac;
            dst.CongCVKhac = src.CongCVKhac;
            dst.LuongCheDoTT = src.LuongCheDoTT;
            dst.LuongLePhepTT = src.LuongLePhepTT;
            dst.LuongHocHopTT = src.LuongHocHopTT;
            dst.TruLuong = src.TruLuong;
            dst.TongLuongQL = src.TongLuongQL;
            dst.LuongSPQL = src.LuongSPQL;
            dst.LuongCheDoQL = src.LuongCheDoQL;
            dst.TLBinhQuanToanPX = src.TLBinhQuanToanPX;
            dst.ThuNhapBQToanPX = src.ThuNhapBQToanPX;
            dst.TLSPBinhQuanLDTT = src.TLSPBinhQuanLDTT;
            dst.ThuNhapBQLDTT = src.ThuNhapBQLDTT;
            dst.BQNgayCongLVTT = src.BQNgayCongLVTT;
            dst.CongAnCa = src.CongAnCa;
            dst.TienAnCa = src.TienAnCa;
            dst.BDDH_Muc1_Cong = src.BDDH_Muc1_Cong;
            dst.BDDH_Muc1_Tien = src.BDDH_Muc1_Tien;
            dst.BDDH_Muc2_Cong = src.BDDH_Muc2_Cong;
            dst.BDDH_Muc2_Tien = src.BDDH_Muc2_Tien;
            dst.BDDH_Muc3_Cong = src.BDDH_Muc3_Cong;
            dst.BDDH_Muc3_Tien = src.BDDH_Muc3_Tien;
            dst.TongTienBDDH = src.TongTienBDDH;
            if (src.ListGNV != null && src.ListGNV.Count > 0)
            {
                dst.ListGNV.Clear();
                dst.ListGNV.AddRange(src.ListGNV);
            }
            if (src.ListNhanVien != null && src.ListNhanVien.Count > 0)
            {
                dst.ListNhanVien.Clear();
                dst.ListNhanVien.AddRange(src.ListNhanVien);
            }
            dst.RecalculateDerived();
        }
    }
}
