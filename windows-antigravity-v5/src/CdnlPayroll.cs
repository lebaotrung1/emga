using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
namespace CiscoLDTLApp
{
    public class CdnlOverride { public double Value; public string Reason; public bool Active=true; public PayrollCell Source; }
    public class CdnlEmployeeRule { public string Group; public double Weight, Rate; public string RateCell; public string Reason; }
    public class CdnlResult
    {
        public readonly Dictionary<string,double> Groups=new Dictionary<string,double>();
        public readonly Dictionary<string,double> Cells=new Dictionary<string,double>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<EmployeeRecord,EmployeeRecord> Employees=new Dictionary<EmployeeRecord,EmployeeRecord>();
        public readonly Dictionary<string,string> Errors=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> Inputs=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<EmployeeRecord> InvalidEmployees=new HashSet<EmployeeRecord>();
        public readonly HashSet<string> InvalidGroups=new HashSet<string>();
        public int EvaluationErrors;
        public bool Valid { get { return Errors.Count==0; } }
    }
    public class CdnlPayroll
    {
        public static readonly string[] GroupCells={"I44","I45","I81","I341","I356","I422"};
        public static readonly string[] GroupNames={"Trưởng ca","Vận hành năng lượng","Vận tải, cẩu, xúc lật","Sàng quặng","Sửa chữa cơ điện","Quản lý"};
        class PointBaseline { public double Point,Weight; public string Rating; }
        readonly Dictionary<string,PointBaseline> pointBaselines=new Dictionary<string,PointBaseline>();
        public readonly PayrollFormulaBook Book;
        public readonly List<WorkshopData> Workshops;
        public readonly Dictionary<string,CdnlOverride> Overrides=new Dictionary<string,CdnlOverride>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string,CdnlEmployeeRule> Rules=new Dictionary<string,CdnlEmployeeRule>();
        public bool Applied;
        public bool Distribute=true;
        public string Status="Chưa áp dụng bộ tính CĐNL. Mở TÍNH LƯƠNG CĐNL TỪNG BỘ PHẬN để kiểm tra.";
        public CdnlResult Last;
        public CdnlPayroll(PayrollFormulaBook book,List<WorkshopData> workshops) { Book=book; Workshops=workshops;
            foreach(var unit in workshops) if(unit.MaPX!="CDNL") foreach(var e in unit.ListNhanVien) pointBaselines[unit.MaPX+"|"+EmployeeKey(e)]=new PointBaseline { Point=e.DiemSanXuat,Weight=e.HeSoHuong,Rating=e.XepLoai };
        }
        public CdnlPayroll Draft() {
            var copy=new CdnlPayroll(Book,Workshops) { Applied=Applied,Distribute=Distribute,Status=Status,Last=Last };
            copy.pointBaselines.Clear(); foreach(var p in pointBaselines) copy.pointBaselines[p.Key]=p.Value;
            foreach(var p in Overrides) copy.Overrides[p.Key]=new CdnlOverride { Value=p.Value.Value,Reason=p.Value.Reason,Active=p.Value.Active,Source=p.Value.Source };
            foreach(var p in Rules) copy.Rules[p.Key]=new CdnlEmployeeRule { Group=p.Value.Group,Weight=p.Value.Weight,Rate=p.Value.Rate,RateCell=p.Value.RateCell,Reason=p.Value.Reason };
            return copy;
        }
        public static string EmployeeKey(EmployeeRecord e) { return e.ExcelRow+"|"+e.SDB+"|"+e.HoTen; }
        public string SourceFormula(string key) { PayrollCell cell; return Book.Cells.TryGetValue(key,out cell)?cell.Formula??"":""; }
        public bool IsLiteralWeight(int row) { PayrollCell c; return Book.Cells.TryGetValue("CĐNL!AB"+row,out c) && c.IsNumber && string.IsNullOrEmpty(c.Formula); }
        public void KeepSource(string key) { if(Overrides.ContainsKey(key) || key.EndsWith("0") && Regex.IsMatch(key,@"![A-Z]+0$")) return; var c=SourceCell(key); Overrides[key]=new CdnlOverride { Source=c,Active=false,Reason="Giữ công thức nguồn trước khi áp dụng chế độ tính riêng" }; }
        public bool IsRatingRow(int row) { PayrollCell cell; if(!Book.Cells.TryGetValue("CĐNL!Z"+row,out cell)) return false; string value=(cell.Text??"").Trim().ToUpperInvariant(); return value=="A"||value=="B"||value=="C"; }
        public static string RateForGroup(string group) { return group=="I45"?"AJ146":group=="I356"?"AJ152":group=="I81"?"AJ158":group=="I422"?"AJ168":""; }
        public string SourceGroup(EmployeeRecord e) {
            CdnlEmployeeRule rule; if(Rules.TryGetValue(EmployeeKey(e),out rule)) return rule.Group;
            string formula=SourceFormula(PayrollFormulaBook.Key("CĐNL","AD"+e.ExcelRow));
            var m=Regex.Match(formula,@"\$?AJ\$?(146|152|158|168)\b",RegexOptions.IgnoreCase);
            return !m.Success?"":m.Groups[1].Value=="146"?"I45":m.Groups[1].Value=="152"?"I356":m.Groups[1].Value=="158"?"I81":"I422";
        }
        public PayrollCell SourceCell(string key)
        {
            PayrollCell cell; if(Book.Cells.TryGetValue(key,out cell)) return cell;
            int split=key.LastIndexOf('!'); if(split<=0) throw new ArgumentException("Nhập ô theo Sheet!A1.");
            string sheet=key.Substring(0,split),address=key.Substring(split+1).ToUpperInvariant();
            if(!Book.Sheets.Contains(sheet) || !Regex.IsMatch(address,@"^[A-Z]{1,3}[1-9][0-9]*$")) throw new ArgumentException("Sheet hoặc địa chỉ ô không hợp lệ.");
            int column=0; foreach(char c in Regex.Match(address,@"^[A-Z]+").Value) column=column*26+c-'A'+1;
            long row; if(column>16384 || !long.TryParse(Regex.Match(address,@"[0-9]+$").Value,out row) || row>1048576) throw new ArgumentException("Ô ngoài giới hạn Excel.");
            cell=new PayrollCell { Sheet=sheet,Address=address }; Book.Cells[cell.Key]=cell; return cell;
        }
        public void SetOverride(string key,double value,string reason) {
            int split=key.LastIndexOf('!');
            if(split<=0 || !Regex.IsMatch(key.Substring(split+1),@"^[A-Z]{1,3}[1-9][0-9]*$",RegexOptions.IgnoreCase)) throw new ArgumentException("Nhập ô theo Sheet!A1.");
            if(string.IsNullOrWhiteSpace(reason)) throw new ArgumentException("Cần lý do/nguồn đối chiếu cho thay đổi đầu vào.");
            var source=SourceCell(key);
            CdnlOverride old; Overrides.TryGetValue(key,out old);
            Overrides[key]=new CdnlOverride { Value=PayrollFormulaBook.Finite(value),Reason=reason.Trim(),Source=old==null?source:old.Source };
        }
        void Bind(WorkshopData ws)
        {
            Book.Values.Clear(); Book.TextValues.Clear(); Book.Reset();
            foreach(var unit in Workshops) {
                string s=unit.MaPX=="TK"?"NT thiêu kết":unit.MaPX=="LG"?"NT gang":unit.MaPX=="LT"?"NT Thép":null;
                if(s==null) continue;
                int fund=unit.MaPX=="TK"?57:unit.MaPX=="LG"?48:35;
                int holiday=unit.MaPX=="TK"?63:unit.MaPX=="LG"?55:43;
                Book.Set(s,"J"+fund,unit.LuongSPTT); Book.Set(s,"J"+holiday,unit.ThanhTienLe);
                if(unit.MaPX=="LG") { Book.Set(s,"D51",unit.SanLuongTinhLuong); Book.Set(s,"D55",unit.SanLuongLe); }
                if(unit.MaPX=="LT") { double le=0; if(unit.CauPhoiDaApDung) foreach(var stage in unit.CauPhoi) le+=stage.Holiday*stage.Rate*2;
                    else { PayrollCell c; if(Book.Cells.TryGetValue(s+"!J45",out c) && c.IsNumber) le=c.Number; }
                    Book.Set(s,"J45",le);
                }
            }
            foreach(var unit in Workshops) if(unit.MaPX!="CDNL") foreach(var e in unit.ListNhanVien) {
                if(e.ExcelRow<=0) continue; string r=e.ExcelRow.ToString(),sheet=unit.SheetBangCong;
                Book.Set(sheet,"G"+r,e.MucLuongBHXH); Book.Set(sheet,"H"+r,e.PhuCapChucVu); Book.Set(sheet,"I"+r,e.HeSoCoDinh);
                PointBaseline previous; if(pointBaselines.TryGetValue(unit.MaPX+"|"+EmployeeKey(e),out previous)) {
                    if(e.DiemSanXuat!=previous.Point) Book.Set(sheet,"W"+r,e.DiemSanXuat);
                    if(e.HeSoHuong!=previous.Weight) { Book.Set(sheet,"Y"+r,e.HeSoHuong); Book.Set(sheet,"Z"+r,e.HeSoHuong); }
                    if(e.XepLoai!=previous.Rating) Book.TextValues[sheet+"!X"+r]=e.XepLoai??"A";
                }
                Book.Set(sheet,"K"+r,e.CongSP); Book.Set(sheet,"M"+r,e.CongLePhep); Book.Set(sheet,"P"+r,e.CongHocHop); Book.Set(sheet,"Q"+r,e.CongLamLe); Book.Set(sheet,"R"+r,e.CongVHTT);
            }
            foreach(var e in ws.ListNhanVien) {
                if(e.ExcelRow<=0) continue; string r=e.ExcelRow.ToString();
                Book.Set("CĐNL","G"+r,e.MucLuongBHXH); Book.Set("CĐNL","H"+r,e.PhuCapChucVu);
                Book.Set("CĐNL","I"+r,e.HeSoCoDinh); Book.Set("CĐNL","K"+r,e.CongSP);
                Book.Set("CĐNL","M"+r,e.CongLePhep); Book.Set("CĐNL","P"+r,e.CongHocHop);
                Book.Set("CĐNL","Q"+r,e.CongLamLe); Book.Set("CĐNL","R"+r,e.CongVHTT);
                Book.Set("CĐNL","S"+r,e.CongBHXH);
                PayrollCell point; if(Book.Cells.TryGetValue("CĐNL!Y"+r,out point) && string.IsNullOrEmpty(point.Formula)) Book.Set("CĐNL","Y"+r,e.DiemSanXuat);
                // Text rating is a source attribute; only the source AB formula decides whether it affects pay.
                if(IsRatingRow(e.ExcelRow)) Book.TextValues["CĐNL!Z"+r]=e.XepLoai??"A";
            }
            foreach(var p in Overrides) if(p.Value.Active) Book.Values[p.Key]=p.Value.Value;
            if(ws.GnvDaChinhSua) { double sum=0; foreach(var item in ws.ListGNV) { sum+=item.ThanhTien; if(item.ExcelRow>=372 && item.ExcelRow<=393) { Book.Set("NT CĐNL","D"+item.ExcelRow,item.KhoiLuong); Book.Set("NT CĐNL","E"+item.ExcelRow,item.KhoiLuong); Book.Set("NT CĐNL","H"+item.ExcelRow,item.DonGia); Book.Set("NT CĐNL","I"+item.ExcelRow,item.ThanhTien); } } Book.Set("NT CĐNL","I371",sum); }
            foreach(var e in ws.ListNhanVien) {
                CdnlEmployeeRule rule; if(!Rules.TryGetValue(EmployeeKey(e),out rule)) continue;
                if(e.ExcelRow>0) Book.Set("CĐNL","AB"+e.ExcelRow,rule.Weight);
            }
        }
        double Eval(CdnlResult result,string sheet,string address) {
            try { double value=PayrollFormulaBook.Finite(Book.Evaluate(sheet,address)); result.Cells[PayrollFormulaBook.Key(sheet,address)]=value; return value; }
            catch(PayrollFormulaError ex) { result.EvaluationErrors++; result.Errors[ex.Cell]=ex.Message; return 0; }
            catch(ArgumentException ex) { result.EvaluationErrors++; string key=sheet+"!"+address; result.Errors[key]=key+": "+ex.Message; return 0; }
        }
        public CdnlResult Preview(WorkshopData ws)
        {
            if(ws.MaPX!="CDNL") throw new ArgumentException("Chỉ tính CĐNL.");
            Bind(ws); var result=new CdnlResult();
            WorkshopData management=null;
            if(ws.TinhQuyQLTuSanLuong) {
                management=new WorkshopData { MaPX="CDNL",LuongSPQL=ws.LuongSPQL,DieuChinhLuongQL=ws.DieuChinhLuongQL,SanLuongLoai1=ws.SanLuongLoai1,SanLuongLoai2=ws.SanLuongLoai2,SanLuongLe=ws.SanLuongLe };
                ManagementProductionPayroll.CopySettings(ws,management);
                foreach(var e in ws.ListNhanVien) if(e.LaQuanLy) { var copy=new EmployeeRecord(); ManagementPayroll.Copy(e,copy); copy.QuanLyTinhMoi=true; management.ListNhanVien.Add(copy); result.Employees[e]=copy; }
                ManagementPayroll.Validate(management); ManagementPayroll.Calculate(management,true);
                Book.Set("NT CĐNL","I422",management.LuongSPQL); Book.Set("NT CĐNL","I432",management.LuongCheDoQL);
                foreach(var pair in result.Employees) { string row=pair.Key.ExcelRow.ToString(); if(pair.Key.ExcelRow<=0) continue; var e=pair.Value;
                    Book.Set("CĐNL","AB"+row,e.HeSoHuong); Book.Set("CĐNL","AD"+row,e.LuongSanPham); Book.Set("CĐNL","AE"+row,e.LuongLePhepCD); Book.Set("CĐNL","AI"+row,e.LuongLamLe); Book.Set("CĐNL","AJ"+row,e.LuongHocVHTT); Book.Set("CĐNL","AK"+row,0); Book.Set("CĐNL","AL"+row,e.PhuCapTien); Book.Set("CĐNL","AC"+row,e.TongTienLuong);
                }
            }
            foreach(string cell in GroupCells) { int errors=result.EvaluationErrors; result.Groups[cell]=Eval(result,"NT CĐNL",cell); if(result.EvaluationErrors>errors) result.InvalidGroups.Add(cell); }
            if(Distribute) foreach(var e in ws.ListNhanVien) {
                if(management!=null && e.LaQuanLy) continue;
                int rowErrors=result.EvaluationErrors;
                var copy=new EmployeeRecord(); ManagementPayroll.Copy(e,copy); string row=e.ExcelRow.ToString();
                CdnlEmployeeRule rule; bool manual=Rules.TryGetValue(EmployeeKey(e),out rule);
                if(e.ExcelRow<=0 && !manual) { result.Errors["Hồ sơ "+e.SDB]="Hồ sơ mới chưa có công thức nguồn: chọn bộ phận, điểm hưởng, đơn giá riêng và lý do tại tab Phân bổ."; continue; }
                if(e.ExcelRow<=0 && (e.CongLePhep!=0 || e.CongLamLe!=0 || e.CongHocHop!=0 || e.CongVHTT!=0)) { result.Errors["Hồ sơ "+e.SDB]="Hồ sơ mới có công chế độ/lễ/học nhưng chưa có dòng công thức Excel. Nạp hồ sơ có công thức đầy đủ trước khi phân bổ, không tự áp mức chung."; continue; }
                if(manual) {
                    if(Array.IndexOf(GroupCells,rule.Group)<0 || rule.Weight<0 || rule.Rate<0 || string.IsNullOrWhiteSpace(rule.Reason)) { result.Errors["Hồ sơ "+e.SDB]="Quy tắc phân bổ riêng cần nhóm, điểm/giá không âm và lý do."; continue; }
                    if(!string.IsNullOrEmpty(rule.RateCell) && rule.RateCell!=RateForGroup(rule.Group)) { result.Errors["Hồ sơ "+e.SDB]="Ô đơn giá không thuộc nhóm đã chọn."; continue; }
                    double rate=string.IsNullOrEmpty(rule.RateCell)?rule.Rate:Eval(result,"CĐNL",rule.RateCell);
                    copy.HeSoHuong=rule.Weight; copy.LuongSanPham=ManagementPayroll.RoundMoney(rule.Weight*rate,100);
                    if(e.ExcelRow>0) Book.Set("CĐNL","AD"+row,copy.LuongSanPham);
                }
                else {
                    PayrollCell sourcePay;
                    bool missingPay=!Book.Cells.TryGetValue("CĐNL!AD"+row,out sourcePay) || (string.IsNullOrEmpty(sourcePay.Formula) && !sourcePay.IsNumber && string.IsNullOrEmpty(sourcePay.Text));
                    CdnlOverride approvedPay; bool approved=Overrides.TryGetValue("CĐNL!AD"+row,out approvedPay) && approvedPay.Active;
                    if(e.CongSP!=0 && missingPay && !approved) { result.Errors["CĐNL!AD"+row]="Có công sản phẩm nhưng thiếu công thức/thành tiền AD nguồn. Khai báo nhóm, điểm, đơn giá và lý do tại tab Phân bổ hoặc nạp nguồn đầy đủ."; result.InvalidEmployees.Add(e); continue; }
                    copy.HeSoHuong=Eval(result,"CĐNL","AB"+row);
                    copy.LuongSanPham=Eval(result,"CĐNL","AD"+row);
                }
                if(e.ExcelRow>0) {
                    copy.LuongLePhepCD=Eval(result,"CĐNL","AE"+row); copy.LuongLamLe=Eval(result,"CĐNL","AI"+row);
                    copy.LuongHocVHTT=Eval(result,"CĐNL","AJ"+row);
                    copy.PhuCapTien=Eval(result,"CĐNL","AK"+row)+Eval(result,"CĐNL","AL"+row);
                    copy.TongTienLuong=Eval(result,"CĐNL","AC"+row);
                }
                else copy.TongTienLuong=copy.LuongSanPham+copy.LuongLePhepCD+copy.LuongLamLe+copy.LuongHocVHTT+copy.PhuCapTien;
                result.Employees[e]=copy;
                if(result.EvaluationErrors>rowErrors) result.InvalidEmployees.Add(e);
            }
            // Roots use the exact NT formulas, including deductions and auxiliary work once.
            foreach(string cell in new[] {"I43","I416","I418","I419","I420","I417","I432","I42","I421","I440"}) result.Cells["NT CĐNL!"+cell]=Eval(result,"NT CĐNL",cell);
            foreach(var p in Book.Results) result.Cells[p.Key]=p.Value;
            foreach(var p in Book.Values) result.Cells[p.Key]=p.Value;
            foreach(string key in Book.Inputs) result.Inputs.Add(key);
            foreach(string key in Book.Values.Keys) result.Inputs.Remove(key);

            return result;
        }
        public void Apply(WorkshopData ws)
        {
            var result=Preview(ws);
            if(!result.Valid) throw new ArgumentException("Chưa áp dụng: "+string.Join("\n",new List<string>(result.Errors.Values).ToArray()));
            if(ws.TinhQuyQLTuSanLuong) {
                KeepSource("NT CĐNL!I422"); KeepSource("NT CĐNL!I432");
                foreach(var e in ws.ListNhanVien) if(e.LaQuanLy && e.ExcelRow>0) foreach(string col in new[] {"AB","AD","AE","AI","AJ","AK","AL","AC"}) KeepSource("CĐNL!"+col+e.ExcelRow);
            }
            Commit(ws,result); Applied=true; ws.Cdnl=this;
            ws.RecalculateDerived();
        }
        void Commit(WorkshopData ws,CdnlResult result) {
            foreach(var pair in result.Employees) ManagementPayroll.Copy(pair.Value,pair.Key);
            Last=result; Status=Distribute?"Đã tính quỹ và phân bổ CĐNL theo công thức từng nhóm. Chưa lưu Excel.":"Đã tính quỹ từng bộ phận; chưa tính lại phân bổ từng người. Chưa lưu Excel.";
            UpdateSummary(ws);
        }
        public bool Recalculate(WorkshopData ws) {
            if(!Applied) return true;
            try {
                var result=Preview(ws);
                if(!result.Valid) { Status="CĐNL chưa tính được sau thay đổi đầu vào. Mở màn hình CĐNL để xử lý: "+string.Join("; ",new List<string>(result.Errors.Keys).ToArray()); return false; }
                Commit(ws,result); return true;
            } catch(ArgumentException ex) { Status="CĐNL chưa tính được: "+ex.Message; return false; }
        }
        public void UpdateSummary(WorkshopData ws) {
            if(!Applied || Last==null || !Last.Valid) return;
            ws.LuongSPTT=Last.Cells["NT CĐNL!I43"];
            ws.LuongCVKhac=Last.Cells["NT CĐNL!I416"];
            ws.LuongLePhepTT=Last.Cells["NT CĐNL!I418"]; ws.LuongVHTTTT=Last.Cells["NT CĐNL!I419"]; ws.LuongHocHopTT=Last.Cells["NT CĐNL!I420"];
            ws.LuongCheDoTT=Last.Cells["NT CĐNL!I417"];
            ws.LuongSPQL=Last.Groups["I422"]; ws.LuongCheDoQL=Last.Cells["NT CĐNL!I432"];
            ws.TongLuongTT=Last.Cells["NT CĐNL!I42"]; ws.TongLuongQL=Last.Cells["NT CĐNL!I421"]+(ws.TinhQuyQLTuSanLuong?ws.DieuChinhLuongQL:0);
            ws.TongLuongToanPX=ws.TongLuongTT+ws.TongLuongQL;
        }
        public static void EnsureCurrent(WorkshopData ws) { if(ws.Cdnl!=null && ws.Cdnl.Applied) { if(!ws.Cdnl.Recalculate(ws)) throw new ArgumentException(ws.Cdnl.Status); ws.RecalculateDerived(); } }
        public static void EnsureReport(WorkshopData ws) { EnsureCurrent(ws); if(ws.Cdnl!=null && ws.Cdnl.Applied && !ws.Cdnl.Distribute && ws.DinhKemBangChiTietNhanVien) throw new ArgumentException("CĐNL mới tính quỹ. Tính lại phân bổ từng người trước khi đính kèm bảng lương cá nhân."); }
        public static void AttendanceOnly(WorkshopData ws) {
            double sp=0,ql=0,cd=0,bh=0,anca=0,cdql=0,le=0,hoc=0,vh=0;
            foreach(var e in ws.ListNhanVien) {
                e.TongCong=e.CongSP+e.CongLePhep+e.CongVHTT; sp+=e.CongSP; bh+=e.CongBHXH;
                cd+=e.CongLePhep+e.CongHocHop+e.CongVHTT; if(e.LaQuanLy) { ql+=e.CongSP; cdql+=e.CongLePhep+e.CongHocHop+e.CongVHTT; } else { le+=e.CongLePhep; hoc+=e.CongHocHop; vh+=e.CongVHTT; }
                e.AnCaCongThuc=e.CongSP+e.CongHocHop+e.CongVHTT;
                if(e.AnCaNguon=="DV") e.CongAnCa=e.AnCaDonVi; else if(e.AnCaNguon=="TC") e.CongAnCa=e.AnCaThuCong; else if(e.AnCaNguon=="CT") e.CongAnCa=e.AnCaCongThuc; else if(!e.CoAnCa && e.CongAnCa<=0 && e.CongSP>0) e.CongAnCa=e.CongSP+e.CongHocHop;
                e.TienAnCa=e.CongAnCa*30000; anca+=e.CongAnCa;
            }
            if(ws.ListNhanVien.Count>0) { ws.CongSP=sp; ws.CongQL=ql; ws.CongLDTT=sp-ql; ws.CongLePhepCD=cd; ws.CongCDQL=cdql; ws.CongLePhepTT=le; ws.CongHocHopTT=hoc; ws.CongVHTTTT=vh; ws.CongBHXH=bh; ws.CongAnCa=anca; ws.LDThucTeBQ=Math.Round(sp/26,2); ws.LDTTThucTeBQ=Math.Round((sp-ql)/26,2); }
        }
        public static IEnumerable<string> ReportLines(WorkshopData ws) {
            if(ws.MaPX!="CDNL" || ws.Cdnl==null || !ws.Cdnl.Applied || ws.Cdnl.Last==null) yield break;
            yield return "Quỹ CĐNL tháng theo từng bộ phận (đã nằm trong tổng quỹ, không cộng thêm):";
            for(int i=0;i<GroupCells.Length;i++) yield return "+ "+GroupNames[i]+": "+WordReportService.FmtMoney(ws.Cdnl.Last.Groups[GroupCells[i]])+" triệu đồng.";
            yield return ws.Cdnl.Distribute?"Đã tính phân bổ từng người theo công thức nguồn/đơn giá nhóm đã đối chiếu. GNV thuộc quỹ sửa chữa không cộng lần nữa.":"Chỉ tính lại quỹ bộ phận; chưa tính lại lương từng người. Không dùng bảng lương cá nhân cũ để chốt thanh toán.";
        }
    }
}
