using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace CiscoLDTLApp
{
    public class CdnlPayrollForm : Form
    {
        readonly WorkshopData workshop; readonly CdnlPayroll draft;
        readonly DataGridView groups,details,inputs,people,issues;
        readonly CheckBox distribute; readonly Label status;
        readonly Button apply;
        CdnlResult preview;
        public CdnlPayrollForm(WorkshopData ws)
        {
            if(ws.Cdnl==null) throw new ArgumentException("Nạp Excel có NT CĐNL và CĐNL trước.");
            workshop=ws; draft=ws.Cdnl.Draft();
            Text="Tính lương CĐNL từng bộ phận — v5.8"; Size=new Size(1320,800); MinimumSize=new Size(1000,650); StartPosition=FormStartPosition.CenterParent;
            Font=new Font("Segoe UI",9);
            var top=new FlowLayoutPanel { Dock=DockStyle.Top,Height=80,Padding=new Padding(8) };
            distribute=new CheckBox { Text="Tính lại lương từng người",Checked=draft.Distribute,AutoSize=true };
            var calc=new Button { Text="Tính / xem trước",AutoSize=true }; apply=new Button { Text="Áp dụng vào nghiệm thu",AutoSize=true };
            var close=new Button { Text="Hủy / đóng",AutoSize=true,DialogResult=DialogResult.Cancel };
            top.Controls.Add(distribute); top.Controls.Add(calc); top.Controls.Add(apply); top.Controls.Add(close);
            status=new Label { Width=1220,Height=40,Text="Quỹ theo đúng công thức từng nhóm; không chia toàn PX bằng một công thức. Áp dụng xong cần bấm Lưu Excel ở màn hình chính." }; top.Controls.Add(status);
            var tabs=new TabControl { Dock=DockStyle.Fill };
            var funds=new TabPage("1. Quỹ từng bộ phận"); var split=new SplitContainer { Dock=DockStyle.Fill,Orientation=Orientation.Horizontal,SplitterDistance=210 };
            groups=Grid(true,"Bộ phận","Ô quỹ","Quỹ tháng (đồng)","Đã phân bổ SP (đồng)","Chênh lệch để đối chiếu","Công thức nguồn");
            details=Grid(true,"Ô thành tiền","Nội dung","Thành tiền (đồng)","Công thức riêng"); split.Panel1.Controls.Add(groups); split.Panel2.Controls.Add(details); funds.Controls.Add(split); tabs.TabPages.Add(funds);
            var inputTab=new TabPage("2. Đầu vào / ô cần đối chiếu");
            inputs=Grid(false,"Ô nguồn","Nội dung","Giá trị nguồn","Giá trị thay thế","Lý do / căn cứ","Áp dụng thay thế","Công thức nguồn");
            ReplaceCheck(inputs,5); for(int i=0;i<inputs.Columns.Count;i++) inputs.Columns[i].ReadOnly=i!=3&&i!=4&&i!=5;
            var addBar=new FlowLayoutPanel { Dock=DockStyle.Top,Height=65 };
            var keyBox=new TextBox { Width=280,Text="NT CĐNL!H46" }; var add=new Button { Text="Thêm ô đối chiếu",AutoSize=true };
            addBar.Controls.Add(keyBox); addBar.Controls.Add(add); addBar.Controls.Add(new Label { Text="Ô lỗi/đầu vào thay thế cần giá trị hữu hạn và lý do. Bỏ chọn để dùng lại nguồn. Công/BHXH nhập tại bảng công.",Width=1100,Height=30 });
            add.Click+=(s,e)=> { string key=keyBox.Text.Trim(); int pos=key.LastIndexOf('!'); if(pos>0) key=PayrollFormulaBook.Key(key.Substring(0,pos),key.Substring(pos+1)); try { draft.SourceCell(key); AddInput(key); } catch(ArgumentException ex) { MessageBox.Show(this,ex.Message); } };
            inputTab.Controls.Add(inputs); inputTab.Controls.Add(addBar); tabs.TabPages.Add(inputTab);
            var employeeTab=new TabPage("3. Phân bổ từng người");
            people=Grid(false,"SDB","Họ tên","Bộ phận","Công thức AD nguồn","Điểm hưởng riêng","Đơn giá / ô nguồn AJ...","Lý do","Dùng đơn giá riêng","Lương SP xem trước","Tổng lương xem trước");
            ReplaceCheck(people,7); people.Columns[0].ReadOnly=people.Columns[1].ReadOnly=people.Columns[3].ReadOnly=people.Columns[8].ReadOnly=people.Columns[9].ReadOnly=true;
            var combo=new DataGridViewComboBoxColumn { HeaderText="Bộ phận",Name="Bộ phận",FlatStyle=FlatStyle.Flat }; combo.Items.Add("Theo công thức nguồn"); combo.Items.AddRange(GroupNames()); people.Columns.RemoveAt(2); people.Columns.Insert(2,combo);
            people.CurrentCellDirtyStateChanged+=(s,e)=> { if(people.IsCurrentCellDirty) people.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            people.CellValueChanged+=(s,e)=> {
                if(e.RowIndex<0) return; var row=people.Rows[e.RowIndex];
                if(e.ColumnIndex==7) { SetEditable(row); if(Checked(row.Cells[7].Value) && Convert.ToString(row.Cells[4].Value)=="" && row.Tag is EmployeeRecord) row.Cells[4].Value=Format(((EmployeeRecord)row.Tag).HeSoHuong); }
                if((e.ColumnIndex==2 || e.ColumnIndex==7) && Checked(row.Cells[7].Value)) { int group=Array.IndexOf(CdnlPayroll.GroupNames,Convert.ToString(row.Cells[2].Value)); string current=Convert.ToString(row.Cells[5].Value); if(group>=0 && (current=="" || current.StartsWith("AJ"))) row.Cells[5].Value=CdnlPayroll.RateForGroup(CdnlPayroll.GroupCells[group]); }
            };
            people.DataError+=(s,e)=> { status.Text="Kiểm tra bộ phận và giá trị vừa nhập."; e.ThrowException=false; };
            var employeeNote=new Label { Dock=DockStyle.Top,Height=65,Text="Mặc định giữ công thức AD/AB và đơn giá nhóm trong Excel (AJ146/AJ152/AJ158/AJ168). Muốn sửa điểm nguồn, thêm ô CĐNL!AB<dòng> ở tab 2.\nHồ sơ mới/chuyển nhóm chưa có công thức phải chọn bộ phận, điểm, đơn giá riêng và lý do. Không tự đoán nhóm theo tên tổ. Quản lý đã bật quỹ theo sản lượng dùng màn hình quản lý riêng." };
            employeeTab.Controls.Add(people); employeeTab.Controls.Add(employeeNote); tabs.TabPages.Add(employeeTab);
            var errorTab=new TabPage("4. Lỗi cần xử lý"); issues=Grid(true,"Ô / hồ sơ","Lý do"); errorTab.Controls.Add(issues); tabs.TabPages.Add(errorTab);
            var note=new Label { Dock=DockStyle.Bottom,Height=42,Text="Quỹ trưởng ca/sàng có thể không có dòng nhân viên riêng trong nguồn: không tự nhập chung sang nhóm khác. Chênh lệch quỹ và lương SP phải được đối chiếu, không tự dồn cho người cuối.\nBỏ chọn tính từng người để chỉ cập nhật quỹ; không xuất phụ lục lương cá nhân khi chưa tính lại phân bổ." };
            Controls.Add(tabs); Controls.Add(note); Controls.Add(top); CancelButton=close;
            foreach(var e in ws.ListNhanVien) {
                CdnlEmployeeRule rule; bool manual=draft.Rules.TryGetValue(CdnlPayroll.EmployeeKey(e),out rule); string code=draft.SourceGroup(e); int index=Array.IndexOf(CdnlPayroll.GroupCells,code);
                int row=people.Rows.Add(e.SDB,e.HoTen,index>=0?CdnlPayroll.GroupNames[index]:e.LaQuanLy?"Quản lý":"Theo công thức nguồn",draft.SourceFormula("CĐNL!AD"+e.ExcelRow),manual?Format(rule.Weight):"",manual?(!string.IsNullOrEmpty(rule.RateCell)?rule.RateCell:Format(rule.Rate)):"",manual?rule.Reason:"",manual,"","");
                people.Rows[row].Tag=e; SetEditable(people.Rows[row]);
            }
            calc.Click+=(s,e)=>Run(false); apply.Click+=(s,e)=>Run(true);
            groups.SelectionChanged+=(s,e)=>ShowDetails();
            Run(false);
        }
        static string[] GroupNames() { return CdnlPayroll.GroupNames; }
        static DataGridView Grid(bool readOnly,params string[] headers) {
            var grid=new DataGridView { Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,ReadOnly=readOnly,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,RowHeadersVisible=false };
            foreach(string h in headers) grid.Columns.Add(h,h); return grid;
        }
        static void ReplaceCheck(DataGridView grid,int i) { string h=grid.Columns[i].HeaderText; grid.Columns.RemoveAt(i); grid.Columns.Insert(i,new DataGridViewCheckBoxColumn { HeaderText=h,Name=h }); }
        static string Format(double value) { return value.ToString("0.########",CultureInfo.GetCultureInfo("vi-VN")); }
        static double Number(object value) { double n; string text=Convert.ToString(value).Trim(); if(text=="") text="0"; if(!double.TryParse(text,NumberStyles.Number,CultureInfo.GetCultureInfo("vi-VN"),out n)) throw new ArgumentException("Nhập số kiểu Việt Nam, dùng dấu phẩy thập phân."); return PayrollFormulaBook.Finite(n); }
        static bool Checked(object value) { return value!=null && value!=DBNull.Value && Convert.ToBoolean(value); }
        void SetEditable(DataGridViewRow row) { bool manual=Checked(row.Cells[7].Value); foreach(int c in new[] {2,4,5,6}) row.Cells[c].ReadOnly=!manual; }
        void AddInput(string key) {
            foreach(DataGridViewRow row in inputs.Rows) if(Convert.ToString(row.Cells[0].Value)==key) return;
            PayrollCell source; try { source=draft.SourceCell(key); } catch(ArgumentException) { return; }
            CdnlOverride item; draft.Overrides.TryGetValue(key,out item);
            string address=source.Address; string rowNumber=Regex.Match(address,@"\d+").Value;
            PayrollCell label; string labelKey=source.Sheet+"!B"+rowNumber;
            string title=draft.Book.Cells.TryGetValue(labelKey,out label)?label.Text??"":"";
            inputs.Rows.Add(key,title,source.IsError?source.Text:source.IsNumber?Format(source.Number):source.Text??"",item!=null&&item.Active?Format(item.Value):"",item==null?"":item.Reason,item!=null&&item.Active,source.Formula??"");
        }
        void KeepSource(string key) { draft.KeepSource(key); }
        void ReadInputs() {
            inputs.EndEdit(); people.EndEdit();
            foreach(DataGridViewRow row in inputs.Rows) {
                string key=Convert.ToString(row.Cells[0].Value); CdnlOverride item;
                if(Checked(row.Cells[5].Value)) draft.SetOverride(key,Number(row.Cells[3].Value),Convert.ToString(row.Cells[4].Value));
                else if(draft.Overrides.TryGetValue(key,out item)) item.Active=false;
            }
            foreach(DataGridViewRow row in people.Rows) {
                var employee=(EmployeeRecord)row.Tag; string key=CdnlPayroll.EmployeeKey(employee);
                if(!Checked(row.Cells[7].Value)) { draft.Rules.Remove(key); continue; }
                if(employee.LaQuanLy && workshop.TinhQuyQLTuSanLuong) throw new ArgumentException("Quản lý đang dùng quỹ theo sản lượng: sửa tại màn hình quản lý, không nhập đơn giá riêng ở đây.");
                string name=Convert.ToString(row.Cells[2].Value); int index=Array.IndexOf(CdnlPayroll.GroupNames,name);
                double weight=Number(row.Cells[4].Value); string rateCell=Convert.ToString(row.Cells[5].Value).Trim().ToUpperInvariant();
                if(rateCell=="") throw new ArgumentException("Chọn ô đơn giá nhóm hoặc nhập đơn giá điểm riêng; không để trống.");
                double rate=rateCell.StartsWith("AJ")?0:Number(row.Cells[5].Value); if(!rateCell.StartsWith("AJ")) rateCell="";
                string reason=Convert.ToString(row.Cells[6].Value);
                if(index<0||weight<0||rate<0||string.IsNullOrWhiteSpace(reason)) throw new ArgumentException("Phân bổ riêng cần bộ phận, điểm/giá không âm và lý do.");
                KeepSource("CĐNL!AB"+employee.ExcelRow); KeepSource("CĐNL!AD"+employee.ExcelRow);
                draft.Rules[key]=new CdnlEmployeeRule { Group=CdnlPayroll.GroupCells[index],Weight=weight,Rate=rate,RateCell=rateCell,Reason=reason.Trim() };
            }
            draft.Distribute=distribute.Checked;
        }
        void Run(bool commit) {
            try {
                ReadInputs(); preview=draft.Preview(workshop); groups.Rows.Clear();
                for(int i=0;i<CdnlPayroll.GroupCells.Length;i++) {
                    string key=CdnlPayroll.GroupCells[i]; double allocated=0;
                    foreach(var pair in preview.Employees) if(draft.SourceGroup(pair.Key)==key || (key=="I422" && pair.Key.LaQuanLy)) allocated+=pair.Value.LuongSanPham;
                    groups.Rows.Add(CdnlPayroll.GroupNames[i],"NT CĐNL!"+key,preview.InvalidGroups.Contains(key)?"Chưa tính được":WordReportService.FmtInt(preview.Groups[key]),draft.Distribute?(preview.Valid?WordReportService.FmtInt(allocated):"Chưa đủ dữ liệu"):"Chưa tính",draft.Distribute?(preview.Valid?WordReportService.FmtInt(preview.Groups[key]-allocated):"Chưa đủ dữ liệu"):"Chưa tính",draft.SourceFormula("NT CĐNL!"+key));
                }
                issues.Rows.Clear(); foreach(var p in preview.Errors) { issues.Rows.Add(p.Key,p.Value); AddInput(p.Key); }
                foreach(string key in preview.Inputs) { PayrollCell c; if(draft.Book.Cells.TryGetValue(key,out c) && c.IsNumber) AddInput(key); }
                foreach(string key in draft.Overrides.Keys) AddInput(key);
                foreach(DataGridViewRow row in people.Rows) { EmployeeRecord result; if(!preview.InvalidEmployees.Contains((EmployeeRecord)row.Tag) && preview.Employees.TryGetValue((EmployeeRecord)row.Tag,out result)) { row.Cells[8].Value=WordReportService.FmtInt(result.LuongSanPham); row.Cells[9].Value=WordReportService.FmtInt(result.TongTienLuong); } else { row.Cells[8].Value=row.Cells[9].Value=draft.Distribute?"Chưa tính được":"Giữ số cũ"; } }
                apply.Enabled=preview.Valid;
                status.Text=preview.Valid?(draft.Distribute?"Đã tính quỹ và từng người. Đối chiếu chênh lệch và bấm Áp dụng.":"Đã tính quỹ từng bộ phận; lương cá nhân giữ số cũ. Bấm Áp dụng nếu chỉ cần cập nhật quỹ."):"Chưa áp dụng: "+preview.Errors.Count+" ô/hồ sơ cần xử lý tại tab 4. Chọn ô ở tab 2, nhập giá trị đã đối chiếu kèm lý do rồi tính lại.";
                ShowDetails();
                if(commit && preview.Valid) { draft.Apply(workshop); DialogResult=DialogResult.OK; Close(); }
            } catch(ArgumentException ex) { status.Text=ex.Message; apply.Enabled=false; }
        }
        void ShowDetails() {
            if(preview==null || groups.CurrentRow==null) return; details.Rows.Clear(); int index=groups.CurrentRow.Index;
            int[] first={44,46,82,342,357,423},last={44,80,340,355,415,435};
            for(int r=first[index];r<=last[index];r++) { string key="NT CĐNL!I"+r; double amount; if(!preview.Cells.TryGetValue(key,out amount)) continue; PayrollCell label; string title=draft.Book.Cells.TryGetValue("NT CĐNL!B"+r,out label)?label.Text??"":""; details.Rows.Add(key,title,WordReportService.FmtInt(amount),draft.SourceFormula(key)); }
        }
    }
}
