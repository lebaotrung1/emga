using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace NghiemThuStudio
{
    public sealed class MainForm : Form
    {
        Project project;
        string projectPath,dependencyFolder,templatePath;
        bool loading,dirty;
        ComboBox unit;
        NumericUpDown year,month;
        Label heading,subtitle,status;
        DataGridView inputs,works,results,history;
        RichTextBox diagnostic,preview,log;
        CheckBox reviewed,cumulative;
        DateTimePicker meeting;
        TextBox location,office,team,signOffice,signTeam,recommend,explanation,quantityUnit;
        TabControl tabs;
        public MainForm(Func<bool> confirmDiscard=null)
        {
            Text="Nghiệm Thu Studio 1.0 — hồ sơ nghiệm thu LĐTL";
            Size=new Size(1440,920); MinimumSize=new Size(1100,720); StartPosition=FormStartPosition.CenterScreen;
            Font=new Font("Segoe UI",10); BackColor=Color.FromArgb(244,247,251);
            var root=new TableLayoutPanel { Dock=DockStyle.Fill,RowCount=4,ColumnCount=1,Padding=new Padding(12) };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,65)); root.RowStyles.Add(new RowStyle(SizeType.Absolute,100)); root.RowStyles.Add(new RowStyle(SizeType.Percent,100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute,28)); Controls.Add(root);
            var banner=new Panel { Dock=DockStyle.Fill,BackColor=Color.FromArgb(20,47,79) };
            heading=new Label { Dock=DockStyle.Top,Height=34,ForeColor=Color.White,Padding=new Padding(14,5,8,0),Font=new Font("Segoe UI",14,FontStyle.Bold),AutoEllipsis=true };
            subtitle=new Label { Dock=DockStyle.Fill,ForeColor=Color.FromArgb(207,223,241),Padding=new Padding(14,0,8,0),Font=new Font("Segoe UI",10) };
            banner.Controls.Add(subtitle); banner.Controls.Add(heading); root.Controls.Add(banner,0,0);
            var actions=new FlowLayoutPanel { Dock=DockStyle.Fill,Padding=new Padding(0,8,0,0),AutoScroll=true };
            actions.Controls.Add(new Label { Text="Đơn vị/kỳ để tạo hoặc nạp:",AutoSize=true,Padding=new Padding(0,6,0,0) });
            unit=new ComboBox { DropDownStyle=ComboBoxStyle.DropDownList,Width=255 }; foreach(var u in Catalog.Units) unit.Items.Add(u[1]); unit.SelectedIndex=0;
            month=new NumericUpDown { Minimum=1,Maximum=12,Value=DateTime.Today.Month,Width=55 };
            year=new NumericUpDown { Minimum=2000,Maximum=2100,Value=DateTime.Today.Year,Width=80 };
            actions.Controls.AddRange(new Control[]{unit,month,year});
            Button(actions,"Tạo hồ sơ trống",()=>{ if(ReplaceAllowed()) { ProjectStore.ArchiveDraft(project); project.Current=Catalog.New(SelectedCode(),(int)year.Value,(int)month.Value); dirty=true; LoadAll(); } });
            Button(actions,"Nạp Excel nguồn",Import);
            Button(actions,"Mở hồ sơ",Open);
            Button(actions,"Lưu hồ sơ",Save);
            Button(actions,"Chốt kỳ",Approve);
            Button(actions,"Xuất Word",Export);
            root.Controls.Add(actions,0,1);
            tabs=new TabControl { Dock=DockStyle.Fill,Padding=new Point(16,8) }; root.Controls.Add(tabs,0,2);
            BuildSource(); BuildInputs(); BuildWorks(); BuildMeeting(); BuildResults(); BuildHistory();
            status=new Label { Dock=DockStyle.Fill,ForeColor=Color.FromArgb(70,89,109),TextAlign=ContentAlignment.MiddleLeft }; root.Controls.Add(status,0,3);
            project=new Project { Current=Catalog.New("TK",DateTime.Today.Year,DateTime.Today.Month) }; LoadAll();
            FormClosing+=(s,e)=>
            {
                if(dirty)
                {
                    bool discard=confirmDiscard!=null?confirmDiscard():MessageBox.Show("Có thay đổi chưa lưu. Bạn muốn đóng mà không lưu?","Đóng hồ sơ",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes;
                    if(!discard) { e.Cancel=true; return; }
                }
                loading=true;
            };
        }
        string SelectedCode() { return Catalog.Units[unit.SelectedIndex][0]; }
        static DataGridView Grid()
        {
            var grid=new DataGridView { Dock=DockStyle.Fill,BackgroundColor=Color.White,BorderStyle=BorderStyle.None,AllowUserToAddRows=false,AllowUserToDeleteRows=false,AutoSizeRowsMode=DataGridViewAutoSizeRowsMode.None,SelectionMode=DataGridViewSelectionMode.FullRowSelect,RowHeadersVisible=false,EnableHeadersVisualStyles=false };
            grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(225,234,245); grid.ColumnHeadersDefaultCellStyle.ForeColor=Color.FromArgb(25,49,78); grid.ColumnHeadersDefaultCellStyle.Font=new Font("Segoe UI",10,FontStyle.Bold);
            grid.AlternatingRowsDefaultCellStyle.BackColor=Color.FromArgb(247,250,253); grid.DefaultCellStyle.Padding=new Padding(4,1,4,1); grid.RowTemplate.Height=31;
            grid.DefaultCellStyle.Font=new Font("Segoe UI",10); grid.DefaultCellStyle.ForeColor=Color.FromArgb(24,44,65); grid.DefaultCellStyle.SelectionForeColor=Color.White;
            return grid;
        }
        static void Col(DataGridView grid,string key,string name,int width,bool readOnly=true)
        { int i=grid.Columns.Add(key,name); grid.Columns[i].Width=width; grid.Columns[i].ReadOnly=readOnly; }
        TabPage Page(string title) { var p=new TabPage(title) { Padding=new Padding(10),BackColor=BackColor }; tabs.TabPages.Add(p); return p; }
        void Button(Control parent,string text,Action action)
        {
            var b=new Button { Text=text,AutoSize=true,Height=34,FlatStyle=FlatStyle.Flat,BackColor=Color.White,Margin=new Padding(4),Padding=new Padding(7,3,7,3) };
            b.Click+=(s,e)=>Run(action); parent.Controls.Add(b);
        }
        void Run(Action action)
        { try { Cursor=Cursors.WaitCursor; action(); } catch(Exception ex) { MessageBox.Show(ex.Message,"Kiểm tra hồ sơ",MessageBoxButtons.OK,MessageBoxIcon.Warning); } finally { Cursor=Cursors.Default; } }
        void BuildSource()
        {
            var page=Page("1  Nguồn & đường dẫn"); var top=new FlowLayoutPanel { Dock=DockStyle.Top,Height=48 };
            Button(top,"Chọn thư mục file tham chiếu",()=>{ using(var d=new FolderBrowserDialog { Description="Chọn thư mục chứa file công/ăn ca liên kết. Chỉ kiểm tra tên, không chạy lại công thức." }) if(d.ShowDialog()==DialogResult.OK) { dependencyFolder=d.SelectedPath; status.Text="Đã chọn thư mục tham chiếu. Nạp lại Excel để kiểm tra liên kết."; } });
            Button(top,"Gắn lại Excel đã di chuyển",()=>
            {
                var d=project.Current;
                if(string.IsNullOrEmpty(d.SourceHash)) throw new InvalidOperationException("Hồ sơ chưa có nguồn. Dùng Nạp Excel nguồn.");
                using(var dialog=new OpenFileDialog { Filter="Excel (*.xlsx)|*.xlsx" }) if(dialog.ShowDialog()==DialogResult.OK)
                {
                    if(Workbook.Hash(dialog.FileName)!=d.SourceHash) throw new InvalidOperationException("Tệp này khác nội dung nguồn đã nhập. Dùng Nạp Excel nguồn để đối chiếu lại, không gắn thay bằng tệp khác.");
                    d.SourcePath=Path.GetFullPath(dialog.FileName); d.Note("Gắn lại đường dẫn","Tệp trùng SHA256; không đổi số liệu"); dirty=true; RefreshViews();
                }
            });
            Button(top,"Đọc kiểm tra Word mẫu",InspectTemplate);
            Button(top,"Xuất nhật ký đối chiếu CSV",ExportAudit);
            diagnostic=new RichTextBox { Dock=DockStyle.Fill,ReadOnly=true,BackColor=Color.White,BorderStyle=BorderStyle.None,Font=new Font("Segoe UI",10) }; page.Controls.Add(diagnostic); page.Controls.Add(top);
        }
        void BuildInputs()
        {
            var page=Page("2  Đầu vào & đơn giá"); var top=new FlowLayoutPanel { Dock=DockStyle.Top,Height=70 };
            top.Controls.Add(new Label { Text="Chỉ sửa cột Đầu vào. Ô trống là CHƯA BIẾT, không phải 0. Giá lễ = 2 × giá L1, riêng cho TT và QL.",AutoSize=true,Padding=new Padding(3,7,3,2) });
            Button(top,"Ghi 0 cho các dòng đã chọn",()=>{ if(inputs.SelectedRows.Count==0) throw new InvalidOperationException("Chọn ít nhất một dòng đầu vào."); Commit(); foreach(DataGridViewRow row in inputs.SelectedRows) project.Current.Change(((Measure)row.Tag).Key,0); Mark(true); LoadInputs(); RefreshViews(); });
            inputs=Grid(); Col(inputs,"label","Chỉ tiêu",420); Col(inputs,"value","Đầu vào",145,false); Col(inputs,"unit","ĐVT",135); Col(inputs,"source","Giá trị nguồn",145); Col(inputs,"state","Trạng thái",260); Col(inputs,"cell","Ô nguồn / đường tính",220);
            inputs.CellValidating+=(s,e)=>{ if(loading || e.ColumnIndex!=1) return; decimal? value; if(!Fmt.Parse(Convert.ToString(e.FormattedValue),out value)) { e.Cancel=true; inputs.Rows[e.RowIndex].ErrorText="Nhập số theo Việt Nam, ví dụ 1.234,5. Để trống nếu chưa có."; } else inputs.Rows[e.RowIndex].ErrorText=""; };
            inputs.CellEndEdit+=(s,e)=>{ if(loading || e.RowIndex<0 || e.ColumnIndex!=1) return; decimal? value; if(Fmt.Parse(Convert.ToString(inputs.Rows[e.RowIndex].Cells[1].Value),out value)) { project.Current.Change(((Measure)inputs.Rows[e.RowIndex].Tag).Key,value); inputs.Rows[e.RowIndex].Cells[4].Value="Đã nhập/chỉnh"; inputs.Rows[e.RowIndex].DefaultCellStyle.BackColor=Color.White; Mark(true); RefreshViews(); } };
            page.Controls.Add(inputs); page.Controls.Add(top);
        }
        void BuildWorks()
        {
            var page=Page("3  GNV & sản phẩm phụ"); var top=new FlowLayoutPanel { Dock=DockStyle.Top,Height=75 };
            top.Controls.Add(new Label { Text="Thành tiền = KL × Đơn giá × Hệ số, làm tròn theo cột cuối. Số tiền nguồn chỉ để so sánh, không chép làm kết quả.",AutoSize=true,Padding=new Padding(3,7,3,2) });
            Button(top,"Thêm GNV",()=>{ Commit(); project.Current.Works.Add(new WorkLine()); project.Current.Note("Thêm GNV","Tạo dòng trống"); Mark(true); LoadWorks(); RefreshViews(); });
            Button(top,"Xóa dòng đã chọn",()=>{ if(works.SelectedRows.Count==0) throw new InvalidOperationException("Chọn dòng cần xóa."); Commit(); foreach(DataGridViewRow row in works.SelectedRows.Cast<DataGridViewRow>().ToArray()) project.Current.Works.Remove((WorkLine)row.Tag); project.Current.Note("Xóa GNV","Xóa các dòng đã chọn"); Mark(true); LoadWorks(); RefreshViews(); });
            works=Grid(); Col(works,"text","Nội dung",300,false); Col(works,"unit","ĐVT",80,false); Col(works,"qty","Khối lượng / công",130,false); Col(works,"price","Đơn giá (đ)",130,false); Col(works,"factor","Hệ số (+/−)",110,false); Col(works,"round","Làm tròn (đ)",110,false); Col(works,"amount","Thành tiền (đ)",145); Col(works,"source","Tiền nguồn (đ)",145); Col(works,"cell","Ô nguồn",170);
            works.Columns.Add(new DataGridViewComboBoxColumn { Name="kind",HeaderText="Loại khoản",Width=115,DataSource=new[]{"GNV","SP phụ"} });
            works.CellValidating+=(s,e)=>{ if(loading || e.ColumnIndex<2 || e.ColumnIndex>5) return; decimal? value; if(!Fmt.Parse(Convert.ToString(e.FormattedValue),out value) || (e.ColumnIndex>=4 && (!value.HasValue || (e.ColumnIndex==5 && value<=0)))) { e.Cancel=true; works.Rows[e.RowIndex].ErrorText="Công/giá có thể để trống; hệ số phải là số, mức làm tròn phải > 0."; } else works.Rows[e.RowIndex].ErrorText=""; };
            works.CellEndEdit+=(s,e)=>{ if(loading || e.RowIndex<0) return; var row=works.Rows[e.RowIndex]; var w=(WorkLine)row.Tag; w.Description=Convert.ToString(row.Cells[0].Value)??""; w.Unit=Convert.ToString(row.Cells[1].Value)??""; w.Kind=Convert.ToString(row.Cells[9].Value)??"GNV"; decimal? value;
                Fmt.Parse(Convert.ToString(row.Cells[2].Value),out value); w.Quantity=value; Fmt.Parse(Convert.ToString(row.Cells[3].Value),out value); w.Price=value;
                Fmt.Parse(Convert.ToString(row.Cells[4].Value),out value); w.Factor=value??1; Fmt.Parse(Convert.ToString(row.Cells[5].Value),out value); w.Rounding=value??1;
                row.Cells[6].Value=w.Amount.HasValue?Fmt.Number(w.Amount):"Chưa đủ"; project.Current.Note("Sửa GNV",w.Description+": KL="+Fmt.Number(w.Quantity)+", giá="+Fmt.Number(w.Price)+", hệ số="+w.Factor); Mark(true); RefreshViews(); };
            page.Controls.Add(works); page.Controls.Add(top);
        }
        void BuildMeeting()
        {
            var page=Page("4  Cuộc họp & giải trình"); var table=new TableLayoutPanel { Dock=DockStyle.Fill,ColumnCount=2,RowCount=11,AutoScroll=true };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,270)); table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            meeting=new DateTimePicker { Format=DateTimePickerFormat.Custom,CustomFormat="dd/MM/yyyy",Width=170 };
            location=new TextBox(); office=new TextBox(); team=new TextBox(); signOffice=new TextBox(); signTeam=new TextBox(); quantityUnit=new TextBox();
            recommend=new TextBox { Multiline=true,Height=90,ScrollBars=ScrollBars.Vertical }; explanation=new TextBox { Multiline=true,Height=110,ScrollBars=ScrollBars.Vertical };
            reviewed=new CheckBox { Text="Tôi đã đối chiếu đầu vào, đơn giá và các chênh lệch với nguồn; đồng ý chốt các số liệu này.",AutoSize=true };
            Control[] fields={meeting,location,quantityUnit,office,team,signOffice,signTeam,recommend,explanation,reviewed};
            string[] labels={"Ngày nghiệm thu","Địa điểm","Đơn vị sản lượng/cơ sở tính","Thành phần phòng TCHC","Thành phần phân xưởng","Người ký phòng TCHC","Người ký phân xưởng","Kiến nghị, đề xuất","Giải trình chênh lệch / ngoại lệ","Xác nhận đối chiếu"};
            for(int i=0;i<fields.Length;i++) { table.Controls.Add(new Label { Text=labels[i],AutoSize=true,Padding=new Padding(2,7,4,0) },0,i); fields[i].Dock=DockStyle.Top; fields[i].Margin=new Padding(3,4,3,7); table.Controls.Add(fields[i],1,i); }
            page.Controls.Add(table);
        }
        void BuildResults()
        {
            var page=Page("5  Kết quả & biên bản"); var split=new SplitContainer { Dock=DockStyle.Fill,Orientation=Orientation.Horizontal,SplitterDistance=265 };
            results=Grid(); Col(results,"label","Chỉ tiêu",330); Col(results,"raw","Kết quả tính (đ/công/người)",220); Col(results,"word","Trình bày Word",210); Col(results,"source","Nguồn tham chiếu (đ)",220); Col(results,"difference","Chênh nguồn (đ)",200);
            var bottom=new Panel { Dock=DockStyle.Fill }; cumulative=new CheckBox { Dock=DockStyle.Top,Height=35,Text="Kèm quý và lũy kế (chỉ dùng kỳ đã chốt; thiếu tháng thì báo thiếu, không bù số)" }; cumulative.CheckedChanged+=(s,e)=>{ if(!loading) RefreshViews(); };
            preview=new RichTextBox { Dock=DockStyle.Fill,ReadOnly=true,BackColor=Color.White,Font=new Font("Segoe UI",10),BorderStyle=BorderStyle.None }; bottom.Controls.Add(preview); bottom.Controls.Add(cumulative);
            split.Panel1.Controls.Add(results); split.Panel2.Controls.Add(bottom); page.Controls.Add(split);
        }
        void BuildHistory()
        {
            var page=Page("6  Hồ sơ & nhật ký"); var split=new SplitContainer { Dock=DockStyle.Fill,Orientation=Orientation.Horizontal,SplitterDistance=240 };
            history=Grid(); Col(history,"state","Trạng thái",105); Col(history,"unit","Phân xưởng",260); Col(history,"period","Kỳ",90); Col(history,"date","Chốt lúc",180); Col(history,"total","Tổng quỹ (triệu đồng)",210); Col(history,"hash","Mã chốt",370);
            var panel=new Panel { Dock=DockStyle.Fill }; var top=new FlowLayoutPanel { Dock=DockStyle.Top,Height=45 }; Button(top,"Mở bản chốt để xem / tạo bản sửa",()=>{ if(history.CurrentRow==null || !ReplaceAllowed()) return; var d=history.CurrentRow.Tag as Dossier; if(d==null) return; ProjectStore.ArchiveDraft(project); project.Current=ProjectStore.Copy(d); dirty=true; LoadAll(); }); panel.Controls.Add(history); panel.Controls.Add(top);
            log=new RichTextBox { Dock=DockStyle.Fill,ReadOnly=true,BackColor=Color.White,Font=new Font("Segoe UI",9),BorderStyle=BorderStyle.None }; split.Panel1.Controls.Add(panel); split.Panel2.Controls.Add(log); page.Controls.Add(split);
        }
        void Mark(bool requireReview=false) { dirty=true; project.Current.ApprovedHash=null; if(requireReview) { project.Current.SourceReviewed=false; reviewed.Checked=false; } }
        void Commit()
        {
            ValidateChildren(); inputs.EndEdit(); works.EndEdit();
            if(inputs.Rows.Cast<DataGridViewRow>().Any(r=>r.ErrorText!="") || works.Rows.Cast<DataGridViewRow>().Any(r=>r.ErrorText!="")) throw new InvalidOperationException("Có ô đang nhập sai định dạng. Sửa ô báo lỗi trước khi lưu/chốt/xuất.");
            var d=project.Current; string before=d.Fingerprint();
            d.MeetingDate=meeting.Value.Date; d.Location=location.Text; d.QuantityUnit=quantityUnit.Text; d.AttendeesOffice=office.Text; d.AttendeesUnit=team.Text; d.SignerOffice=signOffice.Text; d.SignerUnit=signTeam.Text; d.Recommendations=recommend.Text; d.Explanation=explanation.Text; d.SourceReviewed=reviewed.Checked;
            if(before!=d.Fingerprint()) { d.Note("Sửa thông tin hồ sơ","Cập nhật thông tin cuộc họp/giải trình/xác nhận"); Mark(); }
        }
        bool ReplaceAllowed() { Commit(); return !dirty || MessageBox.Show("Hồ sơ hiện tại chưa lưu. Tiếp tục sẽ thay phần đang mở; các bản chốt đã lưu vẫn giữ. Bạn đã lưu phần cần giữ chưa?","Thay hồ sơ đang mở",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes; }
        void Import()
        {
            if(!ReplaceAllowed()) return;
            using(var dialog=new OpenFileDialog { Filter="Excel nguồn (*.xlsx)|*.xlsx",Title="Nạp nguồn chỉ đọc — không ghi ngược Excel" })
                if(dialog.ShowDialog()==DialogResult.OK)
                {
                    var d=Workbook.Import(dialog.FileName,SelectedCode(),(int)year.Value,(int)month.Value,dependencyFolder);
                    string mismatch=d.Diagnostics.FirstOrDefault(x=>x.StartsWith("KỲ NGUỒN"));
                    if(mismatch!=null)
                    {
                        var match=Regex.Match(mismatch,@"(\d{1,2})/(20\d{2})");
                        var answer=MessageBox.Show(mismatch+"\nDùng kỳ nguồn thay kỳ vừa chọn? Yes = dùng kỳ nguồn; No = giữ kỳ chọn và phải giải trình.","Kiểm tra kỳ nguồn",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2);
                        if(answer==DialogResult.Cancel) return;
                        if(answer==DialogResult.Yes && match.Success) { month.Value=int.Parse(match.Groups[1].Value); year.Value=int.Parse(match.Groups[2].Value); d=Workbook.Import(dialog.FileName,SelectedCode(),(int)year.Value,(int)month.Value,dependencyFolder); }
                    }
                    ProjectStore.ArchiveDraft(project); project.Current=d; dirty=true; LoadAll(); tabs.SelectedIndex=0;
                }
        }
        void Open()
        {
            if(!ReplaceAllowed()) return; using(var dialog=new OpenFileDialog { Filter="Hồ sơ nghiệm thu (*.ntproject)|*.ntproject" }) if(dialog.ShowDialog()==DialogResult.OK) { project=ProjectStore.Read(dialog.FileName); projectPath=dialog.FileName; dirty=false; LoadAll(); }
        }
        bool ChooseProjectPath()
        {
            if(!string.IsNullOrEmpty(projectPath)) return true;
            using(var dialog=new SaveFileDialog { Filter="Hồ sơ nghiệm thu (*.ntproject)|*.ntproject",FileName="HoSoNghiemThu.ntproject",DefaultExt="ntproject",AddExtension=true })
                if(dialog.ShowDialog()==DialogResult.OK) { projectPath=dialog.FileName; return true; } return false;
        }
        void Save() { Commit(); if(!ChooseProjectPath()) return; ProjectStore.Save(projectPath,project); dirty=false; RefreshViews(); status.Text="Đã lưu hồ sơ; lần trước được giữ trong .bak. Excel nguồn không thay đổi."; }
        void Approve()
        {
            Commit(); var problems=ProjectStore.ApprovalProblems(project.Current);
            if(problems.Count>0) { tabs.SelectedIndex=4; throw new InvalidOperationException(string.Join("\n",problems.Take(20))); }
            if(!ChooseProjectPath()) return; var backup=ProjectStore.Copy(project);
            try { ProjectStore.Approve(project); ProjectStore.Save(projectPath,project); dirty=false; LoadAll(); status.Text="Đã chốt và lưu ảnh chụp kỳ. Sửa đầu vào sau đó sẽ trở thành bản nháp mới."; }
            catch { project=backup; LoadAll(); throw; }
        }
        void Export()
        {
            Commit(); using(var dialog=new SaveFileDialog { Filter="Biên bản Word (*.docx)|*.docx",DefaultExt="docx",AddExtension=true,FileName="BBNT_"+project.Current.Code+"_"+project.Current.Year+"_"+project.Current.Month.ToString("00")+".docx" })
                if(dialog.ShowDialog()==DialogResult.OK)
                {
                    if(!string.IsNullOrEmpty(templatePath) && string.Equals(Path.GetFullPath(templatePath),Path.GetFullPath(dialog.FileName),StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Không ghi đè Word mẫu tham chiếu. Chọn tên tệp đầu ra khác.");
                    Report.Export(project,dialog.FileName,cumulative.Checked); project.Current.Note("Xuất Word",Path.GetFileName(dialog.FileName)); dirty=true; RefreshViews(); status.Text="Đã xuất "+dialog.FileName+". Báo cáo thiếu dữ liệu/kỳ chốt có nhãn DỰ THẢO.";
                }
        }
        void InspectTemplate()
        {
            using(var dialog=new OpenFileDialog { Filter="Word mẫu tham chiếu (*.docx)|*.docx",Title="Chỉ kiểm tra cấu trúc, không sao chép số liệu cũ hay ghi đè mẫu" }) if(dialog.ShowDialog()==DialogResult.OK)
            {
                templatePath=dialog.FileName;
                project.Current.Diagnostics.AddRange(TemplateInspector.Describe(templatePath)); Mark(); RefreshViews(); tabs.SelectedIndex=0;
            }
        }
        void ExportAudit()
        {
            Commit(); using(var dialog=new SaveFileDialog { Filter="Bảng đối chiếu CSV (*.csv)|*.csv",FileName="DoiChieu_"+project.Current.Code+".csv" }) if(dialog.ShowDialog()==DialogResult.OK)
            {
                ProjectStore.GuardOutput(project,dialog.FileName,".csv");
                using(var writer=new StreamWriter(dialog.FileName,false,new UTF8Encoding(true)))
                { writer.WriteLine("Chỉ tiêu;Đầu vào;Nguồn;Ô nguồn;Công thức;Trạng thái"); foreach(var m in project.Current.Inputs) writer.WriteLine(string.Join(";",new[]{m.Label,Fmt.Number(m.Value),Fmt.Number(m.SourceValue),m.SourceCell,m.Formula,m.SourceState}.Select(Csv))); }
                status.Text="Đã xuất bảng đối chiếu.";
            }
        }
        static string Csv(string text) { text=text??""; if(text.StartsWith("=") || text.StartsWith("+") || text.StartsWith("-") || text.StartsWith("@")) text="'"+text; return "\""+text.Replace("\"","\"\"")+"\""; }
        void LoadInputs()
        {
            bool old=loading; loading=true; inputs.Rows.Clear();
            foreach(var m in project.Current.Inputs) { int i=inputs.Rows.Add(m.Label,m.Value.HasValue?Fmt.Number(m.Value):"",m.Unit,Fmt.Number(m.SourceValue),m.Manual?"Đã nhập/chỉnh — giữ nguồn để đối chiếu":m.SourceState,m.SourceCell); inputs.Rows[i].Tag=m; if(!m.Value.HasValue) inputs.Rows[i].DefaultCellStyle.BackColor=Color.FromArgb(255,244,224); }
            loading=old;
        }
        void LoadWorks()
        {
            bool old=loading; loading=true; works.Rows.Clear(); foreach(var w in project.Current.Works) { int i=works.Rows.Add(w.Description,w.Unit,w.Quantity.HasValue?Fmt.Number(w.Quantity):"",w.Price.HasValue?Fmt.Number(w.Price):"",Fmt.Number(w.Factor),Fmt.Number(w.Rounding),Fmt.Number(w.Amount),Fmt.Number(w.SourceAmount),w.SourceCell,w.Kind); works.Rows[i].Tag=w; } loading=old;
        }
        void LoadAll()
        {
            loading=true; var d=project.Current;
            unit.SelectedIndex=Array.FindIndex(Catalog.Units,x=>x[0]==d.Code); year.Value=d.Year; month.Value=d.Month;
            LoadInputs(); LoadWorks();
            meeting.Value=d.MeetingDate; location.Text=d.Location; quantityUnit.Text=d.QuantityUnit; office.Text=d.AttendeesOffice; team.Text=d.AttendeesUnit; signOffice.Text=d.SignerOffice; signTeam.Text=d.SignerUnit; recommend.Text=d.Recommendations; explanation.Text=d.Explanation; reviewed.Checked=d.SourceReviewed;
            loading=false; RefreshViews();
        }
        void RefreshViews()
        {
            if(project==null) return; var d=project.Current; var r=Calc.Evaluate(d); bool old=loading; loading=true;
            heading.Text="NGHIỆM THU STUDIO 1.0  •  "+d.Name;
            subtitle.Text=d.Month.ToString("00")+"/"+d.Year+"  |  "+(d.Approved?"ĐÃ CHỐT":"BẢN NHÁP")+"  •  Nguồn → đối chiếu → chốt kỳ → Word";
            results.Rows.Clear(); foreach(string key in new[]{"tt1","tt2","tth","productT","ql1","ql2","qlh","productM","extraProduct","gnv","regimeT","adjustT","direct","regimeM","adjustM","management","total","meal","bddh","days","peopleMonths","average"})
            {
                decimal? source; d.References.TryGetValue(key,out source); decimal? value=r.Get(key); bool money=key!="days"&&key!="peopleMonths";
                results.Rows.Add(Calc.Labels(key),Fmt.Number(value),money?Fmt.Money(value)+" triệu đồng"+(key=="average"?"/người/tháng":""):Fmt.Number(value),source.HasValue?Fmt.Number(source):"Không có",source.HasValue&&value.HasValue?Fmt.Number(value-source):"—");
            }
            diagnostic.Text="NGUỒN CHỈ ĐỌC\n"+(d.SourcePath??"Chưa nạp Excel; nhập tay có đối chiếu")+"\nSHA256: "+(d.SourceHash??"—")+"\n\n"+string.Join("\n",d.Diagnostics)+"\n\nCẦN HOÀN THIỆN\n"+string.Join("\n",r.Issues);
            preview.Text=Report.Preview(project,cumulative.Checked);
            history.Rows.Clear(); foreach(var item in project.Closed.Concat(project.Drafts.Where(x=>x.Id!=d.Id)).Reverse())
            {
                decimal? total=item.Approved && item.ApprovedResult!=null ? item.ApprovedResult["total"] : Calc.Evaluate(item).Get("total");
                int i=history.Rows.Add(item.Approved?"Đã chốt":"Nháp",item.Name,item.Month+"/"+item.Year,item.ApprovedAt??(item.Log.LastOrDefault()==null?"—":item.Log.Last().Time),Fmt.Money(total),item.ApprovedHash??("Nháp "+item.Id.Substring(0,8))); history.Rows[i].Tag=item;
            }
            log.Text=string.Join("\n",d.Log.Select(x=>x.Time+"  "+x.Action+"  "+x.Detail));
            status.Text=(dirty?"Có thay đổi chưa lưu. ":"")+r.Issues.Count+" mục tính toán cần đối chiếu. "+(d.SourceReviewed?"":"Chưa xác nhận đối chiếu. ")+(projectPath??"Lưu .ntproject để giữ đầu vào, kỳ chốt và lịch sử."); loading=old;
        }
    }
}
