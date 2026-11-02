using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace CiscoLDTLApp
{
    public static class EmployeeRoster
    {
        private static string Code(string value)
        {int n;string s=(value??"").Trim();return int.TryParse(s,out n)?n.ToString():s.ToUpperInvariant();}
        public static void Validate(IEnumerable<EmployeeRecord> people, EmployeeRecord employee)
        {
            if(string.IsNullOrWhiteSpace(employee.SDB) || string.IsNullOrWhiteSpace(employee.HoTen) || string.IsNullOrWhiteSpace(employee.ToSanXuat) || string.IsNullOrWhiteSpace(employee.ChucDanh))
                throw new ArgumentException("Vui lòng nhập mã SDB, họ tên, tổ/bộ phận và chức danh.");
            foreach(var e in people) if(Code(e.SDB)==Code(employee.SDB))
                throw new ArgumentException("Mã SDB đã tồn tại trong phân xưởng. Kiểm tra người cũ để tránh trùng công / lương.");
            foreach(var n in new[]{employee.MucLuongBHXH,employee.HeSoCoDinh,employee.PhuCapChucVu})
                if(n<0 || double.IsNaN(n) || double.IsInfinity(n))throw new ArgumentException("Lương và hệ số phải là số không âm, hữu hạn.");
        }
        public static void Add(WorkshopData ws, EmployeeRecord employee)
        {
            Validate(ws.ListNhanVien,employee);
            employee.SDB=employee.SDB.Trim();employee.HoTen=employee.HoTen.Trim();employee.ToSanXuat=employee.ToSanXuat.Trim();
            employee.BoPhan=employee.ToSanXuat;employee.ChucDanh=employee.ChucDanh.Trim();employee.ExcelRow=employee.TongHopCongRow=0;
            employee.STT=ws.ListNhanVien.Count+1;employee.XepLoai="A";
            EmployeeClassification.Apply(employee);ws.ListNhanVien.Add(employee);EmployeeClassification.Apply(ws);
        }
    }
    public class EmployeeEditorForm : Form
    {
        public EmployeeRecord Employee { get; private set; }
        private TextBox code,name,team,role;
        private NumericUpDown salary,factor,allowance;
        public EmployeeEditorForm(WorkshopData ws, bool managementOnly)
        {
            Text="Thêm người — "+ws.TenPX;Size=new Size(570,410);StartPosition=FormStartPosition.CenterParent;
            Font=new Font("Segoe UI",9);FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;MinimizeBox=false;
            var table=new TableLayoutPanel { Dock=DockStyle.Fill,ColumnCount=2,RowCount=9,Padding=new Padding(12) };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,175));table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            code=Box();name=Box();team=Box();role=Box();salary=Number();factor=Number();factor.DecimalPlaces=4;allowance=Number();allowance.DecimalPlaces=4;
            string[] labels={"Mã SDB","Họ và tên","Tổ / Bộ phận","Chức danh","Mức lương BHXH (đ)","Hệ số cố định","Hệ số phụ cấp chức vụ"};
            Control[] fields={code,name,team,role,salary,factor,allowance};
            for(int i=0;i<fields.Length;i++){table.Controls.Add(new Label {Text=labels[i],AutoSize=true},0,i);table.Controls.Add(fields[i],1,i);}
            if(managementOnly){team.Text="Quản lý PX";role.Text="KTV";}
            table.Controls.Add(new Label { Text="Quản lý chỉ gồm QĐ, PQĐ, KTV, TK. Sau khi thêm, nhập công ở bảng nhân sự hoặc nạp từ file công; bấm Lưu Excel để lưu hồ sơ.",AutoSize=true,MaximumSize=new Size(510,0) },0,7);table.SetColumnSpan(table.GetControlFromPosition(0,7),2);
            var buttons=new FlowLayoutPanel { Dock=DockStyle.Fill,AutoSize=true };var save=new Button {Text="Thêm người",AutoSize=true};var cancel=new Button {Text="Hủy",DialogResult=DialogResult.Cancel};buttons.Controls.Add(save);buttons.Controls.Add(cancel);table.Controls.Add(buttons,1,8);
            save.Click+=delegate {try {
                var e=new EmployeeRecord {SDB=code.Text,HoTen=name.Text,ToSanXuat=team.Text,ChucDanh=role.Text,MucLuongBHXH=(double)salary.Value,HeSoCoDinh=(double)factor.Value,PhuCapChucVu=(double)allowance.Value};
                EmployeeRoster.Validate(ws.ListNhanVien,e);
                if(managementOnly && !EmployeeClassification.IsManagementRole(e.ChucDanh))throw new ArgumentException("Cửa sổ này chỉ thêm QĐ, PQĐ, KTV, TK. Thêm lao động trực tiếp tại bảng nhân sự.");
                Employee=e;DialogResult=DialogResult.OK;
            } catch(ArgumentException ex){MessageBox.Show(this,ex.Message,"Kiểm tra hồ sơ");} };
            Controls.Add(table);AcceptButton=save;CancelButton=cancel;
        }
        private static TextBox Box(){return new TextBox {Dock=DockStyle.Top};}
        private static NumericUpDown Number(){return new NumericUpDown {Dock=DockStyle.Top,Maximum=1000000000000m,ThousandsSeparator=true};}
    }
}
