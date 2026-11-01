using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace CiscoLDTLApp
{
    public class ManagementPayrollForm : Form
    {
        private WorkshopData original, draft;
        private List<EmployeeRecord> originals = new List<EmployeeRecord>();
        private DataGridView grid;
        private NumericUpDown fund, adjustment;
        private Label summary;
        private BindingSource source;

        public ManagementPayrollForm(WorkshopData ws)
        {
            original = ws;
            draft = new WorkshopData { MaPX = ws.MaPX, LuongSPQL = ws.LuongSPQL, DieuChinhLuongQL = ws.DieuChinhLuongQL };
            draft.ListNhanVien.Clear();
            foreach (var e in ws.ListNhanVien)
                if (e.LaQuanLy)
                {
                    var copy = new EmployeeRecord();
                    ManagementPayroll.Copy(e, copy);
                    draft.ListNhanVien.Add(copy); originals.Add(e);
                }
            Text = "Tính lương quản lý phân xưởng — " + ws.TenPX;
            Size = new Size(1250, 700); MinimumSize = new Size(960, 580);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9);
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 75, Padding = new Padding(8), AutoScroll = true };
            top.Controls.Add(new Label { Text = "Quỹ SP quản lý (đ)", AutoSize = true, Padding = new Padding(0, 6, 0, 0) });
            fund = Money(Math.Max(0, ws.LuongSPQL), false); top.Controls.Add(fund);
            top.Controls.Add(new Label { Text = "Điều chỉnh chung (+/−)", AutoSize = true, Padding = new Padding(0, 6, 0, 0) });
            adjustment = Money(ws.DieuChinhLuongQL, true); top.Controls.Add(adjustment);
            var calc = new Button { Text = "Tính / xem trước", AutoSize = true };
            calc.Click += (s,e) => Preview(false); top.Controls.Add(calc);
            var apply = new Button { Text = "Áp dụng vào bảng nghiệm thu", AutoSize = true };
            apply.Click += (s,e) => Preview(true); top.Controls.Add(apply);
            var cancel = new Button { Text = "Đóng", AutoSize = true, CausesValidation = false };
            cancel.Click += (s,e) => Close(); top.Controls.Add(cancel);
            var note = new Label { Dock = DockStyle.Top, Height = 120, Padding = new Padding(8), Text =
                (string.IsNullOrEmpty(ws.CanhBaoQuyQuanLy) ? "Kiểm tra quỹ SP quản lý được duyệt trước khi phân bổ." : ws.CanhBaoQuyQuanLy) + "\n" +
                "Chỉ gồm người được đánh dấu quản lý của phân xưởng đang chọn.\n" +
                "Hệ số hưởng = ROUND(ROUND(Công SP × Hệ số / 26; 2) × Xếp loại; 2); A=100%, B=90%, C=80%. Lương SP phân bổ theo hệ số hưởng, làm tròn 100đ.\n" +
                "Phép = Công phép × Lương BHXH / 26 (tròn 100đ); Học/họp 250.000đ/công; VHTT 300.000đ/công theo quy tắc bảng NT hiện có. Làm lễ và phụ cấp: nhập số tiền đã duyệt.\n" +
                "Quỹ SP lấy từ bảng NT/nhập tay, không tự tính lại các công thức sản lượng hoặc liên kết ngoài. Điều chỉnh chung chỉ cộng ở tổng PX, chưa phân bổ cho cá nhân." };
            summary = new Label { Dock = DockStyle.Bottom, Height = 82, Padding = new Padding(8), Text = "Nhập số liệu, bấm Tính / xem trước. Chưa áp dụng hoặc ghi Excel." };
            grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
            Add("SDB", "SDB", true); Add("HoTen", "Họ tên", true, 180); Add("ChucDanh", "Chức danh", true, 120);
            Add("HeSoCoDinh", "Hệ số"); Add("CongSP", "Công SP");
            var grade = new DataGridViewComboBoxColumn { DataPropertyName = "XepLoai", HeaderText = "Xếp loại", Width = 75 };
            grade.Items.AddRange("A", "B", "C"); grid.Columns.Add(grade);
            Add("MucLuongBHXH", "Lương BHXH"); Add("CongLePhep", "Công phép"); Add("CongHocHop", "Công học/họp"); Add("CongVHTT", "Công VHTT");
            Add("LuongLamLe", "Tiền làm lễ"); Add("PhuCapTien", "Tiền phụ cấp");
            Add("HeSoHuong", "HS hưởng", true); Add("LuongSanPham", "Lương SP", true);
            Add("LuongLePhepCD", "Lương phép", true); Add("LuongHocVHTT", "Lương học/VHTT", true); Add("TongTienLuong", "Tổng cá nhân", true);
            foreach (var e in draft.ListNhanVien) e.XepLoai = (e.XepLoai ?? "A").Trim().ToUpperInvariant();
            source = new BindingSource { DataSource = new BindingList<EmployeeRecord>(draft.ListNhanVien) };
            grid.DataSource = source;
            grid.DataError += (s,e) => { e.ThrowException = false; e.Cancel = true; MessageBox.Show("Giá trị không hợp lệ. Hãy nhập số theo định dạng của Windows; xếp loại A/B/C."); };
            Controls.Add(grid); Controls.Add(note); Controls.Add(top); Controls.Add(summary);
        }
        private static NumericUpDown Money(double value, bool signed)
        { return new NumericUpDown { Width = 155, Maximum = 1000000000000000m, Minimum = signed ? -1000000000000000m : 0, ThousandsSeparator = true, Value = (decimal)value }; }
        private void Add(string field, string title, bool readOnly = false, int width = 100)
        { grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = field, HeaderText = title, ReadOnly = readOnly, Width = width }); }
        private void Preview(bool apply)
        {
            try
            {
                if (!grid.EndEdit()) return;
                source.EndEdit();
                if (draft.ListNhanVien.Count == 0) throw new ArgumentException("Phân xưởng chưa có người được đánh dấu quản lý. Bổ sung/chỉnh hồ sơ trước.");
                draft.LuongSPQL = (double)fund.Value; draft.DieuChinhLuongQL = (double)adjustment.Value;
                ManagementPayroll.Recalculate(draft);
                source.ResetBindings(false);
                double paid = 0, total = 0;
                foreach (var e in draft.ListNhanVien) { paid += e.LuongSanPham; total += e.TongTienLuong; }
                summary.Text = string.Format("{0} người | Quỹ SP: {1:N0}đ | SP đã phân bổ: {2:N0}đ | Chênh làm tròn: {3:N0}đ\nTổng cá nhân: {4:N0}đ | Khoản ngoài SP: {5:N0}đ | Điều chỉnh chung: {6:N0}đ | Tổng quản lý trên NT: {7:N0}đ\nChênh làm tròn và điều chỉnh chung được hiển thị riêng; bấm Lưu Excel ở màn hình chính để lưu lâu dài.",
                    originals.Count, draft.LuongSPQL, paid, draft.LuongSPQL - paid, total, draft.LuongCheDoQL, draft.DieuChinhLuongQL, draft.TongLuongQL);
                if (apply)
                {
                    for (int i=0; i<originals.Count; i++)
                    { draft.ListNhanVien[i].QuanLyTinhMoi = true; ManagementPayroll.Copy(draft.ListNhanVien[i], originals[i]); }
                    original.LuongSPQL = draft.LuongSPQL; original.LuongCheDoQL = draft.LuongCheDoQL;
                    original.DieuChinhLuongQL = draft.DieuChinhLuongQL;
                    ManagementPayroll.Recalculate(original);
                    DialogResult = DialogResult.OK; Close();
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Kiểm tra lương quản lý", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
    }
}
