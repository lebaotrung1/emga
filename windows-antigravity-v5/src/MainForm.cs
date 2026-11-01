using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace CiscoLDTLApp
{
    public class MainForm : Form
    {
        private List<WorkshopData> _workshops;
        private int _currentIndex = 0;
        private bool _isLoadingUI = false;
        private bool _isFilteringEmp = false;
        private string _excelFilePath;

        // Top controls
        private ComboBox cboWorkshop;
        private NumericUpDown numThang;
        private NumericUpDown numNam;
        private ComboBox cboLoaiKy;
        private CheckBox chkTuDongLuyKe;
        private TextBox txtExcelPath;
        private Label lblStatus;
        private TabControl tabMain;

        // Tab 1: Bảng Công, Ăn Ca & Lương Chi Tiết Từng Người
        private ComboBox cboFilterBoPhan;
        private TextBox txtSearchEmp;
        private Label lblEmpSummaryBanner;
        private DataGridView dgvEmployees;
        private CheckBox chkOnlyAnCaDiff;

        // Tab 2: Nhân Sự, Biến Động LĐ & Tổng Hợp Ngày Công
        private NumericUpDown numNgayHop, numThangHop, numNamHop;
        private TextBox txtDiaDiem, txtGioKetThuc;
        private DataGridView dgvThanhPhanTCHC, dgvThanhPhanPX;
        private NumericUpDown numLDDauKy, numLDTang, numLDGiam, numLDCuoiKy, numLDQuanLy, numLDTrucTiep;
        private TextBox txtLyDoBienDong, txtGhiChuLD;
        private DataGridView dgvNgayCong;

        // Tab 3: Sản lượng, Tiền lương, Ăn ca, Độc hại
        private TextBox txtTenSP, txtDvtSP, txtSanLuongPhuName, txtLyDoTruLuong;
        private TextBox txtSLKeHoach, txtSLThucHien, txtSLTinhLuong, txtSLLoai1, txtDGLoai1, txtTTLoai1;
        private TextBox txtSLLoai2, txtDGLoai2, txtTTLoai2, txtSLLe, txtDGLe, txtTTLe, txtSLPhuKL, txtSLPhuTT;
        private DataGridView dgvTienLuong;
        private DataGridView dgvAnCaBDDH;

        // Tab 4: Công phát sinh (GNV) & Kiến nghị
        private DataGridView dgvGNV;
        private bool _updatingGNV;
        private TextBox txtKienNghi;

        // Tab 5: Xem trước & Điền / Xuất Biên bản Word
        private CheckBox chkDinhKemBangNhanVien;
        private RichTextBox rtbPreview;

        public MainForm()
        {
            Text = "PHẦN MỀM TỰ ĐỘNG HÓA NGHIỆM THU LAO ĐỘNG TIỀN LƯƠNG CHI TIẾT TỪNG NGƯỜI & XUẤT BIÊN BẢN WORD - CISCO (TKV) — v5.5";
            Size = new Size(1360, 870);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            BackColor = Color.FromArgb(244, 247, 252);

            string defaultExcel = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "03.9.2026_NT_T9_TCHC.xlsx");
            _excelFilePath = defaultExcel;

            InitializeLayout();
            LoadInitialData();
        }

        private void InitializeLayout()
        {
            // 1. Header Banner
            var pnlBanner = new Panel
            {
                Dock = DockStyle.Top,
                Height = 62,
                BackColor = Color.FromArgb(15, 52, 96),
                Padding = new Padding(14, 8, 14, 8)
            };

            var lblTitle = new Label
            {
                Text = "CÔNG TY CỔ PHẦN GANG THÉP CAO BẰNG (CISCO) - PHÒNG TỔ CHỨC HÀNH CHÍNH",
                ForeColor = Color.FromArgb(255, 215, 100),
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(14, 6)
            };

            var lblSubTitle = new Label
            {
                Text = "Quản lý Chấm công, Ăn ca, Điểm SX & Lương CHI TIẾT TỪNG NGƯỜI  -> Tự động điền Bảng nghiệm thu Excel (Giữ 100% Công thức) & Biên bản Word",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(14, 28)
            };
            pnlBanner.Controls.Add(lblTitle);
            pnlBanner.Controls.Add(lblSubTitle);

            // 2. Toolbar Panel
            var pnlToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 96,
                BackColor = Color.White,
                Padding = new Padding(12, 8, 12, 8),
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblPX = new Label { Text = "Phân xưởng:", AutoSize = true, Location = new Point(12, 14), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            cboWorkshop = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(102, 10),
                Width = 248,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            cboWorkshop.SelectedIndexChanged += (s, e) =>
            {
                if (!_isLoadingUI && cboWorkshop.SelectedIndex >= 0)
                {
                    SaveUIToCurrentWorkshop();
                    _currentIndex = cboWorkshop.SelectedIndex;
                    LoadCurrentWorkshopToUI();
                }
            };

            var lblThang = new Label { Text = "Tháng:", AutoSize = true, Location = new Point(358, 14) };
            numThang = new NumericUpDown { Minimum = 1, Maximum = 12, Value = 9, Location = new Point(408, 10), Width = 48 };
            var lblNam = new Label { Text = "Năm:", AutoSize = true, Location = new Point(464, 14) };
            numNam = new NumericUpDown { Minimum = 2020, Maximum = 2035, Value = 2026, Location = new Point(506, 10), Width = 65 };

            var lblKy = new Label { Text = "Mẫu:", AutoSize = true, Location = new Point(578, 14) };
            cboLoaiKy = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(618, 10),
                Width = 195
            };
            cboLoaiKy.Items.AddRange(new object[]
            {
                "Tháng thường (Chỉ báo cáo Tháng)",
                "Tháng 3 & Quý I",
                "Tháng 6, Quý II & 6 tháng đầu năm",
                "Tháng 9, Quý III & 9 tháng năm",
                "Tháng 12, Quý IV & Cả năm"
            });
            cboLoaiKy.SelectedIndex = 3;
            cboLoaiKy.SelectedIndexChanged += (s, e) =>
            {
                if (!_isLoadingUI)
                {
                    SaveUIToCurrentWorkshop();
                    RefreshPreview();
                }
            };

            chkTuDongLuyKe = new CheckBox
            {
                Text = "Tự cộng Lũy kế",
                Checked = true,
                AutoSize = true,
                Location = new Point(820, 13),
                ForeColor = Color.FromArgb(0, 102, 51),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            chkTuDongLuyKe.CheckedChanged += (s, e) =>
            {
                if (!_isLoadingUI && _workshops != null && _currentIndex >= 0)
                {
                    _workshops[_currentIndex].TuDongCongLuyKe = chkTuDongLuyKe.Checked;
                    if (chkTuDongLuyKe.Checked)
                    {
                        SaveUIToCurrentWorkshop();
                        LoadCurrentWorkshopToUI();
                    }
                }
            };

            var lblExcel = new Label { Text = "Excel:", AutoSize = true, Location = new Point(940, 14) };
            txtExcelPath = new TextBox { Location = new Point(985, 10), Width = 250, Text = _excelFilePath };
            var btnBrowseExcel = new Button
            {
                Text = "Chọn...",
                Location = new Point(1240, 9),
                Size = new Size(72, 27),
                BackColor = Color.FromArgb(230, 236, 245),
                FlatStyle = FlatStyle.Flat
            };
            btnBrowseExcel.Click += BtnBrowseExcel_Click;

            // Row 2 of Toolbar: Action Buttons
            var btnLoadExcel = CreateActionButton("1. Đọc Excel Gốc ", new Point(12, 48), new Size(190, 36), Color.FromArgb(0, 102, 204));
            btnLoadExcel.Click += BtnLoadExcel_Click;

            var btnImportExternal = CreateActionButton("2. Nạp Bảng Công / Ăn Ca / SX / GNV", new Point(208, 48), new Size(258, 36), Color.FromArgb(106, 27, 154));
            btnImportExternal.Click += BtnImportExternal_Click;

            var btnCalc = CreateActionButton("3. Tính Lương Từng Người & Lũy Kế", new Point(472, 48), new Size(235, 36), Color.FromArgb(230, 126, 34));
            btnCalc.Click += BtnCalc_Click;

            var btnSaveExcel = CreateActionButton("4. Điền Vào Excel (Giữ Công Thức)", new Point(713, 48), new Size(232, 36), Color.FromArgb(39, 174, 96));
            btnSaveExcel.Click += BtnSaveExcel_Click;

            var btnExportWord = CreateActionButton("5. Điền & Xuất Biên Bản Word (.docx)", new Point(951, 48), new Size(232, 36), Color.FromArgb(192, 57, 43));
            btnExportWord.Click += BtnExportWord_Click;

            var btnExportAllWord = CreateActionButton("Điền Tất Cả 5 PX", new Point(1189, 48), new Size(132, 36), Color.FromArgb(142, 68, 173));
            btnExportAllWord.Click += BtnExportAllWord_Click;

            pnlToolbar.Controls.AddRange(new Control[]
            {
                lblPX, cboWorkshop, lblThang, numThang, lblNam, numNam, lblKy, cboLoaiKy, chkTuDongLuyKe,
                lblExcel, txtExcelPath, btnBrowseExcel,
                btnLoadExcel, btnImportExternal, btnCalc, btnSaveExcel, btnExportWord, btnExportAllWord
            });

            // 3. Status Bar
            var pnlStatus = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                BackColor = Color.FromArgb(228, 235, 245),
                Padding = new Padding(10, 5, 10, 5)
            };
            lblStatus = new Label
            {
                Text = "Sẵn sàng. Đã tải dữ liệu chi tiết từng người và bảng nghiệm thu từ 03.9.2026_NT_T9_TCHC.xlsx.",
                AutoSize = true,
                ForeColor = Color.FromArgb(20, 50, 90),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Location = new Point(10, 6)
            };
            pnlStatus.Controls.Add(lblStatus);

            // 4. Main TabControl
            tabMain = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Padding = new Point(12, 6)
            };

            var payrollMenu = new MenuStrip { Dock = DockStyle.Top };
            var crane = new ToolStripMenuItem("LƯƠNG CẨU PHÔI PX THÉP (v5.5)");
            crane.Click += (sender, args) => {
                if (_workshops == null || _currentIndex < 0) return;
                CommitAllGridEdits();
                SaveUIToCurrentWorkshop();
                var ws = _workshops[_currentIndex];
                if (ws.MaPX != "LT") { MessageBox.Show("Chọn PX Luyện thép để nhập lương cẩu phôi."); return; }
                using (var form = new CranePayrollForm(ws))
                    if (form.ShowDialog(this) == DialogResult.OK) { LoadCurrentWorkshopToUI(); RefreshPreview(); }
            };
            var management = new ToolStripMenuItem("TÍNH LƯƠNG QUẢN LÝ PHÂN XƯỞNG (v5.5)");
            management.Click += (s, e) =>
            {
                if (_workshops == null || _currentIndex < 0) return;
                CommitAllGridEdits();
                SaveUIToCurrentWorkshop();
                using (var dialog = new ManagementPayrollForm(_workshops[_currentIndex]))
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadCurrentWorkshopToUI();
                        RefreshPreview();
                        SetStatus("Đã áp dụng lương quản lý. Bấm Lưu Excel để lưu kết quả.");
                    }
            };
            payrollMenu.Items.Add(management);
            payrollMenu.Items.Add(crane);
            Controls.Add(payrollMenu);
            payrollMenu.BringToFront();

            var tabEmp = new TabPage("1. Bảng Công, Ăn Ca & Lương Chi Tiết Từng Người") { BackColor = Color.FromArgb(248, 250, 253), Padding = new Padding(8) };
            var tab1 = new TabPage("2. Bảng Nghiệm Thu Ngày Công & Biến Động LĐ") { BackColor = Color.FromArgb(248, 250, 253), Padding = new Padding(8) };
            var tab2 = new TabPage("3. Nghiệm Thu Sản Xuất, Lương, Ăn Ca & Độc Hại") { BackColor = Color.FromArgb(248, 250, 253), Padding = new Padding(8) };
            var tab3 = new TabPage("4. Công Phát Sinh (GNV) & Kiến Nghị") { BackColor = Color.FromArgb(248, 250, 253), Padding = new Padding(8) };
            var tab4 = new TabPage("5. Biên Bản Nghiệm Thu Word (.docx)") { BackColor = Color.FromArgb(248, 250, 253), Padding = new Padding(8) };

            BuildTabEmployees(tabEmp);
            BuildTab1(tab1);
            BuildTab2(tab2);
            BuildTab3(tab3);
            BuildTab4(tab4);

            tabMain.TabPages.AddRange(new[] { tabEmp, tab1, tab2, tab3, tab4 });
            tabMain.SelectedIndexChanged += (s, e) =>
            {
                if (!_isLoadingUI)
                {
                    SaveUIToCurrentWorkshop();
                    RefreshPreview();
                }
            };

            Controls.Add(tabMain);
            Controls.Add(pnlToolbar);
            Controls.Add(pnlBanner);
            Controls.Add(pnlStatus);
        }

        private Button CreateActionButton(string text, Point location, Size size, Color backColor)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void BuildTabEmployees(TabPage tab)
        {
            var pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 122,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(8)
            };

            var lblFilter = new Label { Text = "Lọc Tổ / Bộ phận:", Location = new Point(10, 12), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            cboFilterBoPhan = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(125, 8),
                Width = 225,
                Font = new Font("Segoe UI", 9f)
            };
            cboFilterBoPhan.SelectedIndexChanged += (s, e) =>
            {
                if (!_isLoadingUI && !_isFilteringEmp)
                    PopulateEmployeeGridRows();
            };

            var lblSearch = new Label { Text = "Tìm Tên / SDB:", Location = new Point(360, 12), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            txtSearchEmp = new TextBox
            {
                Location = new Point(460, 8),
                Width = 165,
                Font = new Font("Segoe UI", 9f)
            };
            txtSearchEmp.TextChanged += (s, e) =>
            {
                if (!_isLoadingUI && !_isFilteringEmp)
                    PopulateEmployeeGridRows();
            };

            var btnImportEmpFile = CreateActionButton("Nạp File Chấm Công / Ăn Ca / Điểm SX Từng Người (.xlsx)", new Point(635, 6), new Size(340, 30), Color.FromArgb(106, 27, 154));
            btnImportEmpFile.Click += BtnImportExternal_Click;

            var btnRecalcEmp = CreateActionButton("Tính Lương Từng Người -> Cập Nhật Bảng NT & Biên Bản", new Point(985, 6), new Size(330, 30), Color.FromArgb(39, 174, 96));
            btnRecalcEmp.Click += (s, e) =>
            {
                CommitAllGridEdits();
                SaveEmployeeGridToWorkshop();
                var ws = _workshops[_currentIndex];
                ws.RecalculateFromEmployees(true);
                LoadCurrentWorkshopToUI();
                SetStatus(string.Format("Đã tính lương chi tiết từng người ({0} LĐ) và tự động đẩy tổng sang Bảng nghiệm thu + Biên bản Word!", ws.ListNhanVien.Count));
                MessageBox.Show(string.Format("Đã tính toán xong ngày công, ăn ca và tiền lương chi tiết cho {0} người lao động của {1}!\n\n- Tổng công phát sinh: {2:N1} công (Công SP: {3:N1} công)\n- Tổng công ăn ca: {4:N1} công ({5:N0} đồng)\n- Tổng quỹ lương toàn PX: {6:N0} đồng\n\nToàn bộ số liệu đã được đẩy tự động sang Bảng nghiệm thu và Biên bản Word!",
                    ws.ListNhanVien.Count, ws.TenPX, ws.TongCong, ws.CongSP, ws.CongAnCa, ws.TienAnCa, ws.TongLuongToanPX),
                    "Tính lương từng người & Đẩy sang Bảng nghiệm thu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            lblEmpSummaryBanner = new Label
            {
                Text = "Tổng hợp từ danh sách chi tiết từng người: Đang tải...",
                Location = new Point(10, 90),
                AutoSize = true,
                ForeColor = Color.FromArgb(15, 52, 96),
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold)
            };

            chkOnlyAnCaDiff = new CheckBox
            {
                Text = "Chỉ hiện người chênh lệch ăn ca",
                Location = new Point(10, 48),
                AutoSize = true,
                ForeColor = Color.FromArgb(192, 57, 43),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            chkOnlyAnCaDiff.CheckedChanged += (s, e) =>
            {
                if (!_isLoadingUI && !_isFilteringEmp) PopulateEmployeeGridRows();
            };

            var btnAnCaCT = CreateActionButton("Ăn ca: dùng CÔNG THỨC (SP+H+VHTT) cho tất cả", new Point(260, 44), new Size(360, 30), Color.FromArgb(41, 128, 185));
            btnAnCaCT.Click += (s, e) => ApplyAnCaChoice("CT", false);
            var btnAnCaDV = CreateActionButton("Ăn ca: lấy theo ĐƠN VỊ gửi (na+tt) cho tất cả", new Point(630, 44), new Size(345, 30), Color.FromArgb(211, 84, 0));
            btnAnCaDV.Click += (s, e) => ApplyAnCaChoice("DV", false);
            var btnAnCaDiffDV = CreateActionButton("Chỉ người LỆCH: lấy số ĐƠN VỊ", new Point(985, 44), new Size(330, 30), Color.FromArgb(142, 68, 173));
            btnAnCaDiffDV.Click += (s, e) => ApplyAnCaChoice("DV", true);

            pnlTop.Controls.AddRange(new Control[]
            {
                lblFilter, cboFilterBoPhan, lblSearch, txtSearchEmp,
                btnImportEmpFile, btnRecalcEmp, lblEmpSummaryBanner,
                chkOnlyAnCaDiff, btnAnCaCT, btnAnCaDV, btnAnCaDiffDV
            });

            dgvEmployees = CreateSimpleGrid(new Point(0, 88), new Size(1300, 560));
            dgvEmployees.Dock = DockStyle.Fill;
            dgvEmployees.AllowUserToAddRows = false;
            dgvEmployees.SelectionMode = DataGridViewSelectionMode.CellSelect;

            dgvEmployees.Columns.Add("STT", "STT");
            dgvEmployees.Columns.Add("ExcelRow", "Dòng XL");
            dgvEmployees.Columns.Add("SDB", "Mã SDB");
            dgvEmployees.Columns.Add("HoTen", "Họ và tên lao động");
            dgvEmployees.Columns.Add("BoPhan", "Tổ / Bộ phận");
            dgvEmployees.Columns.Add("ChucDanh", "C.Danh");
            dgvEmployees.Columns.Add("MucLuongBHXH", "Lương BHXH");
            dgvEmployees.Columns.Add("TongCong", "Tổng công");
            dgvEmployees.Columns.Add("CongSP", "Công SP");
            dgvEmployees.Columns.Add("CongLePhep", "Phép, Lễ");
            dgvEmployees.Columns.Add("CongHocHop", "Học, Họp");
            dgvEmployees.Columns.Add("CongVHTT", "VHTT");
            dgvEmployees.Columns.Add("CongBHXH", "Ốm,BHXH");
            dgvEmployees.Columns.Add("CongAnCa", "Ăn ca");
            dgvEmployees.Columns.Add("DiemSanXuat", "Điểm/HS tháng");
            dgvEmployees.Columns.Add("XepLoai", "Loại");
            dgvEmployees.Columns.Add("HeSoHuong", "Điểm/HS hưởng");
            dgvEmployees.Columns.Add("LuongSanPham", "Lương Sản phẩm (đ)");
            dgvEmployees.Columns.Add("LuongLePhepCD", "Lương Lễ, Phép (đ)");
            dgvEmployees.Columns.Add("PhuCapTien", "Phụ cấp/Khác (đ)");
            dgvEmployees.Columns.Add("TongTienLuong", "TỔNG LƯƠNG (đ)");

            // Cột so sánh công ĂN CA: Công thức (SP + H + VHTT) vs Đơn vị gửi (na + tt) -> chọn nguồn
            dgvEmployees.Columns.Add("AnCaCT", "Ăn ca theo CT (SP+H+VHTT)");
            dgvEmployees.Columns.Add("AnCaDV", "Ăn ca đơn vị gửi (na+tt)");
            dgvEmployees.Columns.Add("AnCaCL", "Chênh lệch (ĐV − CT)");
            var colNguon = new DataGridViewComboBoxColumn { Name = "AnCaNguon", HeaderText = "Chọn nguồn ăn ca", FlatStyle = FlatStyle.Flat };
            colNguon.Items.AddRange(new object[] { "Công thức", "Đơn vị", "Thủ công", "Số cũ (Excel)" });
            dgvEmployees.Columns.Add(colNguon);
            dgvEmployees.DataError += (s, ev) => { ev.ThrowException = false; };

            int[] widths = { 42, 55, 58, 165, 135, 55, 85, 65, 65, 62, 60, 52, 62, 58, 85, 45, 85, 110, 98, 90, 115 };
            for (int i = 0; i < widths.Length && i < dgvEmployees.Columns.Count; i++)
            {
                dgvEmployees.Columns[i].Width = widths[i];
                if (i >= 6 && i != 15)
                    dgvEmployees.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                else if (i <= 2 || i == 5 || i == 15)
                    dgvEmployees.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvEmployees.Columns[0].ReadOnly = true;
            dgvEmployees.Columns[1].ReadOnly = true;
            dgvEmployees.Columns[1].DefaultCellStyle.ForeColor = Color.Gray;
            dgvEmployees.Columns[3].DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            dgvEmployees.Columns[7].DefaultCellStyle.BackColor = Color.FromArgb(240, 244, 250);
            dgvEmployees.Columns[7].DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            // Tô màu vàng nhạt cho các cột nhập liệu chấm công, ăn ca, điểm sản xuất từng người
            int[] editableCols = { 8, 9, 10, 11, 12, 13, 14, 15 };
            foreach (int cIdx in editableCols)
            {
                dgvEmployees.Columns[cIdx].DefaultCellStyle.BackColor = Color.FromArgb(255, 253, 231);
            }
            dgvEmployees.Columns[17].DefaultCellStyle.BackColor = Color.FromArgb(232, 248, 245);
            dgvEmployees.Columns[17].DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            dgvEmployees.Columns[20].DefaultCellStyle.BackColor = Color.FromArgb(234, 242, 248);
            dgvEmployees.Columns[20].DefaultCellStyle.ForeColor = Color.FromArgb(15, 52, 96);
            dgvEmployees.Columns[20].DefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            int[] wNew = { 105, 105, 90, 105 };
            for (int i = 0; i < wNew.Length; i++)
            {
                var cNew = dgvEmployees.Columns[21 + i];
                cNew.Width = wNew[i];
                if (i < 3)
                {
                    cNew.ReadOnly = true;
                    cNew.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    cNew.DefaultCellStyle.BackColor = Color.FromArgb(243, 240, 250);
                }
            }
            dgvEmployees.Columns[24].DefaultCellStyle.BackColor = Color.FromArgb(255, 253, 231);

            dgvEmployees.CellEndEdit += DgvEmployees_CellEndEdit;

            tab.Controls.Add(dgvEmployees);
            tab.Controls.Add(pnlTop);
        }

        private void ApplyAnCaChoice(string nguon, bool onlyDiff)
        {
            CommitAllGridEdits();
            SaveEmployeeGridToWorkshop();
            var ws = _workshops[_currentIndex];
            int n = ws.SetAnCaNguon(nguon, onlyDiff);
            ws.RecalculateFromEmployees(true);
            LoadCurrentWorkshopToUI();
            SetStatus(string.Format("Đã đặt nguồn công ăn ca = {0} cho {1} người. Bảng nghiệm thu & Biên bản đã cập nhật.",
                nguon == "CT" ? "CÔNG THỨC (SP+H+VHTT)" : nguon == "DV" ? "ĐƠN VỊ gửi (na+tt)" : "THỦ CÔNG", n));
        }

        private void DgvEmployees_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (_isLoadingUI || _isFilteringEmp || e.RowIndex < 0) return;
            var row = dgvEmployees.Rows[e.RowIndex];
            var emp = row.Tag as EmployeeRecord;
            if (emp == null) return;

            ReadRowIntoEmployeeRecord(row, emp);
            var ws = _workshops[_currentIndex];
            ws.RecalculateFromEmployees(true);

            // Cập nhật lại các ô tính toán trên dòng vừa sửa
            _isFilteringEmp = true;
            WriteEmployeeRecordToRow(row, emp);
            _isFilteringEmp = false;

            UpdateEmployeeBanner(ws);
            RefreshSummaryGridsFromWorkshop(ws);
            RefreshPreview();
            SetStatus(string.Format("Đã cập nhật công & lương của [{0} - SDB {1}] (Dòng {2} sheet '{3}') và tự động đồng bộ Bảng nghiệm thu + Biên bản!",
                emp.HoTen, emp.SDB, emp.ExcelRow, ws.SheetBangCong));
        }

        private void BuildTab1(TabPage tab)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 540
            };

            var grpInfo = new GroupBox
            {
                Text = "I. Thời gian, Địa điểm & Thành phần nghiệm thu",
                Dock = DockStyle.Top,
                Height = 290,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };

            var lblHop = new Label { Text = "Ngày họp NT:", Location = new Point(14, 28), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            numNgayHop = new NumericUpDown { Minimum = 1, Maximum = 31, Value = 6, Location = new Point(105, 25), Width = 48, Font = new Font("Segoe UI", 9f) };
            numThangHop = new NumericUpDown { Minimum = 1, Maximum = 12, Value = 10, Location = new Point(158, 25), Width = 48, Font = new Font("Segoe UI", 9f) };
            numNamHop = new NumericUpDown { Minimum = 2020, Maximum = 2035, Value = 2026, Location = new Point(212, 25), Width = 62, Font = new Font("Segoe UI", 9f) };

            var lblGio = new Label { Text = "Giờ kết thúc:", Location = new Point(285, 28), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            txtGioKetThuc = new TextBox { Location = new Point(370, 25), Width = 145, Text = "11 giờ 00 phút", Font = new Font("Segoe UI", 9f) };

            var lblDD = new Label { Text = "Địa điểm:", Location = new Point(14, 58), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            txtDiaDiem = new TextBox { Location = new Point(105, 55), Width = 410, Text = "Phòng Tổ chức Hành chính", Font = new Font("Segoe UI", 9f) };

            var lblTP1 = new Label { Text = "1. Đại diện Phòng Tổ chức Hành chính:", Location = new Point(14, 86), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            dgvThanhPhanTCHC = CreateSimpleGrid(new Point(14, 108), new Size(505, 72));
            dgvThanhPhanTCHC.Columns.Add("HoTen", "Họ và tên (nhập người tham dự)");
            dgvThanhPhanTCHC.Columns.Add("ChucVu", "Chức vụ");
            dgvThanhPhanTCHC.Columns[0].Width = 280;
            dgvThanhPhanTCHC.Columns[1].Width = 190;

            var lblTP2 = new Label { Text = "2. Đại diện Phân xưởng:", Location = new Point(14, 186), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            dgvThanhPhanPX = CreateSimpleGrid(new Point(14, 208), new Size(505, 72));
            dgvThanhPhanPX.Columns.Add("HoTen", "Họ và tên (nhập người tham dự)");
            dgvThanhPhanPX.Columns.Add("ChucVu", "Chức vụ");
            dgvThanhPhanPX.Columns[0].Width = 280;
            dgvThanhPhanPX.Columns[1].Width = 190;

            grpInfo.Controls.AddRange(new Control[]
            {
                lblHop, numNgayHop, numThangHop, numNamHop, lblGio, txtGioKetThuc,
                lblDD, txtDiaDiem, lblTP1, dgvThanhPhanTCHC, lblTP2, dgvThanhPhanPX
            });

            var grpBienDong = new GroupBox
            {
                Text = "1.1. Tình hình biến động lao động trong kỳ",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };

            var lblDK = new Label { Text = "LĐ đầu kỳ (người):", Location = new Point(14, 34), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            numLDDauKy = new NumericUpDown { Maximum = 2000, Location = new Point(140, 30), Width = 70, Font = new Font("Segoe UI", 9f) };
            var lblTang = new Label { Text = "Tăng:", Location = new Point(225, 34), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            numLDTang = new NumericUpDown { Maximum = 500, Location = new Point(268, 30), Width = 60, Font = new Font("Segoe UI", 9f) };
            var lblGiam = new Label { Text = "Giảm:", Location = new Point(342, 34), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            numLDGiam = new NumericUpDown { Maximum = 500, Location = new Point(388, 30), Width = 60, Font = new Font("Segoe UI", 9f) };

            var lblCK = new Label { Text = "LĐ cuối kỳ (người):", Location = new Point(14, 68), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            numLDCuoiKy = new NumericUpDown { Maximum = 2000, Location = new Point(140, 64), Width = 70, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            var lblQL = new Label { Text = "LĐ Quản lý:", Location = new Point(225, 68), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            numLDQuanLy = new NumericUpDown { Maximum = 500, Location = new Point(305, 64), Width = 55, Font = new Font("Segoe UI", 9f) };
            var lblTT = new Label { Text = "LĐ Trực tiếp:", Location = new Point(370, 68), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            numLDTrucTiep = new NumericUpDown { Maximum = 2000, Location = new Point(455, 64), Width = 62, Font = new Font("Segoe UI", 9f) };

            var lblLyDo = new Label { Text = "Chi tiết biến động (Tăng / Giảm / Chấm dứt HĐLĐ):", Location = new Point(14, 102), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            txtLyDoBienDong = new TextBox { Location = new Point(14, 124), Width = 505, Height = 48, Multiline = true, Font = new Font("Segoe UI", 9f) };

            var lblGhiChuLD = new Label { Text = "Ghi chú lao động (Hoãn HĐLĐ, ốm dài ngày, biệt phái...):", Location = new Point(14, 180), AutoSize = true, Font = new Font("Segoe UI", 9f) };
            txtGhiChuLD = new TextBox { Location = new Point(14, 202), Width = 505, Height = 48, Multiline = true, Font = new Font("Segoe UI", 9f) };

            grpBienDong.Controls.AddRange(new Control[]
            {
                lblDK, numLDDauKy, lblTang, numLDTang, lblGiam, numLDGiam,
                lblCK, numLDCuoiKy, lblQL, numLDQuanLy, lblTT, numLDTrucTiep,
                lblLyDo, txtLyDoBienDong, lblGhiChuLD, txtGhiChuLD
            });

            split.Panel1.Controls.Add(grpBienDong);
            split.Panel1.Controls.Add(grpInfo);

            var grpNgayCong = new GroupBox
            {
                Text = "1.2. Bảng tổng hợp Ngày công & Lao động bình quân (Tự động tổng hợp từ Tab 1 hoặc sửa trực tiếp)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };

            dgvNgayCong = CreateSimpleGrid(new Point(12, 28), new Size(730, 500));
            dgvNgayCong.Dock = DockStyle.Fill;
            dgvNgayCong.AllowUserToAddRows = false;
            dgvNgayCong.Columns.Add("ChiTieu", "Chỉ tiêu ngày công & lao động");
            dgvNgayCong.Columns.Add("DVT", "ĐVT");
            dgvNgayCong.Columns.Add("Thang", "Tháng báo cáo");
            dgvNgayCong.Columns.Add("Quy", "Quý (I/II/III/IV)");
            dgvNgayCong.Columns.Add("LuyKe", "Lũy kế (6T/9T/Năm)");
            dgvNgayCong.Columns[0].Width = 285;
            dgvNgayCong.Columns[0].ReadOnly = true;
            dgvNgayCong.Columns[1].Width = 75;
            dgvNgayCong.Columns[1].ReadOnly = true;
            dgvNgayCong.Columns[2].Width = 120;
            dgvNgayCong.Columns[3].Width = 120;
            dgvNgayCong.Columns[4].Width = 125;
            dgvNgayCong.CellEndEdit += (s, e) =>
            {
                if (!_isLoadingUI)
                {
                    SaveUIToCurrentWorkshop();
                    RefreshSummaryGridsFromWorkshop(_workshops[_currentIndex]);
                    RefreshPreview();
                }
            };

            grpNgayCong.Controls.Add(dgvNgayCong);
            split.Panel2.Controls.Add(grpNgayCong);

            tab.Controls.Add(split);
        }

        private void BuildTab2(TabPage tab)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 175
            };

            var grpSanLuong = new GroupBox
            {
                Text = "2.1a. Nghiệm thu Sản lượng sản xuất trong tháng & Đơn giá tính lương",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };

            int y1 = 28, y2 = 62, y3 = 96, y4 = 130;
            grpSanLuong.Controls.Add(new Label { Text = "Sản phẩm chính:", Location = new Point(14, y1 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtTenSP = new TextBox { Location = new Point(120, y1), Width = 210, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            grpSanLuong.Controls.Add(new Label { Text = "ĐVT:", Location = new Point(340, y1 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtDvtSP = new TextBox { Location = new Point(380, y1), Width = 65, Font = new Font("Segoe UI", 9f) };

            grpSanLuong.Controls.Add(new Label { Text = "Kế hoạch giao:", Location = new Point(460, y1 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSLKeHoach = new TextBox { Location = new Point(555, y1), Width = 115, Font = new Font("Segoe UI", 9f) };

            grpSanLuong.Controls.Add(new Label { Text = "Thực hiện tháng:", Location = new Point(685, y1 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) });
            txtSLThucHien = new TextBox { Location = new Point(795, y1), Width = 120, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };

            grpSanLuong.Controls.Add(new Label { Text = "SL tính lương:", Location = new Point(930, y1 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) });
            txtSLTinhLuong = new TextBox { Location = new Point(1025, y1), Width = 120, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };

            grpSanLuong.Controls.Add(new Label { Text = "SP Loại 1 (KL):", Location = new Point(14, y2 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSLLoai1 = new TextBox { Location = new Point(120, y2), Width = 110, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Đơn giá L1:", Location = new Point(240, y2 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtDGLoai1 = new TextBox { Location = new Point(315, y2), Width = 110, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Thành tiền L1:", Location = new Point(435, y2 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtTTLoai1 = new TextBox { Location = new Point(530, y2), Width = 140, Font = new Font("Segoe UI", 9f) };

            grpSanLuong.Controls.Add(new Label { Text = "SP Loại 2 (KL):", Location = new Point(685, y2 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSLLoai2 = new TextBox { Location = new Point(795, y2), Width = 95, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Đơn giá L2:", Location = new Point(900, y2 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtDGLoai2 = new TextBox { Location = new Point(975, y2), Width = 95, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Tiền L2:", Location = new Point(1080, y2 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtTTLoai2 = new TextBox { Location = new Point(1135, y2), Width = 120, Font = new Font("Segoe UI", 9f) };

            grpSanLuong.Controls.Add(new Label { Text = "SP ngày Lễ (KL):", Location = new Point(14, y3 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSLLe = new TextBox { Location = new Point(120, y3), Width = 110, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Đơn giá Lễ:", Location = new Point(240, y3 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtDGLe = new TextBox { Location = new Point(315, y3), Width = 110, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Thành tiền Lễ:", Location = new Point(435, y3 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtTTLe = new TextBox { Location = new Point(530, y3), Width = 140, Font = new Font("Segoe UI", 9f) };

            grpSanLuong.Controls.Add(new Label { Text = "SP phụ/Cẩu phôi:", Location = new Point(685, y3 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSanLuongPhuName = new TextBox { Location = new Point(795, y3), Width = 140, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "KL:", Location = new Point(942, y3 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSLPhuKL = new TextBox { Location = new Point(970, y3), Width = 95, Font = new Font("Segoe UI", 9f) };
            grpSanLuong.Controls.Add(new Label { Text = "Tiền:", Location = new Point(1075, y3 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtSLPhuTT = new TextBox { Location = new Point(1115, y3), Width = 140, Font = new Font("Segoe UI", 9f) };

            grpSanLuong.Controls.Add(new Label { Text = "Ghi chú giảm trừ lương (nếu có):", Location = new Point(14, y4 + 3), AutoSize = true, Font = new Font("Segoe UI", 9f) });
            txtLyDoTruLuong = new TextBox { Location = new Point(215, y4), Width = 1040, Font = new Font("Segoe UI", 9f) };

            txtTTLoai1.ReadOnly = txtTTLoai2.ReadOnly = txtTTLe.ReadOnly = true;
            txtTTLoai1.BackColor = txtTTLoai2.BackColor = txtTTLe.BackColor = Color.FromArgb(235, 242, 248);
            txtSLLoai1.TextChanged += (sender, args) => ProductionInputChanged(ProductionLine.Loai1);
            txtDGLoai1.TextChanged += (sender, args) => ProductionInputChanged(ProductionLine.Loai1);
            txtSLLoai2.TextChanged += (sender, args) => ProductionInputChanged(ProductionLine.Loai2);
            txtDGLoai2.TextChanged += (sender, args) => ProductionInputChanged(ProductionLine.Loai2);
            txtSLLe.TextChanged += (sender, args) => ProductionInputChanged(ProductionLine.NgayLe);
            txtDGLe.TextChanged += (sender, args) => ProductionInputChanged(ProductionLine.NgayLe);

            grpSanLuong.Controls.AddRange(new Control[]
            {
                txtTenSP, txtDvtSP, txtSLKeHoach, txtSLThucHien, txtSLTinhLuong,
                txtSLLoai1, txtDGLoai1, txtTTLoai1, txtSLLoai2, txtDGLoai2, txtTTLoai2,
                txtSLLe, txtDGLe, txtTTLe, txtSanLuongPhuName, txtSLPhuKL, txtSLPhuTT, txtLyDoTruLuong
            });

            split.Panel1.Controls.Add(grpSanLuong);

            var splitBottom = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 660
            };

            var grpLuong = new GroupBox
            {
                Text = "2.1b. Bảng tổng hợp Tiền lương & Thu nhập bình quân (Tháng / Quý / Lũy kế)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            dgvTienLuong = CreateSimpleGrid(new Point(10, 25), new Size(635, 360));
            dgvTienLuong.Dock = DockStyle.Fill;
            dgvTienLuong.AllowUserToAddRows = false;
            dgvTienLuong.Columns.Add("ChiTieu", "Chỉ tiêu Tiền lương & Bình quân");
            dgvTienLuong.Columns.Add("DVT", "ĐVT");
            dgvTienLuong.Columns.Add("Thang", "Tháng báo cáo");
            dgvTienLuong.Columns.Add("Quy", "Quý (I/II/III/IV)");
            dgvTienLuong.Columns.Add("LuyKe", "Lũy kế (6T/9T/Năm)");
            dgvTienLuong.Columns[0].Width = 250;
            dgvTienLuong.Columns[0].ReadOnly = true;
            dgvTienLuong.Columns[1].Width = 65;
            dgvTienLuong.Columns[1].ReadOnly = true;
            dgvTienLuong.Columns[2].Width = 105;
            dgvTienLuong.Columns[3].Width = 105;
            dgvTienLuong.Columns[4].Width = 110;
            dgvTienLuong.CellEndEdit += (s, e) =>
            {
                if (!_isLoadingUI)
                {
                    SaveUIToCurrentWorkshop();
                    RefreshSummaryGridsFromWorkshop(_workshops[_currentIndex]);
                    RefreshPreview();
                }
            };
            grpLuong.Controls.Add(dgvTienLuong);

            var grpAnCa = new GroupBox
            {
                Text = "2.2 & 2.3. Ăn giữa ca (30.000đ) & Bồi dưỡng độc hại (Mức 1: 13k, Mức 2: 20k, Mức 3: 26k)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            dgvAnCaBDDH = CreateSimpleGrid(new Point(10, 25), new Size(600, 360));
            dgvAnCaBDDH.Dock = DockStyle.Fill;
            dgvAnCaBDDH.AllowUserToAddRows = false;
            dgvAnCaBDDH.Columns.Add("ChiTieu", "Khoản mục Ăn ca / Độc hại");
            dgvAnCaBDDH.Columns.Add("DVT", "ĐVT");
            dgvAnCaBDDH.Columns.Add("Thang", "Tháng báo cáo");
            dgvAnCaBDDH.Columns.Add("Quy", "Quý (I/II/III/IV)");
            dgvAnCaBDDH.Columns.Add("LuyKe", "Lũy kế (6T/9T/Năm)");
            dgvAnCaBDDH.Columns[0].Width = 225;
            dgvAnCaBDDH.Columns[0].ReadOnly = true;
            dgvAnCaBDDH.Columns[1].Width = 60;
            dgvAnCaBDDH.Columns[1].ReadOnly = true;
            dgvAnCaBDDH.Columns[2].Width = 100;
            dgvAnCaBDDH.Columns[3].Width = 100;
            dgvAnCaBDDH.Columns[4].Width = 105;
            dgvAnCaBDDH.CellEndEdit += (s, e) =>
            {
                if (!_isLoadingUI)
                {
                    SaveUIToCurrentWorkshop();
                    RefreshSummaryGridsFromWorkshop(_workshops[_currentIndex]);
                    RefreshPreview();
                }
            };
            grpAnCa.Controls.Add(dgvAnCaBDDH);

            splitBottom.Panel1.Controls.Add(grpLuong);
            splitBottom.Panel2.Controls.Add(grpAnCa);
            split.Panel2.Controls.Add(splitBottom);

            tab.Controls.Add(split);
        }

        private void ProductionInputChanged(ProductionLine line)
        {
            if (_isLoadingUI || _workshops == null || _currentIndex < 0) return;
            TextBox quantity = line == ProductionLine.Loai1 ? txtSLLoai1 : line == ProductionLine.Loai2 ? txtSLLoai2 : txtSLLe;
            TextBox price = line == ProductionLine.Loai1 ? txtDGLoai1 : line == ProductionLine.Loai2 ? txtDGLoai2 : txtDGLe;
            double q, p;
            if (!TryParseProductionNumber(quantity.Text, out q) || !TryParseProductionNumber(price.Text, out p))
            { SetStatus("Sản lượng/đơn giá chưa hợp lệ. Nhập số không âm theo định dạng Việt Nam."); return; }
            SaveUIToCurrentWorkshop();
            var ws = _workshops[_currentIndex];
            ProductionPayroll.UpdateLine(ws, line, q, p);
            ws.RecalculateFromEmployees(true);
            _isLoadingUI = true;
            try
            {
                txtTTLoai1.Text = WordReportService.FmtInt(ws.ThanhTienLoai1);
                txtTTLoai2.Text = WordReportService.FmtInt(ws.ThanhTienLoai2);
                txtTTLe.Text = WordReportService.FmtInt(ws.ThanhTienLe);
                RefreshSummaryGridsFromWorkshop(ws);
                PopulateEmployeeGridRows();
            }
            finally { _isLoadingUI = false; }
            RefreshPreview();
            SetStatus("Đã tính lại thành tiền = sản lượng × đơn giá, quỹ SP và lương từng người. Chưa lưu Excel.");
        }

        private static bool TryParseProductionNumber(string text, out double value)
        {
            text = (text ?? "").Trim();
            if (text == "") { value=0; return true; }
            return double.TryParse(text.Replace(".", "").Replace(",", "."), NumberStyles.Number,
                CultureInfo.InvariantCulture, out value) && value >= 0 && !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private void BuildTab3(TabPage tab)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 390
            };

            var grpGNV = new GroupBox
            {
                Text = "Danh sách Công việc phát sinh / Giấy giao nhiệm vụ (GNV) nghiệm thu trong tháng",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };

            var pnlGNVButtons = new Panel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(6) };
            var btnAddGNV = CreateActionButton("+ Thêm dòng GNV", new Point(8, 5), new Size(150, 30), Color.FromArgb(41, 128, 185));
            btnAddGNV.Click += (s, e) =>
            {
                int row = dgvGNV.Rows.Add(dgvGNV.Rows.Count, "GNV số ...: Nội dung công việc phát sinh", "Công", "0", "0", "0", "");
                dgvGNV.Rows[row].Tag = new ArisingWorkItem { TuTinhThanhTien=true };
                GnvInputChanged();
            };

            var btnCalcGNV = CreateActionButton("Tính tổng GNV & Đưa vào Lương CV Khác + Biên bản", new Point(168, 5), new Size(340, 30), Color.FromArgb(39, 174, 96));
            btnCalcGNV.Click += (s, e) =>
            {
                SaveUIToCurrentWorkshop();
                _workshops[_currentIndex].GnvDaChinhSua=true;
                _workshops[_currentIndex].RecalculateDerived();
                _workshops[_currentIndex].RecalculateFromEmployees(true);
                LoadCurrentWorkshopToUI();
                MessageBox.Show(string.Format("Đã cập nhật tổng khối lượng ({0:N1} công) và thành tiền ({1:N0} đồng) từ bảng GNV vào Bảng nghiệm thu và Biên bản!", _workshops[_currentIndex].CongCVKhac, _workshops[_currentIndex].LuongCVKhac), "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            var btnClearGNV = CreateActionButton("Xóa dòng đang chọn", new Point(518, 5), new Size(165, 30), Color.FromArgb(192, 57, 43));
            btnClearGNV.Click += (s, e) =>
            {
                if (dgvGNV.CurrentRow != null && !dgvGNV.CurrentRow.IsNewRow)
                    dgvGNV.Rows.Remove(dgvGNV.CurrentRow);
            };

            pnlGNVButtons.Controls.AddRange(new Control[] { btnAddGNV, btnCalcGNV, btnClearGNV });

            dgvGNV = CreateSimpleGrid(new Point(10, 50), new Size(1250, 320));
            dgvGNV.Dock = DockStyle.Fill;
            dgvGNV.Columns.Add("STT", "STT");
            dgvGNV.Columns.Add("NoiDung", "Số Giấy giao nhiệm vụ (GNV) / Nội dung công việc phát sinh");
            dgvGNV.Columns.Add("DVT", "ĐVT");
            dgvGNV.Columns.Add("KhoiLuong", "Khối lượng / Số công");
            dgvGNV.Columns.Add("DonGia", "Đơn giá (đồng)");
            dgvGNV.Columns.Add("ThanhTien", "Thành tiền (đồng)");
            dgvGNV.Columns.Add("GhiChu", "Ghi chú / Dòng Excel");
            dgvGNV.Columns[0].Width = 55;
            dgvGNV.Columns[1].Width = 520;
            dgvGNV.Columns[2].Width = 80;
            dgvGNV.Columns[3].Width = 140;
            dgvGNV.Columns[4].Width = 140;
            dgvGNV.Columns[5].Width = 160;
            dgvGNV.Columns[6].Width = 140;
            dgvGNV.Columns[0].ReadOnly = true;
            dgvGNV.Columns[5].ReadOnly = true;
            dgvGNV.Columns[5].DefaultCellStyle.BackColor = Color.FromArgb(235,242,248);
            dgvGNV.CellEndEdit += (sender,args) =>
            {
                if (_isLoadingUI || _updatingGNV || args.RowIndex < 0) return;
                var row=dgvGNV.Rows[args.RowIndex];
                var item=row.Tag as ArisingWorkItem;
                if(item==null) { item=new ArisingWorkItem { TuTinhThanhTien=true }; row.Tag=item; }
                if(args.ColumnIndex==3 || args.ColumnIndex==4) item.TuTinhThanhTien=true;
                GnvInputChanged();
            };
            dgvGNV.RowsRemoved += (sender,args) => GnvInputChanged();


            grpGNV.Controls.Add(dgvGNV);
            grpGNV.Controls.Add(pnlGNVButtons);

            var grpKienNghi = new GroupBox
            {
                Text = "III. Kiến nghị, đề xuất trong Biên bản nghiệm thu (Mỗi dòng là 1 ý kiến nghị)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            txtKienNghi = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 10f)
            };
            grpKienNghi.Controls.Add(txtKienNghi);

            split.Panel1.Controls.Add(grpGNV);
            split.Panel2.Controls.Add(grpKienNghi);
            tab.Controls.Add(split);
        }

        private void GnvInputChanged()
        {
            if(_isLoadingUI || _updatingGNV || Disposing || IsDisposed || _workshops==null || _currentIndex<0
                || dgvGNV==null || dgvGNV.Disposing || dgvGNV.IsDisposed || dgvGNV.Columns.Count<7
                || dgvTienLuong==null || dgvTienLuong.Columns.Count<5 || dgvNgayCong.Columns.Count<5) return;
            _updatingGNV=true;
            try
            {
                var ws=_workshops[_currentIndex]; ws.GnvDaChinhSua=true;
                SaveUIToCurrentWorkshop();
                ws.RecalculateDerived(); ws.RecalculateFromEmployees(true);
                _isLoadingUI=true;
                foreach(DataGridViewRow row in dgvGNV.Rows)
                    if(!row.IsNewRow && row.Tag is ArisingWorkItem)
                        row.Cells[5].Value=WordReportService.FmtInt(((ArisingWorkItem)row.Tag).ThanhTien);
                RefreshSummaryGridsFromWorkshop(ws); PopulateEmployeeGridRows();
                _isLoadingUI=false; RefreshPreview();
                SetStatus(ws.MaPX=="CDNL" ? "Đã tính tiền GNV = khối lượng × đơn giá. CĐNL giữ cách tổng hợp riêng, không cộng trùng quỹ." : "Đã tính lại tiền GNV, tổng công, lương công việc khác và lương phân bổ. Chưa lưu Excel.");
            }
            finally { _isLoadingUI=false; _updatingGNV=false; }
        }

        private void BuildTab4(TabPage tab)
        {
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 76, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(6) };

            var btnRefreshPreview = CreateActionButton("Làm mới Biên bản", new Point(8, 6), new Size(155, 32), Color.FromArgb(41, 128, 185));
            btnRefreshPreview.Click += (s, e) =>
            {
                SaveUIToCurrentWorkshop();
                RefreshPreview();
            };

            var btnExportSingle = CreateActionButton("Xuất File Biên Bản Word (.docx) PX Này", new Point(170, 6), new Size(275, 32), Color.FromArgb(192, 57, 43));
            btnExportSingle.Click += BtnExportWord_Click;

            var btnFillExistingDocx = CreateActionButton("Điền Trực Tiếp Vào File Word Mẫu Đang Có (.docx)", new Point(452, 6), new Size(340, 32), Color.FromArgb(106, 27, 154));
            btnFillExistingDocx.Click += BtnFillExistingDocx_Click;

            var btnExportBatch = CreateActionButton("Điền & Xuất Tất Cả 5 Phân Xưởng", new Point(799, 6), new Size(245, 32), Color.FromArgb(39, 174, 96));
            btnExportBatch.Click += BtnExportAllWord_Click;

            var btnOpenFolder = CreateActionButton("Mở Thư Mục Báo Cáo", new Point(1051, 6), new Size(175, 32), Color.FromArgb(127, 140, 141));
            btnOpenFolder.Click += (s, e) =>
            {
                string outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Output_BaoCao_Word");
                if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);
                Process.Start("explorer.exe", outDir);
            };

            chkDinhKemBangNhanVien = new CheckBox
            {
                Text = "Đính kèm Phụ lục: Bảng nghiệm thu công, ăn ca & tiền lương CHI TIẾT TỪNG NGƯỜI vào cuối Biên bản Word (.docx)",
                Checked = false,
                AutoSize = true,
                Location = new Point(10, 46),
                ForeColor = Color.FromArgb(192, 57, 43),
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold)
            };
            chkDinhKemBangNhanVien.CheckedChanged += (s, e) =>
            {
                if (!_isLoadingUI && _workshops != null && _currentIndex >= 0)
                {
                    _workshops[_currentIndex].DinhKemBangChiTietNhanVien = chkDinhKemBangNhanVien.Checked;
                }
            };

            pnlTop.Controls.AddRange(new Control[]
            {
                btnRefreshPreview, btnExportSingle, btnFillExistingDocx, btnExportBatch, btnOpenFolder, chkDinhKemBangNhanVien
            });

            rtbPreview = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10.2f),
                BackColor = Color.White,
                ReadOnly = false
            };

            tab.Controls.Add(rtbPreview);
            tab.Controls.Add(pnlTop);
        }

        private DataGridView CreateSimpleGrid(Point location, Size size)
        {
            var dgv = new DataGridView
            {
                Location = location,
                Size = size,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 9f),
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(225, 234, 247);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            return dgv;
        }

        private void LoadInitialData()
        {
            try
            {
                if (File.Exists(_excelFilePath))
                {
                    _workshops = ExcelService.LoadWorkshopsFromExcel(_excelFilePath);
                    int totalEmp = 0;
                    foreach (var w in _workshops) totalEmp += w.ListNhanVien.Count;
                    SetStatus(string.Format("Đã nạp {0} người lao động chi tiết và 5 bảng nghiệm thu từ: {1}", totalEmp, _excelFilePath));
                }
                else
                {
                    _workshops = WorkshopData.CreateDefaultWorkshops();
                    SetStatus("Chưa nạp Excel. Bấm chọn file Excel của đơn vị để bắt đầu.");
                }
            }
            catch (Exception ex)
            {
                _workshops = WorkshopData.CreateDefaultWorkshops();
                SetStatus("Chưa nạp được Excel; dữ liệu đang để trống (Lưu ý: " + ex.Message + ")");
            }

            _isLoadingUI = true;
            cboWorkshop.Items.Clear();
            foreach (var ws in _workshops)
            {
                cboWorkshop.Items.Add(string.Format("{0} ({1} LĐ - Sheet: {2})", ws.TenPX, ws.ListNhanVien.Count, ws.SheetBangCong));
            }
            cboWorkshop.SelectedIndex = 0;
            _isLoadingUI = false;

            LoadCurrentWorkshopToUI();
        }

        private void LoadCurrentWorkshopToUI()
        {
            if (_workshops == null || _currentIndex < 0 || _currentIndex >= _workshops.Count) return;
            _isLoadingUI = true;

            var ws = _workshops[_currentIndex];
            numThang.Value = Clamp(ws.Thang, 1, 12);
            numNam.Value = Clamp(ws.Nam, 2020, 2035);
            cboLoaiKy.SelectedIndex = (int)ws.LoaiKy;
            if (chkTuDongLuyKe != null) chkTuDongLuyKe.Checked = ws.TuDongCongLuyKe;
            if (chkDinhKemBangNhanVien != null) chkDinhKemBangNhanVien.Checked = ws.DinhKemBangChiTietNhanVien;

            // Populate Tab 1: Bộ lọc Tổ/Bộ phận & Danh sách chi tiết từng người
            PopulateBoPhanFilter(ws);
            PopulateEmployeeGridRows();

            numNgayHop.Value = Clamp(ws.NgayHop, 1, 31);
            numThangHop.Value = Clamp(ws.ThangHop, 1, 12);
            numNamHop.Value = Clamp(ws.NamHop, 2020, 2035);
            txtDiaDiem.Text = ws.DiaDiem;
            txtGioKetThuc.Text = ws.GioKetThuc;

            dgvThanhPhanTCHC.Rows.Clear();
            foreach (var p in ws.ThanhPhanTCHC)
                dgvThanhPhanTCHC.Rows.Add(p.HoTen, p.ChucVu);

            dgvThanhPhanPX.Rows.Clear();
            foreach (var p in ws.ThanhPhanPX)
                dgvThanhPhanPX.Rows.Add(p.HoTen, p.ChucVu);

            numLDDauKy.Value = Clamp(ws.LDDauKy, 0, 2000);
            numLDTang.Value = Clamp(ws.LDTang, 0, 500);
            numLDGiam.Value = Clamp(ws.LDGiam, 0, 500);
            numLDCuoiKy.Value = Clamp(ws.LDCuoiKy, 0, 2000);
            numLDQuanLy.Value = Clamp(ws.LDQuanLy, 0, 500);
            numLDTrucTiep.Value = Clamp(ws.LDTrucTiep, 0, 2000);
            txtLyDoBienDong.Text = ws.LyDoBienDong;
            txtGhiChuLD.Text = ws.GhiChuLaoDong;

            RefreshSummaryGridsFromWorkshop(ws);

            // Populate Sản lượng
            txtTenSP.Text = ws.TenSanPhamChinh;
            txtDvtSP.Text = ws.DvtSanPham;
            txtSLKeHoach.Text = WordReportService.FmtDec(ws.SanLuongKH, 2);
            txtSLThucHien.Text = WordReportService.FmtDec(ws.SanLuongTH, 3);
            txtSLTinhLuong.Text = WordReportService.FmtDec(ws.SanLuongTinhLuong, 3);
            txtSLLoai1.Text = WordReportService.FmtDec(ws.SanLuongLoai1, 3);
            txtDGLoai1.Text = WordReportService.FmtDec(ws.DonGiaLoai1, 2);
            txtTTLoai1.Text = WordReportService.FmtInt(ws.ThanhTienLoai1);
            txtSLLoai2.Text = WordReportService.FmtDec(ws.SanLuongLoai2, 3);
            txtDGLoai2.Text = WordReportService.FmtDec(ws.DonGiaLoai2, 2);
            txtTTLoai2.Text = WordReportService.FmtInt(ws.ThanhTienLoai2);
            txtSLLe.Text = WordReportService.FmtDec(ws.SanLuongLe, 3);
            txtDGLe.Text = WordReportService.FmtDec(ws.DonGiaLe, 2);
            txtTTLe.Text = WordReportService.FmtInt(ws.ThanhTienLe);
            txtSanLuongPhuName.ReadOnly = txtSLPhuKL.ReadOnly = txtSLPhuTT.ReadOnly = ws.CauPhoiDaApDung;
            txtSanLuongPhuName.Text = ws.SanLuongPhuName;
            txtSLPhuKL.Text = WordReportService.FmtDec(ws.SanLuongPhuKL, 2);
            txtSLPhuTT.Text = WordReportService.FmtInt(ws.SanLuongPhuTT);
            txtLyDoTruLuong.Text = ws.LyDoTruLuong;

            // Populate GNV & Kiến nghị
            dgvGNV.Rows.Clear();
            foreach (var g in ws.ListGNV)
            {
                int row = dgvGNV.Rows.Add(
                    g.STT,
                    g.NoiDung,
                    g.DVT,
                    WordReportService.FmtDec(g.KhoiLuong, 2),
                    WordReportService.FmtInt(g.DonGia),
                    WordReportService.FmtInt(g.ThanhTien),
                    g.GhiChu
                );
                dgvGNV.Rows[row].Tag=g;
            }

            txtKienNghi.Text = string.Join(Environment.NewLine, ws.KienNghi);

            _isLoadingUI = false;
            RefreshPreview();
        }

        private void PopulateBoPhanFilter(WorkshopData ws)
        {
            _isFilteringEmp = true;
            cboFilterBoPhan.Items.Clear();
            cboFilterBoPhan.Items.Add(string.Format("-- Tất cả các Tổ / Bộ phận ({0} người) --", ws.ListNhanVien.Count));
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in ws.ListNhanVien)
            {
                if (!string.IsNullOrEmpty(e.BoPhan) && !seen.Contains(e.BoPhan))
                {
                    seen.Add(e.BoPhan);
                    cboFilterBoPhan.Items.Add(e.BoPhan);
                }
            }
            cboFilterBoPhan.SelectedIndex = 0;
            _isFilteringEmp = false;
        }

        private void PopulateEmployeeGridRows()
        {
            if (_workshops == null || _currentIndex < 0 || _currentIndex >= _workshops.Count) return;
            _isFilteringEmp = true;
            var ws = _workshops[_currentIndex];

            dgvEmployees.Rows.Clear();
            string filterBP = cboFilterBoPhan.SelectedIndex > 0 ? Convert.ToString(cboFilterBoPhan.SelectedItem) : "";
            string search = (txtSearchEmp.Text ?? "").Trim().ToLowerInvariant();

            foreach (var emp in ws.ListNhanVien)
            {
                if (chkOnlyAnCaDiff != null && chkOnlyAnCaDiff.Checked && Math.Abs(emp.AnCaChenhLech) <= 0.001)
                    continue;
                if (!string.IsNullOrEmpty(filterBP) && !string.Equals(emp.BoPhan, filterBP, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!string.IsNullOrEmpty(search))
                {
                    string h = (emp.HoTen ?? "").ToLowerInvariant();
                    string s = (emp.SDB ?? "").ToLowerInvariant();
                    string cd = (emp.ChucDanh ?? "").ToLowerInvariant();
                    if (!h.Contains(search) && !s.Contains(search) && !cd.Contains(search))
                        continue;
                }

                int rowIdx = dgvEmployees.Rows.Add();
                var row = dgvEmployees.Rows[rowIdx];
                row.Tag = emp;
                WriteEmployeeRecordToRow(row, emp);
            }

            UpdateEmployeeBanner(ws);
            _isFilteringEmp = false;
        }

        private void WriteEmployeeRecordToRow(DataGridViewRow row, EmployeeRecord emp)
        {
            row.Cells[0].Value = emp.STT;
            row.Cells[1].Value = emp.ExcelRow;
            row.Cells[2].Value = emp.SDB;
            row.Cells[3].Value = emp.HoTen;
            row.Cells[4].Value = emp.BoPhan;
            row.Cells[5].Value = emp.ChucDanh;
            row.Cells[6].Value = WordReportService.FmtInt(emp.MucLuongBHXH);
            row.Cells[7].Value = WordReportService.FmtDec(emp.TongCong, 1);
            row.Cells[8].Value = WordReportService.FmtDec(emp.CongSP, 1);
            row.Cells[9].Value = WordReportService.FmtDec(emp.CongLePhep, 1);
            row.Cells[10].Value = WordReportService.FmtDec(emp.CongHocHop, 1);
            row.Cells[11].Value = WordReportService.FmtDec(emp.CongVHTT, 1);
            row.Cells[12].Value = WordReportService.FmtDec(emp.CongBHXH, 1);
            row.Cells[13].Value = WordReportService.FmtDec(emp.CongAnCa, 1);
            row.Cells[14].Value = WordReportService.FmtDec(emp.DiemSanXuat, 2);
            row.Cells[15].Value = emp.XepLoai;
            row.Cells[16].Value = WordReportService.FmtDec(emp.HeSoHuong, 2);
            row.Cells[17].Value = WordReportService.FmtInt(emp.LuongSanPham);
            row.Cells[18].Value = WordReportService.FmtInt(emp.LuongLePhepCD);
            row.Cells[19].Value = WordReportService.FmtInt(emp.PhuCapTien + emp.LuongLamLe + emp.LuongHocVHTT);
            row.Cells[20].Value = WordReportService.FmtInt(emp.TongTienLuong);
            row.Cells[21].Value = WordReportService.FmtDec(emp.AnCaCongThuc, 1);
            row.Cells[22].Value = emp.CoAnCa ? WordReportService.FmtDec(emp.AnCaDonVi, 1) : "";
            double cl = emp.AnCaChenhLech;
            row.Cells[23].Value = emp.CoAnCa ? (Math.Abs(cl) > 0.001 ? cl.ToString("+0.0;-0.0", System.Globalization.CultureInfo.InvariantCulture) : "0") : "";
            row.Cells[23].Style.ForeColor = Math.Abs(cl) > 0.001 ? Color.Red : Color.Black;
            row.Cells[23].Style.Font = Math.Abs(cl) > 0.001 ? new Font("Segoe UI", 9f, FontStyle.Bold) : null;
            row.Cells[24].Value = emp.AnCaNguon == "DV" ? "Đơn vị" : emp.AnCaNguon == "TC" ? "Thủ công" : emp.AnCaNguon == "CT" ? "Công thức" : "Số cũ (Excel)";
        }

        private void ReadRowIntoEmployeeRecord(DataGridViewRow row, EmployeeRecord emp)
        {
            emp.SDB = (Convert.ToString(row.Cells[2].Value) ?? "").Trim();
            emp.HoTen = (Convert.ToString(row.Cells[3].Value) ?? "").Trim();
            emp.BoPhan = (Convert.ToString(row.Cells[4].Value) ?? "").Trim();
            emp.ChucDanh = (Convert.ToString(row.Cells[5].Value) ?? "").Trim();
            emp.MucLuongBHXH = ParseVN(row.Cells[6].Value);
            emp.CongSP = ParseVN(row.Cells[8].Value);
            emp.CongLePhep = ParseVN(row.Cells[9].Value);
            emp.CongHocHop = ParseVN(row.Cells[10].Value);
            emp.CongVHTT = ParseVN(row.Cells[11].Value);
            emp.CongBHXH = ParseVN(row.Cells[12].Value);

            // Công ăn ca: đọc nguồn đã chọn; nếu người dùng gõ tay số mới vào cột "Ăn ca" -> chuyển sang THỦ CÔNG
            string ng = Convert.ToString(row.Cells[24].Value) ?? "";
            emp.AnCaNguon = ng == "Đơn vị" ? "DV" : ng == "Thủ công" ? "TC" : ng == "Công thức" ? "CT" : "";
            double anCaCell = ParseVN(row.Cells[13].Value);
            if (Math.Abs(anCaCell - emp.CongAnCa) > 0.05)
            {
                emp.AnCaThuCong = anCaCell;
                emp.AnCaNguon = "TC";
            }
            else if (emp.AnCaNguon == "TC" && emp.AnCaThuCong <= 0)
            {
                emp.AnCaThuCong = emp.CongAnCa;
            }
            emp.CongAnCa = anCaCell;
            emp.DiemSanXuat = ParseVN(row.Cells[14].Value);
            string xl = (Convert.ToString(row.Cells[15].Value) ?? "A").Trim().ToUpperInvariant();
            emp.XepLoai = string.IsNullOrEmpty(xl) ? "A" : xl;
            emp.HeSoHuong = ParseVN(row.Cells[16].Value);
            emp.LuongSanPham = ParseVN(row.Cells[17].Value);
            emp.LuongLePhepCD = ParseVN(row.Cells[18].Value);
        }

        private void SaveEmployeeGridToWorkshop()
        {
            if (dgvEmployees == null) return;
            foreach (DataGridViewRow row in dgvEmployees.Rows)
            {
                if (row.IsNewRow) continue;
                var emp = row.Tag as EmployeeRecord;
                if (emp != null)
                {
                    ReadRowIntoEmployeeRecord(row, emp);
                }
            }
        }

        private void UpdateEmployeeBanner(WorkshopData ws)
        {
            if (lblEmpSummaryBanner == null || ws == null) return;
            double sCong = 0, sSP = 0, sPhep = 0, sHoc = 0, sBHXH = 0, sAnCa = 0, sLuongSP = 0, sTongLuong = 0;
            foreach (var e in ws.ListNhanVien)
            {
                sCong += e.TongCong;
                sSP += e.CongSP;
                sPhep += e.CongLePhep;
                sHoc += e.CongHocHop;
                sBHXH += e.CongBHXH;
                sAnCa += e.CongAnCa;
                sLuongSP += e.LuongSanPham;
                sTongLuong += e.TongTienLuong;
            }
            lblEmpSummaryBanner.Text = string.Format(
                "Sheet '{0}' ({1} người)  |  Tổng công: {2:N1}  |  Công SP: {3:N1}  |  Phép, Lễ: {4:N1}  |  Học: {5:N1}  |  BHXH: {6:N1}  |  Ăn ca: {7:N1}  |  Lương SP: {8:N0} đ  |  Tổng lương: {9:N0} đ",
                ws.SheetBangCong, ws.ListNhanVien.Count, sCong, sSP, sPhep, sHoc, sBHXH, sAnCa, sLuongSP, sTongLuong);
        }

        private void RefreshSummaryGridsFromWorkshop(WorkshopData ws)
        {
            bool prevLoading = _isLoadingUI;
            _isLoadingUI = true;

            dgvNgayCong.Rows.Clear();
            AddMetricRow(dgvNgayCong, "1. Tổng số công phát sinh", "Công", ws.TongCong, ws.KyQuy.TongCong, ws.KyLuyKe.TongCong, 1);
            AddMetricRow(dgvNgayCong, "2. Công sản phẩm", "Công", ws.CongSP, ws.KyQuy.CongSP, ws.KyLuyKe.CongSP, 1);
            AddMetricRow(dgvNgayCong, "   2.1. Công lao động quản lý", "Công", ws.CongQL, ws.KyQuy.CongQL, ws.KyLuyKe.CongQL, 1);
            AddMetricRow(dgvNgayCong, "   2.2. Công lao động trực tiếp", "Công", ws.CongLDTT, ws.KyQuy.CongLDTT, ws.KyLuyKe.CongLDTT, 1);
            AddMetricRow(dgvNgayCong, "3. Công lễ, phép, tết, học, chế độ", "Công", ws.CongLePhepCD, ws.KyQuy.CongCheDo, ws.KyLuyKe.CongCheDo, 1);
            AddMetricRow(dgvNgayCong, "   - Trong đó: Công lễ, phép LĐTT", "Công", ws.CongLePhepTT, ws.KyQuy.CongCheDoTT, ws.KyLuyKe.CongCheDoTT, 1);
            AddMetricRow(dgvNgayCong, "   - Trong đó: Công học, họp LĐTT", "Công", ws.CongHocHopTT, 0, 0, 1);
            AddMetricRow(dgvNgayCong, "   - Trong đó: Công chế độ Quản lý", "Công", ws.CongCDQL, ws.KyQuy.CongCheDoQL, ws.KyLuyKe.CongCheDoQL, 1);
            AddMetricRow(dgvNgayCong, "4. Công nghỉ ốm, thai sản hưởng BHXH", "Công", ws.CongBHXH, ws.KyQuy.CongBHXH, ws.KyLuyKe.CongBHXH, 1);
            AddMetricRow(dgvNgayCong, "5. Lao động thực tế đi làm BQ", "Người", ws.LDThucTeBQ, ws.KyQuy.LDThucTeBQ, ws.KyLuyKe.LDThucTeBQ, 2);
            AddMetricRow(dgvNgayCong, "   - LĐ trực tiếp đi làm thực tế BQ", "Người", ws.LDTTThucTeBQ, ws.KyQuy.LDTTThucTeBQ, ws.KyLuyKe.LDTTThucTeBQ, 2);
            AddMetricRow(dgvNgayCong, "6. Ngày công thực tế đi làm BQ", "Công/ng", ws.NgayCongBQ, ws.KyQuy.NgayCongBQ, ws.KyLuyKe.NgayCongBQ, 2);
            AddMetricRow(dgvNgayCong, "   - Ngày công BQ Quản lý", "Công/ng", ws.NgayCongBQQL, ws.KyQuy.NgayCongBQQL, ws.KyLuyKe.NgayCongBQQL, 2);
            AddMetricRow(dgvNgayCong, "   - Ngày công BQ LĐ trực tiếp", "Công/ng", ws.NgayCongBQLDTT, ws.KyQuy.NgayCongBQLDTT, ws.KyLuyKe.NgayCongBQLDTT, 2);

            dgvTienLuong.Rows.Clear();
            AddMetricRow(dgvTienLuong, "I. TỔNG TIỀN LƯƠNG TOÀN PX", "Đồng", ws.TongLuongToanPX, ws.KyQuy.TongLuongToanPX, ws.KyLuyKe.TongLuongToanPX, 0);
            AddMetricRow(dgvTienLuong, "1. Tiền lương Lao động trực tiếp", "Đồng", ws.TongLuongTT, ws.KyQuy.TongLuongTT, ws.KyLuyKe.TongLuongTT, 0);
            AddMetricRow(dgvTienLuong, "   1.1. Tiền lương sản phẩm TT", "Đồng", ws.LuongSPTT, ws.KyQuy.LuongSPTT, ws.KyLuyKe.LuongSPTT, 0);
            AddMetricRow(dgvTienLuong, "   1.2. Tiền lương công việc khác (GNV)", "Đồng", ws.LuongCVKhac, ws.KyQuy.LuongCVKhac, ws.KyLuyKe.LuongCVKhac, 0);
            AddMetricRow(dgvTienLuong, "   1.3. Số công việc khác (GNV)", "Công", ws.CongCVKhac, ws.KyQuy.CongCVKhac, ws.KyLuyKe.CongCVKhac, 1);
            AddMetricRow(dgvTienLuong, "   1.4. Tiền lương chế độ (lễ, phép, học)", "Đồng", ws.LuongCheDoTT, ws.KyQuy.LuongCheDoTT, ws.KyLuyKe.LuongCheDoTT, 0);
            AddMetricRow(dgvTienLuong, "   1.5. Tiền lương Trung đại tu (TĐT)", "Đồng", ws.LuongTDT, ws.KyQuy.LuongTDT, ws.KyLuyKe.LuongTDT, 0);
            AddMetricRow(dgvTienLuong, "   1.6. Giảm trừ tiền lương (nếu có)", "Đồng", ws.TruLuong, ws.KyQuy.TruLuong, ws.KyLuyKe.TruLuong, 0);
            AddMetricRow(dgvTienLuong, "2. Tiền lương Lao động quản lý", "Đồng", ws.TongLuongQL, ws.KyQuy.TongLuongQL, ws.KyLuyKe.TongLuongQL, 0);
            AddMetricRow(dgvTienLuong, "   2.1. Tiền lương sản phẩm QL", "Đồng", ws.LuongSPQL, ws.KyQuy.LuongSPQL, ws.KyLuyKe.LuongSPQL, 0);
            AddMetricRow(dgvTienLuong, "   2.2. Tiền lương chế độ QL", "Đồng", ws.LuongCheDoQL, ws.KyQuy.LuongCheDoQL, ws.KyLuyKe.LuongCheDoQL, 0);
            AddMetricRow(dgvTienLuong, "   2.3. Điều chỉnh tiền lương QL", "Đồng", ws.DieuChinhLuongQL, ws.KyQuy.DieuChinhLuongQL, ws.KyLuyKe.DieuChinhLuongQL, 0);
            AddMetricRow(dgvTienLuong, "II. TIỀN LƯƠNG BQ TOÀN PX", "1.000đ", ws.TLBinhQuanToanPX, ws.KyQuy.TLBinhQuanToanPX, ws.KyLuyKe.TLBinhQuanToanPX, 2);
            AddMetricRow(dgvTienLuong, "III. THU NHẬP BQ TOÀN PX", "1.000đ", ws.ThuNhapBQToanPX, ws.KyQuy.ThuNhapBQToanPX, ws.KyLuyKe.ThuNhapBQToanPX, 2);
            AddMetricRow(dgvTienLuong, "IV. TIỀN LƯƠNG SP BQ LĐTT", "1.000đ", ws.TLSPBinhQuanLDTT, ws.KyQuy.TLSPBinhQuanLDTT, ws.KyLuyKe.TLSPBinhQuanLDTT, 2);
            AddMetricRow(dgvTienLuong, "V. THU NHẬP BQ LĐTT", "1.000đ", ws.ThuNhapBQLDTT, ws.KyQuy.ThuNhapBQLDTT, ws.KyLuyKe.ThuNhapBQLDTT, 2);

            dgvAnCaBDDH.Rows.Clear();
            AddMetricRow(dgvAnCaBDDH, "1. Tổng số công Ăn giữa ca", "Công", ws.CongAnCa, ws.KyQuy.CongAnCa, ws.KyLuyKe.CongAnCa, 1);
            AddMetricRow(dgvAnCaBDDH, "2. Thành tiền Ăn giữa ca (30.000đ)", "Đồng", ws.TienAnCa, ws.KyQuy.TienAnCa, ws.KyLuyKe.TienAnCa, 0);
            AddMetricRow(dgvAnCaBDDH, "3. BDĐH Mức 1 (13.000đ) - Số công", "Công", ws.BDDH_Muc1_Cong, ws.KyQuy.BDDH_Muc1_Cong, ws.KyLuyKe.BDDH_Muc1_Cong, 1);
            AddMetricRow(dgvAnCaBDDH, "4. BDĐH Mức 1 (13.000đ) - Thành tiền", "Đồng", ws.BDDH_Muc1_Tien, ws.KyQuy.BDDH_Muc1_Tien, ws.KyLuyKe.BDDH_Muc1_Tien, 0);
            AddMetricRow(dgvAnCaBDDH, "5. BDĐH Mức 2 (20.000đ) - Số công", "Công", ws.BDDH_Muc2_Cong, ws.KyQuy.BDDH_Muc2_Cong, ws.KyLuyKe.BDDH_Muc2_Cong, 1);
            AddMetricRow(dgvAnCaBDDH, "6. BDĐH Mức 2 (20.000đ) - Thành tiền", "Đồng", ws.BDDH_Muc2_Tien, ws.KyQuy.BDDH_Muc2_Tien, ws.KyLuyKe.BDDH_Muc2_Tien, 0);
            AddMetricRow(dgvAnCaBDDH, "7. BDĐH Mức 3 (26.000đ) - Số công", "Công", ws.BDDH_Muc3_Cong, ws.KyQuy.BDDH_Muc3_Cong, ws.KyLuyKe.BDDH_Muc3_Cong, 1);
            AddMetricRow(dgvAnCaBDDH, "8. BDĐH Mức 3 (26.000đ) - Thành tiền", "Đồng", ws.BDDH_Muc3_Tien, ws.KyQuy.BDDH_Muc3_Tien, ws.KyLuyKe.BDDH_Muc3_Tien, 0);
            AddMetricRow(dgvAnCaBDDH, "9. TỔNG TIỀN BỒI DƯỠNG ĐỘC HẠI", "Đồng", ws.TongTienBDDH, ws.KyQuy.TongTienBDDH, ws.KyLuyKe.TongTienBDDH, 0);

            _isLoadingUI = prevLoading;
        }

        private void AddMetricRow(DataGridView dgv, string name, string dvt, double vThang, double vQuy, double vLK, int decimals)
        {
            string sT = decimals == 0 ? WordReportService.FmtInt(vThang) : WordReportService.FmtDec(vThang, decimals);
            string sQ = decimals == 0 ? WordReportService.FmtInt(vQuy) : WordReportService.FmtDec(vQuy, decimals);
            string sLK = decimals == 0 ? WordReportService.FmtInt(vLK) : WordReportService.FmtDec(vLK, decimals);
            dgv.Rows.Add(name, dvt, sT, sQ, sLK);
        }

        private void CommitAllGridEdits()
        {
            if (dgvEmployees != null && dgvEmployees.IsCurrentCellInEditMode) dgvEmployees.EndEdit();
            if (dgvNgayCong != null && dgvNgayCong.IsCurrentCellInEditMode) dgvNgayCong.EndEdit();
            if (dgvTienLuong != null && dgvTienLuong.IsCurrentCellInEditMode) dgvTienLuong.EndEdit();
            if (dgvAnCaBDDH != null && dgvAnCaBDDH.IsCurrentCellInEditMode) dgvAnCaBDDH.EndEdit();
            if (dgvGNV != null && dgvGNV.IsCurrentCellInEditMode) dgvGNV.EndEdit();
            if (dgvThanhPhanTCHC != null && dgvThanhPhanTCHC.IsCurrentCellInEditMode) dgvThanhPhanTCHC.EndEdit();
            if (dgvThanhPhanPX != null && dgvThanhPhanPX.IsCurrentCellInEditMode) dgvThanhPhanPX.EndEdit();
        }

        private void SaveUIToCurrentWorkshop()
        {
            if (_isLoadingUI || _workshops == null || _currentIndex < 0 || _currentIndex >= _workshops.Count) return;
            CommitAllGridEdits();
            SaveEmployeeGridToWorkshop();

            var ws = _workshops[_currentIndex];

            ws.Thang = (int)numThang.Value;
            ws.Nam = (int)numNam.Value;
            ws.LoaiKy = (LoaiKyBaoCao)cboLoaiKy.SelectedIndex;
            if (chkDinhKemBangNhanVien != null) ws.DinhKemBangChiTietNhanVien = chkDinhKemBangNhanVien.Checked;

            ws.NgayHop = (int)numNgayHop.Value;
            ws.ThangHop = (int)numThangHop.Value;
            ws.NamHop = (int)numNamHop.Value;
            ws.DiaDiem = txtDiaDiem.Text.Trim();
            ws.GioKetThuc = txtGioKetThuc.Text.Trim();

            ws.ThanhPhanTCHC.Clear();
            foreach (DataGridViewRow r in dgvThanhPhanTCHC.Rows)
            {
                if (r.IsNewRow) continue;
                string ht = (Convert.ToString(r.Cells[0].Value) ?? "").Trim();
                string cv = (Convert.ToString(r.Cells[1].Value) ?? "").Trim();
                if (!string.IsNullOrEmpty(ht)) ws.ThanhPhanTCHC.Add(new ThanhPhanItem(ht, cv));
            }

            ws.ThanhPhanPX.Clear();
            foreach (DataGridViewRow r in dgvThanhPhanPX.Rows)
            {
                if (r.IsNewRow) continue;
                string ht = (Convert.ToString(r.Cells[0].Value) ?? "").Trim();
                string cv = (Convert.ToString(r.Cells[1].Value) ?? "").Trim();
                if (!string.IsNullOrEmpty(ht)) ws.ThanhPhanPX.Add(new ThanhPhanItem(ht, cv));
            }

            ws.LDDauKy = (int)numLDDauKy.Value;
            ws.LDTang = (int)numLDTang.Value;
            ws.LDGiam = (int)numLDGiam.Value;
            ws.LDCuoiKy = (int)numLDCuoiKy.Value;
            ws.LDQuanLy = (int)numLDQuanLy.Value;
            ws.LDTrucTiep = (int)numLDTrucTiep.Value;
            ws.LyDoBienDong = txtLyDoBienDong.Text.Trim();
            ws.GhiChuLaoDong = txtGhiChuLD.Text.Trim();

            if (dgvNgayCong.Rows.Count >= 14)
            {
                ws.TongCong = ParseVN(dgvNgayCong.Rows[0].Cells[2].Value);
                ws.KyQuy.TongCong = ParseVN(dgvNgayCong.Rows[0].Cells[3].Value);
                ws.KyLuyKe.TongCong = ParseVN(dgvNgayCong.Rows[0].Cells[4].Value);

                ws.CongSP = ParseVN(dgvNgayCong.Rows[1].Cells[2].Value);
                ws.KyQuy.CongSP = ParseVN(dgvNgayCong.Rows[1].Cells[3].Value);
                ws.KyLuyKe.CongSP = ParseVN(dgvNgayCong.Rows[1].Cells[4].Value);

                ws.CongQL = ParseVN(dgvNgayCong.Rows[2].Cells[2].Value);
                ws.KyQuy.CongQL = ParseVN(dgvNgayCong.Rows[2].Cells[3].Value);
                ws.KyLuyKe.CongQL = ParseVN(dgvNgayCong.Rows[2].Cells[4].Value);

                ws.CongLDTT = ParseVN(dgvNgayCong.Rows[3].Cells[2].Value);
                ws.KyQuy.CongLDTT = ParseVN(dgvNgayCong.Rows[3].Cells[3].Value);
                ws.KyLuyKe.CongLDTT = ParseVN(dgvNgayCong.Rows[3].Cells[4].Value);

                ws.CongLePhepCD = ParseVN(dgvNgayCong.Rows[4].Cells[2].Value);
                ws.KyQuy.CongCheDo = ParseVN(dgvNgayCong.Rows[4].Cells[3].Value);
                ws.KyLuyKe.CongCheDo = ParseVN(dgvNgayCong.Rows[4].Cells[4].Value);

                ws.CongLePhepTT = ParseVN(dgvNgayCong.Rows[5].Cells[2].Value);
                ws.CongHocHopTT = ParseVN(dgvNgayCong.Rows[6].Cells[2].Value);
                ws.CongCDQL = ParseVN(dgvNgayCong.Rows[7].Cells[2].Value);

                ws.CongBHXH = ParseVN(dgvNgayCong.Rows[8].Cells[2].Value);
                ws.KyQuy.CongBHXH = ParseVN(dgvNgayCong.Rows[8].Cells[3].Value);
                ws.KyLuyKe.CongBHXH = ParseVN(dgvNgayCong.Rows[8].Cells[4].Value);

                ws.LDThucTeBQ = ParseVN(dgvNgayCong.Rows[9].Cells[2].Value);
                ws.KyQuy.LDThucTeBQ = ParseVN(dgvNgayCong.Rows[9].Cells[3].Value);
                ws.KyLuyKe.LDThucTeBQ = ParseVN(dgvNgayCong.Rows[9].Cells[4].Value);

                ws.LDTTThucTeBQ = ParseVN(dgvNgayCong.Rows[10].Cells[2].Value);
                ws.KyQuy.LDTTThucTeBQ = ParseVN(dgvNgayCong.Rows[10].Cells[3].Value);
                ws.KyLuyKe.LDTTThucTeBQ = ParseVN(dgvNgayCong.Rows[10].Cells[4].Value);

                ws.NgayCongBQ = ParseVN(dgvNgayCong.Rows[11].Cells[2].Value);
                ws.KyQuy.NgayCongBQ = ParseVN(dgvNgayCong.Rows[11].Cells[3].Value);
                ws.KyLuyKe.NgayCongBQ = ParseVN(dgvNgayCong.Rows[11].Cells[4].Value);
            }

            ws.TenSanPhamChinh = txtTenSP.Text.Trim();
            ws.DvtSanPham = txtDvtSP.Text.Trim();
            ws.SanLuongKH = ParseVN(txtSLKeHoach.Text);
            ws.SanLuongTH = ParseVN(txtSLThucHien.Text);
            ws.SanLuongTinhLuong = ParseVN(txtSLTinhLuong.Text);
            ws.SanLuongLoai1 = ParseVN(txtSLLoai1.Text);
            ws.DonGiaLoai1 = ParseVN(txtDGLoai1.Text);
            ws.ThanhTienLoai1 = ParseVN(txtTTLoai1.Text);
            ws.SanLuongLoai2 = ParseVN(txtSLLoai2.Text);
            ws.DonGiaLoai2 = ParseVN(txtDGLoai2.Text);
            ws.ThanhTienLoai2 = ParseVN(txtTTLoai2.Text);
            ws.SanLuongLe = ParseVN(txtSLLe.Text);
            ws.DonGiaLe = ParseVN(txtDGLe.Text);
            ws.ThanhTienLe = ParseVN(txtTTLe.Text);
            if (!ws.CauPhoiDaApDung) ws.SanLuongPhuName = txtSanLuongPhuName.Text.Trim();
            if (!ws.CauPhoiDaApDung) ws.SanLuongPhuKL = ParseVN(txtSLPhuKL.Text);
            if (!ws.CauPhoiDaApDung) ws.SanLuongPhuTT = ParseVN(txtSLPhuTT.Text);
            ws.LyDoTruLuong = txtLyDoTruLuong.Text.Trim();

            if (dgvTienLuong.Rows.Count >= 15)
            {
                ws.TongLuongToanPX = ParseVN(dgvTienLuong.Rows[0].Cells[2].Value);
                ws.KyQuy.TongLuongToanPX = ParseVN(dgvTienLuong.Rows[0].Cells[3].Value);
                ws.KyLuyKe.TongLuongToanPX = ParseVN(dgvTienLuong.Rows[0].Cells[4].Value);

                ws.TongLuongTT = ParseVN(dgvTienLuong.Rows[1].Cells[2].Value);
                ws.KyQuy.TongLuongTT = ParseVN(dgvTienLuong.Rows[1].Cells[3].Value);
                ws.KyLuyKe.TongLuongTT = ParseVN(dgvTienLuong.Rows[1].Cells[4].Value);

                ws.LuongSPTT = ParseVN(dgvTienLuong.Rows[2].Cells[2].Value);
                ws.KyQuy.LuongSPTT = ParseVN(dgvTienLuong.Rows[2].Cells[3].Value);
                ws.KyLuyKe.LuongSPTT = ParseVN(dgvTienLuong.Rows[2].Cells[4].Value);

                ws.LuongCVKhac = ParseVN(dgvTienLuong.Rows[3].Cells[2].Value);
                ws.KyQuy.LuongCVKhac = ParseVN(dgvTienLuong.Rows[3].Cells[3].Value);
                ws.KyLuyKe.LuongCVKhac = ParseVN(dgvTienLuong.Rows[3].Cells[4].Value);

                ws.CongCVKhac = ParseVN(dgvTienLuong.Rows[4].Cells[2].Value);
                ws.KyQuy.CongCVKhac = ParseVN(dgvTienLuong.Rows[4].Cells[3].Value);
                ws.KyLuyKe.CongCVKhac = ParseVN(dgvTienLuong.Rows[4].Cells[4].Value);

                ws.LuongCheDoTT = ParseVN(dgvTienLuong.Rows[5].Cells[2].Value);
                ws.KyQuy.LuongCheDoTT = ParseVN(dgvTienLuong.Rows[5].Cells[3].Value);
                ws.KyLuyKe.LuongCheDoTT = ParseVN(dgvTienLuong.Rows[5].Cells[4].Value);

                ws.LuongTDT = ParseVN(dgvTienLuong.Rows[6].Cells[2].Value);
                ws.KyQuy.LuongTDT = ParseVN(dgvTienLuong.Rows[6].Cells[3].Value);
                ws.KyLuyKe.LuongTDT = ParseVN(dgvTienLuong.Rows[6].Cells[4].Value);

                ws.TruLuong = ParseVN(dgvTienLuong.Rows[7].Cells[2].Value);

                ws.TongLuongQL = ParseVN(dgvTienLuong.Rows[8].Cells[2].Value);
                ws.KyQuy.TongLuongQL = ParseVN(dgvTienLuong.Rows[8].Cells[3].Value);
                ws.KyLuyKe.TongLuongQL = ParseVN(dgvTienLuong.Rows[8].Cells[4].Value);

                ws.LuongSPQL = ParseVN(dgvTienLuong.Rows[9].Cells[2].Value);
                ws.KyQuy.LuongSPQL = ParseVN(dgvTienLuong.Rows[9].Cells[3].Value);
                ws.KyLuyKe.LuongSPQL = ParseVN(dgvTienLuong.Rows[9].Cells[4].Value);

                ws.LuongCheDoQL = ParseVN(dgvTienLuong.Rows[10].Cells[2].Value);
                ws.DieuChinhLuongQL = ParseVN(dgvTienLuong.Rows[11].Cells[2].Value);
                ws.KyQuy.LuongCheDoQL = ParseVN(dgvTienLuong.Rows[10].Cells[3].Value);
                ws.KyLuyKe.LuongCheDoQL = ParseVN(dgvTienLuong.Rows[10].Cells[4].Value);

                ws.DieuChinhLuongQL = ParseVN(dgvTienLuong.Rows[11].Cells[2].Value);
                ws.KyQuy.DieuChinhLuongQL = ParseVN(dgvTienLuong.Rows[11].Cells[3].Value);
                ws.KyLuyKe.DieuChinhLuongQL = ParseVN(dgvTienLuong.Rows[11].Cells[4].Value);

                ws.TLBinhQuanToanPX = ParseVN(dgvTienLuong.Rows[12].Cells[2].Value);
                ws.KyQuy.TLBinhQuanToanPX = ParseVN(dgvTienLuong.Rows[12].Cells[3].Value);
                ws.KyLuyKe.TLBinhQuanToanPX = ParseVN(dgvTienLuong.Rows[12].Cells[4].Value);

                ws.ThuNhapBQToanPX = ParseVN(dgvTienLuong.Rows[13].Cells[2].Value);
                ws.KyQuy.ThuNhapBQToanPX = ParseVN(dgvTienLuong.Rows[13].Cells[3].Value);
                ws.KyLuyKe.ThuNhapBQToanPX = ParseVN(dgvTienLuong.Rows[13].Cells[4].Value);

                ws.TLSPBinhQuanLDTT = ParseVN(dgvTienLuong.Rows[14].Cells[2].Value);
                ws.KyQuy.TLSPBinhQuanLDTT = ParseVN(dgvTienLuong.Rows[14].Cells[3].Value);
                ws.KyLuyKe.TLSPBinhQuanLDTT = ParseVN(dgvTienLuong.Rows[14].Cells[4].Value);

                ws.ThuNhapBQLDTT = ParseVN(dgvTienLuong.Rows[15].Cells[2].Value);
                ws.KyQuy.ThuNhapBQLDTT = ParseVN(dgvTienLuong.Rows[15].Cells[3].Value);
                ws.KyLuyKe.ThuNhapBQLDTT = ParseVN(dgvTienLuong.Rows[15].Cells[4].Value);
            }

            if (dgvAnCaBDDH.Rows.Count >= 9)
            {
                ws.CongAnCa = ParseVN(dgvAnCaBDDH.Rows[0].Cells[2].Value);
                ws.KyQuy.CongAnCa = ParseVN(dgvAnCaBDDH.Rows[0].Cells[3].Value);
                ws.KyLuyKe.CongAnCa = ParseVN(dgvAnCaBDDH.Rows[0].Cells[4].Value);

                ws.TienAnCa = ParseVN(dgvAnCaBDDH.Rows[1].Cells[2].Value);
                ws.KyQuy.TienAnCa = ParseVN(dgvAnCaBDDH.Rows[1].Cells[3].Value);
                ws.KyLuyKe.TienAnCa = ParseVN(dgvAnCaBDDH.Rows[1].Cells[4].Value);

                ws.BDDH_Muc1_Cong = ParseVN(dgvAnCaBDDH.Rows[2].Cells[2].Value);
                ws.KyQuy.BDDH_Muc1_Cong = ParseVN(dgvAnCaBDDH.Rows[2].Cells[3].Value);
                ws.KyLuyKe.BDDH_Muc1_Cong = ParseVN(dgvAnCaBDDH.Rows[2].Cells[4].Value);

                ws.BDDH_Muc1_Tien = ParseVN(dgvAnCaBDDH.Rows[3].Cells[2].Value);
                ws.KyQuy.BDDH_Muc1_Tien = ParseVN(dgvAnCaBDDH.Rows[3].Cells[3].Value);
                ws.KyLuyKe.BDDH_Muc1_Tien = ParseVN(dgvAnCaBDDH.Rows[3].Cells[4].Value);

                ws.BDDH_Muc2_Cong = ParseVN(dgvAnCaBDDH.Rows[4].Cells[2].Value);
                ws.KyQuy.BDDH_Muc2_Cong = ParseVN(dgvAnCaBDDH.Rows[4].Cells[3].Value);
                ws.KyLuyKe.BDDH_Muc2_Cong = ParseVN(dgvAnCaBDDH.Rows[4].Cells[4].Value);

                ws.BDDH_Muc2_Tien = ParseVN(dgvAnCaBDDH.Rows[5].Cells[2].Value);
                ws.KyQuy.BDDH_Muc2_Tien = ParseVN(dgvAnCaBDDH.Rows[5].Cells[3].Value);
                ws.KyLuyKe.BDDH_Muc2_Tien = ParseVN(dgvAnCaBDDH.Rows[5].Cells[4].Value);

                ws.BDDH_Muc3_Cong = ParseVN(dgvAnCaBDDH.Rows[6].Cells[2].Value);
                ws.KyQuy.BDDH_Muc3_Cong = ParseVN(dgvAnCaBDDH.Rows[6].Cells[3].Value);
                ws.KyLuyKe.BDDH_Muc3_Cong = ParseVN(dgvAnCaBDDH.Rows[6].Cells[4].Value);

                ws.BDDH_Muc3_Tien = ParseVN(dgvAnCaBDDH.Rows[7].Cells[2].Value);
                ws.KyQuy.BDDH_Muc3_Tien = ParseVN(dgvAnCaBDDH.Rows[7].Cells[3].Value);
                ws.KyLuyKe.BDDH_Muc3_Tien = ParseVN(dgvAnCaBDDH.Rows[7].Cells[4].Value);

                ws.TongTienBDDH = ParseVN(dgvAnCaBDDH.Rows[8].Cells[2].Value);
                ws.KyQuy.TongTienBDDH = ParseVN(dgvAnCaBDDH.Rows[8].Cells[3].Value);
                ws.KyLuyKe.TongTienBDDH = ParseVN(dgvAnCaBDDH.Rows[8].Cells[4].Value);
            }

            // GNV
            ws.ListGNV.Clear();
            int stt = 1;
            foreach (DataGridViewRow r in dgvGNV.Rows)
            {
                if (r.IsNewRow) continue;
                string nd = (Convert.ToString(r.Cells[1].Value) ?? "").Trim();
                if (string.IsNullOrEmpty(nd)) continue;
                string dvt = (Convert.ToString(r.Cells[2].Value) ?? "").Trim();
                double kl = ParseVN(r.Cells[3].Value);
                double dg = ParseVN(r.Cells[4].Value);
                var previous=r.Tag as ArisingWorkItem;
                bool automatic=previous==null || previous.TuTinhThanhTien;
                double tt=automatic ? ManagementPayroll.RoundMoney(kl*dg,1) : ParseVN(r.Cells[5].Value);
                if(tt==0 && kl>0 && dg>0) tt=ManagementPayroll.RoundMoney(kl*dg,1);
                var item=new ArisingWorkItem(stt++,nd,dvt,kl,dg,tt,Convert.ToString(r.Cells[6].Value) ?? "",previous==null?0:previous.ExcelRow) { TuTinhThanhTien=automatic };
                ws.ListGNV.Add(item); r.Tag=item;
            }

            ws.KienNghi.Clear();
            foreach (string line in txtKienNghi.Lines)
            {
                if (!string.IsNullOrWhiteSpace(line))
                    ws.KienNghi.Add(line.Trim().TrimStart('-', ' '));
            }

            ws.TuDongCongLuyKe = chkTuDongLuyKe != null ? chkTuDongLuyKe.Checked : true;
            if (ws.TuDongCongLuyKe)
            {
                ws.RecalculateDerived();
            }
        }

        private void RefreshPreview()
        {
            if (_workshops == null || _currentIndex < 0) return;
            rtbPreview.Text = WordReportService.GeneratePreviewText(_workshops[_currentIndex]);
        }

        private static double ParseVN(object val)
        {
            if (val == null) return 0;
            string s = Convert.ToString(val).Trim();
            if (string.IsNullOrEmpty(s)) return 0;
            s = s.Replace(".", "").Replace(",", ".");
            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return d;
            return 0;
        }

        private static decimal Clamp(int val, int min, int max)
        {
            if (val < min) return min;
            if (val > max) return max;
            return val;
        }

        private void SetStatus(string msg)
        {
            if (lblStatus != null)
                lblStatus.Text = string.Format("[{0:HH:mm:ss}] {1}", DateTime.Now, msg);
        }

        private void BtnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = "Chọn file Excel tổng hợp nghiệm thu (NT_TCHC.xlsx)";
                ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    _excelFilePath = ofd.FileName;
                    txtExcelPath.Text = _excelFilePath;
                    BtnLoadExcel_Click(sender, e);
                }
            }
        }

        private void BtnLoadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                _excelFilePath = txtExcelPath.Text.Trim();
                if (!File.Exists(_excelFilePath))
                {
                    MessageBox.Show("Không tìm thấy file Excel tại đường dẫn:\n" + _excelFilePath, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _workshops = ExcelService.LoadWorkshopsFromExcel(_excelFilePath);
                int totalEmp = 0;
                foreach (var w in _workshops) totalEmp += w.ListNhanVien.Count;
                LoadCurrentWorkshopToUI();
                SetStatus(string.Format("Đã đọc thành công {0} người lao động & 5 phân xưởng từ file Excel: {1}", totalEmp, Path.GetFileName(_excelFilePath)));
                MessageBox.Show(string.Format("Đã nạp thành công dữ liệu chi tiết từng người ({0} lao động) và bảng nghiệm thu của 5 Phân xưởng từ file:\n{1}", totalEmp, _excelFilePath), "Đọc Excel Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đọc file Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private EmployeeRecord RequestMissingEmployee(EmployeeRecord employee)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "Bổ sung hồ sơ: " + employee.HoTen + " — SDB " + employee.SDB;
                dialog.Size = new Size(640, 440);
                dialog.StartPosition = FormStartPosition.CenterParent;
                var message = new Label { Dock = DockStyle.Top, Height = 55,
                    Text = "Bổ sung tổ/bộ phận, chức danh và hệ số lương cho người chưa có hồ sơ hoặc chưa xác định được tổ." };
                var profile = new EmployeeProfile(employee);
                var properties = new PropertyGrid { Dock = DockStyle.Fill, SelectedObject = profile };
                var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44 };
                var save = new Button { Text = "Thêm và nạp", Width = 130 };
                var skip = new Button { Text = "Bỏ qua người này", Width = 150, DialogResult = DialogResult.Cancel };
                save.Click += delegate {
                    if (string.IsNullOrWhiteSpace(profile.To))
                    { MessageBox.Show(dialog, "Vui lòng nhập tổ/bộ phận."); return; }
                    if (profile.Luong < 0 || profile.HeSo < 0 || profile.PhuCap < 0)
                    { MessageBox.Show(dialog, "Lương và hệ số phải không âm."); return; }
                    profile.Apply(employee);
                    dialog.DialogResult = DialogResult.OK;
                };
                buttons.Controls.Add(save); buttons.Controls.Add(skip);
                dialog.Controls.Add(properties); dialog.Controls.Add(message); dialog.Controls.Add(buttons);
                dialog.AcceptButton = save; dialog.CancelButton = skip;
                return dialog.ShowDialog(this) == DialogResult.OK ? employee : null;
            }
        }

        private class EmployeeProfile
        {
            [System.ComponentModel.DisplayName("Tổ / Bộ phận")]
            public string To { get; set; }
            [System.ComponentModel.DisplayName("Chức danh")]
            public string ChucDanh { get; set; }
            [System.ComponentModel.DisplayName("Mức lương BHXH")]
            public double Luong { get; set; }
            [System.ComponentModel.DisplayName("Hệ số cố định")]
            public double HeSo { get; set; }
            [System.ComponentModel.DisplayName("Phụ cấp chức vụ")]
            public double PhuCap { get; set; }
            [System.ComponentModel.DisplayName("Là lao động quản lý")]
            public bool QuanLy { get; set; }
            public EmployeeProfile(EmployeeRecord e)
            { To = e.ToSanXuat; ChucDanh = e.ChucDanh; Luong = e.MucLuongBHXH; HeSo = e.HeSoCoDinh; PhuCap = e.PhuCapChucVu; QuanLy = e.LaQuanLy; }
            public void Apply(EmployeeRecord e)
            { e.ToSanXuat = e.BoPhan = To.Trim(); e.ChucDanh = ChucDanh; e.MucLuongBHXH = Luong; e.HeSoCoDinh = HeSo; e.PhuCapChucVu = PhuCap; e.LaQuanLy = QuanLy; }
        }

        private void BtnImportExternal_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = string.Format("Chọn file dữ liệu đầu vào cho {0} (Bảng chấm công từng người, Ăn ca, Điểm SX, Nghiệm thu SX, GNV...)", _workshops[_currentIndex].TenPX);
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        SaveUIToCurrentWorkshop();
                        string totalLog = "";
                        foreach (string f in ofd.FileNames)
                        {
                            totalLog += string.Format("File [{0}]:\n", Path.GetFileName(f)) + ExcelService.ImportExternalExcelToWorkshop(f, _workshops[_currentIndex], _excelFilePath, _workshops, RequestMissingEmployee) + "\n";
                        }
                        LoadCurrentWorkshopToUI();
                        SetStatus(string.Format("Đã nạp dữ liệu chi tiết cho {0}.", _workshops[_currentIndex].TenPX));
                        MessageBox.Show(string.Format("Kết quả nạp dữ liệu vào {0}:\n\n{1}", _workshops[_currentIndex].TenPX, totalLog), "Nạp dữ liệu thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khi nạp file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            SaveUIToCurrentWorkshop();
            var ws = _workshops[_currentIndex];
            ws.RecalculateFromEmployees(true);
            LoadCurrentWorkshopToUI();
            SetStatus(string.Format("Đã tính lương chi tiết từng người ({0} LĐ) và tự động cộng lũy kế cho {1}.", ws.ListNhanVien.Count, ws.TenPX));
            MessageBox.Show(string.Format("Đã tính toán xong chi tiết từng người ({0} lao động) và đồng bộ lên Bảng nghiệm thu + Biên bản Word của {1}!\n\n- Tổng công: {2:N1} công (Công SP: {3:N1} công)\n- Ăn ca: {4:N1} công ({5:N0} đ)\n- Tổng tiền lương: {6:N0} đồng",
                ws.ListNhanVien.Count, ws.TenPX, ws.TongCong, ws.CongSP, ws.CongAnCa, ws.TienAnCa, ws.TongLuongToanPX),
                "Tính toán & Đồng bộ thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnSaveExcel_Click(object sender, EventArgs e)
        {
            try
            {
                SaveUIToCurrentWorkshop();
                _excelFilePath = txtExcelPath.Text.Trim();
                ExcelService.SaveWorkshopsToExcel(_excelFilePath, _workshops);
                int totalEmp = 0;
                foreach (var w in _workshops) totalEmp += w.ListNhanVien.Count;
                SetStatus(string.Format("Đã điền chi tiết {0} người lao động & 5 bảng nghiệm thu (giữ nguyên 100% công thức) vào Excel: {1}", totalEmp, _excelFilePath));
                MessageBox.Show(string.Format("Đã tự động điền đầy đủ:\n1. Chi tiết từng người lao động ({0} người) vào các sheet 'Tổng hợp công', 'T kết', 'Gang', 'thép', 'CĐNL', 'K. tuyển'\n2. Các bảng nghiệm thu 'NT thiêu kết', 'NT gang', 'NT Thép', 'NT CĐNL', 'TT BDĐH', '9 tháng ...' (Giữ nguyên 100% công thức Excel!)\n\nFile Excel: {1}", totalEmp, _excelFilePath), "Điền Bảng Nghiệm Thu Excel Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi ghi vào file Excel (Hãy đóng file Excel nếu đang mở trong Microsoft Excel):\n" + ex.Message, "Lỗi ghi Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetDefaultTemplateDocxForWorkshop(WorkshopData ws)
        {
            string tplDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            switch (ws.MaPX)
            {
                case "TK": return Path.Combine(tplDir, "Nt_Thieu_ket_T9.docx");
                case "LG": return Path.Combine(tplDir, "Nt_thang_9_Gang.docx");
                case "LT": return Path.Combine(tplDir, "NT_thep_Thang_9.docx");
                case "CDNL": return Path.Combine(tplDir, "BB_Nt_CDNL_T8.docx");
                default: return "";
            }
        }

        private string GetMasterOutputDocxName(WorkshopData ws)
        {
            switch (ws.MaPX)
            {
                case "TK": return string.Format("Nt_Thieu_ket_T{0}.docx", ws.Thang);
                case "LG": return string.Format("Nt_thang_{0}_Gang.docx", ws.Thang);
                case "LT": return string.Format("NT_thep_Thang_{0}.docx", ws.Thang);
                case "CDNL": return string.Format("BB_Nt_CDNL_T{0}.docx", ws.Thang);
                default: return string.Format("BB_Nt_{0}_T{1}.docx", ws.MaPX, ws.Thang);
            }
        }

        private void BtnExportWord_Click(object sender, EventArgs e)
        {
            try
            {
                SaveUIToCurrentWorkshop();
                var ws = _workshops[_currentIndex];
                string outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Output_BaoCao_Word");
                if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

                string defaultFileName = string.Format("BB_NghiemThu_{0}_Thang_{1}_{2}.docx", ws.MaPX, ws.Thang, ws.Nam);
                using (var sfd = new SaveFileDialog())
                {
                    sfd.InitialDirectory = outDir;
                    sfd.FileName = defaultFileName;
                    sfd.Filter = "Word Document (*.docx)|*.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        WordReportService.ExportWordDocx(ws, sfd.FileName);

                        // Đồng thời tự động điền vào bản sao file Word mẫu lũy kế các tháng của phân xưởng trong Output_BaoCao_Word và Data
                        string tplPath = GetDefaultTemplateDocxForWorkshop(ws);
                        string masterOutPath = Path.Combine(outDir, GetMasterOutputDocxName(ws));
                        if (File.Exists(tplPath) && !string.Equals(Path.GetFullPath(sfd.FileName), Path.GetFullPath(masterOutPath), StringComparison.OrdinalIgnoreCase))
                        {
                            WordReportService.UpdateExistingMasterDocx(ws, tplPath, masterOutPath);
                        }

                        SetStatus(string.Format("Đã điền & xuất Biên bản Word: {0}", sfd.FileName));
                        if (MessageBox.Show(string.Format("Đã điền đầy đủ số liệu nghiệm thu vào Biên bản Word:\n1. File Biên bản tháng: {0}\n2. File Biên bản tổng hợp theo mẫu gốc: {1}\n\nBạn có muốn mở ngay file Biên bản Word vừa điền để kiểm tra không?", sfd.FileName, masterOutPath), "Điền & Xuất Biên Bản Word Thành Công", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        {
                            Process.Start(sfd.FileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất file Word: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnFillExistingDocx_Click(object sender, EventArgs e)
        {
            try
            {
                SaveUIToCurrentWorkshop();
                var ws = _workshops[_currentIndex];
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Title = string.Format("Chọn file Biên bản Word (.docx) của {0} để phần mềm điền trực tiếp số liệu vào", ws.TenPX);
                    ofd.Filter = "Word Document (*.docx)|*.docx";
                    string tplDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
                    if (Directory.Exists(tplDir)) ofd.InitialDirectory = tplDir;

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        WordReportService.UpdateExistingMasterDocx(ws, ofd.FileName, ofd.FileName);
                        SetStatus(string.Format("Đã điền trực tiếp Biên bản tháng {0}/{1} vào file: {2}", ws.Thang, ws.Nam, ofd.FileName));
                        if (MessageBox.Show(string.Format("Đã điền trực tiếp số liệu nghiệm thu của {0} (Tháng {1}/{2}) vào file Word:\n{3}\n\nBạn có muốn mở file Word này ngay không?", ws.TenPX, ws.Thang, ws.Nam, ofd.FileName), "Điền Trực Tiếp Vào File Word Thành Công", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        {
                            Process.Start(ofd.FileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi điền vào file Word (Hãy đóng file Word nếu đang mở trong Microsoft Word):\n" + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportAllWord_Click(object sender, EventArgs e)
        {
            try
            {
                SaveUIToCurrentWorkshop();
                string outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Output_BaoCao_Word");
                if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

                var exportedFiles = new List<string>();
                foreach (var ws in _workshops)
                {
                    string filePath = Path.Combine(outDir, string.Format("BB_NghiemThu_{0}_Thang_{1}_{2}.docx", ws.MaPX, ws.Thang, ws.Nam));
                    WordReportService.ExportWordDocx(ws, filePath);
                    exportedFiles.Add(filePath);

                    string tplPath = GetDefaultTemplateDocxForWorkshop(ws);
                    string masterOutPath = Path.Combine(outDir, GetMasterOutputDocxName(ws));
                    if (File.Exists(tplPath))
                    {
                        WordReportService.UpdateExistingMasterDocx(ws, tplPath, masterOutPath);
                        exportedFiles.Add(masterOutPath);
                    }
                }

                SetStatus(string.Format("Đã điền & xuất thành công {0} file Biên bản Word tại: {1}", exportedFiles.Count, outDir));
                MessageBox.Show(string.Format("Đã điền và xuất thành công trọn bộ {0} file Biên bản nghiệm thu Word (.docx) tại thư mục:\n{1}\n\n- ", exportedFiles.Count, outDir) + string.Join("\n- ", exportedFiles.ConvertAll(Path.GetFileName)), "Điền & Xuất Đồng Loạt Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất hàng loạt: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
