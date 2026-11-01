using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
namespace CiscoLDTLApp
{
    public class CranePayrollForm : Form
    {
        private readonly WorkshopData workshop;
        private readonly DataGridView grid;
        public CranePayrollForm(WorkshopData ws)
        {
            workshop=ws; Text="Lương cẩu phôi PX Thép — v5.5"; Size=new Size(1100,420); StartPosition=FormStartPosition.CenterParent;
            var note=new Label { Dock=DockStyle.Top,Height=65,Text="Nhập SL thường KHÔNG bao gồm SL lễ. Tiền = SL thường × ĐG + SL lễ × ĐG × 2.\nÁp dụng sẽ THAY phần cẩu phôi cũ trong quỹ SP (không cộng thêm). Giữ riêng các khoản cẩu phôi khác từ Excel. Đối chiếu tổng quỹ cũ: "+WordReportService.FmtInt(ws.SanLuongPhuTT)+" đồng. Bấm Lưu Excel sau khi áp dụng." };
            grid=new DataGridView { Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill };
            foreach(string title in new[] {"Công đoạn","SL thường (không gồm lễ)","SL lễ","Đơn giá (đồng/ĐVT)","Thành tiền (đồng)"}) grid.Columns.Add(title,title);
            grid.Columns[0].FillWeight=250; grid.Columns[0].ReadOnly=true; grid.Columns[4].ReadOnly=true;
            for(int i=0;i<3;i++) { var s=ws.CauPhoi[i]; grid.Rows.Add(s.Name,s.Regular.ToString(CultureInfo.GetCultureInfo("vi-VN")),s.Holiday.ToString(CultureInfo.GetCultureInfo("vi-VN")),s.Rate.ToString(CultureInfo.GetCultureInfo("vi-VN")),WordReportService.FmtInt(s.Amount)); }
            grid.CellEndEdit+=(sender,args)=> { try { var values=ReadStages(); for(int i=0;i<3;i++) grid.Rows[i].Cells[4].Value=WordReportService.FmtInt(values[i].Amount); } catch(ArgumentException) { } };
            var buttons=new FlowLayoutPanel { Dock=DockStyle.Bottom,Height=45 };
            var apply=new Button { Text="Áp dụng",AutoSize=true }; var cancel=new Button { Text="Hủy",DialogResult=DialogResult.Cancel };
            apply.Click+=(sender,args)=> { try { grid.EndEdit(); CranePayroll.Apply(workshop,ReadStages()); DialogResult=DialogResult.OK; Close(); } catch(ArgumentException ex) { MessageBox.Show(ex.Message); } };
            buttons.Controls.Add(apply); buttons.Controls.Add(cancel); CancelButton=cancel;
            Controls.Add(grid); Controls.Add(note); Controls.Add(buttons);
        }
        private CraneStage[] ReadStages()
        {
            var stages=CranePayroll.CreateStages();
            for(int i=0;i<3;i++) {
                double[] values=new double[3];
                for(int j=0;j<3;j++) {
                    string text=Convert.ToString(grid.Rows[i].Cells[j+1].Value).Trim();
                    if(text=="") text="0";
                    if(!double.TryParse(text,NumberStyles.Number,CultureInfo.GetCultureInfo("vi-VN"),out values[j]) || values[j]<0 || double.IsInfinity(values[j]) || double.IsNaN(values[j])) throw new ArgumentException("Nhập số không âm, dấu phẩy thập phân; SL thường không gồm SL lễ.");
                }
                stages[i].Regular=values[0]; stages[i].Holiday=values[1]; stages[i].Rate=values[2];
            }
            return stages;
        }
    }
}
