using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using NghiemThuStudio;
class StudioUiTests
{
    static int count;
    static void Check(bool ok,string message) { count++; if(!ok) throw new Exception(message); }
    static FieldInfo Field(string name) { return typeof(MainForm).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance); }
    static void Call(MainForm form,string name) { typeof(MainForm).GetMethod(name,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,null); }
    static void Edit(DataGridView grid,int row,int col,string value)
    { grid.Rows[row].Cells[col].Value=value; typeof(DataGridView).GetMethod("OnCellEndEdit",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(grid,new object[]{new DataGridViewCellEventArgs(col,row)}); Application.DoEvents(); }
    [STAThread] static int Main()
    {
        Application.EnableVisualStyles(); string dir=Path.Combine(Path.GetTempPath(),"nghiemthu-ui-"+Guid.NewGuid().ToString("N")); Directory.CreateDirectory(dir);
        try
        {
            using(var form=new MainForm(()=>false))
            {
                form.Show(); Application.DoEvents();
                var tabs=(TabControl)Field("tabs").GetValue(form); Check(tabs.TabPages.Count==6,"Six clear workflow stages");
                var p=new Project { Current=StudioTests.Complete() }; Field("project").SetValue(form,p); Field("projectPath").SetValue(form,Path.Combine(dir,"ui.ntproject")); Call(form,"LoadAll"); Application.DoEvents();
                var inputs=(DataGridView)Field("inputs").GetValue(form); var works=(DataGridView)Field("works").GetValue(form);
                Check(inputs.Rows.Count==p.Current.Inputs.Count && works.Columns[6].ReadOnly,"Inputs visible and computed GNV amount read-only");
                Call(form,"Approve"); Check(p.Current.Approved && p.Closed.Count==1 && File.Exists(Path.Combine(dir,"ui.ntproject")),"UI approves and atomically persists snapshot");
                int holiday=p.Current.Inputs.FindIndex(m=>m.Key=="qh"); tabs.SelectedIndex=1;
                Edit(inputs,holiday,1,"0"); Check(Calc.Evaluate(p.Current).Get("tth")==0 && Calc.Evaluate(p.Current).Get("qlh")==0,"UI zero quantity recalculates both funds");
                Check(!p.Current.Approved && !((CheckBox)Field("reviewed").GetValue(form)).Checked && p.Closed[0].Approved,"Changing input resets approval/review without damaging closed history");
                Edit(inputs,holiday,1,""); Check(!Calc.Evaluate(p.Current).Get("total").HasValue,"UI empty means missing, not stale money or silently zero");
                Edit(inputs,holiday,1,"5"); tabs.SelectedIndex=2;
                Edit(works,0,2,"0"); Check(p.Current.Works[0].Amount==0 && Calc.Evaluate(p.Current).Get("gnv")==0,"GNV edit updates model/output");
                Edit(works,0,2,"3,5"); Edit(works,0,3,"120.000"); Check(p.Current.Works[0].Amount==420000,"Vietnamese decimal edit works");
                var preview=(RichTextBox)Field("preview").GetValue(form); Check(preview.Text.Contains("DỰ THẢO") && preview.Text.Contains("15,8 triệu đồng"),"Preview immediately shows status and money");
                ((CheckBox)Field("cumulative").GetValue(form)).Checked=true; Check(preview.Text.Contains("Thiếu kỳ đã chốt"),"Incomplete periods explicitly marked");
                var bad=(DataGridViewCellValidatingEventArgs)Activator.CreateInstance(typeof(DataGridViewCellValidatingEventArgs),BindingFlags.Instance|BindingFlags.NonPublic,null,new object[]{1,holiday,"not a number"},null);
                typeof(DataGridView).GetMethod("OnCellValidating",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(inputs,new object[]{bad}); Check(bad.Cancel,"Invalid numeric input prevented"); inputs.Rows[holiday].ErrorText="";
                tabs.SelectedIndex=1; Application.DoEvents();
                string screenshot=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"inputs.png");
                using(var bitmap=new Bitmap(form.Width,form.Height)) { using(var graphics=Graphics.FromImage(bitmap)) graphics.CopyFromScreen(form.Location,Point.Empty,form.Size); bitmap.Save(screenshot); }
                tabs.SelectedIndex=4; Application.DoEvents();
                using(var bitmap=new Bitmap(form.Width,form.Height)) { using(var graphics=Graphics.FromImage(bitmap)) graphics.CopyFromScreen(form.Location,Point.Empty,form.Size); bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"results.png")); }
                form.Close(); Check(!form.IsDisposed && form.Visible,"Canceling close keeps the draft open");
                Edit(inputs,holiday,1,"0"); Check(Calc.Evaluate(p.Current).Get("tth")==0,"Editing still works after canceling close");
                Field("dirty").SetValue(form,false); form.Close();
            }
            Console.WriteLine("PASS: "+count+" Studio UI workflow checks; synthetic screenshots saved."); return 0;
        }
        finally { Directory.Delete(dir,true); }
    }
}
