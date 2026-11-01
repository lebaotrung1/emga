using System;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.Versioning;
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: TargetFramework(".NETFramework,Version=v4.5",FrameworkDisplayName=".NET Framework 4.5")]
namespace NghiemThuStudio
{
    static class Program
    {
        [STAThread] static void Main()
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException+=(s,e)=>MessageBox.Show(e.Exception.Message,"Không hoàn tất thao tác",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            Application.Run(new MainForm());
        }
    }
}
