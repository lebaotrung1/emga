using System;
using System.IO;
using NghiemThuStudio;
class MakeDemo
{
    static int Main(string[] args)
    {
        if(args.Length!=2) throw new ArgumentException("Provide project and sample Word output paths.");
        var d=StudioTests.Complete("LG",9); d.SourceReviewed=false; d.Explanation="DỮ LIỆU GIẢ LẬP ĐỂ THỬ PHẦN MỀM — KHÔNG DÙNG THANH TOÁN.";
        d.Recommendations="Đối chiếu nguồn thật và quy chế trước khi lập hồ sơ thực tế.";
        foreach(var w in d.Works) w.SourceAmount=null;
        var p=new Project { Current=d }; ProjectStore.Save(args[0],p); Report.Export(p,args[1],false);
        Console.WriteLine("Synthetic demonstration files generated; no customer data."); return 0;
    }
}
