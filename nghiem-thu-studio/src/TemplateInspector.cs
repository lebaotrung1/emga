using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
namespace NghiemThuStudio
{
    public static class TemplateInspector
    {
        public static List<string> Describe(string path)
        {
            if(!string.Equals(Path.GetExtension(path),".docx",StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Mẫu tham chiếu phải là .docx.");
            var doc=new XmlDocument { XmlResolver=null };
            using(var zip=ZipFile.OpenRead(path))
            {
                var part=zip.GetEntry("word/document.xml"); if(part==null) throw new InvalidDataException("Word thiếu nội dung document.xml.");
                using(var stream=part.Open()) using(var reader=XmlReader.Create(stream,new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null })) doc.Load(reader);
            }
            var ns=new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("w","http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var paragraphs=doc.SelectNodes("//w:p",ns).Cast<XmlNode>().Select(x=>x.InnerText).ToList();
            int count=paragraphs.Count(x=>x.Trim().StartsWith("BIÊN BẢN",StringComparison.OrdinalIgnoreCase));
            var names=Catalog.Units.Where(u=>doc.InnerText.Contains(u[1])).Select(u=>u[1]);
            var lines=new List<string> { "Word tham chiếu '"+Path.GetFileName(path)+"': "+count+" tiêu đề biên bản; đơn vị xuất hiện: "+string.Join(", ",names)+".","Không chép số/người ký cũ, không ghi đè mẫu. Xuất mới theo cấu trúc I/II/III, đúng đơn vị và kỳ của hồ sơ." };
            if(count>1) lines.Add("Mẫu là tập nhiều kỳ. Chỉ tham khảo cấu trúc; các kỳ lịch sử không tự trở thành số liệu đã chốt.");
            if(Regex.IsMatch(doc.InnerText,@"năm\s*\d{5}",RegexOptions.IgnoreCase)) lines.Add("Có năm 5 chữ số trong mẫu; cần rà soát lỗi sao chép.");
            return lines;
        }
    }
}
