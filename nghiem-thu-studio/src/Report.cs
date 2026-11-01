using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Security;

namespace NghiemThuStudio
{
    public static class Report
    {
        static string Esc(string s) { return SecurityElement.Escape(s??""); }
        static string P(string text,bool bold=false,string align="both")
        { return "<w:p><w:pPr><w:spacing w:after='100' w:line='276' w:lineRule='auto'/><w:jc w:val='"+align+"'/></w:pPr><w:r><w:rPr><w:rFonts w:ascii='Times New Roman' w:hAnsi='Times New Roman'/><w:sz w:val='26'/>"+(bold?"<w:b/>":"")+"</w:rPr><w:t xml:space='preserve'>"+Esc(text).Replace("\r", "").Replace("\n", "</w:t><w:br/><w:t xml:space='preserve'>")+"</w:t></w:r></w:p>"; }
        static string Table(IEnumerable<string[]> rows)
        {
            var sb=new StringBuilder("<w:tbl><w:tblPr><w:tblW w:w='5000' w:type='pct'/><w:tblBorders><w:top w:val='single' w:sz='4'/><w:left w:val='single' w:sz='4'/><w:bottom w:val='single' w:sz='4'/><w:right w:val='single' w:sz='4'/><w:insideH w:val='single' w:sz='4'/><w:insideV w:val='single' w:sz='4'/></w:tblBorders></w:tblPr>");
            bool first=true; foreach(var row in rows) { sb.Append("<w:tr>"); foreach(var cell in row) sb.Append("<w:tc><w:tcPr><w:tcW w:w='"+(5000/row.Length)+"' w:type='pct'/></w:tcPr>"+P(cell,first,"left")+"</w:tc>"); sb.Append("</w:tr>"); first=false; }
            return sb.Append("</w:tbl>").ToString();
        }
        static Result Values(Dossier d)
        { return d.Approved && d.ApprovedResult!=null ? new Result { Values=new Dictionary<string,decimal?>(d.ApprovedResult) } : Calc.Evaluate(d); }
        static string PeriodTitle(Dossier d,bool cumulative)
        {
            string title="THÁNG "+d.Month;
            if(cumulative)
            {
                title+=", QUÝ "+new[]{"I","II","III","IV"}[(d.Month-1)/3];
                if(d.Month%3!=0) title+=" ĐẾN THÁNG "+d.Month;
                if(d.Month>3) title+=d.Month==12?" VÀ CẢ NĂM":" VÀ "+d.Month+" THÁNG ĐẦU NĂM";
            }
            return title+(cumulative && d.Month>3?" ":" NĂM ")+d.Year;
        }
        static string PeriodSentence(Dossier d,bool cumulative)
        {
            string text=PeriodTitle(d,cumulative).ToLower(Fmt.Vi);
            foreach(string roman in new[]{"IV","III","II","I"}) text=text.Replace("quý "+roman.ToLowerInvariant(),"quý "+roman);
            return text;
        }
        public static string Preview(Project project,bool cumulative)
        {
            var d=project.Current; var r=Values(d); var sb=new StringBuilder();
            List<int> qm,ym; var q=ProjectStore.Aggregate(project,((d.Month-1)/3)*3+1,d.Month,out qm); var y=ProjectStore.Aggregate(project,1,d.Month,out ym);
            bool official=d.Approved && Workbook.Unchanged(d) && (!cumulative || (qm.Count==0&&ym.Count==0));
            sb.AppendLine(official?"BẢN ĐÃ CHỐT":"DỰ THẢO — CHƯA ĐỦ ĐIỀU KIỆN CHỐT/ĐỦ KỲ");
            sb.AppendLine("BIÊN BẢN ĐÁNH GIÁ CÔNG TÁC LAO ĐỘNG TIỀN LƯƠNG");
            sb.AppendLine(d.Name+" • Tháng "+d.Month+"/"+d.Year);
            sb.AppendLine("Tiền: triệu đồng, 1 số sau dấu phẩy. Đơn giá ghi đồng/ĐVT; số liệu gốc giữ nguyên đồng.");
            foreach(string key in new[]{"productT","productM","gnv","direct","management","total","meal","bddh","average"})
                sb.AppendLine(Calc.Labels(key)+": "+Fmt.Money(r.Get(key))+" triệu đồng"+(key=="average"?"/người/tháng":""));
            if(cumulative)
            {
                sb.AppendLine("Quý đến tháng hiện tại: "+(qm.Count>0?"Thiếu kỳ đã chốt: "+string.Join(", ",qm):Fmt.Money(q.Get("total"))+" triệu đồng"));
                sb.AppendLine("Lũy kế năm: "+(ym.Count>0?"Thiếu kỳ đã chốt: "+string.Join(", ",ym):Fmt.Money(y.Get("total"))+" triệu đồng"));
            }
            sb.AppendLine("\nCông thức: TT = SL1×ĐG1 + SL2×ĐG2 + SL lễ×(2×ĐG1); QL dùng bộ đơn giá quản lý riêng.");
            sb.AppendLine("GNV = KL×ĐG×hệ số; chế độ và điều chỉnh ghi riêng. Mỗi dòng có mức làm tròn rõ ràng.");
            if(r.Issues.Count>0) sb.AppendLine("\nCẦN ĐỐI CHIẾU:\n"+string.Join("\n",r.Issues));
            return sb.ToString();
        }
        public static void Export(Project project,string path,bool cumulative)
        {
            ProjectStore.GuardOutput(project,path,".docx");
            var d=project.Current; var r=Values(d); List<int> qm,ym;
            var q=ProjectStore.Aggregate(project,((d.Month-1)/3)*3+1,d.Month,out qm); var y=ProjectStore.Aggregate(project,1,d.Month,out ym);
            bool official=d.Approved && Workbook.Unchanged(d) && (!cumulative || (qm.Count==0&&ym.Count==0));
            var body=new StringBuilder();
            body.Append(Table(new[]{new[]{"TỔNG CÔNG TY KHOÁNG SẢN - TKV\nCÔNG TY CP GANG THÉP CAO BẰNG","CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc"}}));
            if(!official) body.Append(P("DỰ THẢO — CHƯA CHỐT HOẶC CHƯA ĐỦ CÁC KỲ LŨY KẾ",true,"center"));
            body.Append(P("BIÊN BẢN ĐÁNH GIÁ",true,"center"));
            body.Append(P("CÔNG TÁC LAO ĐỘNG TIỀN LƯƠNG "+PeriodTitle(d,cumulative),true,"center"));
            body.Append(P("Đơn vị: "+d.Name,true,"center"));
            body.Append(P("Hôm nay, ngày "+d.MeetingDate.ToString("dd/MM/yyyy",Fmt.Vi)+", tại "+d.Location+", tiến hành nghiệm thu công tác lao động tiền lương "+PeriodSentence(d,cumulative)+" của "+d.Name+"."));
            body.Append(P("I. Thành phần",true));
            body.Append(P("1. Phòng Tổ chức Hành chính: "+Empty(d.AttendeesOffice)));
            body.Append(P("2. "+d.Name+": "+Empty(d.AttendeesUnit)));
            body.Append(P("II. Đánh giá kết quả thực hiện công tác LĐTL",true));
            body.Append(P("1. Công tác quản lý, sử dụng lao động",true));
            body.Append(P("Lao động cuối kỳ: "+Fmt.Number(d.V("head"))+" người; trong đó quản lý: "+Fmt.Number(d.V("headm"))+" người. Tổng công: "+Fmt.Number(r.Get("days"))+" công; công sản phẩm trực tiếp: "+Fmt.Number(d.V("dt"))+" công; quản lý: "+Fmt.Number(d.V("dm"))+" công; công chế độ: "+Fmt.Number(d.V("other"))+" công; BHXH: "+Fmt.Number(d.V("bh"))+" công."));
            body.Append(P("Lao động thực tế bình quân theo "+Fmt.Number(d.V("standard"))+" ngày công chuẩn: "+Fmt.Number(r.Get("peopleMonths"))+" người."));
            if(cumulative)
            {
                var labor=new List<string[]> { new[]{"Chỉ tiêu công","Tháng "+d.Month,"Quý đến tháng "+d.Month,"Lũy kế năm"} };
                foreach(var item in new[]{new[]{"days","Tổng công"},new[]{"spdays","Công sản phẩm"},new[]{"otherDays","Công chế độ"},new[]{"bhDays","Công BHXH"}})
                    labor.Add(new[]{item[1],Fmt.Number(r.Get(item[0])),qm.Count==0?Fmt.Number(q.Get(item[0])):"Chưa đủ kỳ",ym.Count==0?Fmt.Number(y.Get(item[0])):"Chưa đủ kỳ"});
                body.Append(Table(labor));
            }
            body.Append(P("Công việc phát sinh trong tháng",true));
            var works=new List<string[]> { new[]{"Nội dung","ĐVT","Khối lượng","Đơn giá (đ/ĐVT)","Hệ số","Thành tiền (triệu đồng)"} };
            foreach(var w in d.Works.Where(x=>x.Kind=="GNV")) works.Add(new[]{w.Description,w.Unit,Fmt.Number(w.Quantity),Fmt.Number(w.Price),Fmt.Number(w.Factor),Fmt.Money(w.Amount)});
            if(!d.Works.Any(x=>x.Kind=="GNV")) works.Add(new[]{"Không có GNV trong hồ sơ","","","","","0,0"}); body.Append(Table(works));
            body.Append(P("2. Tiền lương, ăn ca thực hiện",true));
            body.Append(P("Đơn vị các khoản tiền: triệu đồng, đúng 1 chữ số thập phân. Đơn giá ghi riêng bằng đồng/ĐVT."));
            body.Append(P("2.1. Tiền lương",true));
            var production=new List<string[]> { new[]{"Loại sản phẩm","Sản lượng ("+d.QuantityUnit+")","ĐG trực tiếp (đ)","ĐG quản lý (đ)","Quỹ trực tiếp (triệu)","Quỹ quản lý (triệu)"} };
            production.Add(new[]{"Loại 1",Fmt.Number(d.V("q1")),Fmt.Number(d.V("p1")),Fmt.Number(d.V("m1")),Fmt.Money(r.Get("tt1")),Fmt.Money(r.Get("ql1"))});
            production.Add(new[]{"Loại 2",Fmt.Number(d.V("q2")),Fmt.Number(d.V("p2")),Fmt.Number(d.V("m2")),Fmt.Money(r.Get("tt2")),Fmt.Money(r.Get("ql2"))});
            production.Add(new[]{"Ngày lễ (= 2 × ĐG loại 1)",Fmt.Number(d.V("qh")),Fmt.Number(d.V("p1")*2),Fmt.Number(d.V("m1")*2),Fmt.Money(r.Get("tth")),Fmt.Money(r.Get("qlh"))}); foreach(var w in d.Works.Where(x=>x.Kind=="SP phụ")) production.Add(new[]{w.Description,Fmt.Number(w.Quantity),Fmt.Number(w.Price),"Không áp dụng",Fmt.Money(w.Amount),"Không áp dụng"});
            body.Append(Table(production));
            var totals=new List<string[]> { cumulative?new[]{"Chỉ tiêu (triệu đồng)","Tháng "+d.Month,"Quý đến tháng "+d.Month,"Lũy kế đến tháng "+d.Month}:new[]{"Chỉ tiêu (triệu đồng)","Tháng "+d.Month} };
            foreach(string key in new[]{"direct","productT","extraProduct","gnv","regimeT","adjustT","management","productM","regimeM","adjustM","total","average"})
            {
                string label=Calc.Labels(key)+(key=="average"?" / người / tháng":"");
                totals.Add(cumulative?new[]{label,Fmt.Money(r.Get(key)),qm.Count==0?Fmt.Money(q.Get(key)):"Chưa đủ kỳ",ym.Count==0?Fmt.Money(y.Get(key)):"Chưa đủ kỳ"}:new[]{label,Fmt.Money(r.Get(key))});
            }
            body.Append(Table(totals));
            if(cumulative && (qm.Count>0 || ym.Count>0)) body.Append(P("Chưa đủ kỳ đã chốt. Quý thiếu tháng: "+string.Join(", ",qm)+"; năm thiếu tháng: "+string.Join(", ",ym)+". Không coi tháng thiếu là 0."));
            body.Append(P("2.2. Ăn ca",true)); body.Append(P("Công ăn ca: "+Fmt.Number(d.V("meal"))+" công × "+Fmt.Number(d.V("mealprice"))+" đ/công; thành tiền "+Fmt.Money(r.Get("meal"))+" triệu đồng."));
            if(cumulative) body.Append(Table(new[]{new[]{"Ăn ca","Tháng "+d.Month,"Quý đến tháng "+d.Month,"Lũy kế năm"},new[]{"Công",Fmt.Number(r.Get("mealDays")),qm.Count==0?Fmt.Number(q.Get("mealDays")):"Chưa đủ kỳ",ym.Count==0?Fmt.Number(y.Get("mealDays")):"Chưa đủ kỳ"},new[]{"Tiền (triệu đồng)",Fmt.Money(r.Get("meal")),qm.Count==0?Fmt.Money(q.Get("meal")):"Chưa đủ kỳ",ym.Count==0?Fmt.Money(y.Get("meal")):"Chưa đủ kỳ"}}));
            body.Append(P("2.3. Bồi dưỡng độc hại bằng hiện vật",true));
            var bddh=new List<string[]> { cumulative?new[]{"Mức","ĐG (đ/công)","Công tháng","Tiền tháng (triệu)","Công quý","Tiền quý (triệu)","Công LK","Tiền LK (triệu)"}:new[]{"Mức","Công hưởng","Đơn giá (đ/công)","Tiền (triệu đồng)"} };
            for(int i=1;i<=3;i++) bddh.Add(cumulative?new[]{"Mức "+i,Fmt.Number(d.V("bp"+i)),Fmt.Number(d.V("b"+i)),Fmt.Money(r.Get("bddh"+i)),qm.Count==0?Fmt.Number(q.Get("bddhDays"+i)):"Chưa đủ kỳ",qm.Count==0?Fmt.Money(q.Get("bddh"+i)):"Chưa đủ kỳ",ym.Count==0?Fmt.Number(y.Get("bddhDays"+i)):"Chưa đủ kỳ",ym.Count==0?Fmt.Money(y.Get("bddh"+i)):"Chưa đủ kỳ"}:new[]{"Mức "+i,Fmt.Number(d.V("b"+i)),Fmt.Number(d.V("bp"+i)),Fmt.Money(r.Get("bddh"+i))});
            bddh.Add(cumulative?new[]{"Tổng cộng","","",Fmt.Money(r.Get("bddh")),"",qm.Count==0?Fmt.Money(q.Get("bddh")):"Chưa đủ kỳ","",ym.Count==0?Fmt.Money(y.Get("bddh")):"Chưa đủ kỳ"}:new[]{"Tổng cộng","","",Fmt.Money(r.Get("bddh"))}); body.Append(Table(bddh));
            body.Append(P("III. Kiến nghị, đề xuất",true)); body.Append(P(Empty(d.Recommendations)));
            if(!string.IsNullOrWhiteSpace(d.Explanation)) { body.Append(P("Đối chiếu và giải trình",true)); body.Append(P(d.Explanation)); }
            if(!official) foreach(string issue in ProjectStore.ApprovalProblems(d)) body.Append(P("Cần hoàn thiện: "+issue));
            body.Append(Table(new[]{new[]{"QUẢN ĐỐC\n\n\n"+d.SignerUnit,"PHÒNG TCHC\n\n\n"+d.SignerOffice}}));
            body.Append(P("Mã hồ sơ: "+d.Id+" • Kỳ "+d.Month+"/"+d.Year+" • "+(official?"Đã chốt "+d.ApprovedAt:"Dự thảo")+". Số liệu tính bằng đồng; báo cáo chỉ làm tròn hiển thị."));
            body.Append("<w:sectPr><w:pgSz w:w='11906' w:h='16838'/><w:pgMar w:top='1134' w:right='1134' w:bottom='1134' w:left='1701'/></w:sectPr>");
            string xml="<?xml version='1.0' encoding='UTF-8'?><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>"+body+"</w:body></w:document>";
            string temp=Path.GetFullPath(path)+".tmp-"+Guid.NewGuid().ToString("N");
            try
            {
                using(var zip=ZipFile.Open(temp,ZipArchiveMode.Create))
                {
                    Write(zip,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>");
                    Write(zip,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");
                    Write(zip,"word/document.xml",xml);
                }
                if(File.Exists(path)) File.Replace(temp,path,path+".bak"); else File.Move(temp,path);
            }
            finally { if(File.Exists(temp)) File.Delete(temp); }
        }
        static string Empty(string s) { return string.IsNullOrWhiteSpace(s)?"Chưa nhập":s; }
        static void Write(ZipArchive zip,string path,string content) { using(var writer=new StreamWriter(zip.CreateEntry(path).Open(),new UTF8Encoding(false))) writer.Write(content); }
    }
}
