using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
namespace CiscoLDTLApp
{
    public class PayrollFormulaError : ArgumentException
    {
        public string Cell { get; private set; }
        public PayrollFormulaError(string cell,string message) : base(cell+": "+message) { Cell=cell; }
    }
    public class PayrollCell
    {
        public string Sheet, Address, Formula, Text;
        public double Number;
        public bool IsNumber, IsError;
        public string Key { get { return Sheet+"!"+Address; } }
    }
    // Restricted numeric Excel interpreter: never runs macros, external links, scripts or functions outside this list.
    public class PayrollFormulaBook
    {
        sealed class EmptyCell { public override string ToString() { return ""; } }
        static readonly object Empty=new EmptyCell();
        public readonly Dictionary<string,PayrollCell> Cells=new Dictionary<string,PayrollCell>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string,string> TextValues=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> Sheets=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string,double> Values=new Dictionary<string,double>(StringComparer.OrdinalIgnoreCase);
        public readonly HashSet<string> Inputs=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string,double> Results=new Dictionary<string,double>(StringComparer.OrdinalIgnoreCase);
        readonly HashSet<string> active=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string,object> cache=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase);
        public static string Key(string sheet,string address) { return sheet+"!"+address.Replace("$","").ToUpperInvariant(); }
        public static PayrollFormulaBook Load(ZipArchive zip)
        {
            var book=new PayrollFormulaBook(); var strings=ExcelService.LoadSharedStrings(zip);
            foreach(var pair in ExcelService.GetSheetMap(zip)) {
                book.Sheets.Add(pair.Key);
                // GetSheetMap also supplies trimmed aliases; avoid loading the same cells twice.
                var doc=ExcelService.LoadSheetXml(zip,pair.Value); if(doc==null) continue;
                var shared=new Dictionary<string,PayrollCell>();
                var pending=new Dictionary<string,string>();
                foreach(XmlElement c in doc.GetElementsByTagName("c","http://schemas.openxmlformats.org/spreadsheetml/2006/main")) {
                    var cell=new PayrollCell { Sheet=pair.Key,Address=c.GetAttribute("r") };
                    foreach(XmlNode n in c.ChildNodes) { if(n.LocalName=="f") { cell.Formula=n.InnerText.Length==0?"#SHARED!":n.InnerText; var f=n as XmlElement; if(f.GetAttribute("t")=="shared") { string id=f.GetAttribute("si"); if(n.InnerText.Length>0) shared[id]=cell; else pending[cell.Address]=id; } } if(n.LocalName=="v") cell.Text=n.InnerText; if(n.LocalName=="is") cell.Text=n.InnerText; }
                    string type=c.GetAttribute("t"); int index;
                    if(type=="s" && int.TryParse(cell.Text,out index)) cell.Text=index>=0 && index<strings.Count?strings[index]:"";
                    cell.IsError=type=="e";
                    cell.IsNumber=type!="s" && type!="inlineStr" && type!="str" && !cell.IsError && double.TryParse(cell.Text,NumberStyles.Float,CultureInfo.InvariantCulture,out cell.Number);
                    if(cell.Formula!=null || cell.Text!=null) book.Cells[cell.Key]=cell;
                }
                foreach(var item in pending) {
                    PayrollCell origin;
                    if(!shared.TryGetValue(item.Value,out origin)) continue;
                    string address=item.Key;
                    var cell=book.Cells[Key(pair.Key,address)];
                    int dc=Col(address)-Col(origin.Address),dr=RowIndex(address)-RowIndex(origin.Address);
                    cell.Formula=Regex.Replace(origin.Formula,@"'([^']|'')*'|""([^""]|"""")*""|(?<![\p{L}\d_])(?<col>\$?[A-Z]{1,3})(?<row>\$?\d+)(?![\p{L}\d_])",m=> {
                        if(!m.Groups["col"].Success) return m.Value;
                        string c=m.Groups["col"].Value,r=m.Groups["row"].Value;
                        int cn=Col(c)+ (c.StartsWith("$")?0:dc),rn=RowIndex(r)+(r.StartsWith("$")?0:dr);
                        if(cn<1||rn<1) return "#REF!";
                        return (c.StartsWith("$")?"$":"")+TimesheetImporter.ColName(cn)+(r.StartsWith("$")?"$":"")+rn;
                    },RegexOptions.IgnoreCase);
                }
            }
            return book;
        }
        static int Col(string s) { int n=0; foreach(char c in s.Replace("$", "")) { if(char.IsDigit(c)) break; n=n*26+char.ToUpperInvariant(c)-'A'+1; } return n; }
        static int RowIndex(string s) { return int.Parse(Regex.Match(s,@"\d+").Value); }
        public void Reset() { cache.Clear(); active.Clear(); Inputs.Clear(); Results.Clear(); }
        public void Set(string sheet,string address,double value) { Values[Key(sheet,address)]=Finite(value); }
        public static double Finite(double value) { if(double.IsNaN(value)||double.IsInfinity(value)) throw new ArgumentException("Giá trị phải hữu hạn."); return value; }
        public double Evaluate(string sheet,string address) { return Numeric(Read(Key(sheet,address))); }
        internal object Read(string key)
        {
            double value; object found; PayrollCell cell; string textValue;
            if(TextValues.TryGetValue(key,out textValue)) return textValue;
            if(Values.TryGetValue(key,out value)) return value;
            if(cache.TryGetValue(key,out found)) return found;
            if(!Cells.TryGetValue(key,out cell)) {
                string sheet=key.Substring(0,key.LastIndexOf('!')); bool exists=Sheets.Contains(sheet);
                if(!exists) throw new PayrollFormulaError(key,"Không có sheet nguồn.");
                Inputs.Add(key); return Empty;
            }
            if(!active.Add(key)) throw new PayrollFormulaError(key,"Tham chiếu vòng; cần đối chiếu nguồn.");
            try {
                object result;
                if(!string.IsNullOrEmpty(cell.Formula)) result=new Parser(this,cell.Sheet,cell.Formula).Parse()();
                else { Inputs.Add(key); if(cell.IsError) throw new PayrollFormulaError(key,"Ô nguồn lỗi "+cell.Text); result=cell.IsNumber?(object)cell.Number:cell.Text==null?Empty:(object)cell.Text; }
                if(result is double) { Finite((double)result); Results[key]=(double)result; }
                cache[key]=result; return result;
            }
            catch(PayrollFormulaError) { throw; }
            catch(Exception ex) { if(!(ex is ArgumentException) && !(ex is FormatException) && !(ex is OverflowException)) throw; throw new PayrollFormulaError(key,ex.Message); }
            finally { active.Remove(key); }
        }
        internal static double Numeric(object value) { if(value is double) return (double)value; if(value is bool) return (bool)value?1:0; if(value==null || value is EmptyCell || Convert.ToString(value)=="") return 0; double parsed; if(value is string && double.TryParse((string)value,NumberStyles.Float,CultureInfo.InvariantCulture,out parsed)) return Finite(parsed); throw new ArgumentException("Cần số, nhận chuỗi văn bản."); }
        class Range { public int Rows,Columns; public readonly List<string> Keys=new List<string>(); }
        class Parser
        {
            readonly PayrollFormulaBook book; readonly string sheet,text; int pos;
            public Parser(PayrollFormulaBook b,string s,string t) { book=b; sheet=s; text=t.StartsWith("=")?t.Substring(1):t; }
            void White() { while(pos<text.Length && char.IsWhiteSpace(text[pos])) pos++; }
            bool Take(string token) { White(); if(pos+token.Length<=text.Length && string.Compare(text,pos,token,0,token.Length,StringComparison.OrdinalIgnoreCase)==0) { pos+=token.Length; return true; } return false; }
            string Match(string pattern) { White(); var m=Regex.Match(text.Substring(pos),"^(?:"+pattern+")",RegexOptions.IgnoreCase); if(!m.Success) return null; pos+=m.Length; return m.Value; }
            public Func<object> Parse() { var node=Compare(); White(); if(pos!=text.Length) throw new ArgumentException("Chưa hỗ trợ công thức tại: "+text.Substring(pos)); return node; }
            Func<object> Compare() {
                var left=Add(); string op=null;
                foreach(string s in new[] {"<=",">=","<>","=","<",">"}) if(Take(s)) { op=s; break; }
                if(op==null) return left; var right=Add();
                return ()=> { object a=left(),b=right(); int c;
                    if(a is string || b is string) c=string.Compare(Convert.ToString(a),Convert.ToString(b),StringComparison.OrdinalIgnoreCase);
                    else c=Numeric(a).CompareTo(Numeric(b));
                    return op=="="?c==0:op=="<>"?c!=0:op=="<"?c<0:op==">"?c>0:op=="<="?c<=0:c>=0; };
            }
            Func<object> Add() { var n=Multiply(); while(true) { bool plus=Take("+"); if(!plus&&!Take("-")) break; var a=n; var b=Multiply(); n=()=>plus?(object)(Numeric(a())+Numeric(b())):(object)(Numeric(a())-Numeric(b())); } return n; }
            Func<object> Multiply() { var n=Power(); while(true) { bool times=Take("*"); if(!times&&!Take("/")) break; var a=n; var b=Power(); n=()=> { double divisor=Numeric(b()); if(!times && divisor==0) throw new ArgumentException("Chia cho 0; kiểm tra công, điểm hoặc định mức."); return times?(object)(Numeric(a())*divisor):(object)(Numeric(a())/divisor); }; } return n; }
            Func<object> Power() { var n=Unary(); if(Take("^")) { var a=n; var b=Power(); n=()=>Math.Pow(Numeric(a()),Numeric(b())); } return n; }
            Func<object> Unary() { if(Take("+")) return Unary(); if(Take("-")) { var n=Unary(); return ()=>-Numeric(n()); } var a=Atom(); while(Take("%")) { var n=a; a=()=>Numeric(n())/100; } return a; }
            Func<object> Atom() {
                if(Take("(")) { var n=Compare(); if(!Take(")")) throw new ArgumentException("Thiếu dấu đóng ngoặc."); return n; }
                if(Take("\"")) { string value=""; bool closed=false; while(pos<text.Length) { if(text[pos]=='\"') { pos++; if(pos<text.Length && text[pos]=='\"') { value+='\"';pos++; } else { closed=true;break; } } else value+=text[pos++]; } if(!closed) throw new ArgumentException("Thiếu dấu đóng chuỗi."); return ()=>value; }
                string error=Match(@"#[A-Z0-9/?!]+"); if(error!=null) return ()=> { throw new ArgumentException("Tham chiếu lỗi "+error+"; nhập giá trị đã đối chiếu kèm lý do cho ô này."); };
                string number=Match(@"(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?"); if(number!=null) { double v=double.Parse(number,CultureInfo.InvariantCulture); return ()=>v; }
                int start=pos; string target=sheet; bool hasPrefix=false;
                string quoted=Match(@"'(?:[^']|'')+'!");
                if(quoted!=null) { target=quoted.Substring(1,quoted.Length-3).Replace("''", "'"); hasPrefix=true; }
                else { string prefix=Match(@"[\p{L}_][\p{L}\p{N}_. ]*!"); if(prefix!=null) { target=prefix.Substring(0,prefix.Length-1); hasPrefix=true; } }
                string address=Match(@"\$?[A-Z]{1,3}\$?\d+(?![\p{L}\d_])");
                if(address!=null) {
                    if(Take(":")) {
                        string end=Match(@"\$?[A-Z]{1,3}\$?\d+"); if(end==null) throw new ArgumentException("Vùng ô không hợp lệ.");
                        var range=new Range(); int c1=Column(address),c2=Column(end),r1=Row(address),r2=Row(end);
                        if(c2<c1||r2<r1||(long)(r2-r1+1)*(c2-c1+1)>100000) throw new ArgumentException("Vùng ô vượt giới hạn.");
                        range.Rows=r2-r1+1; range.Columns=c2-c1+1;
                        for(int r=r1;r<=r2;r++) for(int c=c1;c<=c2;c++) range.Keys.Add(Key(target,TimesheetImporter.ColName(c)+r));
                        return ()=>range;
                    }
                    string key=Key(target,address); return ()=>book.Read(key);
                }
                if(hasPrefix) { string refError=Match(@"#[A-Z0-9/?!]+"); if(refError!=null) return ()=> { throw new ArgumentException("Tham chiếu nguồn lỗi "+target+"!"+refError+"; cần đối chiếu giá trị thay thế."); }; }
                pos=start; string name=Match(@"[A-Z_][A-Z0-9_.]*");
                if(name==null) throw new ArgumentException("Chưa hỗ trợ biểu thức tại: "+text.Substring(pos));
                if(name.Equals("TRUE",StringComparison.OrdinalIgnoreCase)) return ()=>true;
                if(name.Equals("FALSE",StringComparison.OrdinalIgnoreCase)) return ()=>false;
                if(!Take("(")) throw new ArgumentException("Không hỗ trợ tên vùng: "+name);
                var args=new List<Func<object>>(); if(!Take(")")) { do { args.Add(Compare()); } while(Take(",")||Take(";")); if(!Take(")")) throw new ArgumentException("Thiếu dấu đóng hàm."); }
                return ()=>Function(name.ToUpperInvariant(),args);
            }
            int Column(string s) { int n=0; foreach(char c in s.Replace("$", "")) { if(char.IsDigit(c)) break; n=n*26+char.ToUpperInvariant(c)-'A'+1; } return n; }
            int Row(string s) { return int.Parse(Regex.Match(s,@"\d+").Value); }
            IEnumerable<object> Flatten(object value) { var range=value as Range; if(range!=null) { foreach(string key in range.Keys) yield return book.Read(key); } else yield return value; }
            object Function(string name,List<Func<object>> a) {
                if(name=="IF") { if(a.Count<2||a.Count>3) throw new ArgumentException("IF thiếu tham số."); return Numeric(a[0]())!=0?a[1]():(a.Count==3?a[2]():(object)false); }
                if(name=="IFERROR") { try { return a[0](); } catch(ArgumentException) { return a[1](); } }
                if(name=="ROUND") { double n=Numeric(a[0]()),digits=Numeric(a[1]()); if(digits!=Math.Truncate(digits)||digits< -12||digits>12) throw new ArgumentException("Độ làm tròn ngoài giới hạn."); double scale=Math.Pow(10,digits); return Finite(Math.Round(n*scale,0,MidpointRounding.AwayFromZero)/scale); }
                if(name=="ABS") return Math.Abs(Numeric(a[0]()));
                if(name=="LEN") return (double)Convert.ToString(a[0]()).Length;
                if(name=="INDEX") {
                    var range=a[0]() as Range; int row=(int)Numeric(a[1]()),col=a.Count>2?(int)Numeric(a[2]()):1;
                    if(range==null||row<1||col<1||row>range.Rows||col>range.Columns) throw new ArgumentException("INDEX ngoài vùng nguồn.");
                    return book.Read(range.Keys[(row-1)*range.Columns+col-1]);
                }
                if(name=="MATCH") {
                    object lookup=a[0](); var range=a[1]() as Range; double kind=a.Count>2?Numeric(a[2]()):1;
                    if(range==null||kind!=0) throw new ArgumentException("Chỉ hỗ trợ MATCH chính xác (0).");
                    for(int i=0;i<range.Keys.Count;i++) { object value=book.Read(range.Keys[i]);
                        if((lookup is double && value is double && (double)lookup==(double)value) || (lookup is string && value is string && string.Equals((string)lookup,(string)value,StringComparison.OrdinalIgnoreCase))) return (double)(i+1);
                    }
                    throw new ArgumentException("MATCH không tìm thấy mã trong nguồn (#N/A).");
                }
                if(name=="SUMIF"||name=="COUNTIF") {
                    var range=a[0]() as Range; var sum=name=="SUMIF" && a.Count>2?a[2]() as Range:range;
                    if(range==null||sum==null||range.Keys.Count!=sum.Keys.Count) throw new ArgumentException("SUMIF/COUNTIF vùng không khớp.");
                    object criteria=a[1](); double total=0;
                    for(int i=0;i<range.Keys.Count;i++) if(Criteria(book.Read(range.Keys[i]),criteria)) total+=name=="COUNTIF"?1:Numeric(book.Read(sum.Keys[i]));
                    return total;
                }
                if(name!="SUM"&&name!="COUNT"&&name!="COUNTA"&&name!="MIN"&&name!="MAX") throw new ArgumentException("Chưa hỗ trợ hàm "+name+"; cần đối chiếu giá trị.");
                double result=0; bool first=true;
                foreach(var arg in a) foreach(var value in Flatten(arg())) {
                    if(name=="COUNTA") { if(!(value is EmptyCell)) result++; continue; }
                    if(!(value is double)) continue; double v=(double)value;
                    if(name=="COUNT") result++; else if(name=="SUM") result+=v; else { result=first?v:name=="MAX"?Math.Max(result,v):Math.Min(result,v); first=false; }
                }
                return Finite(result);
            }
            bool Criteria(object value,object criteria) {
                string text=Convert.ToString(criteria,CultureInfo.InvariantCulture); var m=Regex.Match(text,@"^(<=|>=|<>|<|>|=)?(.*)$"); string op=m.Groups[1].Value,target=m.Groups[2].Value;
                double number; int c;
                if(value is double && double.TryParse(target,NumberStyles.Float,CultureInfo.InvariantCulture,out number)) c=((double)value).CompareTo(number);
                else if((op==""||op=="=")&&(target.Contains("*")||target.Contains("?"))) return Regex.IsMatch(Convert.ToString(value),"^"+Regex.Escape(target).Replace("\\*",".*").Replace("\\?",".")+"$",RegexOptions.IgnoreCase);
                else c=string.Compare(Convert.ToString(value),target,StringComparison.OrdinalIgnoreCase);
                return op=="<>"?c!=0:op=="<"?c<0:op==">"?c>0:op=="<="?c<=0:op==">="?c>=0:c==0;
            }
        }
    }
}
