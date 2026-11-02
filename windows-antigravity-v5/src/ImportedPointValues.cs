using System;
namespace CiscoLDTLApp
{
    // Keep the source's already graded points; applying A/B/C again would grade them twice.
    internal sealed class ImportedPointValues
    {
        internal readonly double Production,Weight;
        internal readonly string Rating;
        internal ImportedPointValues(EmployeeRecord e){Production=e.DiemSanXuat;Weight=e.HeSoHuong;Rating=e.XepLoai??"A";}
        internal bool Matches(EmployeeRecord e){return Math.Abs(e.DiemSanXuat-Production)<0.005 && string.Equals(e.XepLoai??"A",Rating,StringComparison.OrdinalIgnoreCase);}
    }
}
