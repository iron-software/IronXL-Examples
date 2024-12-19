using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section3
    {
        public static void Run()
        {
            /**
            Access Sheet by Index
            anchor-access-specific-worksheet
            **/
            WorkSheet ws = wb.WorkSheets [0]; //by sheet index
        }
    }
}