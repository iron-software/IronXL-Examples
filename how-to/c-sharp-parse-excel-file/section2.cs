using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section2
    {
        public static void Run()
        {
            //specify WorkSheet
            WorkSheet ws = Wb.GetWorkSheet("SheetName");
        }
    }
}