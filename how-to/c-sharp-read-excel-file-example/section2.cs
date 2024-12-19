using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section2
    {
        public static void Run()
        {
            //Open Excel WorkSheet
            WorkSheet ws = wb.GetWorkSheet("SheetName");
        }
    }
}