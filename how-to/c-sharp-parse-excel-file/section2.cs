using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook Wb = IronXL.WorkBook.Load("sample.xlsx");

            //specify WorkSheet
            WorkSheet ws = Wb.GetWorkSheet("SheetName");
        }
    }
}