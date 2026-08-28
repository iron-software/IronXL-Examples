using IronXL;
namespace IronXL.Examples.HowTo.CSharpOpenExcelWorksheet
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook WorkBook = IronXL.WorkBook.Load("sample.xlsx");

            WorkSheet ws = WorkBook.GetWorkSheet("SheetName");
        }
    }
}