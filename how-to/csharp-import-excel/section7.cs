using IronXL;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section7
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            IronXL.Range range = WorkSheet ["starting Cell Address : Ending Cell Address"];
        }
    }
}