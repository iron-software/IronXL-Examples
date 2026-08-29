using IronXL;
namespace IronXL.Examples.HowTo.CSharpParseExcelFile
{
    public static class Section5
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            var array = WorkSheet ["From:To"].ToArray();
        }
    }
}