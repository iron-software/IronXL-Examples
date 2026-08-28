using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section5
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            //Apply range
            IronXL.Range range = WorkSheet ["From Cell Address : To Cell Address"];
        }
    }
}