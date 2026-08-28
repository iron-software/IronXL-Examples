using IronXL;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws1 = workBook.DefaultWorkSheet;

            /**
            Insert WorkSheet Data
            anchor-insert-data-into-worksheets
            **/
            ws1 ["A1"].Value = "Hello World";
        }
    }
}