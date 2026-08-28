using IronXL;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section9
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;

            /**
            Set Font Style
            anchor-set-font-style
            **/
            WorkSheet ["CellAddress"].Style.Font.Bold =true;
            WorkSheet ["CellAddress"].Style.Font.Italic =true;
        }
    }
}