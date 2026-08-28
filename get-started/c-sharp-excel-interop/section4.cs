using IronXL;
namespace IronXL.Examples.GettingStarted.CSharpExcelInterop
{
    public static class Section4
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;
            int RowIndex = 0;

            ws.Rows[RowIndex].Replace("old value", "new value");
        }
    }
}