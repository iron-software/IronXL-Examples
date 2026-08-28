using IronXL;
namespace IronXL.Examples.GettingStarted.CSharpExcelInterop
{
    public static class Section5
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;
            int ColumnIndex = 0;

            ws.Columns[ColumnIndex].Replace("old value", "new Value");
        }
    }
}