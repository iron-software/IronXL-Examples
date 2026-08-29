using IronXL;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section5
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;
            int RowIndex = 0;
            int ColumnIndex = 0;

            IronXL.RangeColumn cell = WorkSheet.Rows [RowIndex].Columns [ColumnIndex];
        }
    }
}