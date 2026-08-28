using System.Data;
using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section4
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.

            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook WorkBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = WorkBook.DefaultWorkSheet;
            int ColumnIndex = 0;
            int RowIndex = 0;

            string val = WorkSheet.Rows [RowIndex].Columns [ColumnIndex].ToString();
        }
    }
}