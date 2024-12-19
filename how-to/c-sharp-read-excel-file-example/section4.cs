using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section4
    {
        public static void Run()
        {
            string val = WorkSheet.Rows [RowIndex].Columns [ColumnIndex].ToString();
        }
    }
}