using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpExcelInterop
{
    public static class Section4
    {
        public static void Run()
        {
            WorkSheet.Rows [RowIndex].Replace("old value", "new value");
        }
    }
}