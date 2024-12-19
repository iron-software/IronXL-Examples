using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section10
    {
        public static void Run()
        {
            workSheet.Rows[RowIndex].Replace("old value", "new value");
        }
    }
}