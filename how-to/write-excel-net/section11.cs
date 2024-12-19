using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section11
    {
        public static void Run()
        {
            workSheet.Columns[ColumnIndex].Replace("old value", "new Value");
        }
    }
}