using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section4
    {
        public static void Run()
        {
            workSheet.Merge("A1:A3");
            
            workSheet.AutoSizeRow(0, false);
            workSheet.AutoSizeRow(1, false);
            workSheet.AutoSizeRow(2, false);
        }
    }
}