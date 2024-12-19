using IronXL.Excel;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section6
    {
        public static void Run()
        {
            workSheet.Merge("A1:B1");
            
            workSheet.AutoSizeColumn(0, false);
            workSheet.AutoSizeColumn(1, false);
        }
    }
}