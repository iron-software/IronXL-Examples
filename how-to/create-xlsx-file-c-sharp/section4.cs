using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section4
    {
        public static void Run()
        {
            /**
            Create WorkSheets
            anchor-create-an-excel-worksheet
            **/
            WorkSheet ws2 = wb.CreateWorkSheet("sheet2");
            WorkSheet ws3 = wb.CreateWorkSheet("sheet3");
        }
    }
}