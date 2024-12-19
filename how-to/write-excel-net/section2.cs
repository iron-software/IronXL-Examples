using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section2
    {
        public static void Run()
        {
            // Open Excel WorkSheet
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
        }
    }
}