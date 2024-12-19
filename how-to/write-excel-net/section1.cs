using IronXL.Excel;
namespace IronXL.Examples.HowTo.WriteExcelNet
{
    public static class Section1
    {
        public static void Run()
        {
            // Load Excel file in the project
            WorkBook workBook = WorkBook.Load("path");
        }
    }
}