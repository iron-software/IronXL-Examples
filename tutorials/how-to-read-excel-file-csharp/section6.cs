using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section6
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            IronXL.Cell cell = workSheet["B1"].First();
        }
    }
}