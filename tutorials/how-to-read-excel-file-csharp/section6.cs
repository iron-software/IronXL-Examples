using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section6
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            IronXL.Cell cell = workSheet["B1"].First();
        }
    }
}