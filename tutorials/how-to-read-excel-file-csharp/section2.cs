using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
        }
    }
}