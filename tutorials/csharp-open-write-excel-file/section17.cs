using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section17
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            decimal avg = workSheet["A2:A4"].Avg();
            Console.WriteLine(avg);
        }
    }
}