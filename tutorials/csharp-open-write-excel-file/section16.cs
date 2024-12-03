using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section16
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            decimal sum = workSheet["A2:A4"].Sum();
            Console.WriteLine(sum);
        }
    }
}