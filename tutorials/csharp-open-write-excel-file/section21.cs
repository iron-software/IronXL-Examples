using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section21
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            decimal min = workSheet["A1:A4"].Min();
            Console.WriteLine(min);
        }
    }
}