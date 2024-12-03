using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section18
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            decimal count = workSheet["A2:A4"].Count();
            Console.WriteLine(count);
        }
    }
}