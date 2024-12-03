using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section5
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}