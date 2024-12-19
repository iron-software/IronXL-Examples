using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section5
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}