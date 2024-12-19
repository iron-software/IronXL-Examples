using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}