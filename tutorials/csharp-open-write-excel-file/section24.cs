using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section24
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
            var range = workSheet["A2:D2"];
            foreach (var cell in range)
            {
                Console.WriteLine(cell.Text);
            }
        }
    }
}