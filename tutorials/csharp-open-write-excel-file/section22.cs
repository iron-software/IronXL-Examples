using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section22
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            foreach (var cell in workSheet["B1:B4"])
            {
                Console.WriteLine(cell.Formula);
            }
            Console.ReadKey();
        }
    }
}