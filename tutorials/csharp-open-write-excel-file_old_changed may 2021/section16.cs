using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section16
    {
        public static void Run()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
            var sheet = workbook.WorkSheets.First();
            foreach(var cell in sheet ["B1:B4"])
            {
                Console.WriteLine(cell.Formula); 
            }
            Console.ReadKey();
        }
    }
}