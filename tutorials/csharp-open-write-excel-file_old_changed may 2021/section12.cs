using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section12
    {
        public static void Run()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            bool max2 =sheet ["A1:A4"].Max(c => c. IsFormula);
            Console.WriteLine(count);
        }
    }
}