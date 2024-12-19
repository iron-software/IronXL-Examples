using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section10
    {
        public static void Run()
        {
            /**
            Function Count
            anchor-count-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            decimal count = sheet ["A2:A4"].Count();
            Console.WriteLine(count);
        }
    }
}