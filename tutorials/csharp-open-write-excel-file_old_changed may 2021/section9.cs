using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section9
    {
        public static void Run()
        {
            /**
            Function AVG
            anchor-avg-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            decimal avg = sheet ["A2:A4"].Avg();
            Console.WriteLine(avg);
        }
    }
}