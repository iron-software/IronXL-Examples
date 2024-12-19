using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section8
    {
        public static void Run()
        {
            /**
            Function SUM
            anchor-sum-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            decimal sum = sheet ["A2:A4"].Sum();
            Console.WriteLine(sum);
        }
    }
}