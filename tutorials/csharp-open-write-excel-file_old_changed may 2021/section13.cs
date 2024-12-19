using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section13
    {
        public static void Run()
        {
            /**
            Function MIN
            anchor-min-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            bool max2 =sheet ["A1:A4"].Min();
            Console.WriteLine(count);
        }
    }
}