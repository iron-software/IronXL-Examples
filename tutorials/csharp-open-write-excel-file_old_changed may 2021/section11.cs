using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section11
    {
        public static void Run()
        {
            /**
            Function MAX
            anchor-max-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            decimal max = sheet ["A2:A4"].Max ();
            Console.WriteLine(max);
        }
    }
}