using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section18
    {
        public static void Run()
        {
            /**
            Name Multiple Sheets
            anchor-read-data-from-multiple-sheets-in-the-same-workbook
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet sheet = workbook.GetWorkSheet("Sheet2");
            var range = sheet ["A2:D2"];
            foreach(var cell in range)
            {
                Console.WriteLine(cell.Text);
            }
        }
    }
}