using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section20
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            workSheet["A1:A4"].SortAscending();
            // workSheet["A1:A4"].SortDescending(); to order descending
            workBook.SaveAs("SortedSheet.xlsx");
        }
    }
}