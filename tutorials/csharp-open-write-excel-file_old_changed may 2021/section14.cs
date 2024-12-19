using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section14
    {
        public static void Run()
        {
            /**
            Function Order Cells
            anchor-order-cells-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            sheet ["A1:A4"].SortAscending(); //or use > sheet ["A1:A4"].SortDescending(); to order descending
            workbook.SaveAs("SortedSheet.xlsx");
        }
    }
}