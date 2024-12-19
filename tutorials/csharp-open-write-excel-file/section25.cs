using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section25
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            workSheet["A1"].Value = "Hello World";
            workBook.SaveAs(@"F:\MY WORK\IronPackage\Xl tutorial\newFile.xlsx");
        }
    }
}